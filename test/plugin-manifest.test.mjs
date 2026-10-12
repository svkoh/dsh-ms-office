// 插件清单与打包面：package.json / cordis.patch.yml / 入口文件里的一致性。
// 这些字段错了不会有任何运行时报错，只会在「装到用户机器上之后」才暴露，所以必须静态测。
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { ROOT, readJson, readParts } from './helpers/paths.mjs';

const pkg = readJson(path.join(ROOT, 'package.json'));
const PLUGIN_ID = 'dsh-office-for-mso';

/** 本机装了 dsh 运行时时，用它自带的 semver 做真检查；没装就跳过（CI 上没有） */
function loadRuntimeSemver() {
  const base = process.env.LOCALAPPDATA || process.env.XDG_DATA_HOME;
  if (!base || !fs.existsSync(base)) return null;
  // 不写死任何目录名：直接扫 <data>/<任意>/runtime/package.json，看它是不是 dsh 运行时
  let dirs = [];
  try {
    dirs = fs.readdirSync(base).filter((d) => fs.existsSync(path.join(base, d, 'runtime', 'package.json')));
  } catch { return null; }
  for (const d of dirs) {
    const manifest = path.join(base, d, 'runtime', 'package.json');
    try {
      const runtimePkg = JSON.parse(fs.readFileSync(manifest, 'utf8'));
      const deps = { ...runtimePkg.dependencies, ...runtimePkg.devDependencies };
      if (!Object.keys(deps).some((n) => n.startsWith('@deepseek-ai/'))) continue;
      const req = createRequire(manifest);
      const semver = req('semver');
      const nm = path.dirname(manifest);
      const versions = {};
      for (const name of Object.keys(pkg.peerDependencies || {})) {
        const p = path.join(nm, 'node_modules', ...name.split('/'), 'package.json');
        if (fs.existsSync(p)) versions[name] = JSON.parse(fs.readFileSync(p, 'utf8')).version;
      }
      return { semver, versions, where: manifest };
    } catch { /* 换下一个候选 */ }
  }
  return null;
}

describe('package.json 清单', () => {
  test('名称与仓库约定一致', () => {
    assert.equal(pkg.name, PLUGIN_ID);
    assert.match(pkg.version, /^\d+\.\d+\.\d+/);
    assert.equal(pkg.type, 'module');
    assert.equal(pkg.license, 'MIT');
  });

  test('dsh 清单字段完整且指向真实文件', () => {
    assert.equal(pkg.dsh.manifestVersion, 1);
    assert.equal(pkg.dsh.bundle.patch, './cordis.patch.yml');
    assert.ok(fs.existsSync(path.join(ROOT, pkg.dsh.bundle.patch)), 'dsh.bundle.patch 指向的文件不存在');
    assert.equal(pkg.dsh.client.platform, 'web');
    assert.ok(Array.isArray(pkg.dsh.client.inject) && pkg.dsh.client.inject.length > 0);
  });

  test('main / exports 指向真实文件', () => {
    for (const rel of [pkg.main, pkg.exports['.'], pkg.exports['./client']]) {
      assert.ok(fs.existsSync(path.join(ROOT, rel)), `${rel} 不存在`);
    }
  });

  test('peerDependencies 只依赖官方 dsh 包，且声明为 peer', () => {
    const peers = Object.keys(pkg.peerDependencies || {});
    assert.ok(peers.length > 0, '至少要声明 cordis 这个 peer');
    for (const name of peers) {
      assert.ok(name.startsWith('@deepseek-ai/'), `${name} 不是官方 dsh 包`);
      assert.equal(pkg.dependencies?.[name], undefined, `${name} 不应是运行时依赖（发布后只依赖官方 dsh）`);
    }
  });

  test('engines 声明 node 与 dsh 下限', () => {
    assert.ok(pkg.engines?.node, '缺 engines.node');
    assert.ok(pkg.engines?.dsh, '缺 engines.dsh');
  });

  test('peer 范围能匹配本机实际安装的运行版本（没装运行时则跳过）', (t) => {
    const rt = loadRuntimeSemver();
    if (!rt) return t.skip('本机没有 dsh 运行时，跳过（这条只在开发机上有效）');
    const checked = [];
    for (const [name, range] of Object.entries(pkg.peerDependencies)) {
      const installed = rt.versions[name];
      if (!installed) continue;
      assert.ok(rt.semver.satisfies(installed, range),
        `${name}@${installed} 不满足声明的范围 "${range}" —— 默认 semver 下预发布版不被 ^ 覆盖`);
      checked.push(`${name}@${installed} ⊨ ${range}`);
    }
    assert.ok(checked.length > 0, '一个 peer 都没比对上，检查路径可能不对');
  });

  test('files 白名单里每一项都存在（防重命名后漏改）', () => {
    for (const entry of pkg.files) {
      assert.ok(fs.existsSync(path.join(ROOT, entry)), `files 里的 ${entry} 不存在`);
    }
  });

  test('files 覆盖全部分片（漏掉会让插件装上却没有 action）', () => {
    assert.ok(pkg.files.includes('server/actions'), 'files 里必须有 server/actions');
    for (const f of readParts()) {
      assert.ok(fs.existsSync(path.join(ROOT, 'server/actions', f)), `分片 ${f} 不存在`);
    }
    assert.ok(fs.existsSync(path.join(ROOT, 'server/actions/parts.json')));
  });

  test('files 不含裸目录 server（会覆盖 .gitignore，把用户的 user-actions.js 发出去）', () => {
    assert.ok(!pkg.files.includes('server'), 'files 里不能有裸的 "server"');
    assert.ok(!pkg.files.includes('.'), 'files 里不能有 "."');
  });

  test('files 不包含任何用户本机文件', () => {
    for (const entry of pkg.files) {
      assert.ok(!entry.includes('user-actions.js'), 'user-actions.js 是本机定制，绝不能入库/发布');
      assert.ok(!entry.startsWith('test'), 'test/ 不该进发布包');
    }
  });

  test('files 覆盖 i18n 字典（漏掉会让插件装上却起不来）', () => {
    assert.ok(pkg.files.includes('i18n'), 'files 里必须有 i18n');
    // 服务端启动时对 LANGS 逐个 JSON.parse：少一个就是启动期直接抛错，
    // 但打包不会报任何错，所以只能在这里静态兜住。
    const src = fs.readFileSync(path.join(ROOT, 'server', 'server.js'), 'utf8');
    const langs = /const LANGS = \[([^\]]*)\]/.exec(src);
    assert.ok(langs, 'server.js 里找不到 LANGS');
    const declared = [...langs[1].matchAll(/'([^']+)'/g)].map((m) => m[1]);
    assert.ok(declared.length >= 2, 'LANGS 至少要有两种语言');
    for (const lang of declared) {
      assert.ok(fs.existsSync(path.join(ROOT, 'i18n', `${lang}.json`)),
        `LANGS 声明了 ${lang} 但没有 i18n/${lang}.json`);
    }
  });

  test('README 里引用的相对图片都存在且在白名单内（否则发布包缺图）', () => {
    for (const readme of ['README.md', 'README.en.md']) {
      const text = fs.readFileSync(path.join(ROOT, readme), 'utf8');
      for (const m of text.matchAll(/!\[[^\]]*\]\(([^)]+)\)/g)) {
        const rel = m[1].trim();
        if (/^[a-z][a-z0-9+.-]*:/i.test(rel)) continue; // 外链 / data: 跳过
        assert.ok(fs.existsSync(path.join(ROOT, rel)), `${readme} 引用了不存在的 ${rel}`);
        const top = rel.split('/')[0];
        assert.ok(pkg.files.includes(top),
          `${readme} 引用的 ${rel} 不在 files 里 —— 发布包会缺这个文件`);
      }
    }
  });

  test('声明了 test / lint 脚本', () => {
    assert.ok(pkg.scripts?.test, '缺 npm test');
    assert.ok(pkg.scripts?.lint, '缺 npm run lint');
  });
});

describe('cordis.patch.yml', () => {
  const yml = fs.readFileSync(path.join(ROOT, 'cordis.patch.yml'), 'utf8');

  test('挂载行的 id / name 都是本插件', () => {
    assert.match(yml, new RegExp(`id:\\s*${PLUGIN_ID}\\b`));
    assert.match(yml, new RegExp(`name:\\s*${PLUGIN_ID}\\b`));
  });

  test('只有一条挂载行（多一条会让插件被加载两次）', () => {
    const ids = yml.match(/^\s*-?\s*id:/gm) || [];
    assert.equal(ids.length, 1, `发现 ${ids.length} 条 id 行`);
  });
});

describe('技能包', () => {
  const skillsDir = path.join(ROOT, 'skills');
  // 约定：技能名 = 插件名去掉 `dsh-` 前缀（`dsh-` 是插件命名空间的事，不该出现在技能名里）
  const SKILL_NAME = PLUGIN_ID.replace(/^dsh-/, '');

  test('只带一个技能，且目录名 / SKILL.md 的 name 都是约定的技能名', () => {
    const dirs = fs.readdirSync(skillsDir);
    assert.equal(dirs.length, 1, `skills/ 下应只有一个技能，实际：${dirs.join(', ')}`);
    const dir = dirs[0];
    assert.equal(dir, SKILL_NAME, `技能目录名应为 ${SKILL_NAME}`);
    const md = fs.readFileSync(path.join(skillsDir, dir, 'SKILL.md'), 'utf8');
    assert.match(md, new RegExp(`^name: ${dir}$`, 'm'),
      `SKILL.md 的 name 与目录名 ${dir} 不一致（发布后的技能名取目录名）`);
  });

  test('SKILL.md 有 description 与 whenToUse（AI 靠这两项决定用不用）', () => {
    const md = fs.readFileSync(path.join(skillsDir, SKILL_NAME, 'SKILL.md'), 'utf8');
    assert.match(md, /^description: .+/m);
    assert.match(md, /^whenToUse: .+/m);
  });

  test('README 里的技能路径与实际目录一致', () => {
    for (const file of ['README.md', 'README.en.md']) {
      const text = fs.readFileSync(path.join(ROOT, file), 'utf8');
      assert.ok(text.includes(`skills/${SKILL_NAME}/SKILL.md`), `${file} 里的技能路径不对`);
    }
  });
});

describe('宿主侧入口', () => {
  test('lib/index.js 用的 pluginName / bridge 与包名一致', () => {
    const src = fs.readFileSync(path.join(ROOT, 'lib/index.js'), 'utf8');
    assert.ok(src.includes(`"${PLUGIN_ID}"`) || src.includes(`'${PLUGIN_ID}'`), 'pluginName 不是包名');
    assert.ok(src.includes('/api/office-for-mso'), 'bridge 路由不对');
    assert.ok(src.includes('server/server.js'), '侧车入口不对');
  });

  test('lib/client.js 的 id 与包名一致', () => {
    const src = fs.readFileSync(path.join(ROOT, 'lib/client.js'), 'utf8');
    assert.ok(src.includes(PLUGIN_ID), 'client id 不是包名');
  });
});

describe('文档与实际一致', () => {
  const capabilities = readJson(path.join(ROOT, 'server/capabilities.json'));
  const count = capabilities.length;

  test('README 里的 action 数就是 capabilities.json 的条数', () => {
    const readme = fs.readFileSync(path.join(ROOT, 'README.md'), 'utf8');
    assert.ok(readme.includes(`${count} 个 action`), `README.md 里没有 "${count} 个 action"`);
    assert.ok(!/\b(34|36) 个 action/.test(readme), 'README.md 里还有过期的 action 数');
  });

  test('README.en.md 里的 action 数与实际一致', () => {
    const readme = fs.readFileSync(path.join(ROOT, 'README.en.md'), 'utf8');
    assert.ok(readme.includes(`${count} actions`), `README.en.md 里没有 "${count} actions"`);
  });

  test('两份 README 里列出的每个 action 都真实存在', () => {
    const names = new Set(capabilities.map((c) => c.name));
    for (const file of ['README.md', 'README.en.md']) {
      const text = fs.readFileSync(path.join(ROOT, file), 'utf8');
      const mentioned = [...text.matchAll(/`([a-z][a-z0-9_]{3,})`/g)].map((m) => m[1]);
      const known = mentioned.filter((n) => n.includes('_') && !names.has(n) && /^(word|ppt|read|write|set|get|add|list|apply|insert|delete|clear|convert|create|format|replace|append|search|locate|validate|evaluate|rename|change)_/.test(n));
      assert.deepEqual([...new Set(known)], [], `${file} 提到了不存在的 action`);
    }
  });
});
