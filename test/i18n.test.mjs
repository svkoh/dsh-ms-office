// i18n 一致性。
//
// 「结构/文案分离」之后最容易出的两类错，这个文件就是冲着它们去的：
//   1. 加了东西只写了一种语言 —— 键集合/占位符比对；
//   2. key 写错（比如文案放在 server.logs 却用 server.messages 查）—— 原样吐出 key 而不是报错，
//      这类错不跑一次根本看不出来，所以这里把「代码里引用的 key」逐个对回字典。
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ROOT, SERVER, readJson, readCapabilities } from './helpers/paths.mjs';
import { startBridge } from './helpers/server.mjs';

const I18N_DIR = path.join(ROOT, 'i18n');
const LANGS = ['zh', 'en'];
const dict = Object.fromEntries(LANGS.map((l) => [l, readJson(path.join(I18N_DIR, `${l}.json`))]));

const serverSrc = fs.readFileSync(path.join(SERVER, 'server.js'), 'utf8');
const paneJs = fs.readFileSync(path.join(SERVER, 'taskpane.js'), 'utf8');
const paneHtml = fs.readFileSync(path.join(SERVER, 'taskpane.html'), 'utf8');
const manifestXml = fs.readFileSync(path.join(SERVER, 'manifest.xml'), 'utf8');

/** 把嵌套结构拍平成 "a.b[0].c" → 叶子值 */
function leaves(node, prefix = '', out = {}) {
  if (Array.isArray(node)) {
    node.forEach((v, i) => leaves(v, `${prefix}[${i}]`, out));
  } else if (node && typeof node === 'object') {
    for (const [k, v] of Object.entries(node)) leaves(v, prefix ? `${prefix}.${k}` : k, out);
  } else {
    out[prefix] = node;
  }
  return out;
}
const LEAVES = Object.fromEntries(LANGS.map((l) => [l, leaves(dict[l])]));
const holes = (s) => [...String(s).matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort();

/** taskpane.js 用到的 pane 文案 key：t('key')、set('id','key')、setStatusKey('cls','key') 都算 */
function paneKeysUsed() {
  const used = new Set();
  for (const m of paneJs.matchAll(/\bt\(\s*'([A-Za-z0-9_]+)'/g)) used.add(m[1]);
  for (const m of paneJs.matchAll(/\bset\(\s*'[\w-]+'\s*,\s*'([\w-]+)'\s*\)/g)) used.add(m[1]);
  for (const m of paneJs.matchAll(/\bsetStatusKey\(\s*'[\w-]+'\s*,\s*'([\w-]+)'\s*\)/g)) used.add(m[1]);
  return used;
}

describe('i18n 字典', () => {
  test('两种语言的键集合完全相同', () => {
    const a = Object.keys(LEAVES.zh).sort();
    const b = Object.keys(LEAVES.en).sort();
    assert.deepEqual(b.filter((k) => !a.includes(k)), [], 'en 里有 zh 没有的键');
    assert.deepEqual(a.filter((k) => !b.includes(k)), [], 'zh 里有 en 没有的键');
  });

  test('同一个键的占位符在两种语言里一致', () => {
    for (const key of Object.keys(LEAVES.zh)) {
      assert.deepEqual(holes(LEAVES.en[key]), holes(LEAVES.zh[key]), `${key} 的占位符不一致`);
    }
  });

  test('没有空值（漏译会留下空字符串）', () => {
    for (const lang of LANGS) {
      for (const [key, value] of Object.entries(LEAVES[lang])) {
        assert.equal(typeof value, 'string', `${lang}:${key} 不是字符串`);
        assert.ok(value.trim().length > 0, `${lang}:${key} 是空的`);
      }
    }
  });

  test('en 里没有残留的中文（漏译成原样的典型症状）', () => {
    const cjk = /[\u4e00-\u9fff]/;
    for (const [key, value] of Object.entries(LEAVES.en)) {
      // 例外：pane/logs 里本来就要展示中文的场景不存在，所以这里一律不许有汉字。
      assert.doesNotMatch(value, cjk, `en:${key} 里还有中文：${value}`);
    }
  });
});

describe('服务端文案与代码引用一致', () => {
  test('server.js 里 msg() 引用的每个 key 都在两种语言里存在', () => {
    const used = new Set([...serverSrc.matchAll(/\bmsg\(\s*[A-Za-z_$]+\s*,\s*'([A-Za-z0-9_]+)'/g)].map((m) => m[1]));
    assert.ok(used.size > 10, `只抓到 ${used.size} 个 msg() 调用，正则可能没匹配上`);
    for (const key of used) {
      for (const lang of LANGS) {
        assert.ok(dict[lang].server.messages[key], `${lang} 的 server.messages 缺 key "${key}"`);
      }
    }
  });

  test('没有无人引用的服务端文案（改完代码留下的死文案）', () => {
    const used = new Set([...serverSrc.matchAll(/\bmsg\(\s*[A-Za-z_$]+\s*,\s*'([A-Za-z0-9_]+)'/g)].map((m) => m[1]));
    for (const key of Object.keys(dict.zh.server.messages)) {
      assert.ok(used.has(key), `server.messages."${key}" 没有任何 msg() 引用它`);
    }
  });

  test('/office/errors 的每个错误码在两种语言里都有非空含义与建议', async () => {
    const srv = await startBridge();
    try {
      for (const lang of LANGS) {
        const r = await srv.req(`/office/errors?lang=${lang}`);
        assert.equal(r.body.lang, lang);
        assert.ok(r.body.codes.length > 0);
        for (const c of r.body.codes) {
          assert.ok(c.meaning && c.meaning.trim(), `${lang} 的 ${c.code} 缺 meaning`);
          assert.ok(c.aiAction && c.aiAction.trim(), `${lang} 的 ${c.code} 缺 aiAction`);
          assert.ok(typeof c.retry === 'boolean', `${lang} 的 ${c.code} 缺 retry`);
        }
        assert.ok(r.body.semantics.length > 0, `${lang} 缺结果语义约定`);
      }
    } finally { srv.stop(); }
  });

  test('错误码的结构字段（code/layer/retry）不随语言变', async () => {
    const srv = await startBridge();
    try {
      const strip = (codes) => JSON.stringify(codes.map(({ code, layer, retry }) => ({ code, layer, retry })));
      const zh = await srv.req('/office/errors?lang=zh');
      const en = await srv.req('/office/errors?lang=en');
      assert.equal(strip(en.body.codes), strip(zh.body.codes));
    } finally { srv.stop(); }
  });
});

describe('能力表：结构不随语言变，文案随语言变', () => {
  test('/office/capabilities 两种语言的 name/hosts/destructive/requirement 完全相同', async () => {
    const srv = await startBridge();
    try {
      const strip = (actions) => JSON.stringify(actions.map(({ name, hosts, destructive, requirement }) => ({ name, hosts, destructive, requirement })));
      const zh = await srv.req('/office/capabilities?lang=zh');
      const en = await srv.req('/office/capabilities?lang=en');
      assert.equal(zh.body.actions.length, readCapabilities().length);
      assert.equal(strip(en.body.actions), strip(zh.body.actions));
      // 文案确实换了（否则「双语」是假的）
      const pick = (b) => b.actions.find((a) => a.name === 'insert_table').args.rows;
      assert.notEqual(pick(en.body), pick(zh.body));
    } finally { srv.stop(); }
  });

  test('字典里没有 capabilities.json 里不存在的 action（改名后的残留文案）', () => {
    const names = new Set(readCapabilities().map((c) => c.name));
    for (const lang of LANGS) {
      for (const name of Object.keys(dict[lang].capabilities)) {
        assert.ok(names.has(name), `${lang} 的 capabilities 里有 "${name}"，但 capabilities.json 里没有它`);
      }
    }
  });
});

describe('窗格 UI 文案', () => {
  test('taskpane.js 里引用的每个 pane key 都在字典里', () => {
    const used = paneKeysUsed();
    assert.ok(used.size > 15, `只抓到 ${used.size} 个 key，正则可能没匹配上`);
    for (const key of used) {
      for (const lang of LANGS) {
        assert.ok(dict[lang].pane[key], `${lang} 的 pane 缺 key "${key}"`);
      }
    }
  });

  test('pane 字典里没有无人引用的文案', () => {
    const used = paneKeysUsed();
    for (const key of Object.keys(dict.zh.pane)) {
      assert.ok(used.has(key), `pane."${key}" 没有被 taskpane.js 引用`);
    }
  });

  test('applyI18n 填的每个元素 id 都真实存在于 taskpane.html', () => {
    const pairs = [...paneJs.matchAll(/\bset\(\s*'([\w-]+)'\s*,\s*'([\w-]+)'\s*\)/g)].map((m) => ({ id: m[1], key: m[2] }));
    assert.ok(pairs.length > 10, `只抓到 ${pairs.length} 个 set() 调用`);
    for (const { id, key } of pairs) {
      assert.ok(paneHtml.includes(`id="${id}"`), `taskpane.html 里没有 id="${id}"`);
      assert.ok(dict.zh.pane[key], `pane 字典里没有 key "${key}"`);
    }
  });

  test('taskpane.html 里没有硬编码的可见文案（注释除外）', () => {
    // HTML 注释和 CSS 注释都不是「给人看的文案」，去掉它们之后 HTML 里不该再出现汉字。
    const stripped = paneHtml.replace(/<!--[\s\S]*?-->/g, '').replace(/\/\*[\s\S]*?\*\//g, '');
    const cjk = /[\u4e00-\u9fff]/;
    const found = stripped.match(new RegExp(`.{0,40}${cjk.source}.{0,40}`, 'g')) || [];
    assert.deepEqual(found, [], `HTML 里出现了硬编码中文：${found.join(' | ')}`);
  });

  test('/office/i18n 返回 pane 字典，且默认语言与 ?lang= 都能生效', async () => {
    const srv = await startBridge();
    try {
      const en = await srv.req('/office/i18n?lang=en');
      assert.equal(en.body.lang, 'en');
      assert.equal(Object.keys(en.body.pane).sort().join(), Object.keys(dict.en.pane).sort().join());
      assert.equal(en.body.pane.title, dict.en.pane.title);

      // Accept-Language 也能选语言（?lang= 优先级更高）
      const viaHeader = await srv.req('/office/i18n', { headers: { 'accept-language': 'en-US,en;q=0.9' } });
      assert.equal(viaHeader.body.lang, 'en');
      const override = await srv.req('/office/i18n?lang=zh', { headers: { 'accept-language': 'en-US' } });
      assert.equal(override.body.lang, 'zh');
    } finally { srv.stop(); }
  });
});

describe('技能包双语', () => {
  const SKILLS_DIR = path.join(ROOT, 'skills');
  const skillDirs = fs.readdirSync(SKILLS_DIR).filter((d) => fs.existsSync(path.join(SKILLS_DIR, d, 'SKILL.md')));
  const frontmatter = (text) => {
    const m = text.match(/^---\r?\n([\s\S]*?)\r?\n---/);
    assert.ok(m, 'SKILL 文件缺 frontmatter');
    return Object.fromEntries(m[1].split(/\r?\n/).map((line) => {
      const i = line.indexOf(':');
      return [line.slice(0, i).trim(), line.slice(i + 1).trim()];
    }).filter(([k]) => k));
  };

  test('至少有一个技能包', () => {
    assert.ok(skillDirs.length > 0);
  });

  for (const name of skillDirs) {
    test(`${name}：中文与英文版本都存在且 name 一致`, () => {
      const zh = fs.readFileSync(path.join(SKILLS_DIR, name, 'SKILL.md'), 'utf8');
      const enPath = path.join(SKILLS_DIR, name, 'SKILL.en.md');
      assert.ok(fs.existsSync(enPath), `缺 SKILL.en.md（发布时按宿主语言二选一，不能只留中文）`);
      const en = fs.readFileSync(enPath, 'utf8');
      const fmZh = frontmatter(zh);
      const fmEn = frontmatter(en);
      // 技能是按目录名发布的，两种语言的 name 必须都等于目录名，否则发布出去的技能名会对不上
      assert.equal(fmZh.name, name);
      assert.equal(fmEn.name, name);
      for (const field of ['description', 'whenToUse']) {
        assert.ok(fmZh[field] && fmZh[field].length > 20, `${name} 的 zh.${field} 太短或缺失`);
        assert.ok(fmEn[field] && fmEn[field].length > 20, `${name} 的 en.${field} 太短或缺失`);
      }
    });
  }

  test('技能文档里不写死 action 数量（会随能力表漂移）', () => {
    for (const name of skillDirs) {
      for (const file of ['SKILL.md', 'SKILL.en.md']) {
        const text = fs.readFileSync(path.join(SKILLS_DIR, name, file), 'utf8');
        const found = text.match(/\d+\s*(?:个\s*)?actions?\b|\d+\s*个\s*action/gi) || [];
        assert.deepEqual(found, [], `${name}/${file} 写死了 action 数量：${found.join(', ')}（改成「以 /office/capabilities 返回为准」）`);
      }
    }
  });

  test('技能不给出无法执行的端口指令，且指向真实存在的公告文件', () => {
    // 文件名从服务端代码里读，别在测试里再写一遍 —— 否则两边可以各自漂移。
    const serverSrc = fs.readFileSync(path.join(ROOT, 'server', 'server.js'), 'utf8');
    const announced = /writeBridgeAnnouncement[\s\S]*?path\.join\(stateDir\(\), '([^']+)'\)/.exec(serverSrc);
    assert.ok(announced, 'server.js 里找不到端口公告的文件名');
    const announcementFile = announced[1];
    for (const name of skillDirs) {
      for (const file of ['SKILL.md', 'SKILL.en.md']) {
        const text = fs.readFileSync(path.join(SKILLS_DIR, name, file), 'utf8');
        // 「先找端口」如果不给出可查的落点，agent 只能猜 —— 那等于没写
        assert.ok(!/<port>/i.test(text),
          `${name}/${file} 出现了 <port> 占位符：agent 无法填它，只会退化成猜 3000`);
        assert.ok(text.includes(announcementFile),
          `${name}/${file} 必须告诉 agent 去读 ${announcementFile}`);
        assert.ok(text.includes('$DSH_HOME'),
          `${name}/${file} 必须说明 $DSH_HOME（否则那个路径不可知）`);
      }
    }
  });

  test('技能里写的默认端口 == 插件声明的默认端口', () => {
    const indexSrc = fs.readFileSync(path.join(ROOT, 'lib', 'index.js'), 'utf8');
    const m = /defaultPort:\s*(\d+)/.exec(indexSrc);
    assert.ok(m, 'lib/index.js 里找不到 defaultPort');
    for (const name of skillDirs) {
      for (const file of ['SKILL.md', 'SKILL.en.md']) {
        const text = fs.readFileSync(path.join(SKILLS_DIR, name, file), 'utf8');
        assert.ok(text.includes(`127.0.0.1:${m[1]}`),
          `${name}/${file} 没写默认端口 ${m[1]}（改了默认端口要同步文档）`);
      }
    }
  });
});

describe('加载项清单多语言', () => {
  test('DisplayName / Description 都有 DefaultValue 且带 en-US Override', () => {
    for (const tag of ['DisplayName', 'Description']) {
      const block = manifestXml.match(new RegExp(`<${tag}\\b[\\s\\S]*?</${tag}>|<${tag}\\b[^>]*/>`));
      assert.ok(block, `manifest.xml 里没有 <${tag}>`);
      assert.match(block[0], /DefaultValue="/, `${tag} 缺 DefaultValue`);
      assert.match(block[0], /<Override\s+Locale="en-US"\s+Value="/, `${tag} 缺 en-US Override`);
    }
  });

  test('DefaultLocale 与项目语言规则一致（en* 用英文，其余中文）', () => {
    const m = manifestXml.match(/<DefaultLocale>\s*([\w-]+)\s*<\/DefaultLocale>/);
    assert.ok(m, 'manifest.xml 里没有 <DefaultLocale>');
    assert.equal(m[1], 'zh-CN');
    // Override 里出现的语言必须在 i18n 目录里都有字典，否则「支持了某种语言」是假的
    const overrides = [...manifestXml.matchAll(/<Override\s+Locale="([\w-]+)"/g)].map((x) => x[1].split('-')[0]);
    for (const lang of overrides) {
      assert.ok(LANGS.includes(lang), `manifest 声明了 ${lang}，但 i18n/ 下没有对应字典`);
    }
  });

  test('清单里的显示名与窗格标题一致（同一件东西只有一个名字）', () => {
    const zhName = manifestXml.match(/<DisplayName\s+DefaultValue="([^"]+)"/)[1];
    const enName = manifestXml.match(/<DisplayName[\s\S]*?<Override\s+Locale="en-US"\s+Value="([^"]+)"/)[1];
    assert.equal(zhName, dict.zh.pane.title);
    assert.equal(enName, dict.en.pane.title);
  });
});
