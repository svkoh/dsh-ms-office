// 这个仓库是独立开源插件，不能出现父项目的任何标识。
// 关键词在下面按片段拼出来，免得本文件自己被自己的规则命中。
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { ROOT } from './helpers/paths.mjs';

const PARENT = ['deepseek', 'harness', 'desktop'].join('-');
const OLD_NAME = ['dsh', 'office', 'bridge'].join('-');

const PATTERNS = [
  [OLD_NAME, '旧插件名（改名前的残留）'],
  [PARENT, '父仓库名'],
  ['dsh-desktop', '父项目简称'],
  ['companions-state', '父项目的状态文件名'],
  ['yuanliu', '品牌名'],
  ['oneai', '品牌名'],
  ['39100', '父项目的端口区间'],
  ['壳 Supervisor', '父项目的侧车表述'],
  ['登录计划任务', '父项目的安装期机制'],
  ['platform.config.json', '父项目的构建期产物'],
];

// Local AST indexes are generated binary analysis artifacts, not distributed plugin source.
const SKIP_DIRS = new Set(['.git', 'node_modules', '.tmp', '.codegraph', 'dist']);
// 本文件按设计包含这些关键词
const SKIP_FILES = new Set([path.basename(fileURLToPath(import.meta.url))]);

function containsPattern(text,needle){return /^\d+$/.test(needle)?new RegExp('(^|[^0-9])'+needle+'([^0-9]|$)').test(text):text.includes(needle);}

function walk(dir, acc = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    if (e.isDirectory()) {
      if (SKIP_DIRS.has(e.name)) continue;
      walk(path.join(dir, e.name), acc);
    } else {
      acc.push(path.join(dir, e.name));
    }
  }
  return acc;
}

const files = walk(ROOT).filter((f) => !SKIP_FILES.has(path.basename(f)));

describe('不泄露父项目信息', () => {
  test('仓库里扫到足够多的文件（防止扫描逻辑本身失效）', () => {
    assert.ok(files.length > 10, `只扫到 ${files.length} 个文件，路径可能不对`);
  });

  for (const [needle, why] of PATTERNS) {
    test(`没有出现「${needle}」（${why}）`, () => {
      const hits = [];
      for (const f of files) {
        let text;
        try { text = fs.readFileSync(f, 'utf8'); } catch { continue; } // 二进制跳过
        if (!containsPattern(text, needle)) continue;
        const line = text.split('\n').findIndex((l) => containsPattern(l, needle)) + 1;
        hits.push(`${path.relative(ROOT, f)}:${line}`);
      }
      assert.deepEqual(hits, [], `以下位置出现「${needle}」：\n${hits.join('\n')}`);
    });
  }

  test('.gitignore 把本机定制文件挡在外面', () => {
    const gi = fs.readFileSync(path.join(ROOT, '.gitignore'), 'utf8');
    assert.ok(gi.includes('user-actions.js'), '.gitignore 必须忽略 user-actions.js');
    assert.ok(gi.includes('node_modules'));
  });
});

function numericPortRegression(){const port=PATTERNS.find(([n])=>/^\d+$/.test(n))[0];assert.equal(containsPattern('12.3750'+port+'646973',port),false);for(const text of [':'+port,'port='+port,JSON.stringify({port:Number(port)})])assert.equal(containsPattern(text,port),true);}
test("numeric port scan accepts native Office decimals and still rejects endpoint references", numericPortRegression);
