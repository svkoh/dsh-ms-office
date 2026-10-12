// 分片边界：拆分的意义全在这个文件里 —— 如果这些断言不成立，拆分就只是把文件切碎了。
//
// 三条硬边界：
//   1. 宿主分片里不能出现「按宿主分叉」的判断（那就等于没拆）
//   2. Word 分片不能引用 Excel/PowerPoint 的对象模型（反之亦然）
//   3. 宿主无关的东西不能混进宿主分片（要放 common.js）
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readParts, readPartsSource } from './helpers/paths.mjs';

const src = readPartsSource();

/** 去掉行注释与块注释，避免注释里提到这些写法就误报 */
function stripComments(code) {
  return code.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^[ \t]*\/\/.*$/gm, '');
}

const HOST_OBJECT_MODEL = { Word: 'Word.run', Excel: 'Excel.run', PowerPoint: 'PowerPoint.run' };
const HOST_PARTS = { 'word.js': 'Word', 'excel.js': 'Excel', 'ppt.js': 'PowerPoint' };

describe('分片边界', () => {
  test('清单里的分片都存在且非空', () => {
    for (const f of readParts()) {
      assert.ok(src[f] && src[f].trim().length > 0, `${f} 缺失或为空`);
    }
  });

  test('宿主分片里没有「按宿主分叉」的判断', () => {
    for (const f of Object.keys(HOST_PARTS)) {
      const code = stripComments(src[f]);
      assert.doesNotMatch(code, /\bhost\s*[!=]==\s*'/,
        `${f} 里还有按宿主分叉的判断 —— 分片之后函数签名里已经没有 host 参数`);
      assert.doesNotMatch(code, /\bhost\s*[!=]==\s*[A-Za-z]/,
        `${f} 里还有 host 相关的比较`);
    }
  });

  test('宿主分片不引用别的宿主的对象模型', () => {
    for (const [f, ownHost] of Object.entries(HOST_PARTS)) {
      const code = stripComments(src[f]);
      for (const [host, api] of Object.entries(HOST_OBJECT_MODEL)) {
        if (host === ownHost) continue;
        assert.ok(!code.includes(api), `${f}（${ownHost}）不该引用 ${api}`);
      }
      assert.ok(code.includes(HOST_OBJECT_MODEL[ownHost]),
        `${f} 里找不到 ${HOST_OBJECT_MODEL[ownHost]} —— 可能内容搬错了`);
    }
  });

  test('common 分片不引用任何宿主对象模型', () => {
    const code = stripComments(src['common.js']);
    for (const api of Object.values(HOST_OBJECT_MODEL)) {
      assert.ok(!code.includes(api), `common.js 不该引用 ${api}（那它就不是宿主无关的）`);
    }
  });

  test('runtime 分片不引用任何宿主对象模型', () => {
    const code = stripComments(src['runtime.js']);
    for (const api of Object.values(HOST_OBJECT_MODEL)) {
      assert.ok(!code.includes(api), `runtime.js 不该引用 ${api}`);
    }
  });

  test('宿主分片把跨宿主 action 也只实现本宿主那一份（不含别的宿主的实现痕迹）', () => {
    // 拆分前这些函数的共同特征是「一个函数里好几段 if (host === ...)」，
    // 拆分后每个宿主的文件里都只应出现一次本宿主的实现。
    const code = stripComments(src['word.js']);
    assert.ok(!code.includes("host: 'Excel'"), 'word.js 里出现了 Excel 的结果标记');
    assert.ok(!code.includes("host: 'PowerPoint'"), 'word.js 里出现了 PowerPoint 的结果标记');
    const excel = stripComments(src['excel.js']);
    assert.ok(!excel.includes("host: 'Word'"), 'excel.js 里出现了 Word 的结果标记');
    assert.ok(!excel.includes("host: 'PowerPoint'"), 'excel.js 里出现了 PowerPoint 的结果标记');
    const ppt = stripComments(src['ppt.js']);
    assert.ok(!ppt.includes("host: 'Word'"), 'ppt.js 里出现了 Word 的结果标记');
    assert.ok(!ppt.includes("host: 'Excel'"), 'ppt.js 里出现了 Excel 的结果标记');
  });

  test('每个宿主分片都挂到 A.hosts.<宿主> 上，且 registry 只装配不实现', () => {
    assert.match(src['word.js'], /A\.hosts\.Word\s*=/);
    assert.match(src['excel.js'], /A\.hosts\.Excel\s*=/);
    assert.match(src['ppt.js'], /A\.hosts\.PowerPoint\s*=/);
    assert.match(src['common.js'], /A\.hosts\.Common\s*=/);
    const registry = stripComments(src['registry.js']);
    for (const api of Object.values(HOST_OBJECT_MODEL)) {
      assert.ok(!registry.includes(api), `registry.js 不该引用 ${api}（它只做装配）`);
    }
  });

  test('registry 是最后一个分片（装配必须在所有实现之后）', () => {
    const parts = readParts();
    assert.equal(parts[parts.length - 1], 'registry.js');
    assert.equal(parts[0], 'runtime.js', 'runtime 必须最先，否则分片拿不到容器');
  });
});
