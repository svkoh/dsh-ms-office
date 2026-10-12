// 分片装配校验：把 actions 分片在 Node 里真实加载一遍，跟 capabilities.json 对账。
//
// 这一层能挡住的最重要的一类回归是「静默缺失」：
//   - 某个 action 声明支持 PowerPoint，但 ppt.js 里忘了实现 → AI 能看到、调用永远失败
//   - 某个覆盖类操作忘了预览读取器 → confirm 分支永远走不通，等于功能不可用
//   - 某个 action 有实现但 capabilities.json 没登记 → 平台校验和审批拦截同时失效
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readCapabilities } from './helpers/paths.mjs';
import { HOSTS, implOf, loadActions, previewOf } from './helpers/actions.mjs';

const capabilities = readCapabilities();
const byName = new Map(capabilities.map((e) => [e.name, e]));

describe('actions 分片装配', () => {
  test('分片能加载，且注册表自校验通过（与 capabilities.json 一致）', () => {
    assert.doesNotThrow(() => loadActions(), '分片加载失败：说明注册表自校验拒绝了某处不一致');
  });

  test('注册表里的 action 集合与 capabilities.json 完全相同', () => {
    const { win } = loadActions();
    const registered = Object.keys(win.__ACTIONS__).sort();
    const declared = capabilities.map((e) => e.name).sort();
    assert.deepEqual(registered, declared,
      '两个方向的差异都会在这里显形：有实现没登记 / 登记了没实现');
  });

  test('每个 action 的 hosts 与 capabilities.json 一致', () => {
    const { win } = loadActions();
    for (const [name, entry] of Object.entries(win.__ACTIONS__)) {
      assert.deepEqual(entry.hosts, byName.get(name).hosts, `${name} 的 hosts 不一致`);
      assert.equal(entry.destructive, byName.get(name).destructive === true, `${name} 的 destructive 不一致`);
    }
  });

  test('capabilities.json 声明的每个 (action, host) 都有实现', () => {
    const { A } = loadActions();
    const missing = [];
    for (const entry of capabilities) {
      for (const host of entry.hosts) {
        if (typeof implOf(A, host, entry.name) !== 'function') missing.push(`${entry.name}@${host}`);
      }
    }
    assert.deepEqual(missing, [], '这些 (action, host) 声明了但没有实现');
  });

  test('每个覆盖类操作的每个宿主都有预览读取器', () => {
    const { A } = loadActions();
    const missing = [];
    for (const entry of capabilities) {
      if (!entry.destructive) continue;
      for (const host of entry.hosts) {
        if (typeof previewOf(A, host, entry.name) !== 'function') missing.push(`${entry.name}@${host}`);
      }
    }
    assert.deepEqual(missing, [], '这些覆盖类操作缺预览读取器，confirm 分支会走不通');
  });

  test('每个宿主分片只挂自己的宿主名', () => {
    const { A } = loadActions();
    assert.deepEqual(Object.keys(A.hosts).sort(), [...HOSTS, 'Common'].sort());
    for (const host of HOSTS) {
      assert.equal(typeof A.hosts[host].actions, 'object', `${host} 分片没有 actions 表`);
      assert.ok(Object.keys(A.hosts[host].actions).length > 0, `${host} 分片的 actions 表是空的`);
    }
  });

  test('跨宿主 action 要么各宿主独立实现，要么明确走宿主无关分片', () => {
    const { A } = loadActions();
    const multiHost = capabilities.filter((e) => e.hosts.length > 1);
    assert.ok(multiHost.length > 0, 'capabilities.json 里应该有跨宿主 action');
    for (const entry of multiHost) {
      const fromCommon = entry.hosts.every((h) => A.hosts[h].actions[entry.name] === undefined);
      const impls = entry.hosts.map((h) => implOf(A, h, entry.name));
      assert.ok(impls.every((f) => typeof f === 'function'), `${entry.name} 在某宿主上缺实现`);
      if (fromCommon) {
        // 宿主无关：三宿主共用 common.js 的实现，这是设计而不是漏拆
        assert.equal(new Set(impls).size, 1, `${entry.name} 走 common 分片，实现应该只有一个`);
      } else {
        // 宿主专属：每个宿主必须有**自己的**函数对象，共用 = 宿主分支没拆干净
        assert.equal(new Set(impls).size, impls.length,
          `${entry.name} 在多个宿主上共用同一个实现函数 —— 宿主分支没拆干净`);
      }
    }
  });

  test('write_selection 三个宿主都可用，且实现来自宿主无关分片', () => {
    const { A } = loadActions();
    for (const host of HOSTS) {
      assert.equal(typeof implOf(A, host, 'write_selection'), 'function', `write_selection 在 ${host} 上不可用`);
    }
    assert.equal(implOf(A, 'Word', 'write_selection'), implOf(A, 'PowerPoint', 'write_selection'));
  });

  test('每个宿主都有 read_selection / read_document 的独立实现', () => {
    const { A } = loadActions();
    for (const host of HOSTS) {
      assert.equal(typeof implOf(A, host, 'read_selection'), 'function', `${host} 缺 read_selection`);
      assert.equal(typeof implOf(A, host, 'read_document'), 'function', `${host} 缺 read_document`);
    }
    const rs = HOSTS.map((h) => implOf(A, h, 'read_selection'));
    assert.equal(new Set(rs).size, 3, 'read_selection 应该三个宿主各一份实现');
  });

  test('窗口全局：__ACTIONS__ / __EXECUTE__ / __AFTER_CONTENT__ 都已暴露', () => {
    const { win } = loadActions();
    assert.equal(typeof win.__ACTIONS__, 'object');
    assert.equal(typeof win.__EXECUTE__, 'function');
    assert.equal(typeof win.__AFTER_CONTENT__, 'function');
  });

  test('__AFTER_CONTENT__ 对已登记与未登记的 action 都能给出文案', () => {
    const { win } = loadActions();
    for (const name of Object.keys(win.__ACTIONS__)) {
      const text = win.__AFTER_CONTENT__(name, {});
      assert.equal(typeof text, 'string', `${name} 的「修改后内容」不是字符串`);
    }
    assert.equal(typeof win.__AFTER_CONTENT__('__不存在的 action__', { a: 1 }), 'string');
  });

  test('装配阶段不调用任何 Office 对象模型（桩会抛错）', () => {
    // 只要 loadActions 不抛，就说明分片在加载期没有真的去跑 Word.run 之类
    assert.doesNotThrow(() => loadActions({ host: 'PowerPoint' }));
  });
});
