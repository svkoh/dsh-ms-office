// capabilities.json 的形状与自洽性：它是 hosts / destructive 的唯一真源，
// 格式错一处就会让「平台校验」或「审批拦截」静默失效，所以形状本身要测。
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readCapabilities } from './helpers/paths.mjs';

const capabilities = readCapabilities();
const VALID_HOSTS = ['Word', 'Excel', 'PowerPoint'];

describe('capabilities.json', () => {
  test('是数组且非空', () => {
    assert.ok(Array.isArray(capabilities));
    assert.ok(capabilities.length > 0);
  });

  test('每条都有 name / hosts / destructive，且类型正确', () => {
    for (const e of capabilities) {
      assert.equal(typeof e.name, 'string', `${JSON.stringify(e)} 缺 name`);
      assert.ok(e.name.length > 0);
      assert.ok(Array.isArray(e.hosts) && e.hosts.length > 0, `${e.name} 的 hosts 必须是非空数组`);
      assert.equal(typeof e.destructive, 'boolean', `${e.name} 的 destructive 必须是布尔值`);
    }
  });

  test('hosts 里只有合法宿主名，且不重复', () => {
    for (const e of capabilities) {
      for (const h of e.hosts) {
        assert.ok(VALID_HOSTS.includes(h), `${e.name} 的 hosts 里有非法值 ${h}`);
      }
      assert.equal(new Set(e.hosts).size, e.hosts.length, `${e.name} 的 hosts 有重复`);
    }
  });

  test('action 名唯一', () => {
    const seen = new Set();
    const dup = [];
    for (const e of capabilities) {
      if (seen.has(e.name)) dup.push(e.name);
      seen.add(e.name);
    }
    assert.deepEqual(dup, [], 'action 名重复会让注册表互相覆盖');
  });

  test('每个 action 至少支持一个宿主', () => {
    assert.deepEqual(capabilities.filter((e) => e.hosts.length === 0).map((e) => e.name), []);
  });

  test('description / args 若存在则类型正确', () => {
    for (const e of capabilities) {
      if (e.description !== undefined) assert.equal(typeof e.description, 'string', `${e.name} 的 description`);
      if (e.requirement !== undefined) assert.equal(typeof e.requirement, 'string', `${e.name} 的 requirement`);
      if (e.args !== undefined) assert.equal(typeof e.args, 'object', `${e.name} 的 args`);
    }
  });
});
