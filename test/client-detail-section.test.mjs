// 插件详情页的区块：注册到 plugins.detail.section 的组件会拿到「当前打开的是哪一页」，
// 认不出自己就渲染 null。认错了不会有任何报错，页面只是空着 —— 所以只能把组件真跑一遍。
//
// 两种主体形状都要认：
//   { kind: "bundle", pkg }  —— 已安装的包打开时的页面（本插件就是这种）
//   { kind: "item",   id  }  —— 官方条目的页面
// 行（row）页面属于某一行，不属于包，不该在这里出设置区。
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { ROOT } from './helpers/paths.mjs';

const PLUGIN_ID = 'dsh-ms-office';

/** 在假的浏览器环境里跑 lib/client.js，拿回它交给模块加载器的那份声明。 */
function loadDeclaration() {
  const source = fs.readFileSync(path.join(ROOT, 'lib/client.js'), 'utf8');
  let declaration;
  const context = vm.createContext({
    window: { __ModuleLoader__: { load: (value) => { declaration = value; } } },
    navigator: { language: 'zh-CN' },
    fetch: () => Promise.reject(new Error('these tests never reach the bridge')),
    setInterval: () => 0,
    clearInterval: () => {},
    console,
  });
  vm.runInContext(source, context, { filename: 'lib/client.js' });
  assert.ok(declaration, 'lib/client.js 没有调用 window.__ModuleLoader__.load');
  return declaration;
}

/** 只要够组件跑完一次渲染的 react 替身：state 取初值，effect 不执行。 */
const react = {
  useState: (initial) => [typeof initial === 'function' ? initial() : initial, () => {}],
  useEffect: () => {},
};

const element = (type, props) => ({ type, props });

/** 装好组件，返回它注册到槽位里的那一份。 */
function mountSection() {
  const declaration = loadDeclaration();
  assert.equal(declaration.id, PLUGIN_ID);
  const exports = declaration.factory((name) => {
    if (name === 'react') return react;
    if (name === 'react/jsx-runtime') return { jsx: element, jsxs: element, Fragment: 'fragment' };
    throw new Error(`客户端不应该 require(${JSON.stringify(name)})`);
  });
  // 这份数组来自 vm 里的另一个 realm，先搬回本 realm 再比，否则原型不同直接判不等
  assert.deepEqual([...exports.inject], ['slots']);

  const registered = [];
  const ctx = {
    slots: {
      inject(key, callback) {
        assert.equal(key, 'plugins.detail.section');
        callback();
      },
      register(spec, component) {
        registered.push({ spec, component });
        return () => {};
      },
    },
  };
  exports.apply(ctx);
  assert.equal(registered.length, 1, '应当只注册一个详情页区块');
  assert.equal(registered[0].spec.name, 'plugins.detail.section');
  assert.equal(registered[0].spec.id, `sidecar-settings.${PLUGIN_ID}`);
  return registered[0].component;
}

describe('详情页区块只在自己的页面上出现', () => {
  const Section = mountSection();
  const render = (subject) => Section({ subject });

  test('已安装的包页面：包名是本插件就渲染', () => {
    const tree = render({ kind: 'bundle', pkg: { name: PLUGIN_ID } });
    assert.ok(tree, '包页面什么都不渲染 —— 设置区会整块消失');
    assert.equal(tree.type, 'section');
    assert.equal(tree.props['data-sidecar-settings'], PLUGIN_ID);
  });

  test('官方条目页面：id 是本插件就渲染', () => {
    const tree = render({ kind: 'item', id: PLUGIN_ID });
    assert.ok(tree);
    assert.equal(tree.props['data-sidecar-settings'], PLUGIN_ID);
  });

  test('别人的包、别人的条目、行页面都不渲染', () => {
    assert.equal(render({ kind: 'bundle', pkg: { name: 'dsh-better-sidebar' } }), null);
    assert.equal(render({ kind: 'item', id: 'agent-loop' }), null);
    assert.equal(render({ kind: 'row', pkg: { name: PLUGIN_ID }, row: { rowId: 'dsh-ms-office' } }), null);
    assert.equal(render(undefined), null);
    assert.equal(render({ kind: 'bundle' }), null);
  });
});
