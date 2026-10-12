// shapeType 的合法值只有 Office 知道，插件不能凭空枚举，所以校验层不拦未验证的名字。
// 但执行失败时 Office 只回一个 "InvalidArgument"，调用方看不出问题出在哪项参数——
// 这个文件测的是：执行报错必须把「已验证可用」的清单补进消息里。
//
// 仍然不模拟 Office 对象模型：校验放行后走到执行层桩，桩抛错，
// 我们断言的是「报错文本长什么样」，不是桩本身。
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { loadActions, implOf } from './helpers/actions.mjs';

const { A } = loadActions({ host: 'PowerPoint', supported: true });
const addElements = implOf(A, 'PowerPoint', 'ppt_add_elements');

const PASSED = 'execution'; // 走到执行层桩 = 校验放行了

function elements(list) {
  return {
    slideId: 'slide-1',
    coordinateSpace: { width: 960, height: 540 },
    targetSize: { width: 960, height: 540 },
    elements: list,
  };
}

describe('ppt_add_elements：shapeType 的执行报错必须带合法值提示', () => {
  test('未验证的名字照样放行（不能因为插件没收录就拒绝）', async () => {
    const r = await addElements(elements([{ type: 'shape', shapeType: 'Oval', x: 0, y: 0, w: 100, h: 100 }]));
    assert.equal(r.code, PASSED, `未验证的 shapeType 被校验层拒了：${r.error}`);
  });

  test('执行报错里补上已验证清单，调用方不用再猜', async () => {
    const r = await addElements(elements([{ type: 'shape', shapeType: 'Oval', x: 0, y: 0, w: 100, h: 100 }]));
    assert.match(r.error, /shapeType "Oval" is accepted but unverified/);
    assert.match(r.error, /verified values: Rectangle, Ellipse, Triangle/);
  });

  test('已验证的名字不加提示，报错保持干净', async () => {
    const r = await addElements(elements([{ type: 'shape', shapeType: 'RoundRectangle', x: 0, y: 0, w: 100, h: 100 }]));
    assert.equal(r.code, PASSED);
    assert.equal(r.error, 'stub: Office run must not execute');
  });

  test('多个未验证名字去重后一起列出', async () => {
    const r = await addElements(elements([
      { type: 'shape', shapeType: 'Oval', x: 0, y: 0, w: 10, h: 10 },
      { type: 'shape', shapeType: 'oval', x: 20, y: 0, w: 10, h: 10 },
      { type: 'shape', shapeType: 'FivePointStar', x: 40, y: 0, w: 10, h: 10 },
    ]));
    assert.match(r.error, /shapeType "Oval, FivePointStar"/);
  });

  test('没有 shape 元素时不加提示', async () => {
    const r = await addElements(elements([{ type: 'text', text: 'x', x: 0, y: 0, w: 10, h: 10 }]));
    assert.equal(r.code, PASSED);
    assert.equal(r.error, 'stub: Office run must not execute');
  });

  test('非 PascalCase 仍然在校验层被拒，并给出示例名', async () => {
    const r = await addElements(elements([{ type: 'shape', shapeType: 'round-rect', x: 0, y: 0, w: 10, h: 10 }]));
    assert.equal(r.code, 'bad_args');
    assert.match(r.error, /shapeType is invalid: "round-rect"/);
    assert.match(r.error, /must be an Office shape name, e\.g\. Rectangle, Ellipse, Triangle, RightTriangle/);
  });
});
