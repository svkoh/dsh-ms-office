// PowerPoint 参数校验的两个真实坑，这个文件冲着它们去：
//   1. 枚举以前只接受规范大小写：调用方写 align:"left" 被拒，而错误文本没列出合法值；
//   2. 字段名写错（kind / width / height / fontSize）只能靠猜，报错也不提示正确名称。
//
// 这里跑的是真分片（vm + 最小桩），仍然不模拟 Office 对象模型：
// 校验发生在执行之前，所以「参数被放行」的断言是 code==='execution'（执行层桩会抛），
// 「参数被拒」的断言是 code==='bad_args'。这样测的是校验本身，不是桩的行为。
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { loadActions, implOf } from './helpers/actions.mjs';

const { A } = loadActions({ host: 'PowerPoint', supported: true });
const addElements = implOf(A, 'PowerPoint', 'ppt_add_elements');
const setZOrder = implOf(A, 'PowerPoint', 'ppt_set_z_order');
const alignShapes = implOf(A, 'PowerPoint', 'ppt_align_shapes');
const distributeShapes = implOf(A, 'PowerPoint', 'ppt_distribute_shapes');

const PASSED = 'execution'; // 走到执行层桩 = 校验放行了

function elements(list) {
  return {
    slideId: 'slide-1',
    coordinateSpace: { width: 960, height: 540 },
    targetSize: { width: 960, height: 540 },
    elements: list,
  };
}

describe('ppt_add_elements：枚举大小写不敏感', () => {
  test('align / valign 写小写也能通过', async () => {
    const r = await addElements(elements([{ type: 'text', text: '标题', x: 48, y: 34, w: 864, h: 46, align: 'left', valign: 'middlecentered' }]));
    assert.equal(r.code, PASSED, `小写枚举被拒了：${r.error}`);
  });

  test('connectorType / dashStyle 写小写也能通过', async () => {
    const r = await addElements(elements([{ type: 'line', x: 0, y: 10, w: 100, h: 1, connectorType: 'straight', dashStyle: 'dashdot' }]));
    assert.equal(r.code, PASSED, `小写枚举被拒了：${r.error}`);
  });

  test('shapeType 首字母小写也能通过', async () => {
    const r = await addElements(elements([{ type: 'shape', shapeType: 'rectangle', x: 0, y: 0, w: 100, h: 100 }]));
    assert.equal(r.code, PASSED, `小写 shapeType 被拒了：${r.error}`);
  });

  test('非法枚举必须报出「实际值 + 合法值清单」，不能只说 invalid', async () => {
    const r = await addElements(elements([{ type: 'text', text: 'x', x: 0, y: 0, w: 10, h: 10, align: 'centre' }]));
    assert.equal(r.code, 'bad_args');
    assert.match(r.error, /align is invalid: "centre"/);
    assert.match(r.error, /\(valid: Left, Center, Right, Justify, JustifyLow, Distributed, ThaiDistributed\)/);
  });
});

describe('ppt_add_elements：字段别名', () => {
  test('kind / width / height / fontSize 这些常见写法被接受', async () => {
    const r = await addElements(elements([{ kind: 'text', text: 'x', x: 0, y: 0, width: 100, height: 20, fontSize: 14 }]));
    assert.equal(r.code, PASSED, `别名没生效：${r.error}`);
  });

  test('type 非法时仍然报错，且列出可选值', async () => {
    const r = await addElements(elements([{ type: 'textbox', text: 'x', x: 0, y: 0, w: 10, h: 10 }]));
    assert.equal(r.code, 'bad_args');
    assert.match(r.error, /\(valid: text, shape, line\)/);
  });

  // 一条水平线 h=0、一条竖线 w=0 是合法的几何形状，不该被「最小 0.1」卡住。
  test('line 元素允许 w 或 h 为 0（其它类型仍要求正数）', async () => {
    const horizontal = await addElements(elements([{ type: 'line', x: 0, y: 10, w: 100, h: 0 }]));
    assert.equal(horizontal.code, PASSED, `h=0 的线被拒了：${horizontal.error}`);
    const vertical = await addElements(elements([{ type: 'line', x: 10, y: 0, w: 0, h: 100 }]));
    assert.equal(vertical.code, PASSED, `w=0 的线被拒了：${vertical.error}`);
    const text = await addElements(elements([{ type: 'text', text: 'x', x: 0, y: 0, w: 0, h: 10 }]));
    assert.equal(text.code, 'bad_args', '非 line 元素不该放宽到 0');
  });
});

describe('z-order / align / distribute：mode 与 operation', () => {
  test('operation 大小写不敏感', async () => {
    const r = await setZOrder({ slideId: 'slide-1', shapeIds: ['1'], operation: 'bringtofront' });
    assert.equal(r.code, PASSED, `小写 operation 被拒了：${r.error}`);
  });

  test('非法 operation 报出合法值清单', async () => {
    const r = await setZOrder({ slideId: 'slide-1', shapeIds: ['1'], operation: 'front' });
    assert.equal(r.code, 'bad_args');
    assert.match(r.error, /\(valid: BringForward, BringToFront, SendBackward, SendToBack\)/);
  });

  test('align mode 大小写不敏感', async () => {
    const r = await alignShapes({ slideId: 'slide-1', shapeIds: ['1', '2'], mode: 'Center' });
    assert.equal(r.code, PASSED, `大写 mode 被拒了：${r.error}`);
  });

  test('distribute mode 与 align mode 各自的合法值不混用', async () => {
    const ok = await distributeShapes({ slideId: 'slide-1', shapeIds: ['1', '2', '3'], mode: 'horizontal' });
    assert.equal(ok.code, PASSED, `合法 distribute mode 被拒了：${ok.error}`);
    const bad = await distributeShapes({ slideId: 'slide-1', shapeIds: ['1', '2', '3'], mode: 'center' });
    assert.equal(bad.code, 'bad_args');
    assert.match(bad.error, /\(valid: horizontal, vertical\)/);
  });
});
