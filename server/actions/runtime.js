// runtime.js — actions 分片的共享运行时（必须最先加载）
//
// 分片契约：服务端把 ACTIONS_PARTS 里列出的分片按序拼成**一个**脚本响应给 /actions.js
// （浏览器侧仍是「一个 <script>」，加载/热更新契约不变），因此分片之间**不共享词法作用域**，
// 每个分片都是自包含 IIFE，只通过 window.__OFFICE_ACTIONS__ 这个容器交换东西。
//
// 放这里的必须是**真的跨宿主**的东西。任何只服务某一个宿主的代码都属于那个宿主的分片。
(function () {
  const U = window.__UTIL__;
  const A = (window.__OFFICE_ACTIONS__ = window.__OFFICE_ACTIONS__ || { util: U, hosts: {} });
  if (!A.util) A.util = U;
  if (!A.hosts) A.hosts = {};

  // 宿主能力探测：requirements 不存在时一律视为不支持（老宿主）
  A.isSetSupported = function (name, version) {
    try {
      return !!Office.context.requirements.isSetSupported(name, version);
    } catch (e) { return false; }
  };

  // 二维矩阵校验（Excel 写区域与 PPT 插表格共用）
  A.validateMatrix = function (value, name, maxCells) {
    if (!Array.isArray(value) || !value.length || !Array.isArray(value[0]) || !value[0].length) throw new Error(`${name} must be a non-empty 2D array`);
    const cols = value[0].length;
    if (value.some((row) => !Array.isArray(row) || row.length !== cols)) throw new Error(`${name} must be rectangular`);
    if (value.length * cols > maxCells) throw new Error(`${name} exceeds ${maxCells} cells`);
    return value;
  };

  // locate_select 的闪烁间隔
  A.sleep = function (ms) { return new Promise((r) => setTimeout(r, ms)); };

  // 定位失败（目标不存在）专用错误：区别于执行异常，让 AI 明确知道是「找不到」而不是「出故障」
  A.notFound = function (msg) { const e = new Error(msg); e.notFound = true; return e; };
})();
