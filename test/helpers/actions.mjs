// 在 Node 里加载 actions 分片：用最小的 Office / window 桩跑真实的分片代码。
//
// 这里**不模拟 Office 对象模型**（Word.run 之类不实现）—— 目的是验证「装配」而不是「执行」：
// 分片能不能加载、注册表能不能通过自校验、每个宿主的表里有没有该有的东西。
// 想验证行为就得有真 Office，那是加载项侧的事。
import vm from 'node:vm';
import { actionsMeta, buildActionsBundle } from './paths.mjs';

function makeUtil() {
  const ok = (result) => ({ ok: true, result });
  const err = (code, error) => ({ ok: false, code, error });
  const notImplemented = () => { throw new Error('测试桩：该工具函数不应在装配阶段被调用'); };
  return {
    okResult: ok,
    errResult: err,
    hostName: () => 'Word',
    excelRange: notImplemented,
    getSelectionTextAsync: notImplemented,
    setSelectionTextAsync: notImplemented,
    mapBuiltInStyle: notImplemented,
    mapAlignment: notImplemented,
    readPptFile: notImplemented,
    unzip: notImplemented,
    inflateEntry: notImplemented,
  };
}

function makeOffice(host) {
  return {
    context: {
      host,
      platform: 'PC',
      requirements: { isSetSupported: () => false },
      document: { mode: 'edit', url: '', title: '', settings: {} },
    },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    CoercionType: { Image: 'image', Text: 'text' },
    HostType: { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint' },
    GoToType: { Bookmark: 'bookmark' },
    InsertLocation: { replace: 'replace', end: 'end' },
  };
}

/**
 * 加载全部分片，返回 { sandbox, win, A }。
 * 分片自校验失败（与 capabilities.json 不一致）会在这里直接抛错 —— 这正是我们要测的。
 */
export function loadActions({ host = 'Word', meta = actionsMeta(), util = null } = {}) {
  const win = {
    __UTIL__: util || makeUtil(),
    __ACTIONS_META__: meta,
  };
  const sandbox = {
    window: win,
    Office: makeOffice(host),
    console,
    fetch: () => Promise.reject(new Error('测试桩：不应发起网络请求')),
    setTimeout,
  };
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(buildActionsBundle(), sandbox, { filename: 'actions.bundle.js' });
  return { sandbox, win, A: win.__OFFICE_ACTIONS__ };
}

export const HOSTS = ['Word', 'Excel', 'PowerPoint'];

/** 某宿主上某 action 的实现（先看宿主分片，再看宿主无关的 common） */
export function implOf(A, host, name) {
  const t = A.hosts[host];
  return (t && t.actions && t.actions[name]) || (A.hosts.Common.actions && A.hosts.Common.actions[name]) || null;
}

export function previewOf(A, host, name) {
  const t = A.hosts[host];
  return (t && t.previewReaders && t.previewReaders[name]) || (A.hosts.Common.previewReaders && A.hosts.Common.previewReaders[name]) || null;
}
