// 测试共用路径与读取器。所有测试都从这里拿「仓库里的东西」，避免各自拼路径拼错。
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
export const SERVER = path.join(ROOT, 'server');
export const ACTIONS = path.join(SERVER, 'actions');

const require = createRequire(import.meta.url);

export function readJson(absPath) {
  return JSON.parse(fs.readFileSync(absPath, 'utf8'));
}

/** actions 分片清单（唯一真源：server/actions/parts.json） */
export function readParts() {
  return readJson(path.join(ACTIONS, 'parts.json')).parts;
}

/** capabilities.json（hosts / destructive 的唯一真源） */
export function readCapabilities() {
  return require(path.join(SERVER, 'capabilities.json'));
}

/** 服务端注入给窗格的能力元数据，与 server.js#actionsMeta() 同构 */
export function actionsMeta() {
  const meta = {};
  for (const entry of readCapabilities()) {
    meta[entry.name] = { hosts: entry.hosts, destructive: entry.destructive === true };
  }
  return meta;
}

/** 按 parts.json 的顺序拼出窗格实际收到的 actions 脚本 */
export function buildActionsBundle() {
  return readParts()
    .map((f) => fs.readFileSync(path.join(ACTIONS, f), 'utf8'))
    .join('\n');
}

/** 逐分片读取源码（键 = 文件名），用于「分片边界」类断言 */
export function readPartsSource() {
  const out = {};
  for (const f of readParts()) out[f] = fs.readFileSync(path.join(ACTIONS, f), 'utf8');
  return out;
}
