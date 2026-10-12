// pako.min.js 是第三方 vendored 代码：出处、许可、以及「我们到底改了什么」必须可验证。
// 这条测的就是 THIRD_PARTY_NOTICES.md 里记录的那个哈希 —— 记录与实际不符时，许可声明就是假的。
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { ROOT } from './helpers/paths.mjs';

const PAKO = path.join(ROOT, 'server/pako.min.js');
const NOTICES = path.join(ROOT, 'THIRD_PARTY_NOTICES.md');

const raw = fs.readFileSync(PAKO);
const text = raw.toString('utf8');
const bannerEnd = text.indexOf('*/');
const body = Buffer.from(text.slice(bannerEnd + 3).replace(/^\n/, ''), 'utf8');
const bodySha256 = crypto.createHash('sha256').update(body).digest('hex');

const notices = fs.readFileSync(NOTICES, 'utf8');
// 记录里为了可读性写了千位分隔（45,988），比对前先去掉分隔符
const noticesDigits = notices.replace(/(\d)[,_ ](?=\d{3}\b)/g, '$1');

describe('pako 出处与许可', () => {
  test('文件头有出处注释（上游 dist 本身不带 banner，必须自己加）', () => {
    assert.equal(bannerEnd > 0, true, '没有找到 /*! ... */ 形式的出处注释');
    const banner = text.slice(0, bannerEnd);
    assert.match(banner, /pako 1\.0\.11/);
    assert.match(banner, /nodeca\/pako/);
    assert.match(banner, /MIT AND Zlib/);
  });

  test('正文哈希与 THIRD_PARTY_NOTICES.md 记录的一致（记录不是摆设）', () => {
    assert.ok(notices.includes(bodySha256),
      `正文实际 sha256=${bodySha256}，但 THIRD_PARTY_NOTICES.md 里没有这个值 —— 要么换了版本要么改了正文`);
  });

  test('正文之外的唯一改动就是那段 banner（记录里声明的字节数要对得上）', () => {
    assert.ok(noticesDigits.includes(String(body.length)), `正文字节数 ${body.length} 没写在 THIRD_PARTY_NOTICES.md 里`);
    assert.ok(noticesDigits.includes(String(raw.length)), `文件总字节数 ${raw.length} 没写在 THIRD_PARTY_NOTICES.md 里`);
  });

  test('许可正文（MIT 与 Zlib）都在 THIRD_PARTY_NOTICES.md 里', () => {
    assert.match(notices, /MIT License/);
    assert.match(notices, /zlib License/i);
    assert.match(notices, /Permission is hereby granted, free of charge/);
  });

  test('LICENSE 是本项目的 MIT', () => {
    const lic = fs.readFileSync(path.join(ROOT, 'LICENSE'), 'utf8');
    assert.match(lic, /MIT License/);
    assert.match(lic, /Copyright \(c\) \d{4} svkoh/);
  });
});
