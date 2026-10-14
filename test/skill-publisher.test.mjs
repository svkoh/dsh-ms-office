import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { publishSkillTree, SKILL_MANIFEST } from "../lib/skill-publisher.mjs";

function fixture(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "dsh-skill-publisher-"));
  // This is an exclusively owned temporary tree, never a real skill directory.
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const src = path.join(root, "source");
  const dest = path.join(root, "destination");
  fs.mkdirSync(src);
  const put = (base, relative, body) => {
    const file = path.join(base, relative);
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, body);
    return file;
  };
  const read = (base, relative) => fs.readFileSync(path.join(base, relative), "utf8");
  put(src, "SKILL.md", "中文说明\n");
  return { root, src, dest, put, read };
}

function symlinkOrSkip(t, target, link, type) {
  try { fs.symlinkSync(target, link, type); return true; }
  catch (error) {
    if (error.code === "EPERM" || error.code === "EACCES" || error.code === "ENOSYS") {
      t.skip(`Symlinks unavailable: ${error.code}`);
      return false;
    }
    throw error;
  }
}

test("publishes full relative paths and binary bytes; updates/adds/removes resources with unchanged SKILL", (t) => {
  const { src, dest, put, read } = fixture(t);
  put(src, "references/shared.md", "reference");
  put(src, "resources/shared.md", "resource");
  put(src, "resources/data.bin", Buffer.from([0, 255, 1]));
  publishSkillTree(src, dest);
  assert.equal(read(dest, "references/shared.md"), "reference");
  assert.equal(read(dest, "resources/shared.md"), "resource");
  assert.deepEqual(fs.readFileSync(path.join(dest, "resources/data.bin")), Buffer.from([0, 255, 1]));
  put(src, "resources/shared.md", "changed");
  put(src, "scripts/check.mjs", "export {};\n");
  fs.unlinkSync(path.join(src, "references/shared.md"));
  const result = publishSkillTree(src, dest);
  assert.equal(result.copied, 2);
  assert.equal(result.removed, 1);
  assert.equal(read(dest, "resources/shared.md"), "changed");
  assert.equal(read(dest, "scripts/check.mjs"), "export {};\n");
  assert.equal(fs.existsSync(path.join(dest, "references/shared.md")), false);
  assert.equal(read(dest, "SKILL.md"), "中文说明\n");
});

test("identical trees are idempotent and do not rewrite files or manifest", (t) => {
  const { src, dest, put } = fixture(t);
  put(src, "resources/a.json", "{}");
  publishSkillTree(src, dest);
  const files = ["SKILL.md", "resources/a.json", SKILL_MANIFEST];
  const before = files.map((file) => fs.statSync(path.join(dest, file)).mtimeMs);
  assert.deepEqual(publishSkillTree(src, dest), {
    copied: 0, removed: 0, manifestChanged: false, changed: false, lang: "zh",
  });
  assert.deepEqual(files.map((file) => fs.statSync(path.join(dest, file)).mtimeMs), before);
});

test("English selection, language switching and missing variant fallback", (t) => {
  const { src, dest, put, read } = fixture(t);
  put(src, "SKILL.en.md", "English\n");
  publishSkillTree(src, dest, { lang: "en-US" });
  assert.equal(read(dest, "SKILL.md"), "English\n");
  assert.equal(read(dest, "SKILL.en.md"), "English\n");
  publishSkillTree(src, dest, { lang: "zh" });
  assert.equal(read(dest, "SKILL.md"), "中文说明\n");
  put(src, "SKILL.zh.md", "中文变体\n");
  publishSkillTree(src, dest, { lang: "zh-CN" });
  assert.equal(read(dest, "SKILL.md"), "中文变体\n");
  fs.unlinkSync(path.join(src, "SKILL.en.md"));
  publishSkillTree(src, dest, { lang: "en" });
  assert.equal(read(dest, "SKILL.md"), "中文说明\n");
  assert.equal(fs.existsSync(path.join(dest, "SKILL.en.md")), false);
});

test("repairs missing/corrupt resources and manifest without losing user files", (t) => {
  const { src, dest, put, read } = fixture(t);
  put(src, "resources/a.json", "original");
  put(src, "references/guide.md", "guide");
  publishSkillTree(src, dest);
  put(dest, "notes/user.md", "keep me");
  put(dest, "resources/a.json", "corrupt");
  fs.unlinkSync(path.join(dest, "references/guide.md"));
  put(dest, SKILL_MANIFEST, "not json");
  assert.equal(publishSkillTree(src, dest).copied, 2);
  assert.equal(read(dest, "resources/a.json"), "original");
  assert.equal(read(dest, "references/guide.md"), "guide");
  assert.equal(read(dest, "notes/user.md"), "keep me");
  fs.unlinkSync(path.join(dest, SKILL_MANIFEST));
  assert.equal(publishSkillTree(src, dest).manifestChanged, true);
  assert.equal(read(dest, "notes/user.md"), "keep me");
});

test("removes only previously managed files and preserves unknown files in the same directory", (t) => {
  const { src, dest, put, read } = fixture(t);
  put(src, "resources/obsolete.json", "old");
  publishSkillTree(src, dest);
  put(dest, "resources/user.json", "mine");
  fs.unlinkSync(path.join(src, "resources/obsolete.json"));
  publishSkillTree(src, dest);
  assert.equal(read(dest, "resources/user.json"), "mine");
  assert.equal(fs.existsSync(path.join(dest, "resources/obsolete.json")), false);
});

test("malformed and unsafe manifests have no deletion authority", (t) => {
  const { root, src, dest, put, read } = fixture(t);
  publishSkillTree(src, dest);
  const outside = put(root, "outside.txt", "outside");
  put(dest, "user.txt", "user");
  for (const manifest of [
    { version: 99, files: ["user.txt"] },
    { version: 1, files: "user.txt" },
    { version: 1, files: ["user.txt", "../outside.txt"] },
    { version: 1, files: ["user.txt", outside] },
    { version: 1, files: ["user.txt", "C:\\outside.txt"] },
    { version: 1, files: ["user.txt", "..\\outside.txt"] },
    { version: 1, files: ["user.txt", "a//b"] },
    { version: 1, files: ["user.txt", SKILL_MANIFEST] },
    { version: 1, files: ["user.txt", 123] },
  ]) {
    put(dest, SKILL_MANIFEST, JSON.stringify(manifest));
    publishSkillTree(src, dest);
    assert.equal(read(dest, "user.txt"), "user");
    assert.equal(fs.readFileSync(outside, "utf8"), "outside");
  }
});

test("failed publishing keeps the previous manifest and can be retried", (t) => {
  const { src, dest, put, read } = fixture(t);
  publishSkillTree(src, dest);
  const manifest = read(dest, SKILL_MANIFEST);
  put(src, "resources/new.txt", "new");
  fs.mkdirSync(path.join(dest, "resources/new.txt"), { recursive: true });
  assert.throws(() => publishSkillTree(src, dest), /Expected regular skill file/);
  assert.equal(read(dest, SKILL_MANIFEST), manifest);
  // Only the known empty conflict directory in this fixture is removed.
  fs.rmdirSync(path.join(dest, "resources/new.txt"));
  publishSkillTree(src, dest);
  assert.equal(read(dest, "resources/new.txt"), "new");
  assert.ok(JSON.parse(read(dest, SKILL_MANIFEST)).files.includes("resources/new.txt"));
});

test("copy failure leaves old ownership intact until a successful retry", (t) => {
  const { src, dest, put, read } = fixture(t);
  put(src, "old.txt", "obsolete");
  publishSkillTree(src, dest);
  const before = read(dest, SKILL_MANIFEST);
  fs.unlinkSync(path.join(src, "old.txt"));
  put(src, "a.txt", "first");
  put(src, "z.txt", "second");
  const originalWrite = fs.writeFileSync;
  const injected = t.mock.method(fs, "writeFileSync", (file, ...args) => {
    if (file === path.join(dest, "z.txt")) throw Object.assign(new Error("simulated disk failure"), { code: "EIO" });
    return originalWrite(file, ...args);
  });
  try {
    assert.throws(() => publishSkillTree(src, dest), /simulated disk failure/);
    assert.equal(read(dest, SKILL_MANIFEST), before);
    assert.equal(read(dest, "old.txt"), "obsolete");
    assert.equal(read(dest, "a.txt"), "first");
  } finally {
    injected.mock.restore();
  }
  publishSkillTree(src, dest);
  assert.equal(read(dest, "z.txt"), "second");
  assert.equal(fs.existsSync(path.join(dest, "old.txt")), false);
});

test("stale symlink ancestor cannot cause an outside file to be removed", (t) => {
  const { root, src, dest, put, read } = fixture(t);
  put(src, "resources/a.txt", "managed");
  publishSkillTree(src, dest);
  fs.unlinkSync(path.join(src, "resources/a.txt"));
  fs.unlinkSync(path.join(dest, "resources/a.txt"));
  fs.rmdirSync(path.join(dest, "resources"));
  const outside = path.join(root, "outside");
  put(outside, "a.txt", "keep");
  if (!symlinkOrSkip(t, outside, path.join(dest, "resources"), "junction")) return;
  publishSkillTree(src, dest);
  assert.equal(read(outside, "a.txt"), "keep");
  assert.ok(fs.lstatSync(path.join(dest, "resources")).isSymbolicLink());
});

test("active destination symlinks and source symlinks are refused", (t) => {
  const { root, src, dest, put, read } = fixture(t);
  put(src, "resources/a.txt", "managed");
  publishSkillTree(src, dest);
  fs.unlinkSync(path.join(dest, "resources/a.txt"));
  fs.rmdirSync(path.join(dest, "resources"));
  const outside = path.join(root, "outside");
  put(outside, "a.txt", "keep");
  if (!symlinkOrSkip(t, outside, path.join(dest, "resources"), "junction")) return;
  assert.throws(() => publishSkillTree(src, dest), /Symlink/);
  assert.equal(read(outside, "a.txt"), "keep");
  if (!symlinkOrSkip(t, outside, path.join(src, "linked"), "junction")) return;
  assert.throws(() => publishSkillTree(src, path.join(root, "another")), /Symlink/);
});

test("destination root and manifest symlinks cannot redirect publishing", (t) => {
  const { root, src, dest, put, read } = fixture(t);
  const outside = path.join(root, "outside");
  put(outside, "SKILL.md", "keep");
  if (!symlinkOrSkip(t, outside, dest, "junction")) return;
  assert.throws(() => publishSkillTree(src, dest), /Symlink/);
  assert.equal(read(outside, "SKILL.md"), "keep");
  const safeDest = path.join(root, "safe");
  fs.mkdirSync(safeDest);
  const other = put(root, "other.json", "{}");
  if (!symlinkOrSkip(t, other, path.join(safeDest, SKILL_MANIFEST), "file")) return;
  assert.throws(() => publishSkillTree(src, safeDest), /Symlink/);
  assert.equal(fs.readFileSync(other, "utf8"), "{}");
});

test("source and destination overlap is refused", (t) => {
  const { src, dest } = fixture(t);
  assert.throws(() => publishSkillTree(src, src), /overlap/);
  assert.throws(() => publishSkillTree(src, path.join(src, "nested")), /overlap/);
  assert.throws(() => publishSkillTree(path.join(dest, "nested"), dest), /overlap/);
});
