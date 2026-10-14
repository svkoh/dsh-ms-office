import fs from "node:fs";
import path from "node:path";
import { randomUUID } from "node:crypto";

export const SKILL_MANIFEST = ".dsh-skill-manifest.json";

function statIfPresent(file) {
  try { return fs.lstatSync(file); }
  catch (error) { if (error.code === "ENOENT") return null; throw error; }
}

// Reject portable escape spellings even when running on another platform.
function validRelative(file) {
  return typeof file === "string" && file.length > 0 &&
    !file.includes("\\") && !file.includes(":") && !file.includes("\0") &&
    !path.posix.isAbsolute(file) && !path.win32.isAbsolute(file) &&
    file.split("/").every((part) => part && part !== "." && part !== "..") &&
    file !== SKILL_MANIFEST;
}

function absoluteFile(root, relative) {
  if (!validRelative(relative)) throw pathConflict(`Unsafe skill path: ${relative}`);
  const file = path.resolve(root, ...relative.split("/"));
  const inside = path.relative(root, file);
  if (inside.startsWith(`..${path.sep}`) || inside === ".." || path.isAbsolute(inside)) {
    throw pathConflict(`Skill path escapes destination: ${relative}`);
  }
  return file;
}

function pathConflict(message) {
  return Object.assign(new Error(message), { code: "SKILL_PATH_CONFLICT" });
}

// Check every existing ancestor, including the destination root. Never traverse
// a symlink/junction for reads, writes or removals, even if it points inward.
function checkPath(file, { leafFile = false } = {}) {
  const resolved = path.resolve(file);
  const { root } = path.parse(resolved);
  let cursor = root;
  const parts = resolved.slice(root.length).split(path.sep).filter(Boolean);
  for (let index = 0; index < parts.length; index += 1) {
    cursor = path.join(cursor, parts[index]);
    const stat = statIfPresent(cursor);
    if (!stat) continue;
    if (stat.isSymbolicLink()) throw pathConflict(`Symlink in skill path: ${cursor}`);
    const leaf = index === parts.length - 1;
    if (!leaf || !leafFile) {
      if (!stat.isDirectory()) throw pathConflict(`Expected skill directory: ${cursor}`);
    } else if (!stat.isFile()) {
      throw pathConflict(`Expected regular skill file: ${cursor}`);
    }
  }
}

function readTree(root) {
  checkPath(root);
  const files = new Map();
  function visit(directory, prefix) {
    for (const entry of fs.readdirSync(directory, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
      const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
      const file = absoluteFile(root, relative);
      if (entry.isSymbolicLink()) throw pathConflict(`Symlink in skill source: ${relative}`);
      if (entry.isDirectory()) visit(file, relative);
      else if (entry.isFile()) files.set(relative, fs.readFileSync(file));
      else throw pathConflict(`Unsupported skill source entry: ${relative}`);
    }
  }
  visit(root, "");
  return files;
}

function readManifest(file) {
  try {
    const value = JSON.parse(fs.readFileSync(file, "utf8"));
    // A malformed manifest has no authority to delete anything.
    if (value?.version !== 1 || !Array.isArray(value.files) || !value.files.every(validRelative)) return [];
    return [...new Set(value.files)];
  } catch (error) {
    if (error instanceof SyntaxError || error.code === "ENOENT") return [];
    throw error;
  }
}

/**
 * Publish one complete skill tree using only filesystem operations.
 * Unknown destination files are preserved. Only regular files recorded by a
 * valid previous manifest may be removed; directories are never recursively
 * deleted. Copy failures leave the old manifest in place for the next retry.
 */
export function publishSkillTree(source, destination, { lang = "zh" } = {}) {
  const src = path.resolve(source);
  const dest = path.resolve(destination);
  for (const [parent, child] of [[src, dest], [dest, src]]) {
    const relative = path.relative(parent, child);
    if (!relative || (!relative.startsWith(`..${path.sep}`) && relative !== ".." && !path.isAbsolute(relative))) {
      throw pathConflict("Skill source and destination must not overlap");
    }
  }
  const files = readTree(src);
  if (!files.has("SKILL.md")) throw pathConflict(`Missing SKILL.md: ${src}`);
  const language = String(lang).toLowerCase().startsWith("en") ? "en" : "zh";
  const variant = `SKILL.${language}.md`;
  if (files.has(variant)) files.set("SKILL.md", files.get(variant));

  checkPath(dest);
  const manifestFile = path.join(dest, SKILL_MANIFEST);
  checkPath(manifestFile, { leafFile: true });
  const previous = readManifest(manifestFile);
  const obsolete = previous.filter((relative) => !files.has(relative));
  const removals = [];
  for (const relative of obsolete) {
    const file = absoluteFile(dest, relative);
    // User-replaced symlinks/directories are not ours to remove. In particular,
    // do not follow an obsolete parent directory that has become a symlink.
    try { checkPath(file, { leafFile: true }); }
    catch (error) {
      if (error.code === "SKILL_PATH_CONFLICT") continue;
      throw error;
    }
    if (statIfPresent(file)?.isFile()) removals.push(file);
  }
  // Validate active targets before making changes. Directory/file conflicts are
  // reported without removing user data; resolving the conflict allows retry.
  for (const relative of files.keys()) checkPath(absoluteFile(dest, relative), { leafFile: true });
  fs.mkdirSync(dest, { recursive: true });
  let copied = 0;
  for (const [relative, bytes] of files) {
    const target = absoluteFile(dest, relative);
    checkPath(target, { leafFile: true });
    const current = statIfPresent(target);
    if (current && fs.readFileSync(target).equals(bytes)) continue;
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, bytes);
    copied += 1;
  }
  let removed = 0;
  for (const file of removals) {
    checkPath(file, { leafFile: true });
    if (statIfPresent(file)?.isFile()) {
      fs.unlinkSync(file);
      removed += 1;
    }
  }
  const manifest = `${JSON.stringify({ version: 1, files: [...files.keys()].sort() }, null, 2)}\n`;
  checkPath(manifestFile, { leafFile: true });
  const manifestChanged = !statIfPresent(manifestFile) || fs.readFileSync(manifestFile, "utf8") !== manifest;
  if (manifestChanged) {
    const temporary = path.join(dest, `${SKILL_MANIFEST}.${randomUUID()}.tmp`);
    try {
      fs.writeFileSync(temporary, manifest, { flag: "wx" });
      fs.renameSync(temporary, manifestFile);
    } finally {
      if (statIfPresent(temporary)?.isFile()) fs.unlinkSync(temporary);
    }
  }
  return { copied, removed, manifestChanged, changed: copied > 0 || removed > 0 || manifestChanged, lang: language };
}
