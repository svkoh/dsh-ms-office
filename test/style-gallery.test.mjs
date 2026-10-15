import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { buildGallery } from '../skills/office-document-styles/scripts/gallery.mjs';

for (const lang of ['zh', 'en']) test(`published ${lang} gallery matches its generator and all assets exist`, () => {
  const document = new URL(`../docs/style-gallery${lang === 'en' ? '.en' : ''}.md`, import.meta.url);
  const markdown = buildGallery(lang);
  assert.equal(fs.readFileSync(document, 'utf8').replaceAll('\r\n', '\n'), markdown);
  assert.equal([...markdown.matchAll(/!\[/g)].length, 17);
  for (const [, href] of markdown.matchAll(/!?\[[^\]]*\]\(([^)]+)\)/g)) {
    assert(fs.existsSync(new URL(href, document)), `Missing gallery asset: ${href}`);
  }
});
