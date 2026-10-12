# Third-party notices

This project is MIT-licensed (see `LICENSE`). It redistributes one third-party
component; its license and provenance are recorded below.

---

## pako 1.0.11

- **File**: `server/pako.min.js` (loaded by the task pane; used to inflate the
  deflate-compressed entries inside Office `.docx` / `.xlsx` / `.pptx` ZIP
  containers)
- **Upstream**: <https://github.com/nodeca/pako> (`dist/pako.min.js`)
- **Version**: 1.0.11
- **License**: `(MIT AND Zlib)` — MIT for pako itself, Zlib for the
  inflate/deflate port derived from zlib
- **Upstream body SHA-256**:
  `29cfa0542b9864b6640cfdcc8f15b184db6431b532db2702a3ce97e5f2a7fc1c`
  (45,988 bytes — the exact bytes of `dist/pako.min.js` as published on npm)
- **Modification**: a 4-line `/*! ... */` provenance comment was prepended to
  the upstream bytes. This is the only change; the 45,988-byte body is
  unmodified, and the hash above is the verification handle for that claim.
- **Vendored file size**: 46,303 bytes (45,988 upstream + 315 bytes of
  provenance comment). `test/pako-provenance.test.mjs` re-derives both numbers
  and the SHA-256 from the file on disk, so this record cannot silently drift.

Upstream 1.0.11 ships `dist/pako.min.js` with no banner at all, which is why the
provenance comment had to be added here rather than preserved.

### MIT License (pako)

```
(The MIT License)

Copyright (C) 2014-2017 by Vitaly Puzrin and Andrei Tuputcyn

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```

### Zlib License (zlib-derived deflate/inflate code)

```
Copyright notice:

 (C) 1995-2026 Jean-loup Gailly and Mark Adler

  This software is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.

  Jean-loup Gailly        Mark Adler
  jloup@gzip.org          madler@alumni.caltech.edu
```

---

## Not bundled

Nothing else is vendored. `server/server.js` uses only the Node.js standard
library; the host/client halves under `lib/` use only Node built-ins and the
`react` / `react/jsx-runtime` modules that the dsh web client supplies at
runtime. The Office.js library is loaded by `taskpane.html` from Microsoft's
CDN (`appsforoffice.microsoft.com`) and is not redistributed here.
