# scripts/

Build scripts for the resume project. Run everything from the **repo root**.

---

## build-docx.js

Converts `resume.json` → `resume.docx`.

`resume.json` is the single source of truth. Edit it, then run this script to
regenerate the Word document before committing.

### First-time setup

```bash
npm install
```

This installs the one dependency (`docx`) into a local `node_modules/` folder.
You only need to do this once (or after a fresh clone).

### Running it

```bash
node scripts/build-docx.js
```

Output: `resume.docx` in the repo root.

### Typical workflow

1. Edit `resume.json`
2. `node scripts/build-docx.js`
3. Verify `resume.docx` looks right
4. `git add resume.json resume.docx && git commit -m "update resume"`
5. `git push` — GitHub Pages picks up the rest

---

## .gitignore notes

Make sure your `.gitignore` includes:

```text
node_modules/
```
