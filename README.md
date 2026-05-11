# Resume

Static resume website hosted on GitHub Pages.

**Live site:** [https://heros-tempus.github.io/Resume/](https://heros-tempus.github.io/Resume/)

## Editing content

All resume content lives in [`resume.json`](resume.json). Edit that file — the page rebuilds itself on load.

## Previewing locally

`fetch()` is blocked on `file://` URLs, so opening `index.html` directly will show a blank page. Run a local server instead:

```bash
python -m http.server
```

Then open

```text
http://localhost:8000
```
