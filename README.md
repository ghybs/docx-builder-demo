# DocX builder demo

## Demo

[Online Demo](https://ghybs.github.io/docx-builder-demo/index.html)

## Build

```bash
npm run install
# this script assumes webpack (^3.4.1) is installed globally.
# if not, you can add it locally: npm install webpack
npm run build
```

This will bundle the `src/main.js` file with its dependencies into `docs/js/index.js` file.

Then open the file `docs/js/index.html` in your browser.
