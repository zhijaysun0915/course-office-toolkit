const fs = require('fs');
const path = require('path');

const rootDir = path.resolve(__dirname, '..');
const indexPath = path.join(rootDir, 'index.html');
const cssPath = path.join(rootDir, 'css', 'style.css');
const demoDataPath = path.join(rootDir, 'js', 'demo-data.js');
const appPath = path.join(rootDir, 'js', 'app.js');
const outputPath = path.join(rootDir, '辦公室萬能工具箱_移動辦公單一檔.html');

let html = fs.readFileSync(indexPath, 'utf-8');
const css = fs.readFileSync(cssPath, 'utf-8');
const demoData = fs.readFileSync(demoDataPath, 'utf-8');
const app = fs.readFileSync(appPath, 'utf-8');

// 替換 CSS
html = html.replace(
  '<link rel="stylesheet" href="css/style.css" />',
  `<style>\n${css}\n</style>`
);

// 替換 JS
html = html.replace(
  '<script src="js/demo-data.js"></script>',
  `<script>\n${demoData}\n</script>`
);

html = html.replace(
  '<script src="js/app.js"></script>',
  `<script>\n${app}\n</script>`
);

fs.writeFileSync(outputPath, html, 'utf-8');
console.log(`✅ 成功打包單一隨身檔: ${outputPath} (${(fs.statSync(outputPath).size / 1024).toFixed(1)} KB)`);
