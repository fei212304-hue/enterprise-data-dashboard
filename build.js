const fs = require('fs');
const path = require('path');
const root = __dirname;

const tpl = fs.readFileSync(path.join(root, 'template.html'), 'utf8');
const xlsx = fs.readFileSync(path.join(root, 'libs', 'xlsx.full.min.js'), 'utf8');
const plotly = fs.readFileSync(path.join(root, 'libs', 'plotly.min.js'), 'utf8');

// 关键：把库代码中任何 </script 都转义成 <\/script，避免浏览器提前结束脚本块
const escapeScript = s => s.replace(/<\/script/gi, '<\\/script');

let out = tpl;
out = out.replace('<!--XLSX_LIB-->', '<script>\n' + escapeScript(xlsx) + '\n</script>');
out = out.replace('<!--PLOTLY_LIB-->', '<script>\n' + escapeScript(plotly) + '\n</script>');

const dest = path.join(root, 'dashboard-web.html');
fs.writeFileSync(dest, out, 'utf8');

// ---------- 自检 ----------
const final = fs.readFileSync(dest, 'utf8');
const openers = (final.match(/<script>/g) || []).length;
let unescaped = 0; let m; const re = /<\/script/gi;
while ((m = re.exec(final))) {
  if (final[m.index - 1] !== '\\') unescaped++;
}
const hasXlsx = final.includes('SheetJS');
const hasPlotly = final.includes('plotly.js v2.35.2') && final.includes('Plotly.newPlot');
console.log('WROTE', dest, (fs.statSync(dest).size / 1024 / 1024).toFixed(2), 'MB');
console.log('真实 <script> 起始标签数:', openers, '(应 = 3)');
console.log('未转义 </script> 数:', unescaped, '(应 = 3)');
console.log('含 xlsx(SheetJS):', hasXlsx, '| 含 plotly:', hasPlotly);
if (openers !== 3 || unescaped !== 3 || !hasXlsx || !hasPlotly) {
  console.error('校验失败：脚本标签结构或库注入异常！');
  process.exit(1);
}
console.log('校验通过 ✅');
