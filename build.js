const fs = require('fs');
const path = require('path');
const vm = require('vm');
const root = __dirname;
const BUILD_ID = 'build-' + new Date().toISOString().replace(/[:.]/g, '-');

const tpl = fs.readFileSync(path.join(root, 'template.html'), 'utf8');
const xlsx = fs.readFileSync(path.join(root, 'libs', 'xlsx.full.min.js'), 'utf8');
const plotly = fs.readFileSync(path.join(root, 'libs', 'plotly.min.js'), 'utf8');

// 把库代码中任何 </script 转义成 <\/script，避免浏览器提前结束脚本块
const escapeScript = s => s.replace(/<\/script/gi, '<\\/script');

// 【关键】必须用函数式替换：库代码里含有 $& $' $` $1 等序列，
// 若直接作为 replace 的字符串参数，会被当成特殊替换模式，把库代码改坏。
const inject = (html, marker, label, lib) => {
  const parts = html.split(marker);
  if (parts.length !== 2) throw new Error(`占位符 ${marker} 出现 ${parts.length - 1} 次（应恰好 1 次）`);
  const block = '\n<script>\n/* ===== ' + label + ' (inlined) ===== */\n' + escapeScript(lib) + '\n</script>\n';
  return parts.join(block);
};

let out = tpl;
out = inject(out, '<!--XLSX_LIB-->', 'SheetJS xlsx.full.min.js 0.18.5', xlsx);
out = inject(out, '<!--PLOTLY_LIB-->', 'plotly.js 2.35.2', plotly);

// 写入构建标识，方便确认浏览器不是缓存的旧版本
out = out.replace('<head>', '<head>\n<meta name="build-id" content="' + BUILD_ID + '">');

const dest = path.join(root, 'dashboard-web.html');
fs.writeFileSync(dest, out, 'utf8');

// ---------------- 自检 ----------------
const final = fs.readFileSync(dest, 'utf8');
const fail = [];

// 1) 标签结构
const openers = (final.match(/<script>/g) || []).length;
let unescaped = 0, m, re = /<\/script/gi;
while ((m = re.exec(final))) if (final[m.index - 1] !== '\\') unescaped++;
if (openers !== 3) fail.push(`<script> 数=${openers}（应 3）`);
if (unescaped !== 3) fail.push(`未转义 </script> 数=${unescaped}（应 3）`);
if (/<script[^>]+src=/.test(final)) fail.push('仍存在外链 <script src>');
if (final.includes('<!--XLSX_LIB-->') || final.includes('<!--PLOTLY_LIB-->')) fail.push('占位符未被替换');

// 2) 取出每个脚本体，原样（不做任何反转义）做语法解析 —— 能抓住"库被改坏"这类致命问题
const starts = [];
re = /<script[^>]*>/g;
while ((m = re.exec(final))) starts.push(m.index + m[0].length);
const bodies = starts.map(s => final.slice(s, final.indexOf('</script>', s)));

bodies.forEach((b, i) => {
  try { new vm.Script(b, { filename: 'block' + i }); }
  catch (e) { fail.push(`块${i} 语法错误: ${e.message}`); }
});

// 3) 真跑一遍表格解析库，确认 XLSX 可用（Plotly 依赖真实 document，只在浏览器里跑，这里只验语法与关键接口）
const sb = { console, setTimeout, clearTimeout, setInterval, clearInterval, TextDecoder, TextEncoder };
sb.window = sb; sb.self = sb; sb.globalThis = sb;
vm.createContext(sb);
try { vm.runInContext(bodies[0], sb, { timeout: 60000 }); }
catch (e) { fail.push(`XLSX 库执行报错: ${e.message}`); }
if (!sb.XLSX) fail.push('执行后 XLSX 未定义');
else if (typeof sb.XLSX.read !== 'function') fail.push('XLSX.read 不是函数');
else if (typeof sb.XLSX.utils.sheet_to_json !== 'function') fail.push('XLSX.utils.sheet_to_json 不是函数');
if (!bodies[1].includes('newPlot') || !bodies[1].includes('Plotly')) fail.push('Plotly 库内容不完整');

console.log('WROTE', dest, (fs.statSync(dest).size / 1024 / 1024).toFixed(2), 'MB');
console.log('BUILD_ID:', BUILD_ID);
console.log('脚本块长度:', bodies.map(b => b.length).join(' / '));
console.log('XLSX.version =', sb.XLSX && sb.XLSX.version);
if (fail.length) {
  console.error('\n校验失败：\n - ' + fail.join('\n - '));
  process.exit(1);
}
console.log('校验通过 ✅');
