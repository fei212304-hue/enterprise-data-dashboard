#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
通用数据看板生成器
==================

把任意一份 Excel / CSV 数据，自动生成一份「可筛选、可导出、离线可看」的交互式看板 HTML。

用法（最简单）：
    python build_dashboard.py 数据文件.xlsx

更多用法：
    python build_dashboard.py 销售表.xlsx -o 我的看板.html
    python build_dashboard.py 问卷.csv --main 满意度
    python build_dashboard.py 数据.xlsx --sheet 2024年 --title 2024年度经营看板
    python build_dashboard.py 数据.xlsx --filters 行业 地区

产出：一个独立的 .html 文件，双击用浏览器打开即可，无需联网、无需安装任何东西。

依赖：pandas、numpy、plotly（用于取 plotly.min.js 并内嵌）、openpyxl（读 xlsx 需要）
"""

import argparse
import io
import json
import os
import sys
from datetime import datetime

try:
    import numpy as np
    import pandas as pd
except ImportError:
    sys.stderr.write(
        "缺少依赖。请先安装：\n"
        "    pip install pandas numpy plotly openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple\n"
    )
    sys.exit(1)

MAX_CATEGORY_VALUE = 12      # 分类列取值超过这个数量就不做成筛选器
MAX_FILTERS = 4              # 最多几个筛选器
MAX_TOP_ROWS = 12            # 排行榜显示多少行
MAX_HEATMAP_COLS = 12        # 相关性热力图最多放多少列


# ---------------------------------------------------------------- 读取源文件
def find_plotly_js():
    """找到本地的 plotly.min.js（优先用当前 Python 环境里装的那份）。"""
    try:
        import plotly
        base = os.path.dirname(os.path.abspath(plotly.__file__))
        p = os.path.join(base, "package_data", "plotly.min.js")
        if os.path.exists(p):
            return p
    except Exception:
        pass
    for p in [os.path.join(os.path.dirname(os.path.abspath(__file__)), "plotly.min.js"),
              os.path.join(os.getcwd(), "plotly.min.js")]:
        if os.path.exists(p):
            return p
    return None


def read_source(path, sheet=None, csv_encoding=None):
    """读 Excel 或 CSV，返回 DataFrame。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        if sheet:
            return pd.read_excel(path, sheet_name=sheet)
        # 没指定工作表就读第一个
        return pd.read_excel(path, sheet_name=0)
    if ext == ".csv":
        if csv_encoding:
            return pd.read_csv(path, encoding=csv_encoding)
        for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
            try:
                return pd.read_csv(path, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise RuntimeError("CSV 编码识别失败，请用 --encoding 手动指定（常见：utf-8、gbk）")
    raise RuntimeError("不支持的文件类型：%s（只支持 .xlsx / .xls / .csv）" % ext)


# ---------------------------------------------------------------- 自动识别字段
def is_id_column(col, series, n_rows, n_unique):
    """
    判断某列是不是「编号」类字段（企业ID、订单号、序号……）。
    这类字段算均值、算相关系数都没有意义，要从数值字段里剔除。

    注意：只认「列名像编号」或「整数且取值不重复」两种情况。
    浮点数列即使每行都不重复（比如一堆随机数），也不能当成编号。
    """
    low = str(col).strip().lower()
    for kw in ("id", "编号", "序号", "流水号", "单号", "订单号"):
        if low == kw or low.endswith(kw) or low.startswith(kw):
            return True
    if n_rows <= 5 or n_unique != n_rows:
        return False
    if not pd.api.types.is_integer_dtype(series):
        return False
    s = series.dropna()
    if len(s) == 0:
        return False
    # 整数、每行都不同、且取值跨度接近行数 → 是编号（如 1,2,3...1010）
    return (s.max() - s.min()) <= n_rows * 1.2


def pick_main_metric(numeric, n_rows, unique_counts, specified=None):
    """挑主指标：用户指定优先 > 像金额/销量的字段 > 第一个非编号数值字段。"""
    if specified:
        return str(specified)
    hot = ("营业额", "营收", "销售额", "销售收入", "金额", "交易额", "销量", "销售额度",
           "利润", "收入", "gmv", "amount", "sales", "revenue", "total")
    for c in numeric:
        low = str(c).strip().lower()
        for kw in hot:
            if kw in low:
                return c
    return numeric[0] if numeric else None


def classify_columns(df, exclude):
    """把字段分成数值列、候选分类列、其它列（编号类字段会被单独剔除）。"""
    numeric, candidates, other, ids = [], [], [], []
    n = len(df)
    for c in df.columns:
        if str(c) in exclude:
            continue
        s = df[c]
        if pd.api.types.is_numeric_dtype(s):
            nun = int(s.nunique(dropna=True))
            if is_id_column(c, s, n, nun):
                ids.append(str(c))
            else:
                numeric.append(str(c))
        elif pd.api.types.is_datetime64_any_dtype(s):
            other.append(str(c))
        else:
            nun = s.nunique(dropna=True)
            if 1 < nun <= MAX_CATEGORY_VALUE:
                candidates.append((str(c), int(nun)))
            else:
                other.append(str(c))
    # 分类列按取值数量从少到多排列，取值少的更适合做筛选器
    candidates.sort(key=lambda x: x[1])
    return numeric, [c for c, _ in candidates], other, ids


def pick_filters(df, cats, want=None):
    """挑选用于筛选的字段（最多 MAX_FILTERS 个）。"""
    if want:
        chosen = [str(w) for w in want]
        missing = [w for w in chosen if w not in list(df.columns)]
        if missing:
            raise RuntimeError("指定的筛选字段在表里找不到：%s\n表里的字段有：%s"
                               % ("、".join(missing), "、".join(str(c) for c in df.columns)))
        return chosen[:MAX_FILTERS]
    return cats[:MAX_FILTERS]


def clean_data(df, numeric, cats):
    """
    做一遍最基本的清洗：
      - 数值列的缺失值用该列中位数填补（并记录填补了几个）
      - 分类型的缺失值标成「未填写」
      - 不去重，但统计完全重复行的数量
    """
    report = {"filled": {}, "filled_total": 0, "duplicates": int(df.duplicated().sum()),
              "missing_before": int(df.isna().sum().sum())}
    df = df.copy()

    for c in numeric:
        n = int(df[c].isna().sum())
        if n:
            med = df[c].median()
            if pd.isna(med):
                med = 0
            df[c] = df[c].fillna(med)
            report["filled"][str(c)] = n
            report["filled_total"] += n

    for c in cats:
        if df[c].isna().any():
            df[c] = df[c].astype(object).where(df[c].notna(), "未填写")

    return df, report


# ---------------------------------------------------------------- 准备数据
def to_records(df, numeric, cats):
    """转成干净的 JSON 记录：数值是真数字，分类是字符串。"""
    cols = list(df.columns)
    recs = []
    for row in df.itertuples(index=False):
        item = {}
        for col, v in zip(cols, row):
            col = str(col)
            if col in numeric:
                item[col] = None if (v is None or pd.isna(v)) else float(v)
            else:
                item[col] = "未填写" if (v is None or (not isinstance(v, str) and pd.isna(v))) else str(v)
        recs.append(item)
    return recs


def build_spec(df, numeric, cats, filters, main_metric, title, source_name):
    catsorder = {}
    for c in filters:
        vals = sorted([str(v) for v in df[c].dropna().unique().tolist()])
        catsorder[c] = vals
    return {
        "title": title,
        "source": source_name,
        "numeric": numeric,
        "cats": cats,
        "filters": [{"col": c, "values": catsorder[c]} for c in filters],
        "main": main_metric,
        "rows": int(len(df)),
        "cols": int(df.shape[1]),
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def corr_pair(df, numeric, a, b):
    s = df[[a, b]].dropna()
    if len(s) < 3:
        return None
    if s[a].std() == 0 or s[b].std() == 0:
        return None
    return float(s[a].corr(s[b]))


def compute_missing(df):
    """按列统计原始缺失数量，用于在 HTML 里画那张静态的缺失值图。"""
    s = df.isna().sum()
    s = s[s > 0].sort_values(ascending=False)
    return [{"col": str(k), "n": int(v)} for k, v in s.items()]


# ---------------------------------------------------------------- HTML 模板
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<style>
  :root{ --ink:#1f2937; --sub:#6b7280; --line:#e5e7eb; --brand:#2563eb;
         --brand-soft:#eff6ff; --warn-bg:#fef2f2; --warn:#b91c1c; --bg:#f5f7fa; }
  *{box-sizing:border-box;}
  body{margin:0;background:var(--bg);color:var(--ink);
       font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei","PingFang SC",sans-serif;
       -webkit-font-smoothing:antialiased;}
  .wrap{max-width:1180px;margin:0 auto;padding:28px 22px 60px;}
  h1{font-size:22px;margin:0 0 4px;font-weight:650;letter-spacing:-.2px;}
  .sub{color:var(--sub);font-size:13px;margin:0 0 18px;line-height:1.7;}
  .toolbar{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin-bottom:16px;}
  .toolbar .row{display:flex;flex-wrap:wrap;gap:14px 26px;align-items:flex-start;}
  .fgroup .lab{font-size:12px;color:var(--sub);margin-bottom:6px;letter-spacing:.3px;}
  .chip{display:inline-block;padding:4px 11px;margin:0 6px 6px 0;border-radius:999px;font-size:12px;
        border:1px solid var(--line);background:#fff;color:var(--sub);cursor:pointer;user-select:none;
        transition:all .12s ease;}
  .chip:hover{border-color:#93c5fd;}
  .chip.on{background:var(--brand-soft);border-color:var(--brand);color:var(--brand);font-weight:600;}
  .btn{border:1px solid var(--line);background:#fff;border-radius:8px;padding:7px 14px;font-size:12.5px;
       cursor:pointer;color:var(--ink);transition:all .12s ease;}
  .btn:hover{border-color:var(--brand);color:var(--brand);}
  .btn.primary{background:var(--brand);border-color:var(--brand);color:#fff;}
  .btn.primary:hover{background:#1d4ed8;color:#fff;}
  .acts{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px;padding-top:12px;border-top:1px dashed var(--line);}
  .meta{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:16px;}
  .mi{background:#fff;border:1px solid var(--line);border-radius:10px;padding:9px 14px;flex:1 1 160px;}
  .mi .k{font-size:11.5px;color:var(--sub);letter-spacing:.3px;}
  .mi .v{font-size:17px;font-weight:650;margin-top:3px;}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:18px;}
  .kpi{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px 16px;}
  .kpi .k{font-size:12px;color:var(--sub);}
  .kpi .v{font-size:21px;font-weight:650;margin:5px 0 3px;}
  .kpi .d{font-size:11.5px;color:var(--sub);}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
  @media(max-width:900px){ .grid{grid-template-columns:1fr;} }
  .card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px 16px 12px;min-height:330px;}
  .card.full{grid-column:1 / -1;}
  .chd{display:flex;justify-content:space-between;align-items:baseline;gap:10px;margin-bottom:6px;}
  .ct{font-size:13.5px;font-weight:650;}
  .cn{font-size:11px;color:var(--sub);margin:0 0 6px;line-height:1.6;}
  .save{border:1px solid var(--line);background:#fff;border-radius:6px;font-size:11.5px;
        padding:4px 9px;cursor:pointer;color:var(--sub);white-space:nowrap;}
  .save:hover{border-color:var(--brand);color:var(--brand);}
  .alert{display:none;background:var(--warn-bg);border:1px solid #fecaca;color:var(--warn);
         border-radius:10px;padding:12px 16px;margin-bottom:16px;font-size:13px;}
  table{width:100%;border-collapse:collapse;font-size:12.5px;}
  th,td{padding:7px 10px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap;}
  th:first-child,td:first-child{text-align:left;}
  th{color:var(--sub);font-weight:600;cursor:pointer;user-select:none;}
  th:hover{color:var(--brand);}
  .wraptable{overflow:auto;max-height:420px;}
  footer{margin-top:26px;font-size:11.5px;color:#9ca3af;line-height:1.9;}
  @media print{ .toolbar,.save{display:none;} .grid{grid-template-columns:1fr;} }
</style>
</head>
<body>
<div class="wrap">
  <h1>__TITLE__</h1>
  <p class="sub">__SOURCE__ &nbsp;·&nbsp; 共 __ROWS__ 行 × __COLS__ 列 &nbsp;·&nbsp; 生成于 __GEN__<br>点击标签可筛选，鼠标移到图上可保存该图为 PNG</p>

  <div class="toolbar">
    <div class="row" id="filters"></div>
    <div class="acts">
      <button class="btn primary" id="exportAll">一键导出全部图片</button>
      <button class="btn" id="reset">重置全部筛选</button>
      <span id="state" style="font-size:12px;color:var(--sub);align-self:center;"></span>
    </div>
  </div>

  <div class="alert" id="alert"></div>

  <div class="meta">
    <div class="mi"><div class="k">当前样本量</div><div class="v" id="mRows">-</div></div>
    <div class="mi"><div class="k">占总样本</div><div class="v" id="mPct">-</div></div>
    <div class="mi"><div class="k">数值字段 / 分类字段</div><div class="v" id="mCols">-</div></div>
    <div class="mi"><div class="k">缺失值填补（原始）</div><div class="v" id="mMiss">-</div></div>
    <div class="mi"><div class="k">完全重复行</div><div class="v" id="mDup">-</div></div>
  </div>

  <div class="kpis" id="kpis"></div>
  <div class="grid" id="grid"></div>

  <footer>
    本看板由 <b>build_dashboard.py</b> 自动生成。数据已内嵌在文件中，断网也能打开。<br>
    __FOOTNOTE__
  </footer>
</div>

<script>__PLOTLYJS__</script>
<script>
var SPEC = __SPEC__;
var DATA = __DATA__;
var MISS = __MISS__;
var CLEAN = __CLEAN__;
var FN = {};

function fmt(n){
  if(n===null||n===undefined||!isFinite(n)) return "-";
  var a=Math.abs(n);
  if(a>=100000000) return (n/100000000).toFixed(2)+"亿";
  if(a>=10000) return (n/10000).toFixed(2)+"万";
  if(a>=1000) return n.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g,",");
  if(a>=1) return n.toFixed(2);
  return n.toFixed(3);
}
function num(x){ return typeof x==='number' && isFinite(x); }
function mean(a){ a=a.filter(num); return a.length? a.reduce(function(s,x){return s+x;},0)/a.length : NaN; }
function corr(xs,ys){
  var i,p=[],n=Math.min(xs.length,ys.length);
  for(i=0;i<n;i++){ if(num(xs[i])&&num(ys[i])) p.push([xs[i],ys[i]]); }
  if(p.length<3) return null;
  var mx=p.reduce(function(s,q){return s+q[0];},0)/p.length;
  var my=p.reduce(function(s,q){return s+q[1];},0)/p.length;
  var sxy=0,sxx=0,syy=0;
  for(i=0;i<p.length;i++){
    sxy+=(p[i][0]-mx)*(p[i][1]-my); sxx+=Math.pow(p[i][0]-mx,2); syy+=Math.pow(p[i][1]-my,2);
  }
  if(sxx===0||syy===0) return null;
  return sxy/Math.sqrt(sxx*syy);
}
function col(rows,c){ return rows.map(function(d){return d[c];}); }
function unique(a){ var s={},o=[]; for(var i=0;i<a.length;i++){ if(!s[a[i]]){s[a[i]]=1;o.push(a[i]);} } return o; }

var sel={};
SPEC.filters.forEach(function(f){ sel[f.col]=new Set(f.values); });

function buildFilters(){
  var host=document.getElementById('filters'); host.innerHTML='';
  SPEC.filters.forEach(function(f){
    var g=document.createElement('div'); g.className='fgroup';
    var lab=document.createElement('div'); lab.className='lab'; lab.textContent=f.col;
    g.appendChild(lab);
    f.values.forEach(function(v){
      var s=document.createElement('span');
      s.className='chip on'; s.textContent=v;
      s.onclick=function(){
        if(sel[f.col].has(v)){ sel[f.col].delete(v); s.className='chip'; }
        else { sel[f.col].add(v); s.className='chip on'; }
        redraw();
      };
      g.appendChild(s);
    });
    host.appendChild(g);
  });
}

function current(){
  return DATA.filter(function(d){
    for(var i=0;i<SPEC.filters.length;i++){
      var f=SPEC.filters[i];
      if(!sel[f.col].has(String(d[f.col]))) return false;
    }
    return true;
  });
}

function groupMean(rows, cat, metric){
  var m={};
  rows.forEach(function(d){
    if(!num(d[metric])) return;
    var k=String(d[cat]); if(!m[k]) m[k]=[]; m[k].push(d[metric]);
  });
  var out=[];
  for(var k in m) out.push([k, mean(m[k]), m[k].length]);
  out.sort(function(a,b){ return b[1]-a[1]; });
  return out;
}

var charts=[];
function addCard(o){
  var g=document.getElementById('grid');
  var card=document.createElement('div'); card.className='card'+(o.full?' full':'');
  var idx=charts.length+1;
  var no=('0'+idx).slice(-2);
  FN[o.id]='图'+no+'_'+String(o.title).replace(/[\\/:*?"<>|]/g,'_');
  var head=document.createElement('div'); head.className='chd';
  var t=document.createElement('div'); t.className='ct';
  t.textContent='图'+no+' · '+o.title;
  var b=document.createElement('button'); b.className='save'; b.textContent='保存图片';
  b.onclick=function(){ Plotly.downloadImage(document.getElementById(o.id),{format:'png',scale:3,filename:FN[o.id]}); };
  head.appendChild(t); head.appendChild(b);
  var note=document.createElement('p'); note.className='cn'; note.textContent=o.note||'';
  var box=document.createElement('div'); box.id=o.id; box.style.height=o.height||'318px';
  if(o.type==='table'){ box.className='wraptable'; box.style.height='420px'; }
  card.appendChild(head); card.appendChild(note); card.appendChild(box);
  g.appendChild(card);
  charts.push(o);
}

var BASE_LAYOUT={
  paper_bgcolor:'#ffffff', plot_bgcolor:'#ffffff',
  font:{family:'-apple-system,"Microsoft YaHei","PingFang SC",sans-serif', size:11.5, color:'#1f2937'},
  margin:{l:70,r:26,t:12,b:44}, hovermode:'closest',
  xaxis:{gridcolor:'#f1f5f9', zerolinecolor:'#e5e7eb'},
  yaxis:{gridcolor:'#f1f5f9', zerolinecolor:'#e5e7eb'}
};
var CFG={displaylogo:false, responsive:true, displayModeBar:false, toImageButtonOptions:{format:'png',scale:3}};
function lay(extra){
  var l=JSON.parse(JSON.stringify(BASE_LAYOUT));
  for(var k in extra){
    if((k==='xaxis'||k==='yaxis'||k==='margin') && Object.prototype.toString.call(extra[k])==='[object Object]'){
      for(var k2 in extra[k]) l[k][k2]=extra[k][k2];   // 合并而不是覆盖，保住网格线等默认样式
    } else { l[k]=extra[k]; }
  }
  return l;
}
var PALETTE=['#2563eb','#0d9488','#f59e0b','#db2777','#7c3aed','#0891b2','#65a30d','#dc2626','#4b5563','#c2410c'];

function drawEmpty(){
  document.getElementById('alert').style.display='block';
  document.getElementById('alert').textContent='当前筛选条件下没有任何数据。请至少勾选每个筛选器里的一个选项。';
  charts.forEach(function(o){
    if(o.type==='table') document.getElementById(o.id).innerHTML='';
    else Plotly.purge(document.getElementById(o.id));
  });
  document.getElementById('kpis').innerHTML='';
  ['mRows','mPct','mCols','mMiss','mDup'].forEach(function(i){ document.getElementById(i).textContent='-'; });
}

function redraw(){
  var rows=current();
  var alert=document.getElementById('alert');
  if(!rows.length){ drawEmpty(); return; }
  alert.style.display='none';

  document.getElementById('mRows').textContent=rows.length.toLocaleString();
  document.getElementById('mPct').textContent=(rows.length/SPEC.rows*100).toFixed(1)+'%';
  document.getElementById('mCols').textContent=SPEC.numeric.length+' / '+SPEC.cats.length;
  document.getElementById('mMiss').textContent=CLEAN.filled_total.toLocaleString()+' 个';
  document.getElementById('mDup').textContent=CLEAN.duplicates.toLocaleString()+' 条';

  var stxt=[];
  SPEC.filters.forEach(function(f){
    if(sel[f.col].size<f.values.length) stxt.push(f.col+': '+Array.from(sel[f.col]).join('/'));
  });
  document.getElementById('state').textContent= stxt.length? '当前筛选 — '+stxt.join('　|　') : '未筛选，显示全部数据';

  drawKpis(rows);
  charts.forEach(function(o){
    var el=document.getElementById(o.id);
    if(o.type==='table'){ o.render(rows,el); return; }
    var r;
    try{ r=o.render(rows); }
    catch(e){ Plotly.purge(el); return; }
    if(!r){ Plotly.purge(el); return; }
    Plotly.react(el, r.data, r.layout, CFG);
  });
}

function drawKpis(rows){
  var host=document.getElementById('kpis'); host.innerHTML='';
  var main=SPEC.main;
  function card(label, val, base, unit){
    var d=document.createElement('div'); d.className='kpi';
    var k=document.createElement('div'); k.className='k'; k.textContent=label;
    var v=document.createElement('div'); v.className='v'; v.textContent=fmt(val);
    var dd=document.createElement('div'); dd.className='d';
    if(num(base)&&base!==0){
      var p=(val-base)/Math.abs(base)*100;
      dd.textContent=(p>=0?'▲ 高于全量 ':'▼ 低于全量 ')+Math.abs(p).toFixed(1)+'%';
      dd.style.color = p>=0? '#2563eb' : '#64748b';
    } else { dd.textContent='全量基准 '+fmt(base); }
    d.appendChild(k); d.appendChild(v); d.appendChild(dd); host.appendChild(d);
  }
  card('样本量（当前筛选）', rows.length, SPEC.rows, '');
  if(main) card(main+' · 均值', mean(col(rows,main)), mean(col(DATA,main)), '');
  SPEC.numeric.filter(function(c){ return c!==main; }).slice(0,4).forEach(function(c){
    card(c+' · 均值', mean(col(rows,c)), mean(col(DATA,c)), '');
  });
}

buildFilters();

// ---- 图表 ----
var mainCat = SPEC.filters.length? SPEC.filters[0].col : (SPEC.cats.length? SPEC.cats[0] : null);

SPEC.filters.slice(0,3).forEach(function(f, i){
  if(SPEC.main===null) return;
  addCard({
    id:'catmean'+i, title:'各'+f.col+'的'+SPEC.main+'均值', height:'318px',
    note:'按'+f.col+'分组计算'+SPEC.main+'的平均值，从高到低排列。柱子上的数字是样本数。',
    render:function(rows){
      var g=groupMean(rows,f.col,SPEC.main);
      if(!g.length) return null;
      return { data:[{ type:'bar', orientation:'h',
        y:g.map(function(x){return x[0];}), x:g.map(function(x){return x[1];}),
        text:g.map(function(x){return 'n='+x[2];}), textposition:'outside', cliponaxis:false,
        marker:{color:PALETTE[i%PALETTE.length]},
        hovertemplate:'%{y}<br>'+SPEC.main+'均值 %{x:.2f}<extra></extra>' }],
        layout:lay({ height:Math.max(280,g.length*34+70), margin:{l:130,r:60,t:12,b:44},
          yaxis:{autorange:'reversed', gridcolor:'#f1f5f9'},
          xaxis:{title:{text:SPEC.main+' 均值'}, gridcolor:'#f1f5f9'} }) };
    }
  });
});

if(mainCat){
  addCard({
    id:'donut', title:mainCat+'构成占比', height:'318px',
    note:'当前筛选结果按'+mainCat+'统计数量占比。',
    render:function(rows){
      var m={};
      rows.forEach(function(d){ var k=String(d[mainCat]); m[k]=(m[k]||0)+1; });
      var ks=Object.keys(m), vs=ks.map(function(k){return m[k];});
      if(!ks.length) return null;
      return { data:[{ type:'pie', labels:ks, values:vs, hole:0.55, textinfo:'label+percent',
        textposition:'outside', marker:{colors:PALETTE}, automargin:true }],
        layout:lay({ height:330, showlegend:false, margin:{l:60,r:60,t:16,b:30} }) };
    }
  });
}

if(SPEC.main){
  addCard({
    id:'hist', title:SPEC.main+'分布', height:'318px',
    note:'直方图看分布形状：单峰还是双峰、有没有明显偏斜。虚线是均值位置。',
    render:function(rows){
      var xs=col(rows,SPEC.main).filter(num);
      if(!xs.length) return null;
      var m=mean(xs);
      return { data:[{ type:'histogram', x:xs, nbinsx:30, marker:{color:'#2563eb',opacity:0.82},
        hovertemplate:SPEC.main+' %{x}<br>频数 %{y}<extra></extra>' }],
        layout:lay({ height:318,
          xaxis:{title:{text:SPEC.main}, gridcolor:'#f1f5f9'},
          yaxis:{title:{text:'频数'}, gridcolor:'#f1f5f9'},
          shapes:[{type:'line', x0:m, x1:m, y0:0, y1:1, yref:'paper',
                   line:{color:'#f59e0b', width:1.5, dash:'dash'}}],
          annotations:[{x:m, y:1, yref:'paper', text:'均值 '+fmt(m), showarrow:false,
                        yanchor:'bottom', font:{size:11,color:'#b45309'}, xanchor:'left'}] }) };
    }
  });
}

if(SPEC.main && mainCat){
  addCard({
    id:'box', title:mainCat+'的'+SPEC.main+'分布（箱线图）', height:'318px',
    note:'箱子覆盖中间一半数据，须线外的小点是离群值。这是识别异常值最常用的图。',
    render:function(rows){
      var cats=unique(col(rows,mainCat));
      if(!cats.length) return null;
      // 按中位数排序，让图更好读
      cats.sort(function(a,b){
        return mean(col(rows.filter(function(d){return String(d[mainCat])===a;}),SPEC.main))
             - mean(col(rows.filter(function(d){return String(d[mainCat])===b;}),SPEC.main));
      });
      return { data:cats.map(function(k,i){
        return { type:'box', name:String(k), y:col(rows.filter(function(d){return String(d[mainCat])===k;}),SPEC.main),
                 boxpoints:'outliers', boxmean:'sd',
                 marker:{color:PALETTE[i%PALETTE.length]}, line:{width:1.4} };
      }), layout:lay({ height:Math.max(300,cats.length*30+90),
        xaxis:{title:{text:mainCat}}, yaxis:{title:{text:SPEC.main}, gridcolor:'#f1f5f9'},
        margin:{l:70,r:26,t:12,b:60} }) };
    }
  });
}

// 与主指标最相关的那个数值列 → 散点图
(function(){
  if(!SPEC.main) return;
  var others=SPEC.numeric.filter(function(c){ return c!==SPEC.main; });
  if(!others.length) return;
  var best=null,bv=0;
  others.forEach(function(c){
    var r=corr(col(DATA,SPEC.main), col(DATA,c));
    if(r!==null && Math.abs(r)>Math.abs(bv)){ bv=r; best=c; }
  });
  if(!best) return;
  addCard({
    id:'scatter', title:SPEC.main+' × '+best, height:'318px',
    note:'看两个数值字段是否同步变化。全量相关系数 r = '+bv.toFixed(3)+
         '（'+ (Math.abs(bv)<0.3?'基本无关':(Math.abs(bv)<0.6?'弱相关':'较强相关')) +'）。',
    render:function(rows){
      var xs=col(rows,best), ys=col(rows,SPEC.main);
      var groups = mainCat? unique(col(rows,mainCat)) : [null];
      if(!mainCat){
        return { data:[{ type:'scatter', mode:'markers', x:xs, y:ys,
          marker:{color:'#2563eb', size:6, opacity:0.55},
          hovertemplate:best+' %{x}<br>'+SPEC.main+' %{y}<extra></extra>' }],
          layout:lay({ height:318, xaxis:{title:{text:best}, gridcolor:'#f1f5f9'},
            yaxis:{title:{text:SPEC.main}, gridcolor:'#f1f5f9'} }) };
      }
      return { data:groups.slice(0,8).map(function(k,i){
        var sub=rows.filter(function(d){ return String(d[mainCat])===k; });
        return { type:'scatter', mode:'markers', name:String(k),
          x:col(sub,best), y:col(sub,SPEC.main),
          marker:{size:6, opacity:0.6, color:PALETTE[i%PALETTE.length]},
          hovertemplate:'%{fullData.name}<br>'+best+' %{x}<br>'+SPEC.main+' %{y}<extra></extra>' };
      }), layout:lay({ height:318,
        xaxis:{title:{text:best}, gridcolor:'#f1f5f9'},
        yaxis:{title:{text:SPEC.main}, gridcolor:'#f1f5f9'},
        showlegend:true, legend:{font:{size:10}, orientation:'h', y:-0.22},
        margin:{l:70,r:26,t:12,b:74} }) };
    }
  });
})();

(function(){
  if(!SPEC.main) return;
  var others=SPEC.numeric.filter(function(c){ return c!==SPEC.main; });
  if(others.length<1) return;
  addCard({
    id:'corrbar', title:'各数值字段与'+SPEC.main+'的关系', height:'318px',
    note:'相关系数取值 -1 ~ +1。红向右表示正相关（同步上升），绿向左表示负相关，接近 0 表示基本无关。随筛选实时变化。',
    render:function(rows){
      var arr=[];
      others.forEach(function(c){ var r=corr(col(rows,c), col(rows,SPEC.main)); if(r!==null) arr.push([c,r]); });
      if(!arr.length) return null;
      arr.sort(function(a,b){ return a[1]-b[1]; });
      return { data:[{ type:'bar', orientation:'h',
        y:arr.map(function(a){return a[0];}), x:arr.map(function(a){return a[1];}),
        marker:{color:arr.map(function(a){ return a[1]>=0? '#e24b4a' : '#1d9e75'; })},
        text:arr.map(function(a){return a[1].toFixed(3);}), textposition:'outside', cliponaxis:false,
        hovertemplate:'%{y}<br>r = %{x:.3f}<extra></extra>' }],
        layout:lay({ height:Math.max(280,arr.length*30+70),
          xaxis:{title:{text:'相关系数 r'}, range:[-1,1], zeroline:true, zerolinecolor:'#9ca3af', gridcolor:'#f1f5f9'},
          margin:{l:130,r:50,t:12,b:44} }) };
    }
  });
})();

(function(){
  var use=SPEC.numeric.slice(0,MAXHEAT);
  if(use.length<2) return;
  addCard({
    id:'heat', title:'数值字段相关性矩阵', height:'500px', full:true,
    note:'越接近红=正相关越强，越接近绿=负相关越强。对角线是自己跟自己，恒为 1。',
    render:function(rows){
      var z=[];
      for(var i=0;i<use.length;i++){
        var row=[];
        for(var j=0;j<use.length;j++){ var r=corr(col(rows,use[i]),col(rows,use[j])); row.push(r===null?0:r); }
        z.push(row);
      }
      var n=use.length;
      return { data:[{ type:'heatmap', z:z, x:use, y:use, zmin:-1, zmax:1,
        colorscale:[[0,'#1d9e75'],[0.5,'#ffffff'],[1,'#e24b4a']],
        hovertemplate:'%{y} × %{x}<br>r = %{z:.3f}<extra></extra>' }],
        layout:lay({ height:Math.max(420, n*44+120),
          margin:{l:150,r:40,t:20,b:150},
          xaxis:{tickangle:-40, automargin:true}, yaxis:{automargin:true, autorange:'reversed'} }) };
    }
  });
})();

if(MISS.length){
  addCard({
    id:'miss', title:'原始缺失值分布', height:'318px',
    note:'生成看板前检测到 '+CLEAN.filled_total+' 个缺失值，已用该列中位数填补（分类型字段标为「未填写」）。此图不随筛选变化。',
    render:function(rows){
      MISS.sort(function(a,b){ return b.n-a.n; });
      var d=MISS.slice(0,15);
      return { data:[{ type:'bar', orientation:'h',
        y:d.map(function(x){return x.col;}), x:d.map(function(x){return x.n;}),
        text:d.map(function(x){return (x.n/SPEC.rows*100).toFixed(1)+'%';}),
        textposition:'outside', cliponaxis:false, marker:{color:'#f59e0b'},
        hovertemplate:'%{y}<br>缺失 %{x} 个<extra></extra>' }],
        layout:lay({ height:Math.max(280,d.length*30+70),
          xaxis:{title:{text:'缺失个数'}, gridcolor:'#f1f5f9'},
          yaxis:{autorange:'reversed'}, margin:{l:130,r:60,t:12,b:44} }) };
    }
  });
}

if(SPEC.main){
  addCard({
    id:'top', title:SPEC.main+' 排行 TOP'+MAX_TOP, height:'420px', full:true, type:'table',
    note:'点表头可以换排序字段。',
    render:function(rows, el){
      var key=SPEC.main, asc=false;
      function make(){
        var arr=rows.slice().sort(function(a,b){
          var x=a[key], y=b[key];
          if(typeof x==='number'&&typeof y==='number') return asc? x-y : y-x;
          return asc? String(x).localeCompare(String(y)) : String(y).localeCompare(String(x));
        }).slice(0,MAX_TOP);
        var cols=SPEC.filters.map(function(f){return f.col;}).concat(SPEC.numeric.slice(0,6));
        cols=unique(cols);
        var h='<table><thead><tr>';
        h+='<th>名次</th>';
        cols.forEach(function(c){ h+='<th data-c="'+c+'">'+c+'</th>'; });
        h+='</tr></thead><tbody>';
        arr.forEach(function(r,i){
          h+='<tr><td>'+(i+1)+'</td>';
          cols.forEach(function(c){
            var v=r[c];
            h+='<td>'+(typeof v==='number'? fmt(v) : (v===null?'':String(v)))+'</td>';
          });
          h+='</tr>';
        });
        h+='</tbody></table>';
        el.innerHTML=h;
        Array.prototype.forEach.call(el.querySelectorAll('th[data-c]'), function(th){
          th.onclick=function(){ key=th.getAttribute('data-c'); asc=!asc; make(); };
        });
      }
      make();
    }
  });
}

document.getElementById('reset').onclick=function(){
  SPEC.filters.forEach(function(f){ sel[f.col]=new Set(f.values); });
  Array.prototype.forEach.call(document.querySelectorAll('.chip'),function(c){ c.className='chip on'; });
  redraw();
};
document.getElementById('exportAll').onclick=function(){
  charts.forEach(function(o,i){
    if(o.type==='table') return;
    setTimeout(function(){
      Plotly.downloadImage(document.getElementById(o.id),{format:'png',scale:3,filename:FN[o.id]});
    }, i*450);
  });
};

redraw();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------- 主流程
def build_dashboard(src, out=None, sheet=None, title=None, main=None, filters=None,
                    csv_encoding=None, no_open=False):
    path = os.path.abspath(src)
    if not os.path.exists(path):
        raise RuntimeError("找不到文件：%s" % path)

    df_raw = read_source(path, sheet, csv_encoding)
    if df_raw.shape[0] == 0:
        raise RuntimeError("这个表是空的（0 行）")
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    exclude = set(filters or [])
    numeric, cats, other, idcols = classify_columns(df_raw, set())
    if not numeric:
        # 极端情况：所有数值列都被当成编号了，那就把它们都用回来
        numeric = [str(c) for c in df_raw.columns
                   if pd.api.types.is_numeric_dtype(df_raw[c])]
        idcols = []
    if not numeric:
        raise RuntimeError("表里没识别出任何数值字段，无法生成看板。请检查表格内容。")

    # 主指标：用户指定优先，否则挑一个像金额/销量的字段，再不行用第一个非编号数值字段
    if main:
        if str(main) not in numeric:
            raise RuntimeError("指定的主指标「%s」不是可用的数值字段。\n表里可用的数值字段有：%s"
                               % (main, "、".join(numeric)))
        main_metric = str(main)
    else:
        main_metric = pick_main_metric(numeric, len(df_raw), None)

    chosen_filters = pick_filters(df_raw, cats, filters)
    df, clean = clean_data(df_raw, numeric, cats)

    spec = build_spec(df, numeric, cats, chosen_filters, main_metric,
                      title or os.path.splitext(os.path.basename(path))[0],
                      os.path.basename(path))
    records = to_records(df, numeric, cats)
    miss = compute_missing(df_raw)

    js_path = find_plotly_js()
    if js_path:
        with io.open(js_path, "r", encoding="utf-8") as f:
            plotly_js = f.read()
    else:
        plotly_js = ""
        sys.stderr.write("警告：没找到本地 plotly.min.js，将改用在线 CDN（离线打不开图）。\n")

    html = HTML_TEMPLATE
    if plotly_js:
        html = html.replace("<script>__PLOTLYJS__</script>", "<script>" + plotly_js + "</script>")
    else:
        html = html.replace("<script>__PLOTLYJS__</script>",
                            '<script src="https://cdn.jsdelivr.net/npm/'
                            'plotly.js-dist-min@2.27.0/plotly.min.js"></script>')
    html = html.replace("__TITLE__", spec["title"])
    html = html.replace("__SOURCE__", "数据来源：" + spec["source"])
    html = html.replace("__ROWS__", "{:,}".format(spec["rows"]))
    html = html.replace("__COLS__", str(spec["cols"]))
    html = html.replace("__GEN__", spec["generated"])
    html = html.replace("__FOOTNOTE__",
                        "数值字段缺失值用列中位数填补，分类型字段缺失值标为「未填写」；未删除任何原始记录。"
                        if clean["filled_total"] or clean["duplicates"] else
                        "本表未检测到缺失值或完全重复行。")
    html = html.replace("__SPEC__", json.dumps(spec, ensure_ascii=False))
    html = html.replace("__DATA__", json.dumps(records, ensure_ascii=False))
    html = html.replace("__MISS__", json.dumps(miss, ensure_ascii=False))
    html = html.replace("__CLEAN__", json.dumps(clean, ensure_ascii=False))
    html = html.replace("MAXHEAT", str(MAX_HEATMAP_COLS))
    html = html.replace("MAX_TOP", str(MAX_TOP_ROWS))

    out_path = out or (os.path.splitext(path)[0] + "_看板.html")
    with io.open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    return {
        "out": out_path,
        "rows": spec["rows"],
        "cols": spec["cols"],
        "numeric": numeric,
        "cats": cats,
        "filters": chosen_filters,
        "main": main_metric,
        "filled": clean["filled_total"],
        "duplicates": clean["duplicates"],
        "idcols": idcols,
        "charts": None,
        "size_mb": round(os.path.getsize(out_path) / 1024 / 1024, 2),
    }


def main():
    ap = argparse.ArgumentParser(
        description="把任意 Excel / CSV 生成一份交互式数据看板（HTML）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="例子：\n"
               "  python build_dashboard.py 数据.xlsx\n"
               "  python build_dashboard.py 问卷.csv --main 满意度\n"
               "  python build_dashboard.py 数据.xlsx --sheet 2024 --title 年度经营看板\n")
    ap.add_argument("source", help="数据文件路径（.xlsx / .xls / .csv）")
    ap.add_argument("-o", "--out", help="输出的 HTML 文件名（默认：数据文件同名的 _看板.html）")
    ap.add_argument("--sheet", help="Excel 里要读的工作表名（默认读第一个）")
    ap.add_argument("--title", help="看板标题（默认用文件名）")
    ap.add_argument("--main", help="主指标列名，比如 年营业额（默认自动猜一个像金额/销量的字段）")
    ap.add_argument("--filters", nargs="*", help="指定用作筛选器的列名，最多 4 个")
    ap.add_argument("--encoding", help="CSV 文件的编码（例如 gbk）")
    args = ap.parse_args()

    info = build_dashboard(args.source, args.out, args.sheet, args.title,
                           args.main, args.filters, args.encoding)

    sys.stdout.reconfigure(encoding="utf-8")
    print("")
    print("看板生成完成")
    print("  输出文件 : %s （%.2f MB）" % (info["out"], info["size_mb"]))
    print("  数据量   : %s 行 × %s 列" % (info["rows"], info["cols"]))
    print("  主指标   : %s" % info["main"])
    print("  数值字段 : %d 个 —— %s" % (len(info["numeric"]), "、".join(info["numeric"][:8])))
    if info["idcols"]:
        print("  已忽略编号: %s（编号类字段不参与均值与相关性计算）" % "、".join(info["idcols"]))
    if info["cats"]:
        print("  分类字段 : %d 个 —— %s" % (len(info["cats"]), "、".join(info["cats"][:8])))
    print("  筛选器   : %s" % ("、".join(info["filters"]) if info["filters"] else "无（没找到合适的分类字段）"))
    print("  缺失填补 : %d 个    完全重复行：%d 条" % (info["filled"], info["duplicates"]))
    print("")
    print("双击那个 html 文件，用浏览器打开就能看到。")


if __name__ == "__main__":
    main()
