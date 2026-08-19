#!/usr/bin/env python3
"""
E-Catalogue Builder
Reads the Excel price list + data folder structure
and generates a fully static index.html site.
"""

import os
import re
import json
import openpyxl

# ─── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_DIR   = os.path.join(BASE_DIR, "data")
EXCEL_FILE = os.path.join(BASE_DIR, "Price List 2026-2027.xlsx")
OUT_FILE   = os.path.join(BASE_DIR, "index.html")
SHEET_NAME = "Full Price List"

# ─── 1.  Read Excel ──────────────────────────────────────────────────────────
print("Reading Excel …")
wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True)
ws = wb[SHEET_NAME]
rows = list(ws.iter_rows(values_only=True))
headers = [str(h).strip() if h else "" for h in rows[0]]

# Column indices (0-based)
COL_ART     = headers.index("Art no.")
COL_ITEM    = headers.index("Item")
COL_QUALITY = headers.index("Quality")
COL_MRP     = headers.index("Mrp")

# Build lookup: numeric_art_no -> {item, quality, mrp}
price_data = {}
for row in rows[1:]:
    art = row[COL_ART]
    if art is None:
        continue
    try:
        art_int = int(art)
    except (ValueError, TypeError):
        continue
    mrp     = row[COL_MRP]
    item    = row[COL_ITEM] or ""
    quality = row[COL_QUALITY] or ""
    item    = str(item).strip().title()
    quality = str(quality).strip().title()
    price_data[art_int] = {
        "art": art_int,
        "item": item,
        "quality": quality,
        "mrp": int(mrp) if mrp else 0,
    }

print(f"  Loaded {len(price_data)} price-list entries.")

# ─── 2.  Scan image folders ───────────────────────────────────────────────────
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".JPG", ".JPEG", ".PNG"}

ART_RE = re.compile(r'Art\s*no[.\s]*(\d+)', re.IGNORECASE)

categories = {}

for folder_name in sorted(os.listdir(DATA_DIR)):
    folder_path = os.path.join(DATA_DIR, folder_name)
    if not os.path.isdir(folder_path):
        continue

    products = []
    for fname in sorted(os.listdir(folder_path)):
        ext = os.path.splitext(fname)[1]
        if ext not in IMAGE_EXTS:
            continue

        m = ART_RE.search(fname)
        if not m:
            print(f"  [SKIP] Cannot parse art no from: {fname}")
            continue

        art_int = int(m.group(1))
        info    = price_data.get(art_int, {})

        # Extract variant suffix (e.g. "A", "B", "Mix Designs", "Jari", etc.)
        raw_after = fname[m.end():]
        raw_after = re.sub(r'^\s*', '', raw_after)
        raw_after = re.sub(r'\.[^.]+$', '', raw_after)
        variant = raw_after.strip()

        rel_path = f"data/{folder_name}/{fname}".replace("\\", "/")
        rel_path_url = rel_path.replace(" ", "%20")

        products.append({
            "artNo":    art_int,
            "variant":  variant,
            "filename": fname,
            "imgPath":  rel_path_url,
            "item":     info.get("item", "Shawl"),
            "quality":  info.get("quality", ""),
            "mrp":      info.get("mrp", 0),
            "category": folder_name,
        })

    if products:
        categories[folder_name] = products
        print(f"  {folder_name}: {len(products)} images")

# ─── 3.  Flatten all products for "All" view ─────────────────────────────────
all_products = []
for prods in categories.values():
    all_products.extend(prods)
all_products.sort(key=lambda p: (p["artNo"], p["variant"]))

# ─── 4.  Determine global MRP range ──────────────────────────────────────────
mrp_vals = [p["mrp"] for p in all_products if p["mrp"] > 0]
mrp_min  = min(mrp_vals) if mrp_vals else 0
mrp_max  = max(mrp_vals) if mrp_vals else 99999

# ─── 5.  Embed data as JSON ──────────────────────────────────────────────────
catalogue_json = json.dumps({
    "categories": {k: v for k, v in categories.items()},
    "all": all_products,
    "mrpMin": mrp_min,
    "mrpMax": mrp_max,
}, ensure_ascii=False, separators=(",", ":"))

print(f"\nTotal products: {len(all_products)}")
print(f"MRP range: Rs.{mrp_min} -- Rs.{mrp_max}")

# ─── 6.  Build HTML ──────────────────────────────────────────────────────────
html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Shawl E-Catalogue 2026-2027</title>
<meta name="description" content="Browse our exclusive collection of premium shawls, stoles and lohi — Polywool, Fine Wool, Kashmiri and more."/>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Playfair+Display:wght@700&display=swap" rel="stylesheet"/>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0f0f14;
  --surface:#16161e;
  --surface2:#1e1e2a;
  --border:#2a2a3a;
  --accent:#c9a96e;
  --accent2:#e8c987;
  --text:#f0ede8;
  --text-muted:#9090a8;
  --sidebar-w:270px;
  --radius:14px;
}
html{scroll-behavior:smooth}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;display:flex;flex-direction:column;}

/* Header */
header{
  position:sticky;top:0;z-index:100;
  background:rgba(15,15,20,.88);
  backdrop-filter:blur(16px);
  border-bottom:1px solid var(--border);
  padding:0 20px;height:64px;
  display:flex;align-items:center;gap:14px;
}
.logo{
  font-family:'Playfair Display',serif;font-size:1.25rem;
  background:linear-gradient(135deg,var(--accent),var(--accent2));
  -webkit-background-clip:text;-webkit-text-fill-color:transparent;
  white-space:nowrap;flex-shrink:0;
}
.header-search{
  flex:1;max-width:360px;
  display:flex;align-items:center;gap:8px;
  background:var(--surface2);border:1px solid var(--border);
  border-radius:999px;padding:8px 16px;
}
.header-search input{background:none;border:none;outline:none;color:var(--text);font-size:.88rem;width:100%;}
.header-search input::placeholder{color:var(--text-muted)}
.header-search svg{flex-shrink:0;color:var(--text-muted)}
.count-badge{font-size:.82rem;color:var(--text-muted);white-space:nowrap;margin-left:auto;}

/* Layout */
.layout{display:flex;flex:1;min-height:0;}

/* Sidebar */
aside{
  width:var(--sidebar-w);flex-shrink:0;
  background:var(--surface);border-right:1px solid var(--border);
  overflow-y:auto;padding:16px 10px;
  display:flex;flex-direction:column;gap:4px;
}
aside::-webkit-scrollbar{width:4px}
aside::-webkit-scrollbar-thumb{background:var(--border);border-radius:99px}

.sidebar-label{
  font-size:.7rem;font-weight:600;letter-spacing:.1em;
  text-transform:uppercase;color:var(--text-muted);
  padding:6px 12px 8px;
}
.cat-btn{
  display:flex;align-items:center;justify-content:space-between;
  width:100%;background:none;border:none;cursor:pointer;
  color:var(--text-muted);font-size:.88rem;font-weight:500;
  padding:9px 13px;border-radius:10px;
  transition:background .18s,color .18s;text-align:left;
}
.cat-btn:hover{background:var(--surface2);color:var(--text)}
.cat-btn.active{
  background:linear-gradient(135deg,rgba(201,169,110,.15),rgba(232,201,135,.06));
  color:var(--accent2);border:1px solid rgba(201,169,110,.22);
}
.cat-count{font-size:.7rem;background:var(--surface2);padding:2px 7px;border-radius:999px;color:var(--text-muted);}

/* Price filter */
.pf-wrap{margin-top:auto;border-top:1px solid var(--border);padding:14px 10px 10px;}
.pf-label{font-size:.7rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:var(--text-muted);margin-bottom:10px;}
.pf-inputs{display:flex;gap:7px;margin-bottom:12px;}
.pf-input-group{flex:1;}
.pf-input-group label{display:block;font-size:.7rem;color:var(--text-muted);margin-bottom:3px;}
.pf-input-group input{
  width:100%;background:var(--surface2);border:1px solid var(--border);
  border-radius:8px;padding:6px 9px;color:var(--text);font-size:.82rem;outline:none;
  transition:border-color .18s;
}
.pf-input-group input:focus{border-color:var(--accent)}

/* Dual range slider */
.rw{position:relative;height:26px;margin-bottom:5px;}
.rt{position:absolute;top:50%;transform:translateY(-50%);left:0;right:0;height:4px;background:var(--border);border-radius:99px;}
.rf{position:absolute;top:50%;transform:translateY(-50%);height:4px;background:linear-gradient(90deg,var(--accent),var(--accent2));border-radius:99px;pointer-events:none;}
.rs{position:absolute;top:50%;transform:translateY(-50%);width:100%;height:4px;-webkit-appearance:none;appearance:none;background:transparent;outline:none;cursor:pointer;}
.rs::-webkit-slider-thumb{-webkit-appearance:none;width:17px;height:17px;border-radius:50%;background:var(--accent2);border:3px solid var(--bg);box-shadow:0 0 0 2px var(--accent);cursor:pointer;transition:transform .15s;}
.rs::-webkit-slider-thumb:hover{transform:scale(1.18)}
#smin{z-index:3}#smax{z-index:4}
.pf-vals{display:flex;justify-content:space-between;font-size:.76rem;color:var(--accent2);font-weight:500;}

.reset-btn{
  width:100%;margin-top:10px;
  background:linear-gradient(135deg,var(--accent),var(--accent2));
  border:none;border-radius:9px;padding:9px;
  color:#111;font-weight:600;font-size:.85rem;cursor:pointer;
  transition:opacity .18s,transform .12s;
}
.reset-btn:hover{opacity:.86;transform:translateY(-1px)}

/* Main */
main{flex:1;overflow-y:auto;padding:22px 20px;}
main::-webkit-scrollbar{width:6px}
main::-webkit-scrollbar-thumb{background:var(--border);border-radius:99px}

.ch{display:flex;align-items:baseline;gap:10px;margin-bottom:18px;flex-wrap:wrap;}
.ch h1{font-family:'Playfair Display',serif;font-size:1.5rem;background:linear-gradient(135deg,var(--accent),var(--accent2));-webkit-background-clip:text;-webkit-text-fill-color:transparent;}
.ch .rc{font-size:.83rem;color:var(--text-muted);}

/* Grid */
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:18px;}

/* Card */
.card{
  background:var(--surface);border:1px solid var(--border);
  border-radius:var(--radius);overflow:hidden;cursor:pointer;
  transition:transform .22s,box-shadow .22s,border-color .22s;
}
.card:hover{transform:translateY(-5px);box-shadow:0 14px 38px rgba(0,0,0,.55);border-color:rgba(201,169,110,.38);}
.ciw{aspect-ratio:3/4;overflow:hidden;background:var(--surface2);}
.ciw img{width:100%;height:100%;object-fit:cover;transition:transform .38s ease;display:block;}
.card:hover .ciw img{transform:scale(1.05)}
.ciw .placeholder{width:100%;height:100%;display:flex;align-items:center;justify-content:center;color:var(--text-muted);font-size:.8rem;}
.cb{padding:11px 13px 13px;}
.ca{font-size:.72rem;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:var(--text-muted);margin-bottom:2px;}
.ct{font-size:.9rem;font-weight:600;color:var(--text);line-height:1.3;margin-bottom:4px;}
.cq{font-size:.76rem;color:var(--text-muted);margin-bottom:7px;}
.cp{font-size:1.02rem;font-weight:700;background:linear-gradient(135deg,var(--accent),var(--accent2));-webkit-background-clip:text;-webkit-text-fill-color:transparent;}
.cna{font-size:.8rem;color:var(--text-muted);font-style:italic;}

/* Empty */
.empty{grid-column:1/-1;text-align:center;padding:70px 0;color:var(--text-muted);}
.empty svg{margin:0 auto 14px;display:block;opacity:.3}
.empty p{font-size:.95rem;}

/* Lightbox */
#lb{position:fixed;inset:0;z-index:200;background:rgba(0,0,0,.88);display:none;align-items:center;justify-content:center;backdrop-filter:blur(6px);}
#lb.open{display:flex}
.lbi{position:relative;max-width:92vw;max-height:92vh;display:flex;flex-direction:column;align-items:center;gap:10px;}
#lb-img{max-width:90vw;max-height:76vh;border-radius:10px;box-shadow:0 20px 56px rgba(0,0,0,.8);object-fit:contain;}
#lb-cap{color:var(--text);font-size:.92rem;text-align:center;background:rgba(22,22,30,.82);padding:9px 22px;border-radius:999px;backdrop-filter:blur(8px);}
#lb-x{position:absolute;top:-12px;right:-12px;background:var(--surface2);border:1px solid var(--border);color:var(--text);width:34px;height:34px;border-radius:50%;cursor:pointer;font-size:1rem;display:flex;align-items:center;justify-content:center;transition:background .18s;}
#lb-x:hover{background:var(--border)}

/* Mobile */
@media(max-width:750px){
  aside{position:fixed;left:0;top:64px;bottom:0;z-index:50;transform:translateX(-100%);transition:transform .28s;width:285px;}
  aside.open{transform:translateX(0)}
  .mbt{display:flex!important}
}
.mbt{display:none;background:none;border:none;cursor:pointer;color:var(--text);padding:5px;border-radius:8px;}
</style>
</head>
<body>

<header>
  <button class="mbt" id="mbt" aria-label="Open categories">
    <svg width="21" height="21" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>
  </button>
  <div class="logo">&#10022; ShawlCraft E-Catalogue 2026&ndash;27</div>
  <div class="header-search">
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/></svg>
    <input type="search" id="q" placeholder="Search art no, quality&#8230;" autocomplete="off"/>
  </div>
  <span class="count-badge" id="cb2">0 items</span>
</header>

<div class="layout">
  <aside id="sb">
    <div class="sidebar-label">Categories</div>
    <div id="cats"></div>
    <div class="pf-wrap">
      <div class="pf-label">Price Filter (MRP)</div>
      <div class="pf-inputs">
        <div class="pf-input-group"><label for="pmin">Min &#8377;</label><input type="number" id="pmin"/></div>
        <div class="pf-input-group"><label for="pmax">Max &#8377;</label><input type="number" id="pmax"/></div>
      </div>
      <div class="rw">
        <div class="rt"></div>
        <div class="rf" id="rf"></div>
        <input type="range" class="rs" id="smin" step="50"/>
        <input type="range" class="rs" id="smax" step="50"/>
      </div>
      <div class="pf-vals"><span id="vmin"></span><span id="vmax"></span></div>
      <button class="reset-btn" id="rbtn">Reset Filter</button>
    </div>
  </aside>

  <main>
    <div class="ch">
      <h1 id="cht">All Products</h1>
      <span class="rc" id="rc"></span>
    </div>
    <div class="grid" id="grid"></div>
  </main>
</div>

<div id="lb" role="dialog" aria-modal="true">
  <div class="lbi">
    <button id="lb-x" aria-label="Close">&times;</button>
    <img id="lb-img" src="" alt=""/>
    <div id="lb-cap"></div>
  </div>
</div>

<script>
const D=CATALOGUE_DATA;
const GMIN=D.mrpMin,GMAX=D.mrpMax;
let activeCat="All",fmin=GMIN,fmax=GMAX,sq="";

// Categories
const catNames=["All",...Object.keys(D.categories).sort()];
const catsEl=document.getElementById("cats");
function getCnt(n){return n==="All"?D.all.length:(D.categories[n]||[]).length;}
function buildCats(){
  catsEl.innerHTML="";
  catNames.forEach(n=>{
    const b=document.createElement("button");
    b.className="cat-btn"+(n===activeCat?" active":"");
    b.dataset.c=n;
    b.innerHTML=`<span>${n}</span><span class="cat-count">${getCnt(n)}</span>`;
    b.onclick=()=>{activeCat=n;document.getElementById("sb").classList.remove("open");render();};
    catsEl.appendChild(b);
  });
}

// Sliders
const smin=document.getElementById("smin"),smax=document.getElementById("smax");
const rf=document.getElementById("rf"),vmin=document.getElementById("vmin"),vmax=document.getElementById("vmax");
const pmin=document.getElementById("pmin"),pmax=document.getElementById("pmax");
function ini(){smin.min=smax.min=pmin.min=pmax.min=GMIN;smin.max=smax.max=pmin.max=pmax.max=GMAX;smin.value=GMIN;smax.value=GMAX;pmin.value=GMIN;pmax.value=GMAX;}
ini();
function fmt(n){return "\\u20b9"+n.toLocaleString("en-IN");}
function fill(){
  const mn=+smin.value,mx=+smax.value,r=GMAX-GMIN;
  rf.style.left=((mn-GMIN)/r*100)+"%";
  rf.style.right=((GMAX-mx)/r*100)+"%";
  vmin.textContent=fmt(mn);vmax.textContent=fmt(mx);
  pmin.value=mn;pmax.value=mx;
}
fill();
smin.oninput=()=>{if(+smin.value>+smax.value)smin.value=smax.value;fmin=+smin.value;fmax=+smax.value;fill();render();};
smax.oninput=()=>{if(+smax.value<+smin.value)smax.value=smin.value;fmin=+smin.value;fmax=+smax.value;fill();render();};
pmin.onchange=()=>{let v=Math.max(GMIN,Math.min(+pmin.value||GMIN,fmax));smin.value=v;fmin=v;fill();render();};
pmax.onchange=()=>{let v=Math.min(GMAX,Math.max(+pmax.value||GMAX,fmin));smax.value=v;fmax=v;fill();render();};
document.getElementById("rbtn").onclick=()=>{fmin=GMIN;fmax=GMAX;smin.value=GMIN;smax.value=GMAX;fill();render();};

// Search
document.getElementById("q").oninput=e=>{sq=e.target.value.trim().toLowerCase();render();};

// Render
const grid=document.getElementById("grid"),cht=document.getElementById("cht"),rc=document.getElementById("rc"),cb2=document.getElementById("cb2");
function mkLabel(p){const s=p.variant?" "+p.variant:"";return `Art no. ${p.artNo}${s} (${p.item||"Shawl"})`;}
function render(){
  document.querySelectorAll(".cat-btn").forEach(b=>b.classList.toggle("active",b.dataset.c===activeCat));
  cht.textContent=activeCat==="All"?"All Products":activeCat;
  let prods=activeCat==="All"?D.all:(D.categories[activeCat]||[]);
  prods=prods.filter(p=>p.mrp===0||(p.mrp>=fmin&&p.mrp<=fmax));
  if(sq)prods=prods.filter(p=>{
    const h=["art no "+p.artNo,p.artNo+"",p.variant||"",p.item||"",p.quality||"",p.filename||""].join(" ").toLowerCase();
    return h.includes(sq);
  });
  rc.textContent=prods.length+" items";cb2.textContent=prods.length+" items";
  if(!prods.length){
    grid.innerHTML=`<div class="empty"><svg width="60" height="60" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1"><path d="M21 21l-6-6m2-5a7 7 0 1 1-14 0 7 7 0 0 1 14 0z"/></svg><p>No products match your filter.</p></div>`;
    return;
  }
  grid.innerHTML=prods.map((p,i)=>{
    const s=p.variant?" "+p.variant:"";
    const price=p.mrp>0?`<div class="cp">&#8377; ${p.mrp.toLocaleString("en-IN")}</div>`:`<div class="cna">Price on request</div>`;
    return `<div class="card" data-i="${i}" tabindex="0" role="button">
      <div class="ciw"><img src="${p.imgPath}" alt="${mkLabel(p)}" loading="lazy" onerror="this.outerHTML='<div class=placeholder>Image not found</div>'"/></div>
      <div class="cb">
        <div class="ca">Art no. ${p.artNo}${s}</div>
        <div class="ct">${p.item||"Shawl"}${s?" &mdash; "+s:""}</div>
        <div class="cq">${p.quality||""}</div>
        ${price}
      </div>
    </div>`;
  }).join("");
  grid.querySelectorAll(".card").forEach((c,i)=>{
    const p=prods[i];
    c.onclick=()=>openLb(p);
    c.onkeydown=e=>{if(e.key==="Enter"||e.key===" ")openLb(p);};
  });
}

// Lightbox
const lb=document.getElementById("lb"),lbImg=document.getElementById("lb-img"),lbCap=document.getElementById("lb-cap");
function openLb(p){
  const s=p.variant?" "+p.variant:"";
  lbImg.src=p.imgPath;lbImg.alt=mkLabel(p);
  const pr=p.mrp>0?"  \\u20b9"+p.mrp.toLocaleString("en-IN"):"Price on request";
  lbCap.textContent=`Art no. ${p.artNo}${s} (${p.item||"Shawl"}) \u2014 ${p.quality||""}  \u2022  ${pr}`;
  lb.classList.add("open");document.body.style.overflow="hidden";
}
function closeLb(){lb.classList.remove("open");lbImg.src="";document.body.style.overflow="";}
document.getElementById("lb-x").onclick=closeLb;
lb.onclick=e=>{if(e.target===lb)closeLb();};
document.onkeydown=e=>{if(e.key==="Escape")closeLb();};

// Mobile
document.getElementById("mbt").onclick=()=>document.getElementById("sb").classList.toggle("open");

buildCats();render();
</script>
</body>
</html>"""

html = html.replace("CATALOGUE_DATA", catalogue_json)

# ─── 7.  Write output ─────────────────────────────────────────────────────────
with open(OUT_FILE, "w", encoding="utf-8") as f:
    f.write(html)

print(f"\n[OK] Generated: {OUT_FILE}")
print("     Open index.html in a browser or push the whole folder to GitHub Pages.")
