#!/usr/bin/env python3
"""Build the client review artifact from the REAL formatter output, with an
interactive front-end: the real DOCX render (image) + precise per-section
overlay regions (coordinates from the PDF) + a left legend. Hovering/focusing a
legend item or a letter region highlights both and draws a connector line that
tracks on scroll.

Pipeline: fictional content -> LetterDocument -> render_letter (DOCX) ->
LibreOffice (PDF) -> pdftoppm (PNG pages) + pdftotext -bbox (word coords) ->
HTML with overlay regions + SVG connectors.

Run:  .venv/bin/python tool/scripts/build_review_html.py
"""
from __future__ import annotations
import os, sys, html, glob, re, shutil, subprocess, tempfile
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
from docx import Document
from docx.shared import Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from formatters.letters import (LetterDocument, Span, render_letter,
                                paragraph, section_heading, subheading, list_item)

OUT = os.path.join(ROOT, "out"); OUTH = os.path.join(OUT, "html"); os.makedirs(OUTH, exist_ok=True)
TEMPLATE = os.path.join(ROOT, "assets", "Template_Vuoto.docx")
DOCX = os.path.join(OUT, "lettera_fittizia.docx")
PDF = os.path.join(OUT, "lettera_fittizia.pdf")
HTMLF = os.path.join(OUTH, "lettera_fittizia_review.html")
PNG_PREFIX = os.path.join(OUTH, "lettera_fittizia_review")

# ---------- 1) fictional content ----------
DELIVERY = "A mezzo PEC, anticipata a mezzo posta elettronica"
DATE = "Bergamo, 24 giugno 2026"
RECIPIENT = ["Egr.", Span("Sig. Mario Rossi", bold=True), "Via dei Tigli, n. 14",
             "24121 Bergamo (BG)", Span("PEC: mario.rossi@esempio-pec.it", bold=True)]
SUBJECT = "OGGETTO: Diffida ad adempiere - contratto di fornitura del 10 gennaio 2026"
OPENING = "Egregio Signore,"
CLOSING = "In attesa di un cortese e sollecito riscontro, si porgono distinti saluti."
SIGNATURE = ["Bergamo Legal Società tra Avvocati S.r.l.",
             Span("Avv. Giulia Conti", bold=True), "_______________________________"]
ATTACH = ["All. 1 - copia del contratto di fornitura del 10 gennaio 2026;",
          "All. 2 - copia della fattura n. 42."]
POST = "(documento fittizio, predisposto a solo scopo di revisione formale)"
BODY = [
    ("body",    paragraph("lo Studio scrivente agisce in nome e per conto di Aurora Servizi S.r.l., con sede in Bergamo, in forza di incarico ricevuto, in relazione al rapporto di fornitura di seguito descritto.")),
    ("subhead", subheading("Fatto")),
    ("body",    paragraph("1. In data 10 gennaio 2026 le parti sottoscrivevano un contratto di fornitura di materiale per ufficio, con pagamento a 30 giorni dalla consegna.")),
    ("body",    paragraph("2. La fornitura veniva regolarmente eseguita e fatturata (fattura n. 42), ma il relativo corrispettivo è rimasto integralmente impagato alla scadenza convenuta.")),
    ("ritual",  section_heading("DIFFIDA")),
    ("body",    paragraph("la S.V., come sopra individuata, a voler:")),
    ("list",    list_item("(i) provvedere al pagamento della somma di euro 3.500,00 entro e non oltre 15 giorni dalla ricezione della presente;")),
    ("list",    list_item("(ii) trasmettere allo Studio scrivente la contabile dell'avvenuto pagamento.")),
    ("body",    paragraph("In difetto di adempimento nel termine indicato, si procederà nelle competenti sedi giudiziarie per il recupero del credito, oltre interessi e spese.")),
]
letter = LetterDocument(delivery_method=DELIVERY, date_place=DATE, recipient_block=RECIPIENT,
    subject=SUBJECT, opening=OPENING, body_blocks=[b for _, b in BODY],
    closing=CLOSING, signature_block=SIGNATURE, attachments=ATTACH, postscript=POST)
res = render_letter(letter, TEMPLATE, DOCX)
print(f"[1] DOCX {DOCX} ({res.paragraph_count} par)")

# ---------- 2) classify rendered paragraphs ----------
def norm(s): return re.sub(r"\s+", " ", (s or "").replace("–","-").replace("—","-")).strip().lower()
def blk_text(b):
    c=b.content
    if isinstance(c,str): return c
    if isinstance(c,Span): return c.text
    if isinstance(c,list): return "".join(x.text if isinstance(x,Span) else str(x) for x in c)
    return str(c)
doc = Document(DOCX)
AL_NAME={AL.LEFT:"left",AL.RIGHT:"right",AL.CENTER:"center",AL.JUSTIFY:"justify",None:"left"}
AL_IT={"left":"sinistra","right":"destra","center":"centrato","justify":"giustificato"}
known={norm(DELIVERY):"delivery",norm(DATE):"date","oggetto:":"ogglabel",
       norm(SUBJECT.split(":",1)[1]):"oggbody",norm(OPENING):"opening",
       norm(CLOSING):"closing","allegati:":"attach",norm(POST):"post"}
for k,b in BODY: known[norm(blk_text(b))]=k
for ln in SIGNATURE: known[norm(ln.text if isinstance(ln,Span) else ln)]="signature"
for ln in ATTACH: known[norm(ln)]="attach"
def classify(p):
    t=norm(p.text)
    if t in known: return known[t]
    li=p.paragraph_format.left_indent
    if li is not None and abs(Emu(li).cm-8.5)<0.2: return "recipient"
    return "body"
def pinfo(p):
    sizes=[r.font.size.pt for r in p.runs if r.text and r.font.size]
    fonts=[r.font.name for r in p.runs if r.text and r.font.name]; pf=p.paragraph_format
    return dict(al=AL_NAME.get(p.alignment,"left"),size=max(sizes) if sizes else 12,
        font=fonts[0] if fonts else "Times New Roman",anybold=any(r.bold for r in p.runs if r.text),
        italic=any(r.italic for r in p.runs if r.text),
        li=round(Emu(pf.left_indent).cm,2) if pf.left_indent is not None else 0,
        sb=pf.space_before.pt if pf.space_before is not None else 0,
        sa=pf.space_after.pt if pf.space_after is not None else 0,kwn=bool(pf.keep_with_next))
body_seq=[(classify(p),norm(p.text)) for p in doc.paragraphs if p.text.strip()]
agg={}
for p in doc.paragraphs:
    if p.text.strip(): agg.setdefault(classify(p),pinfo(p))
sec=doc.sections[0]
def part_lines(part):
    out=[]
    for p in part.paragraphs:
        if p.text.strip():
            szs=[r.font.size.pt for r in p.runs if r.text and r.font.size]
            out.append((p.text,max(szs) if szs else None,p.runs[0].font.name if p.runs and p.runs[0].font.name else None))
    return out
header_lines, footer_lines = part_lines(sec.header), part_lines(sec.footer)
top_cm=round(Emu(sec.top_margin).cm,2); bot_cm=round(Emu(sec.bottom_margin).cm,2)
hdr_set=[norm(t) for t,_,_ in header_lines]; ftr_set=[norm(t) for t,_,_ in footer_lines]

# ---------- 3) DOCX -> PDF -> PNG + bbox ----------
def soffice_pdf(src,outdir):
    prof=tempfile.mkdtemp(prefix="lo_")
    subprocess.run(["soffice","--headless",f"-env:UserInstallation=file://{prof}",
        "--convert-to","pdf","--outdir",outdir,src],check=True,
        stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=180)
    shutil.rmtree(prof,ignore_errors=True)
    return os.path.join(outdir,os.path.splitext(os.path.basename(src))[0]+".pdf")
soffice_pdf(DOCX,OUT)
print(f"[2] PDF {PDF}")
for old in glob.glob(PNG_PREFIX+"-*.png"): os.remove(old)
subprocess.run(["pdftoppm","-png","-r","150",PDF,PNG_PREFIX],check=True,timeout=120)
pages_png=sorted(glob.glob(PNG_PREFIX+"-*.png"))
BBOX=os.path.join(OUTH,"_bbox.xhtml")
subprocess.run(["pdftotext","-bbox",PDF,BBOX],check=True,timeout=60)
xml=open(BBOX,encoding="utf-8").read(); os.remove(BBOX)
print(f"[3] PNG {[os.path.basename(p) for p in pages_png]}")

# ---------- 4) words -> lines -> regions (% coords) per page ----------
PT_CM=28.3464567
top_pt=top_cm*PT_CM;
def line_key(text, yMin, yMax, pageH):
    if yMax <= top_pt + 8:
        for t in hdr_set:
            if t and (t in text or text in t): return "header"
        return "header"
    if yMin >= pageH - bot_cm*PT_CM - 8:
        return "footer"
    # exact paragraph match first (so short headings like "DIFFIDA" are not
    # swallowed by a longer paragraph that merely contains the word)
    for key,ptext in body_seq:
        if text and text == ptext: return key
    # then containment, for wrapped lines (a paragraph split over several lines)
    for key,ptext in body_seq:
        if len(text) >= 4 and text in ptext: return key
    return "body"

COLORS={"header":(100,116,139),"delivery":(202,138,4),"date":(13,148,136),
    "recipient":(79,70,229),"ogglabel":(225,29,72),"oggbody":(219,39,119),
    "opening":(22,163,74),"body":(37,99,235),"subhead":(124,58,237),"ritual":(234,88,12),
    "list":(8,145,178),"closing":(101,163,13),"signature":(146,64,14),"attach":(162,28,175),
    "post":(71,85,105),"footer":(100,116,139)}

pages=re.findall(r'<page width="([\d.]+)" height="([\d.]+)">(.*?)</page>', xml, re.S)
regions=[]   # dict(page, key, l,t,w,h in %)
for pidx,(pw,ph,bodyx) in enumerate(pages):
    pw=float(pw); ph=float(ph)
    words=[(float(a),float(b),float(c),float(d),html.unescape(e))
           for a,b,c,d,e in re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', bodyx)]
    words.sort(key=lambda w:(round(w[1],1), w[0]))
    # group into visual lines
    lines=[]; cur=[]
    for w in words:
        if cur and abs(w[1]-cur[-1][1])>5:
            lines.append(cur); cur=[]
        cur.append(w)
    if cur: lines.append(cur)
    L=[]
    for ln in lines:
        xs=[w[0] for w in ln]+[w[2] for w in ln]; ys=[w[1] for w in ln]+[w[3] for w in ln]
        txt=norm(" ".join(w[4] for w in sorted(ln,key=lambda w:w[0])))
        L.append((min(xs),min(ys),max(xs),max(ys),txt))
    # classify + group consecutive same-key (split on vertical gap)
    run=None
    for (x0,y0,x1,y1,txt) in L:
        k=line_key(txt,y0,y1,ph)
        lh=y1-y0
        if run and run["k"]==k and (y0-run["y1"])<=lh*1.8:
            run["x0"]=min(run["x0"],x0); run["y0"]=min(run["y0"],y0)
            run["x1"]=max(run["x1"],x1); run["y1"]=max(run["y1"],y1)
        else:
            if run: regions.append(run)
            run=dict(k=k,page=pidx,x0=x0,y0=y0,x1=x1,y1=y1,pw=pw,ph=ph)
    if run: regions.append(run)

def pct(r):
    pad=1.2  # small px-pt padding around region
    l=(r["x0"]-pad)/r["pw"]*100; t=(r["y0"]-pad)/r["ph"]*100
    w=(r["x1"]-r["x0"]+2*pad)/r["pw"]*100; h=(r["y1"]-r["y0"]+2*pad)/r["ph"]*100
    return max(0,l),max(0,t),w,h
present=set(r["k"] for r in regions)

# ---------- 5) legend ----------
LABEL={"header":"Intestazione (letterhead)","delivery":"Metodo di trasmissione","date":"Data e luogo",
 "recipient":"Destinatario","ogglabel":"OGGETTO — etichetta","oggbody":"OGGETTO — contenuto",
 "opening":"Apertura (saluto)","body":"Paragrafo di corpo","subhead":"Titoletto a sinistra",
 "ritual":"Titolo rituale centrato","list":"Voce di elenco","closing":"Congedo",
 "signature":"Blocco firma","attach":"Allegati","post":"Postilla / disclaimer","footer":"Piè di pagina"}
TAG={"header":"Ereditato","footer":"Ereditato","ogglabel":"Punto sensibile","oggbody":"Punto sensibile","signature":"Da approvare"}
NOTE={"header":"Carta intestata ereditata dal template; il corpo non la duplica.",
 "delivery":"Reso solo se fornito. Può anche essere la prima riga del destinatario.",
 "date":"Resa solo se fornita; lo script non inventa una data assente.",
 "recipient":"Appellativo personale unito al nome (“Egr. Sig. Mario Rossi”); “Spett.le” davanti a società resta su riga propria.",
 "ogglabel":"Punto sensibile (feedback 027): etichetta come titolo centrato, lasciata come scritta.",
 "oggbody":"La contestazione su 027 era il contenuto centrato invece che giustificato: ora è giustificato.",
 "opening":"Formale → grassetto + giustificato; e-mail → tondo + sinistra. Reso solo se fornito.",
 "body":"Normalizzazione: “–/—” → “-”; divisori “***” rimossi.",
 "subhead":"Es. “Fatto”, “MOTIVA …”; i verbi dispositivi usano titoletto + contenuto a capo.",
 "ritual":"Es. “IN FATTO”, “DICHIARA”, “DIFFIDA”, “PREMESSO CHE”, “INVITA”.",
 "list":"Marcatore ((i)/(a)/•) incluso nel testo.","closing":"Stesso lato della firma (destra).",
 "signature":"Nomi firmatari in grassetto; “(Firma digitale)” in corsivo; spaziatura ariosa — da approvare.",
 "attach":"“Allegati:” in grassetto; voci rientrate.","post":"Piccola, in corsivo.",
 "footer":"Dati di registrazione dello Studio, ereditati dal template."}
ORDER=["header","delivery","date","recipient","ogglabel","oggbody","opening","body",
 "subhead","ritual","list","closing","signature","attach","post","footer"]
def esc(s): return html.escape(s or "",quote=True)
def details(key):
    rows=[]
    if key in ("header","footer"):
        lines=header_lines if key=="header" else footer_lines
        szs=sorted({s for _,s,_ in lines if s}); font=(lines[0][2] if lines else "Garamond") or "Garamond"
        rows=[("Provenienza","Template (non generato dallo script)"),("Font",esc(font))]
        if szs: rows.append(("Dimensioni"," / ".join(f"{s:g} pt" for s in szs)))
        rows+= [("Allineamento","centrato"),("Pagina",f"margine sup. {top_cm:g} cm" if key=="header" else f"margine inf. {bot_cm:g} cm")]
    else:
        i=agg.get(key)
        if i:
            rows=[("Provenienza","Generato dallo script"),("Font",esc(i["font"])),("Dimensione",f"{i['size']:g} pt"),("Allineamento",AL_IT[i["al"]])]
            if i["li"]: rows.append(("Rientro sx",f"{i['li']:g} cm"))
            rows+=[("Grassetto","sì" if i["anybold"] else "no"),("Corsivo","sì" if i["italic"] else "no"),
                   ("Spaziatura",f"prima {i['sb']:g} / dopo {i['sa']:g} pt"),("Anti-orfano","keep-with-next" if i["kwn"] else "no")]
    dl="".join(f"<dt>{k}</dt><dd>{v}</dd>" for k,v in rows)
    note=f'<div class="note">{NOTE.get(key,"")}</div>' if NOTE.get(key) else ""
    return f"<dl>{dl}</dl>{note}"
def summary(key):
    if key in ("header","footer"): return "Ereditato dal template (Garamond)."
    i=agg.get(key)
    if not i: return ""
    bits=[AL_IT[i["al"]],f"{i['size']:g} pt"]
    if i["anybold"]: bits.append("grassetto")
    if i["italic"]: bits.append("corsivo")
    if i["li"]: bits.append(f"rientro {i['li']:g} cm")
    return ", ".join(bits)+"."
legend=[]
for key in ORDER:
    if key not in present: continue
    r,g,b=COLORS[key]; tag=f' <span class="tag">{TAG[key]}</span>' if key in TAG else ""
    legend.append(f'''    <div class="item" data-k="{key}" tabindex="0" style="--c:{r},{g},{b}">
      <span class="dot" aria-hidden="true"></span>
      <div class="meta"><div class="label">{LABEL[key]}{tag}</div>
        <div class="sum">{summary(key)}</div><div class="details">{details(key)}</div></div>
    </div>''')
legend_html="\n".join(legend)

# ---------- 6) pages with overlay regions ----------
def col(k): r,g,b=COLORS[k]; return f"{r},{g},{b}"
pages_html=[]
for pidx,png in enumerate(pages_png):
    rg=[]
    for r in regions:
        if r["page"]!=pidx: continue
        l,t,w,h=pct(r)
        rg.append(f'<div class="rg" data-k="{r["k"]}" tabindex="0" title="{esc(LABEL[r["k"]])}" '
                  f'style="left:{l:.2f}%;top:{t:.2f}%;width:{w:.2f}%;height:{h:.2f}%;--c:{col(r["k"])}"></div>')
    pages_html.append(f'''      <div class="pagewrap">
        <img class="page" src="{esc(os.path.basename(png))}" alt="Lettera fittizia — pagina {pidx+1} (render reale)">
{os.linesep.join("        "+x for x in rg)}
      </div>''')
pages_block="\n".join(pages_html)
legend_keys=[k for k in ORDER if k in present]

PAGE=f"""<!DOCTYPE html>
<html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Revisione formattazione — Bergamo Legal (lettera fittizia)</title>
<style>
  :root{{--ink:#1b2330;--muted:#64707f;--line:#e2e7ee;--bg:#eef1f5;--paper:#fff}}
  *{{box-sizing:border-box}} html,body{{margin:0}}
  body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:var(--ink);background:var(--bg);line-height:1.45}}
  .top{{padding:16px 24px;border-bottom:1px solid var(--line);background:#fff}}
  .top h1{{font-size:17px;margin:0 0 3px;letter-spacing:.01em}} .top p{{margin:0;color:var(--muted);font-size:12.5px;max-width:80ch}}
  .layout{{display:grid;grid-template-columns:340px 1fr;gap:0;align-items:start}}
  .legend{{position:sticky;top:0;align-self:start;max-height:100vh;overflow:auto;padding:14px;border-right:1px solid var(--line);background:#fff}}
  .legend h2{{font-size:12px;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);margin:2px 0 10px}}
  .paper-wrap{{padding:26px 26px 60px;display:flex;flex-direction:column;align-items:center;gap:18px}}
  .pagewrap{{position:relative;width:740px;max-width:100%;line-height:0}}
  .page{{width:100%;display:block;border:1px solid var(--line);box-shadow:0 3px 18px rgba(20,30,45,.10);background:#fff}}
  .rg{{position:absolute;border-radius:3px;cursor:pointer;outline:none;background:transparent;transition:background .12s,box-shadow .12s}}
  .rg:hover,.rg:focus,.rg.on{{background:rgba(var(--c),.20);box-shadow:0 0 0 2px rgba(var(--c),.85)}}
  .item{{display:flex;gap:10px;padding:8px;border-radius:9px;outline:none;border:1px solid transparent;cursor:pointer}}
  .item+.item{{margin-top:2px}}
  .item:hover,.item:focus,.item.on{{background:rgba(var(--c),.10);border-color:rgba(var(--c),.5)}}
  .item:focus{{box-shadow:0 0 0 2px rgba(var(--c),.55)}}
  .dot{{flex:0 0 12px;height:12px;margin-top:4px;border-radius:3px;background:rgb(var(--c))}}
  .meta{{min-width:0}} .label{{font-weight:600;font-size:13px}}
  .label .tag{{font-weight:600;font-size:9.5px;text-transform:uppercase;letter-spacing:.04em;color:rgb(var(--c));margin-left:6px}}
  .sum{{font-size:11.5px;color:var(--muted);margin-top:1px}}
  .details{{display:none;margin-top:7px;border-top:1px dashed rgba(var(--c),.5);padding-top:7px;font-size:11.5px}}
  .item:hover .details,.item:focus .details,.item.on .details{{display:block}}
  .details dl{{display:grid;grid-template-columns:auto 1fr;gap:1px 10px;margin:0}}
  .details dt{{color:var(--muted)}} .details dd{{margin:0}} .details .note{{margin-top:5px;color:#475569}}
  #wires{{position:fixed;inset:0;width:100%;height:100%;pointer-events:none;z-index:50}}
  @media (max-width:900px){{.layout{{grid-template-columns:1fr}} .legend{{position:static;max-height:none;border-right:none;border-bottom:1px solid var(--line)}} #wires{{display:none}}}}
  @media print{{body{{background:#fff}} .layout{{display:block}} .legend{{position:static;max-height:none;border:none}} .details{{display:block !important}} #wires{{display:none}} .page{{box-shadow:none}}}}
</style></head>
<body>
<header class="top">
  <h1>Revisione formattazione — Lettere Bergamo Legal</h1>
  <p>Lettera <strong>fittizia</strong> resa dal formatter (DOCX → PDF → immagine): font e
     impaginazione sono quelli reali. Passa il mouse o usa <kbd>Tab</kbd> su una voce della
     legenda o su una zona della lettera: le parti corrispondenti si evidenziano a vicenda e
     una linea le collega. Documenti puliti: <code>out/lettera_fittizia.docx</code> / <code>.pdf</code>.</p>
</header>
<svg id="wires" aria-hidden="true"></svg>
<div class="layout">
  <aside class="legend" aria-label="Legenda delle sezioni">
    <h2>Sezioni della lettera</h2>
{legend_html}
  </aside>
  <main class="paper-wrap">
{pages_block}
  </main>
</div>
<script>
(function(){{
  var KEYS={legend_keys!r};
  function all(s){{return Array.prototype.slice.call(document.querySelectorAll(s));}}
  var svg=document.getElementById('wires'), active=null;
  function items(k){{return all('.item[data-k="'+k+'"]');}}
  function regs(k){{return all('.rg[data-k="'+k+'"]');}}
  function rgb(k){{var el=items(k)[0]; return el?getComputedStyle(el).getPropertyValue('--c'):'30,40,60';}}
  function clearWire(){{svg.innerHTML='';}}
  function nearestReg(k){{
    var vh=window.innerHeight, best=null, bd=1e9;
    regs(k).forEach(function(r){{var b=r.getBoundingClientRect();
      var d=(b.top<0)? -b.top : (b.top>vh? b.top-vh : 0);
      if(d<bd){{bd=d;best=r;}}}});
    return best;
  }}
  function draw(){{
    clearWire(); if(!active) return;
    var it=items(active)[0], rg=nearestReg(active); if(!it||!rg) return;
    var a=it.getBoundingClientRect(), b=rg.getBoundingClientRect();
    var x1=a.right-6, y1=a.top+a.height/2, x2=b.left+2, y2=b.top+b.height/2;
    var mx=(x1+x2)/2;
    var p=document.createElementNS('http://www.w3.org/2000/svg','path');
    p.setAttribute('d','M'+x1+','+y1+' C'+mx+','+y1+' '+mx+','+y2+' '+x2+','+y2);
    p.setAttribute('fill','none'); p.setAttribute('stroke','rgb('+rgb(active)+')');
    p.setAttribute('stroke-width','2'); p.setAttribute('stroke-dasharray','5 4'); p.setAttribute('opacity','.9');
    var dot=document.createElementNS('http://www.w3.org/2000/svg','circle');
    dot.setAttribute('cx',x2);dot.setAttribute('cy',y2);dot.setAttribute('r','3.5');dot.setAttribute('fill','rgb('+rgb(active)+')');
    svg.appendChild(p); svg.appendChild(dot);
  }}
  function set(k,on){{
    items(k).forEach(function(e){{e.classList.toggle('on',on);}});
    regs(k).forEach(function(e){{e.classList.toggle('on',on);}});
  }}
  function activate(k){{ if(active && active!==k) set(active,false); active=k; set(k,true); draw(); }}
  function deactivate(k){{ if(active===k){{set(k,false); active=null; clearWire();}} }}
  function wire(el){{var k=el.getAttribute('data-k');
    el.addEventListener('mouseenter',function(){{activate(k);}});
    el.addEventListener('mouseleave',function(){{deactivate(k);}});
    el.addEventListener('focus',function(){{activate(k);}});
    el.addEventListener('blur',function(){{deactivate(k);}});
  }}
  all('.item').forEach(wire); all('.rg').forEach(wire);
  var raf=null;
  function onScroll(){{ if(raf) return; raf=requestAnimationFrame(function(){{raf=null; if(active) draw();}}); }}
  window.addEventListener('scroll',onScroll,true); window.addEventListener('resize',onScroll);
}})();
</script>
</body></html>
"""
open(HTMLF,"w",encoding="utf-8").write(PAGE)
print(f"[6] HTML {HTMLF}")
print("regioni per pagina:", {i:sum(1 for r in regions if r['page']==i) for i in range(len(pages_png))})
print("sezioni:", legend_keys)
miss=[k for k in present if k not in COLORS]
print("chiavi senza colore:", miss or "nessuna")
