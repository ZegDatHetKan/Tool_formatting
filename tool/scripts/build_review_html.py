#!/usr/bin/env python3
"""Build the client review artifact from the REAL formatter output.

Pipeline (this is the point — use the formatter's true pagination, not a CSS
approximation):

  invent fictional content
    -> LetterDocument -> render_letter -> out/lettera_fittizia.docx   (real DOCX)
    -> LibreOffice headless            -> out/lettera_fittizia.pdf    (real PDF)
    -> a *shaded copy* (per-section background colors) -> PDF -> PNG  (real layout,
       with each section highlighted in the actual rendering)
    -> out/html/lettera_fittizia_review.html embeds those page images + a legend
       whose details are read from the real DOCX.

Run:  .venv/bin/python tool/scripts/build_review_html.py
"""
from __future__ import annotations
import os, sys, html, glob, shutil, subprocess, tempfile
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from docx import Document
from docx.shared import Emu
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from formatters.letters import (
    LetterDocument, Span, render_letter,
    paragraph, section_heading, subheading, list_item,
)

OUT = os.path.join(ROOT, "out")
OUTH = os.path.join(OUT, "html")
os.makedirs(OUTH, exist_ok=True)
TEMPLATE = os.path.join(ROOT, "assets", "Template_Vuoto.docx")
DOCX = os.path.join(OUT, "lettera_fittizia.docx")          # real, clean (deliverable)
PDF = os.path.join(OUT, "lettera_fittizia.pdf")            # real, clean (deliverable)
SHADED = os.path.join(OUTH, "_lettera_fittizia_highlight.docx")
HTMLF = os.path.join(OUTH, "lettera_fittizia_review.html")
PNG_PREFIX = os.path.join(OUTH, "lettera_fittizia_review")  # -1.png, -2.png ...

# --------------------------------------------------------------------------
# 1) Fictional content (invented; no real client data)
# --------------------------------------------------------------------------
DELIVERY = "A mezzo PEC, anticipata a mezzo posta elettronica"
DATE = "Bergamo, 24 giugno 2026"
RECIPIENT = ["Egr.", Span("Sig. Mario Rossi", bold=True),
             "Via dei Tigli, n. 14", "24121 Bergamo (BG)",
             Span("PEC: mario.rossi@esempio-pec.it", bold=True)]
SUBJECT = "OGGETTO: Diffida ad adempiere - contratto di fornitura del 10 gennaio 2026"
OPENING = "Egregio Signore,"
CLOSING = "In attesa di un cortese e sollecito riscontro, si porgono distinti saluti."
SIGNATURE = ["Bergamo Legal Società tra Avvocati S.r.l.",
             Span("Avv. Giulia Conti", bold=True),
             "_______________________________"]
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

letter = LetterDocument(
    delivery_method=DELIVERY, date_place=DATE, recipient_block=RECIPIENT,
    subject=SUBJECT, opening=OPENING, body_blocks=[b for _, b in BODY],
    closing=CLOSING, signature_block=SIGNATURE, attachments=ATTACH, postscript=POST,
)
res = render_letter(letter, TEMPLATE, DOCX)
print(f"[1] render DOCX: {DOCX} ({res.paragraph_count} par) needs_review={res.needs_review}")

# --------------------------------------------------------------------------
# 2) classify each rendered paragraph into a legend section
# --------------------------------------------------------------------------
def norm(s): return (s or "").replace("–", "-").replace("—", "-").strip()
def blk_text(blk):
    c = blk.content
    if isinstance(c, str): return c
    if isinstance(c, Span): return c.text
    if isinstance(c, list): return "".join(x.text if isinstance(x, Span) else str(x) for x in c)
    return str(c)

known = {norm(DELIVERY): "delivery", norm(DATE): "date", "OGGETTO:": "ogglabel",
         norm(SUBJECT.split(":", 1)[1]): "oggbody", norm(OPENING): "opening",
         norm(CLOSING): "closing", "Allegati:": "attach", norm(POST): "post"}
for k, b in BODY: known[norm(blk_text(b))] = k
for ln in SIGNATURE: known[norm(ln.text if isinstance(ln, Span) else ln)] = "signature"
for ln in ATTACH: known[norm(ln)] = "attach"

AL_NAME = {AL.LEFT: "left", AL.RIGHT: "right", AL.CENTER: "center", AL.JUSTIFY: "justify", None: "left"}
AL_IT = {"left": "sinistra", "right": "destra", "center": "centrato", "justify": "giustificato"}

def classify(p):
    t = norm(p.text)
    if t in known: return known[t]
    li = p.paragraph_format.left_indent
    if li is not None and abs(Emu(li).cm - 8.5) < 0.2: return "recipient"
    return "body"

def pinfo(p):
    sizes = [r.font.size.pt for r in p.runs if r.text and r.font.size]
    fonts = [r.font.name for r in p.runs if r.text and r.font.name]
    pf = p.paragraph_format
    return dict(al=AL_NAME.get(p.alignment, "left"),
                size=max(sizes) if sizes else 12,
                font=fonts[0] if fonts else "Times New Roman",
                anybold=any(r.bold for r in p.runs if r.text),
                italic=any(r.italic for r in p.runs if r.text),
                li=round(Emu(pf.left_indent).cm, 2) if pf.left_indent is not None else 0,
                sb=pf.space_before.pt if pf.space_before is not None else 0,
                sa=pf.space_after.pt if pf.space_after is not None else 0,
                kwn=bool(pf.keep_with_next))

doc = Document(DOCX)
body_keys = [classify(p) for p in doc.paragraphs if p.text.strip()]
agg = {}
for p in doc.paragraphs:
    if not p.text.strip(): continue
    agg.setdefault(classify(p), pinfo(p))

sec = doc.sections[0]
def part_lines(part):
    out = []
    for p in part.paragraphs:
        if p.text.strip():
            szs = [r.font.size.pt for r in p.runs if r.text and r.font.size]
            out.append((p.text, max(szs) if szs else None,
                        p.runs[0].font.name if p.runs and p.runs[0].font.name else None))
    return out
header_lines, footer_lines = part_lines(sec.header), part_lines(sec.footer)
top_cm = round(Emu(sec.top_margin).cm, 2); bot_cm = round(Emu(sec.bottom_margin).cm, 2)

# --------------------------------------------------------------------------
# 3) colors + a shaded copy of the DOCX (highlights baked into REAL layout)
# --------------------------------------------------------------------------
COLORS = {  # solid RGB per section
    "header": (100,116,139), "delivery": (202,138,4), "date": (13,148,136),
    "recipient": (79,70,229), "ogglabel": (225,29,72), "oggbody": (219,39,119),
    "opening": (22,163,74), "body": (37,99,235), "subhead": (124,58,237),
    "ritual": (234,88,12), "list": (8,145,178), "closing": (101,163,13),
    "signature": (146,64,14), "attach": (162,28,175), "post": (71,85,105),
    "footer": (100,116,139),
}
def tint_hex(rgb, f=0.78):  # blend toward white for a readable page highlight
    r, g, b = rgb
    return "{:02X}{:02X}{:02X}".format(int(r+(255-r)*f), int(g+(255-g)*f), int(b+(255-b)*f))

def shade(p, hex6):
    pPr = p._p.get_or_add_pPr()
    for old in pPr.findall(qn('w:shd')): pPr.remove(old)
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex6)
    pPr.append(shd)

sh = Document(DOCX)
for p in sh.paragraphs:
    if not p.text.strip(): continue
    shade(p, tint_hex(COLORS[classify(p)]))
shsec = sh.sections[0]
for part, key in ((shsec.header, "header"), (shsec.footer, "footer")):
    for p in part.paragraphs:
        if p.text.strip(): shade(p, tint_hex(COLORS[key]))
sh.save(SHADED)
print(f"[2] shaded copy: {SHADED}")

# --------------------------------------------------------------------------
# 4) DOCX -> PDF (LibreOffice headless) -> PNG (pdftoppm)
# --------------------------------------------------------------------------
def soffice_pdf(src, outdir):
    profile = tempfile.mkdtemp(prefix="lo_")
    subprocess.run(["soffice", "--headless", f"-env:UserInstallation=file://{profile}",
                    "--convert-to", "pdf", "--outdir", outdir, src],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    shutil.rmtree(profile, ignore_errors=True)
    return os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")

# clean PDF deliverable
clean_pdf = soffice_pdf(DOCX, OUT)
print(f"[3] clean PDF: {clean_pdf}")
# shaded PDF -> PNG pages
shaded_pdf = soffice_pdf(SHADED, OUTH)
for old in glob.glob(PNG_PREFIX + "-*.png"): os.remove(old)
subprocess.run(["pdftoppm", "-png", "-r", "150", shaded_pdf, PNG_PREFIX],
               check=True, timeout=120)
pages = sorted(glob.glob(PNG_PREFIX + "-*.png"))
os.remove(shaded_pdf); os.remove(SHADED)
print(f"[4] page images: {[os.path.basename(p) for p in pages]}")

# --------------------------------------------------------------------------
# 5) legend + HTML (embeds the real-layout page images)
# --------------------------------------------------------------------------
LABEL = {"header":"Intestazione (letterhead)","delivery":"Metodo di trasmissione","date":"Data e luogo",
         "recipient":"Destinatario","ogglabel":"OGGETTO — etichetta","oggbody":"OGGETTO — contenuto",
         "opening":"Apertura (saluto)","body":"Paragrafo di corpo","subhead":"Titoletto a sinistra",
         "ritual":"Titolo rituale centrato","list":"Voce di elenco","closing":"Congedo",
         "signature":"Blocco firma","attach":"Allegati","post":"Postilla / disclaimer","footer":"Piè di pagina"}
TAG = {"header":"Ereditato dal template","footer":"Ereditato dal template",
       "ogglabel":"Punto sensibile · Da approvare","oggbody":"Punto sensibile · Da approvare",
       "signature":"Da approvare"}
NOTE = {"header":"Carta intestata (logo + “BERGAMO LEGAL”) ereditata dal template; il corpo non la duplica.",
        "delivery":"Reso solo se fornito. Variante: può essere la prima riga del destinatario.",
        "date":"Resa solo se fornita; lo script non inventa una data assente.",
        "recipient":"Appellativo personale unito al nome (“Egr. Sig. Mario Rossi”); “Spett.le” davanti a società resta su riga propria.",
        "ogglabel":"Punto sensibile (feedback Trustpilot v3 / 027): etichetta come titolo centrato; lasciata come scritta, non forzata in maiuscolo.",
        "oggbody":"La contestazione su 027 era il contenuto centrato invece che giustificato: ora è giustificato.",
        "opening":"Formale → grassetto + giustificato; stile e-mail → tondo + sinistra. Reso solo se fornito.",
        "body":"Normalizzazione: “–/—” → “-”; divisori “***” rimossi.",
        "subhead":"Es. “Fatto”, “Diritto”, “MOTIVA …”; i verbi dispositivi (“CHIEDE ALTRESÌ”) usano titoletto + contenuto a capo.",
        "ritual":"Es. “IN FATTO”, “DICHIARA”, “DIFFIDA”, “PREMESSO CHE”, “INVITA”.",
        "list":"Marcatore ((i)/(a)/•) incluso nel testo.",
        "closing":"Stesso lato della firma (destra).",
        "signature":"Nomi firmatari in grassetto; “(Firma digitale)” in corsivo; spaziatura ariosa — da approvare.",
        "attach":"“Allegati:” in grassetto; voci rientrate. Reso solo se presente.",
        "post":"Piccola, in corsivo. Resa solo se presente.",
        "footer":"Dati di registrazione dello Studio, ereditati dal template (non generati)."}
ORDER = ["header","delivery","date","recipient","ogglabel","oggbody","opening","body",
         "subhead","ritual","list","closing","signature","attach","post","footer"]

def esc(s): return html.escape(s or "", quote=True)

def details_for(key):
    rows = []
    if key in ("header","footer"):
        lines = header_lines if key=="header" else footer_lines
        szs = sorted({s for _,s,_ in lines if s})
        font = (lines[0][2] if lines else "Garamond") or "Garamond"
        rows = [("Provenienza","Template_Vuoto.docx — <strong>non</strong> generato dallo script"),
                ("Font", esc(font))]
        if szs: rows.append(("Dimensioni", " / ".join(f"{s:g} pt" for s in szs)))
        rows.append(("Allineamento","centrato"))
        rows.append(("Pagina", f"A4; margine superiore {top_cm:g} cm" if key=="header" else f"margine inferiore {bot_cm:g} cm"))
    else:
        i = agg.get(key)
        if i:
            rows = [("Provenienza","Generato dallo script"),("Font",esc(i["font"])),
                    ("Dimensione",f"{i['size']:g} pt"),("Allineamento",AL_IT[i["al"]])]
            if i["li"]: rows.append(("Rientro sx",f"{i['li']:g} cm"))
            rows += [("Grassetto","sì" if i["anybold"] else "no"),
                     ("Corsivo","sì" if i["italic"] else "no"),
                     ("Spaziatura",f"prima {i['sb']:g} pt / dopo {i['sa']:g} pt"),
                     ("Anti-orfano","keep-with-next" if i["kwn"] else "no")]
    dl = "".join(f"<dt>{k}</dt><dd>{v}</dd>" for k,v in rows)
    note = f'<div class="note">{NOTE.get(key,"")}</div>' if NOTE.get(key) else ""
    return f"<dl>{dl}</dl>{note}"

def summary_for(key):
    if key in ("header","footer"): return "Ereditato dal template (Garamond). Non generato dal corpo."
    i = agg.get(key)
    if not i: return ""
    bits = [AL_IT[i["al"]], f"{i['size']:g} pt"]
    if i["anybold"]: bits.append("grassetto")
    if i["italic"]: bits.append("corsivo")
    if i["li"]: bits.append(f"rientro {i['li']:g} cm")
    return ", ".join(bits) + "."

present = set(body_keys) | {"header","footer"}
legend = []
for key in ORDER:
    if key not in present: continue
    r,g,b = COLORS[key]; tint = tint_hex(COLORS[key])
    tag = f' <span class="tag">{TAG[key]}</span>' if key in TAG else ""
    legend.append(f'''    <div class="item" tabindex="0" style="--c:{r},{g},{b}">
      <span class="swatch" style="background:#{tint};border-color:rgb({r},{g},{b})" aria-hidden="true"></span>
      <div class="meta"><div class="label">{LABEL[key]}{tag}</div>
        <div class="sum">{summary_for(key)}</div>
        <div class="details">{details_for(key)}</div></div>
    </div>''')
legend_html = "\n".join(legend)

imgs = "\n        ".join(
    f'<img class="page" src="{esc(os.path.basename(p))}" alt="Lettera fittizia — pagina {n+1} (render reale del DOCX)">'
    for n, p in enumerate(pages))

swatches_inline = "".join(
    f'<span class="lk" style="background:#{tint_hex(COLORS[k])};border-color:rgb{COLORS[k]}"></span>{LABEL[k]} &nbsp;'
    for k in ORDER if k in present)

PAGE = f"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Revisione formattazione lettere — Bergamo Legal (lettera fittizia)</title>
<style>
  :root{{--ink:#1c2430;--muted:#5b6675;--line:#d8dde5;--bg:#f4f6f9}}
  *{{box-sizing:border-box}} html,body{{margin:0}}
  body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:var(--ink);background:var(--bg);line-height:1.45}}
  .topbar{{padding:18px 24px;border-bottom:1px solid var(--line);background:#fff}}
  .topbar h1{{font-size:18px;margin:0 0 4px}} .topbar p{{margin:0;color:var(--muted);font-size:13px;max-width:78ch}}
  .hint{{margin-top:8px;font-size:12px;color:var(--muted)}}
  .layout{{display:grid;grid-template-columns:minmax(300px,360px) 1fr}}
  .legend{{position:sticky;top:0;max-height:100vh;overflow:auto;padding:16px;border-right:1px solid var(--line);background:#fff}}
  .legend h2{{font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin:4px 0 10px}}
  .paper-wrap{{padding:24px;display:flex;flex-direction:column;align-items:center;gap:16px}}
  .page{{width:760px;max-width:100%;border:1px solid var(--line);box-shadow:0 2px 14px rgba(20,30,45,.08);background:#fff}}
  .cap{{font-size:12px;color:var(--muted);max-width:760px}}
  .item{{display:flex;gap:10px;padding:9px 8px;border-radius:8px;outline:none;border:1px solid transparent}}
  .item+.item{{margin-top:2px}}
  .item:hover,.item:focus,.item:focus-within{{background:rgba(var(--c),.10);border-color:rgba(var(--c),.45)}}
  .item:focus{{box-shadow:0 0 0 2px rgba(var(--c),.55)}}
  .swatch{{flex:0 0 14px;height:14px;margin-top:3px;border-radius:4px;border:1px solid}}
  .meta{{min-width:0}} .label{{font-weight:600;font-size:13.5px}}
  .label .tag{{font-weight:500;font-size:10.5px;color:var(--muted);margin-left:6px}}
  .sum{{font-size:12px;color:var(--muted);margin-top:1px}}
  .details{{display:none;margin-top:8px;border-top:1px dashed rgba(var(--c),.5);padding-top:8px;font-size:12px}}
  .item:hover .details,.item:focus .details,.item:focus-within .details{{display:block}}
  .details dl{{display:grid;grid-template-columns:auto 1fr;gap:2px 10px;margin:0}}
  .details dt{{color:var(--muted)}} .details dd{{margin:0}}
  .details .note{{margin-top:6px;color:#475569}}
  .lkbar{{font-size:11px;color:var(--muted);line-height:1.9}}
  .lk{{display:inline-block;width:10px;height:10px;border-radius:2px;border:1px solid;margin:0 4px 0 10px;vertical-align:middle}}
  @media (max-width:880px){{.layout{{grid-template-columns:1fr}} .legend{{position:static;max-height:none;border-right:none;border-bottom:1px solid var(--line)}} .page{{width:100%}}}}
  @media print{{body{{background:#fff}} .layout{{display:block}} .legend{{position:static;max-height:none;border:none}} .details{{display:block !important}} .page{{box-shadow:none}}}}
</style>
</head>
<body>
<header class="topbar">
  <h1>Revisione formattazione — Lettere Bergamo Legal</h1>
  <p>Anteprima di una <strong>lettera fittizia</strong> (contenuto inventato).
     <strong>L'immagine è il render reale del DOCX prodotto dal formatter</strong>
     (DOCX → PDF via LibreOffice → immagine): font, impaginazione, intestazione e
     piè di pagina sono quelli effettivi. Le sezioni sono evidenziate con i colori
     della legenda; passa il mouse o porta il focus su una voce per i dettagli
     (letti dal documento reale).</p>
  <div class="lkbar">{swatches_inline}</div>
  <p class="hint">Documento pulito: <code>out/lettera_fittizia.docx</code> e
     <code>out/lettera_fittizia.pdf</code>. In stampa/PDF i dettagli della legenda sono espansi.</p>
</header>
<div class="layout">
  <aside class="legend" aria-label="Legenda delle sezioni di formattazione">
    <h2>Sezioni della lettera</h2>
{legend_html}
  </aside>
  <main class="paper-wrap">
        {imgs}
    <div class="cap">Render reale del DOCX (LibreOffice, 150 dpi). Le bande colorate
      sono applicate a una copia per la sola revisione; il documento consegnabile
      (<code>out/lettera_fittizia.docx</code> / <code>.pdf</code>) è senza evidenziazioni.</div>
  </main>
</div>
</body>
</html>
"""
with open(HTMLF, "w", encoding="utf-8") as f:
    f.write(PAGE)
print(f"[5] html: {HTMLF}")
print("sezioni:", [k for k in ORDER if k in present])
