#!/usr/bin/env python3
"""Generate the client review HTML from a REAL formatter render.

Flow (this is the point): invent fictional content -> build a LetterDocument ->
render it with formatters.letters.render_letter into a real .docx -> read that
.docx back -> emit an HTML review page whose preview and legend reflect the
*actual* per-paragraph formatting the script produced (alignment, size,
bold/italic, indent, spacing, keep-with-next). The header/footer are read from
the template and shown as inherited (not generated).

Run:  .venv/bin/python tool/scripts/build_review_html.py
Outputs: out/html/lettera_fittizia.docx
         out/html/lettera_fittizia_review.html
"""
from __future__ import annotations
import os, sys, html
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from docx import Document
from docx.shared import Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from formatters.letters import (
    LetterDocument, Span, render_letter,
    paragraph, section_heading, subheading, list_item,
)

OUTDIR = os.path.join(ROOT, "out", "html")
DOCX = os.path.join(OUTDIR, "lettera_fittizia.docx")
HTMLF = os.path.join(OUTDIR, "lettera_fittizia_review.html")
TEMPLATE = os.path.join(ROOT, "assets", "Template_Vuoto.docx")

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

# body blocks, tagged with the legend key each one belongs to
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
print(f"render: {DOCX} ({res.paragraph_count} paragrafi) needs_review={res.needs_review}")

# --------------------------------------------------------------------------
# 2) Read the rendered docx back (real per-paragraph properties)
# --------------------------------------------------------------------------
def norm(s): return (s or "").replace("–", "-").replace("—", "-").strip()

doc = Document(DOCX)

# text -> legend key (built from the fictional content we supplied)
known = {}
known[norm(DELIVERY)] = "delivery"
known[norm(DATE)] = "date"
known["OGGETTO:"] = "ogglabel"
known[norm(SUBJECT.split(":", 1)[1])] = "oggbody"
known[norm(OPENING)] = "opening"
for key, blk in BODY:
    known[norm(_plain := "".join(s.text if hasattr(s, "text") else s for s in ([blk.content] if isinstance(blk.content, str) else (blk.content if isinstance(blk.content, list) else [blk.content]))))] = key
known[norm(CLOSING)] = "closing"
for ln in SIGNATURE:
    known[norm(ln.text if isinstance(ln, Span) else ln)] = "signature"
known["Allegati:"] = "attach"
for ln in ATTACH:
    known[norm(ln)] = "attach"
known[norm(POST)] = "post"

AL_NAME = {AL.LEFT: "left", AL.RIGHT: "right", AL.CENTER: "center", AL.JUSTIFY: "justify", None: "left"}
AL_IT = {"left": "sinistra", "right": "destra", "center": "centrato", "justify": "giustificato"}

def classify(p):
    t = norm(p.text)
    if t in known:
        return known[t]
    li = p.paragraph_format.left_indent
    if li is not None and abs(Emu(li).cm - 8.5) < 0.2:
        return "recipient"   # merged appellative line ("Egr. Sig. Mario Rossi")
    return "body"

def pinfo(p):
    al = AL_NAME.get(p.alignment, "left")
    sizes = [r.font.size.pt for r in p.runs if r.text and r.font.size]
    size = max(sizes) if sizes else 12
    fonts = [r.font.name for r in p.runs if r.text and r.font.name]
    font = fonts[0] if fonts else "Times New Roman"
    bold = any(r.bold for r in p.runs if r.text) and all((r.bold or not r.text) for r in p.runs)
    anybold = any(r.bold for r in p.runs if r.text)
    italic = any(r.italic for r in p.runs if r.text)
    pf = p.paragraph_format
    li = round(Emu(pf.left_indent).cm, 2) if pf.left_indent is not None else 0
    sb = pf.space_before.pt if pf.space_before is not None else 0
    sa = pf.space_after.pt if pf.space_after is not None else 0
    kwn = bool(pf.keep_with_next)
    return dict(al=al, size=size, font=font, bold=bold, anybold=anybold,
                italic=italic, li=li, sb=sb, sa=sa, kwn=kwn, text=p.text)

paras = [(classify(p), pinfo(p)) for p in doc.paragraphs if p.text.strip()]

# template header/footer (inherited)
sec = doc.sections[0]
def part_lines(part):
    out = []
    for p in part.paragraphs:
        if p.text.strip():
            szs = [r.font.size.pt for r in p.runs if r.text and r.font.size]
            out.append((p.text, max(szs) if szs else None,
                        (p.runs[0].font.name if p.runs and p.runs[0].font.name else None)))
    return out
header_lines = part_lines(sec.header)
footer_lines = part_lines(sec.footer)
top_cm = round(Emu(sec.top_margin).cm, 2)
bot_cm = round(Emu(sec.bottom_margin).cm, 2)
lr_cm = round(Emu(sec.left_margin).cm, 2)

# --------------------------------------------------------------------------
# 3) Legend metadata (colors + rationale notes; values come from the render)
# --------------------------------------------------------------------------
COLORS = {
    "header":"100,116,139","delivery":"202,138,4","date":"13,148,136",
    "recipient":"79,70,229","ogglabel":"225,29,72","oggbody":"219,39,119",
    "opening":"22,163,74","body":"37,99,235","subhead":"124,58,237",
    "ritual":"234,88,12","list":"8,145,178","closing":"101,163,13",
    "signature":"146,64,14","attach":"162,28,175","post":"71,85,105",
    "footer":"100,116,139",
}
LABEL = {
    "header":"Intestazione (letterhead)","delivery":"Metodo di trasmissione",
    "date":"Data e luogo","recipient":"Destinatario","ogglabel":"OGGETTO — etichetta",
    "oggbody":"OGGETTO — contenuto","opening":"Apertura (saluto)","body":"Paragrafo di corpo",
    "subhead":"Titoletto a sinistra","ritual":"Titolo rituale centrato","list":"Voce di elenco",
    "closing":"Congedo","signature":"Blocco firma","attach":"Allegati",
    "post":"Postilla / disclaimer","footer":"Piè di pagina",
}
TAG = {
    "header":"Ereditato dal template","footer":"Ereditato dal template",
    "ogglabel":"Punto sensibile · Da approvare","oggbody":"Punto sensibile · Da approvare",
    "signature":"Da approvare",
}
NOTE = {
    "header":"Carta intestata (logo + “BERGAMO LEGAL”) ereditata dal template; il corpo non la duplica.",
    "delivery":"Reso solo se fornito. Variante: può essere la prima riga del destinatario.",
    "date":"Resa solo se fornita; lo script non inventa una data assente.",
    "recipient":"Appellativo personale unito al nome sulla stessa riga (es. “Egr. Sig. Mario Rossi”); “Spett.le” davanti a società resta su riga propria.",
    "ogglabel":"Punto sensibile (feedback Trustpilot v3 / 027): etichetta come titolo centrato. Lasciata come scritta, non forzata in maiuscolo.",
    "oggbody":"La contestazione su 027 era il contenuto centrato invece che giustificato: ora è giustificato.",
    "opening":"Formale → grassetto + giustificato; stile e-mail → tondo + sinistra. Reso solo se fornito.",
    "body":"Normalizzazione: trattini “–/—” → “-”; divisori “***” rimossi. Grassetto/corsivo leggero inline ammesso.",
    "subhead":"Es. “Fatto”, “Diritto”, “MOTIVA …”. I verbi dispositivi (“CHIEDE ALTRESÌ”) usano questo titoletto + contenuto a capo.",
    "ritual":"Es. “IN FATTO”, “DICHIARA”, “DIFFIDA”, “PREMESSO CHE”, “INVITA”.",
    "list":"Marcatore ((i)/(a)/•) incluso nel testo.",
    "closing":"Stesso lato della firma (destra).",
    "signature":"Nomi firmatari in grassetto; “(Firma digitale)” in corsivo; spaziatura ariosa — da approvare.",
    "attach":"“Allegati:” in grassetto; voci rientrate. Reso solo se presente.",
    "post":"Piccola, in corsivo. Resa solo se presente.",
    "footer":"Dati di registrazione dello Studio, ereditati dal template (non generati).",
}
ORDER = ["header","delivery","date","recipient","ogglabel","oggbody","opening",
         "body","subhead","ritual","list","closing","signature","attach","post","footer"]

# aggregate the real values observed per key (first paragraph of that key)
agg = {}
for key, info in paras:
    agg.setdefault(key, info)

CMPX = 34.0  # px per cm in the preview
def px_pt(pt): return round(pt * CMPX * 2.54 / 72, 1)
def px_cm(cm): return round(cm * CMPX, 1)

def esc(s): return html.escape(s, quote=True)

# --------------------------------------------------------------------------
# 4) Build the letter preview (each paragraph styled from REAL values)
# --------------------------------------------------------------------------
def seg_style(info):
    st = [f"text-align:{info['al']}", f"font-size:{px_pt(info['size'])}px"]
    if info["bold"]: st.append("font-weight:bold")
    if info["italic"]: st.append("font-style:italic")
    if info["li"]: st.append(f"margin-left:{px_cm(info['li'])}px")
    st.append(f"margin-top:{px_pt(info['sb'])}px")
    st.append(f"margin-bottom:{px_pt(info['sa'])}px")
    return ";".join(st)

# group consecutive paragraphs of the same key into one visual block (recipient,
# signature, attachments span several paragraphs but are one legend section)
letter_html = []
GROUPED = {"recipient", "signature", "attach"}
i = 0
while i < len(paras):
    key, info = paras[i]
    if key in GROUPED:
        j = i
        inner = []
        while j < len(paras) and paras[j][0] == key:
            inner.append(f'<div style="{seg_style(paras[j][1])}">{esc(paras[j][1]["text"])}</div>')
            j += 1
        letter_html.append(
            f'<div class="seg c-{key}" data-k="{key}" tabindex="0">{"".join(inner)}</div>')
        i = j
    else:
        letter_html.append(
            f'<div class="seg c-{key}" data-k="{key}" tabindex="0" '
            f'style="{seg_style(info)}">{esc(info["text"])}</div>')
        i += 1
letter_body = "\n        ".join(letter_html)

# header/footer simulated blocks (inherited)
hdr = "<br>".join(esc(t) for t, _, _ in header_lines) or "BERGAMO LEGAL"
ftr = "<br>".join(esc(t) for t, _, _ in footer_lines) or "Dati Studio"
hdr_font = (header_lines[0][2] if header_lines else "Garamond") or "Garamond"
ftr_font = (footer_lines[0][2] if footer_lines else "Garamond") or "Garamond"

# --------------------------------------------------------------------------
# 5) Build the legend (details from the real render)
# --------------------------------------------------------------------------
def details_for(key):
    rows = []
    if key in ("header", "footer"):
        lines = header_lines if key == "header" else footer_lines
        szs = sorted({s for _, s, _ in lines if s})
        font = (lines[0][2] if lines else "Garamond") or "Garamond"
        rows.append(("Provenienza", "Template_Vuoto.docx — <strong>non</strong> generato dallo script"))
        rows.append(("Font", esc(font)))
        if szs: rows.append(("Dimensioni", " / ".join(f"{s:g} pt" for s in szs)))
        rows.append(("Allineamento", "centrato"))
        if key == "header":
            rows.append(("Pagina", f"A4; margine superiore {top_cm:g} cm"))
        else:
            rows.append(("Pagina", f"margine inferiore {bot_cm:g} cm"))
    else:
        info = agg.get(key)
        if info:
            rows.append(("Provenienza", "Generato dallo script"))
            rows.append(("Font", esc(info["font"])))
            rows.append(("Dimensione", f"{info['size']:g} pt"))
            rows.append(("Allineamento", AL_IT[info["al"]]))
            if info["li"]:
                rows.append(("Rientro sx", f"{info['li']:g} cm"))
            rows.append(("Grassetto", "sì" if info["anybold"] else "no"))
            rows.append(("Corsivo", "sì" if info["italic"] else "no"))
            rows.append(("Spaziatura", f"prima {info['sb']:g} pt / dopo {info['sa']:g} pt"))
            rows.append(("Anti-orfano", "keep-with-next" if info["kwn"] else "no"))
    dl = "".join(f"<dt>{k}</dt><dd>{v}</dd>" for k, v in rows)
    note = f'<div class="note">{NOTE.get(key,"")}</div>' if NOTE.get(key) else ""
    return f"<dl>{dl}</dl>{note}"

def summary_for(key):
    if key in ("header", "footer"):
        return "Ereditato dal template (Garamond). Non generato dal corpo."
    info = agg.get(key)
    if not info:
        return ""
    bits = [AL_IT[info["al"]], f"{info['size']:g} pt"]
    if info["anybold"]: bits.append("grassetto")
    if info["italic"]: bits.append("corsivo")
    if info["li"]: bits.append(f"rientro {info['li']:g} cm")
    return ", ".join(bits) + "."

legend_html = []
present = set(k for k, _ in paras) | {"header", "footer"}
for key in ORDER:
    if key not in present:
        continue
    tag = f' <span class="tag">{TAG[key]}</span>' if key in TAG else ""
    legend_html.append(f'''    <div class="item c-{key}" data-k="{key}" tabindex="0">
      <span class="swatch" aria-hidden="true"></span>
      <div class="meta">
        <div class="label">{LABEL[key]}{tag}</div>
        <div class="sum">{summary_for(key)}</div>
        <div class="details">{details_for(key)}</div>
      </div>
    </div>''')
legend_body = "\n".join(legend_html)

color_css = "\n  ".join(f".c-{k}{{--c:{v}}}" for k, v in COLORS.items())

# --------------------------------------------------------------------------
# 6) Emit HTML
# --------------------------------------------------------------------------
PAGE = f"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Revisione formattazione lettere — Bergamo Legal (lettera fittizia)</title>
<style>
  :root{{--ink:#1c2430;--muted:#5b6675;--line:#d8dde5;--bg:#f4f6f9;--paper:#fff}}
  *{{box-sizing:border-box}} html,body{{margin:0}}
  body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:var(--ink);background:var(--bg);line-height:1.45}}
  {color_css}
  .topbar{{padding:18px 24px;border-bottom:1px solid var(--line);background:#fff}}
  .topbar h1{{font-size:18px;margin:0 0 4px}}
  .topbar p{{margin:0;color:var(--muted);font-size:13px;max-width:74ch}}
  .hint{{margin-top:8px;font-size:12px;color:var(--muted)}}
  .badges{{margin-top:8px;display:flex;flex-wrap:wrap;gap:6px}}
  .badge{{font-size:11px;padding:2px 8px;border-radius:999px;border:1px solid var(--line);background:#fff;color:var(--muted)}}
  .layout{{display:grid;grid-template-columns:minmax(300px,360px) 1fr}}
  .legend{{position:sticky;top:0;max-height:100vh;overflow:auto;padding:16px;border-right:1px solid var(--line);background:#fff}}
  .legend h2{{font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin:4px 0 10px}}
  .paper-wrap{{padding:28px;display:flex;justify-content:center}}
  .item{{display:flex;gap:10px;padding:9px 8px;border-radius:8px;outline:none;border:1px solid transparent}}
  .item+.item{{margin-top:2px}}
  .item:hover,.item:focus,.item.active{{background:rgba(var(--c),.10);border-color:rgba(var(--c),.45)}}
  .item:focus{{box-shadow:0 0 0 2px rgba(var(--c),.55)}}
  .swatch{{flex:0 0 14px;height:14px;margin-top:3px;border-radius:4px;background:rgb(var(--c));border:1px solid rgba(0,0,0,.15)}}
  .meta{{min-width:0}} .label{{font-weight:600;font-size:13.5px}}
  .label .tag{{font-weight:500;font-size:10.5px;color:var(--muted);margin-left:6px}}
  .sum{{font-size:12px;color:var(--muted);margin-top:1px}}
  .details{{display:none;margin-top:8px;border-top:1px dashed rgba(var(--c),.5);padding-top:8px;font-size:12px}}
  .item:hover .details,.item:focus .details,.item:focus-within .details,.item.active .details{{display:block}}
  .details dl{{display:grid;grid-template-columns:auto 1fr;gap:2px 10px;margin:0}}
  .details dt{{color:var(--muted)}} .details dd{{margin:0}}
  .details .note{{margin-top:6px;color:#475569}}
  .paper{{width:{px_cm(21):g}px;max-width:100%;background:var(--paper);border:1px solid var(--line);box-shadow:0 2px 14px rgba(20,30,45,.08)}}
  .tpl-header{{text-align:center;padding:22px 48px 14px;border-bottom:1px solid #eef0f3;font-family:"{esc(hdr_font)}",Garamond,"Times New Roman",serif}}
  .tpl-tag{{display:inline-block;font-family:-apple-system,Segoe UI,sans-serif;font-size:10px;color:#64748b;background:#f1f5f9;border:1px solid #e2e8f0;border-radius:999px;padding:1px 8px;margin-bottom:6px}}
  .tpl-header .firm{{font-size:22px;letter-spacing:.04em}}
  .tpl-footer{{margin-top:14px;border-top:1px solid #eef0f3;padding:12px 48px 20px;text-align:center;font-family:"{esc(ftr_font)}",Garamond,"Times New Roman",serif;color:#64748b;font-size:11.5px;line-height:1.5}}
  .letter{{font-family:"Times New Roman",Georgia,serif;color:#10151c;padding:20px {px_cm(lr_cm):g}px 26px}}
  .seg{{display:block;border-radius:3px;padding:1px 3px;outline:none;background:rgba(var(--c),.14)}}
  .seg:hover,.seg:focus,.seg.hl{{background:rgba(var(--c),.30);box-shadow:0 0 0 2px rgba(var(--c),.55)}}
  @media (max-width:880px){{.layout{{grid-template-columns:1fr}} .legend{{position:static;max-height:none;border-right:none;border-bottom:1px solid var(--line)}}}}
  @media print{{body{{background:#fff}} .topbar .hint{{display:none}} .layout{{display:block}} .legend{{position:static;max-height:none;border:none}} .details{{display:block !important}} .paper{{box-shadow:none;border:none}} .seg{{background:rgba(var(--c),.18) !important;box-shadow:none !important}}}}
</style>
</head>
<body>
<header class="topbar">
  <h1>Revisione formattazione — Lettere Bergamo Legal</h1>
  <p>Anteprima di una <strong>lettera fittizia</strong> (contenuto inventato).
     <strong>La lettera è stata generata dal formatter</strong> (<code>render_letter</code>)
     e poi riversata qui: le proprietà mostrate (allineamento, dimensione, spaziatura,
     rientro, grassetto/corsivo, keep-with-next) sono <em>lette dal documento prodotto</em>,
     non riscritte a mano. Passa il mouse o porta il focus con la tastiera su una voce
     della legenda o su una sezione della lettera per i dettagli.</p>
  <div class="badges" aria-hidden="true">
    <span class="badge">Ereditato dal template</span>
    <span class="badge">Generato dallo script</span>
    <span class="badge">Punto sensibile</span>
    <span class="badge">Da approvare</span>
  </div>
  <p class="hint">In stampa/PDF i dettagli sono espansi. Documento sorgente: <code>out/html/lettera_fittizia.docx</code>.</p>
</header>
<div class="layout">
  <aside class="legend" aria-label="Legenda delle sezioni di formattazione">
    <h2>Sezioni della lettera</h2>
{legend_body}
  </aside>
  <main class="paper-wrap">
    <div class="paper" aria-label="Anteprima lettera fittizia generata dallo script">
      <div class="seg c-header" data-k="header" tabindex="0">
        <div class="tpl-header"><div class="tpl-tag">Ereditato dal template</div>
          <div class="firm">{hdr}</div></div>
      </div>
      <div class="letter">
        {letter_body}
      </div>
      <div class="seg c-footer" data-k="footer" tabindex="0">
        <div class="tpl-footer"><div class="tpl-tag">Ereditato dal template</div>{ftr}</div>
      </div>
    </div>
  </main>
</div>
<script>
  (function(){{
    function all(s){{return Array.prototype.slice.call(document.querySelectorAll(s));}}
    function setActive(k,on){{
      all('.item[data-k="'+k+'"]').forEach(function(e){{e.classList.toggle('active',on);}});
      all('.seg[data-k="'+k+'"]').forEach(function(e){{e.classList.toggle('hl',on);}});
    }}
    function wire(el){{var k=el.getAttribute('data-k');
      ['mouseenter','focus'].forEach(function(ev){{el.addEventListener(ev,function(){{setActive(k,true);}});}});
      ['mouseleave','blur'].forEach(function(ev){{el.addEventListener(ev,function(){{setActive(k,false);}});}});
    }}
    all('.item').forEach(wire); all('.seg').forEach(wire);
  }})();
</script>
</body>
</html>
"""

os.makedirs(OUTDIR, exist_ok=True)
with open(HTMLF, "w", encoding="utf-8") as f:
    f.write(PAGE)
print(f"html:   {HTMLF}")
print("sezioni in legenda:", [k for k in ORDER if k in present])
