# app.py — Zoom → Customer Success Tab (.docx)
# Single-file Streamlit app. No runtime pip installs.
# - Upload transcript (.txt/.vtt/.srt/.docx/.pdf) + optional prior CS-Tab .docx
# - Extract per-section {sentiment, new_history_entry}, append-only history, recompute summary <= N words
# - Relationship sentiment mirrors Overall; Consumption summary fallback
# - Output beautiful .docx with true Word bullets + bold-italics section headings
# - In-app preview + download
#
# Extension point (later): add a sidebar panel to append-only update Salesforce fields
# (never overwrite long-text fields; upsert tasks via stable hashes).

import os, io, re, json, zipfile, html, xml.etree.ElementTree as ET
from datetime import datetime
from typing import Dict, List, Optional

import streamlit as st

# ---------------------------
# Streamlit shell
# ---------------------------
st.set_page_config(page_title="Zoom → Customer Success Tab (Word)", layout="wide")
st.title("Zoom → Customer Success Tab (Word)")
st.caption("Upload Zoom transcript (+ optional prior CS-Tab .docx). Appends new history, recomputes summaries, and outputs a polished Word doc with proper bullets. No runtime installs.")

# ---------------------------
# Sections & defaults
# ---------------------------
SECTION_ORDER = [
    ("overall", "Overall Customer Sentiment"),
    ("relationship", "Relationship Sentiment"),
    ("consumption", "Consumption Sentiment"),
    ("prod_eng", "Prod & Eng Sentiment"),
    ("network", "Network Sentiment"),
    ("support", "Support Sentiment"),
    ("implementation", "Implementation Sentiment"),
]
DEFAULT_WORD_LIMIT = 50
MONTH_RX = r"^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}$"

# ---------------------------
# File readers (stdlib only)
# ---------------------------
def read_txt_bytes(raw: bytes) -> str:
    return raw.decode("utf-8", errors="ignore")

def read_vtt_srt(raw: bytes) -> str:
    text = raw.decode("utf-8", errors="ignore")
    lines = []
    for ln in text.splitlines():
        if re.match(r"^\d+$", ln.strip()):
            continue
        if re.search(r"\d{2}:\d{2}:\d{2}", ln) or re.search(r"\d{2}:\d{2}\.\d{3}", ln):
            continue
        lines.append(ln.strip())
    return "\n".join([l for l in lines if l.strip()])

def extract_docx_paragraphs(raw: bytes) -> List[str]:
    """Parse .docx with stdlib zipfile + WordprocessingML."""
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            xml_bytes = z.read("word/document.xml")
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(xml_bytes)
        paragraphs = []
        for p in root.findall(".//w:body/w:p", ns):
            texts = []
            for t in p.findall(".//w:t", ns):
                texts.append(t.text or "")
            for _ in p.findall(".//w:br", ns):
                texts.append("\n")
            para = "".join(texts).replace("\xa0", " ").strip()
            paragraphs.append(para)
        return paragraphs
    except Exception:
        return []

def read_docx_as_text(raw: bytes) -> str:
    paras = extract_docx_paragraphs(raw)
    return "\n".join(paras)

def read_pdf_as_text(raw: bytes) -> str:
    """Very light best-effort PDF text grab (no external libs)."""
    try:
        txt = raw.decode("latin-1", errors="ignore")
        chunks = re.findall(r"BT(.*?)ET", txt, flags=re.S)
        out = []
        for c in chunks:
            pieces = re.findall(r"\((.*?)\)", c, flags=re.S)
            line = " ".join(p.replace("\\)", ")").replace("\\(", "(").replace("\\n", " ") for p in pieces)
            out.append(line)
        return "\n".join(out)
    except Exception:
        return ""

def read_uploaded_text(file) -> str:
    """Reads .txt, .vtt, .srt, .docx, .pdf using only stdlib."""
    name = (file.name or "").lower()
    raw = file.read(); file.seek(0)
    if name.endswith(".txt"):  return read_txt_bytes(raw)
    if name.endswith(".vtt") or name.endswith(".srt"): return read_vtt_srt(raw)
    if name.endswith(".docx"): return read_docx_as_text(raw)
    if name.endswith(".pdf"):  return read_pdf_as_text(raw)
    return read_txt_bytes(raw)

# ---------------------------
# Parse existing CS-Tab .docx (append-only preservation)
# ---------------------------
def parse_existing_cstab_docx(file) -> Dict[str, Dict[str, str]]:
    """
    Recover section blocks from a previous doc produced by this app:
      For each section returns dict with {sentiment, summary, history_text}
    """
    result = {k: {"sentiment": "", "summary": "", "history_text": ""} for k, _ in SECTION_ORDER}
    try:
        raw = file.read(); file.seek(0)
        paras = extract_docx_paragraphs(raw)
        key, field = None, None
        labels = {label: k for k, label in SECTION_ORDER}
        for p in paras:
            t = p.strip()
            if t in labels:
                key, field = labels[t], None
                continue
            if t == "Sentiment":          field = "sentiment";      continue
            if t == "Sentiment Summary":  field = "summary";        continue
            if t == "Sentiment History":  field = "history_text";   continue
            if key and field is not None:
                prev = result[key].get(field, "")
                result[key][field] = (prev + ("\n" if prev else "") + t).strip()
        return result
    except Exception:
        return result

# ---------------------------
# Heuristic extraction & summarization (offline fallback)
# ---------------------------
POSITIVE_KWS = ["happy","satisfied","good","improving","stable","resolved","green","renewed","auto-renew"]
NEGATIVE_KWS = ["blocked","delay","issue","problem","risk","concern","bad","red","escalat","degrad","churn"]
YELLOW_KWS  = ["working on","needs to","pending","investigat","monitor","follow up","gap","accuracy","partial","improve"]

SECTION_HINTS = {
    "overall":       ["overall","exec","business","roi","renew","escalation","account","sentiment"],
    "relationship":  ["relationship","stakeholder","sponsor","champion","engage","trust","communication","buy-in"],
    "consumption":   ["consumption","usage","adoption","tracked","milestone completeness","coverage","license","utilization"],
    "prod_eng":      ["engineering","product","bug","feature","roadmap","api","accuracy","latency","release"],
    "network":       ["network","carrier","forwarder","integration","connectivity","api access","expeditor","ffw","partner"],
    "support":       ["support","ticket","sla","incident","case","csat","helpdesk"],
    "implementation":["implement","onboard","go live","integration","blocked","project","timeline","cutover","uat"],
}

def simple_sentiment(text: str) -> str:
    t = text.lower()
    pos = sum(t.count(k) for k in POSITIVE_KWS)
    neg = sum(t.count(k) for k in NEGATIVE_KWS)
    yel = sum(t.count(k) for k in YELLOW_KWS)
    if max(pos,neg,yel) == 0: return "NA"
    if neg > max(pos,yel): return "Red"
    if yel >= max(pos,neg): return "Yellow"
    return "Green"

def heuristic_extract_new_entries(transcript: str) -> Dict[str, Dict[str,str]]:
    """Extract up to 1–4 short sentences per section heuristically."""
    sents = re.split(r'(?<=[.!?])\s+', transcript.strip())
    buckets = {k: [] for k,_ in SECTION_ORDER}
    for s in sents:
        ls = s.lower()
        for k, hints in SECTION_HINTS.items():
            if any(h in ls for h in hints):
                buckets[k].append(s.strip())
    out = {}
    for k,_ in SECTION_ORDER:
        chunk = " ".join(buckets[k])[:1000]
        senti = simple_sentiment(chunk)
        entry_sents = []
        for s in buckets[k][:4]:
            s = s.strip()
            if not s: continue
            if not re.search(r'[.!?]$', s): s += '.'
            entry_sents.append(s)
        entry = " ".join(entry_sents)
        out[k] = {"sentiment": (senti or "NA"), "new_history_entry": entry}
    return out

def heuristic_summary(full_history: str, word_limit: int) -> str:
    if not full_history.strip(): return "NA"
    sents = re.split(r'(?<=[.!?])\s+', full_history.strip())
    norm = []
    for s in sents[:6]:
        s = s.strip()
        if not s: continue
        if not re.search(r'[.!?]$', s): s += '.'
        norm.append(s)
        if sum(len(x.split()) for x in norm) >= word_limit: break
    text = " ".join(norm)
    words = text.split()
    return " ".join(words[:word_limit]) if words else "NA"

def transcript_mentions_consumption(text: str) -> bool:
    t = text.lower()
    return any(k in t for k in SECTION_HINTS["consumption"])

# ---------------------------
# Optional OpenAI (no installs; safe fallback)
# ---------------------------
def get_openai_client():
    try:
        from openai import OpenAI
        key = st.session_state.get("openai_key") or os.getenv("OPENAI_API_KEY")
        if not key: return None
        os.environ["OPENAI_API_KEY"] = key
        return OpenAI()
    except Exception:
        return None

SYSTEM_INSTRUCTIONS = """You convert Zoom transcripts into CS-Tab updates.
For each of these exact sections:
- Overall Customer Sentiment
- Relationship Sentiment
- Consumption Sentiment
- Prod & Eng Sentiment
- Network Sentiment
- Support Sentiment
- Implementation Sentiment
Return JSON with for each section:
  - sentiment: Green|Yellow|Red|NA
  - new_history_entry: 1–4 concise sentences describing THIS meeting only.
If unclear: sentiment='NA', new_history_entry=''."""

def llm_extract_new_entries(client, transcript: str, model="gpt-4o-mini") -> Dict[str, Dict[str,str]]:
    labels = [label for _, label in SECTION_ORDER]
    prompt = f"""Sections:\n{json.dumps(labels)}\n\nTranscript:\n\"\"\"{transcript[:120000]}\"\"\"\n
Return JSON exactly as:
{{"sections": {{ "<label>": {{"sentiment":"Green|Yellow|Red|NA","new_history_entry":"..."}} }} }}
Use the exact labels. Keep entries concise. If nothing relevant, use sentiment='NA' and new_history_entry=''."""
    try:
        resp = client.chat.completions.create(
            model=model, temperature=0.2,
            messages=[{"role":"system","content":SYSTEM_INSTRUCTIONS},
                      {"role":"user","content":prompt}]
        )
        content = resp.choices[0].message.content or "{}"
        m = re.search(r"\{[\s\S]*\}", content)
        data = json.loads(m.group(0) if m else content)
    except Exception:
        data = {"sections": {}}
    out = {}
    for key, label in SECTION_ORDER:
        sec = data.get("sections", {}).get(label, {}) or {}
        # enforce punctuation in the new entry
        entry = (sec.get("new_history_entry") or "").strip()
        if entry and not re.search(r'[.!?]$', entry): entry += '.'
        out[key] = {
            "sentiment": (sec.get("sentiment") or "NA").strip(),
            "new_history_entry": entry
        }
    return out

def llm_summary_from_history(client, history: str, word_limit: int, model="gpt-4o-mini") -> str:
    if not history.strip(): return "NA"
    prompt = f"""Write a single-paragraph 'Sentiment Summary' (≤ {word_limit} words) based on the entire history below (newest first).
Be specific, concise, and reflect current state/trends/next steps. Ensure proper sentence punctuation. If trivial, return 'NA'.
History:
\"\"\"{history[:120000]}\"\"\""""
    try:
        resp = client.chat.completions.create(
            model=model, temperature=0.2,
            messages=[{"role":"user","content":prompt}]
        )
        text = (resp.choices[0].message.content or "").strip()
    except Exception:
        text = ""
    if not text: return "NA"
    # enforce punctuation and limit
    sents = re.split(r'(?<=[.!?])\s+', text)
    norm = []
    for s in sents:
        s = s.strip()
        if not s: continue
        if not re.search(r'[.!?]$', s): s += '.'
        norm.append(s)
    words = " ".join(norm).split()
    return " ".join(words[:word_limit]) if words else "NA"

# ---------------------------
# Word (.docx) builder (true bullets + bold-italics headings)
# ---------------------------
CONTENT_TYPES_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
"""
RELS_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
# Relationship to numbering (helps Google Docs/Word render bullets)
DOC_RELS_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
"""
# Bold + italics Heading2
STYLES_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/><w:qFormat/>
    <w:pPr><w:spacing w:after="200"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="40"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/><w:qFormat/>
    <w:pPr><w:spacing w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:i/><w:sz w:val="28"/></w:rPr>
  </w:style>
</w:styles>
"""
# Bullets (numId 1) — UTF-8 for "•"
NUMBERING_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>
      <w:pPr/><w:rPr><w:sz w:val="24"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
""".encode("utf-8")

def par(style: Optional[str], text: str, bold=False, italic=False) -> str:
    text = html.escape(text).replace("\n", "<w:br/>")
    b = "<w:b/>" if bold else ""
    i = "<w:i/>" if italic else ""
    rpr = f"<w:rPr>{b}{i}</w:rPr>" if (bold or italic) else ""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

def bullet_par(text: str, level: int = 0) -> str:
    text = html.escape(text).replace("\n", "<w:br/>")
    num = f'<w:pPr><w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="1"/></w:numPr></w:pPr>'
    return f'<w:p>{num}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

def ensure_sentence_end(s: str) -> str:
    s = s.strip()
    if not s: return s
    if not re.search(r'[.!?]$', s): s += '.'
    return s

def history_to_bullets(history_text: str) -> List[str]:
    """
    Convert raw history text to bullets.
    Each bullet:
      first line: 'Month YYYY'
      following lines: statements (ensure periods, preserve line breaks)
    """
    if not history_text.strip(): return []
    lines = [l.rstrip() for l in history_text.splitlines() if l.strip()]
    bullets, current_date, buf = [], None, []
    for ln in lines:
        if re.match(MONTH_RX, ln):
            if current_date or buf:
                statements = [ensure_sentence_end(x) for x in buf if x.strip()]
                bullets.append((current_date or "").strip() + ("\n" + "\n".join(statements) if statements else ""))
            current_date, buf = ln, []
        else:
            buf.append(ln)
    if current_date or buf:
        statements = [ensure_sentence_end(x) for x in buf if x.strip()]
        bullets.append((current_date or "").strip() + ("\n" + "\n".join(statements) if statements else ""))
    return bullets

def build_docx_bytes(title: str, assembled: Dict[str, Dict[str, str]]) -> bytes:
    """Compose a minimal Word docx with styles + numbering for bullets."""
    body = [par("Title", title, bold=True)]
    for key, label in SECTION_ORDER:
        s = assembled.get(key, {})
        body.append(par("Heading2", label, bold=True, italic=True))
        body.append(par(None, "Sentiment", bold=True))
        body.append(par(None, (s.get("sentiment") or "NA")))
        body.append(par(None, "Sentiment Summary", bold=True))
        body.append(par(None, (s.get("summary") or "NA")))
        body.append(par(None, "Sentiment History", bold=True))
        items = history_to_bullets(s.get("history_text",""))
        if items:
            for item in items:
                body.append(bullet_par(item))
        else:
            body.append(bullet_par(""))  # consistent look even if empty

    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {''.join(body)}
    <w:sectPr></w:sectPr>
  </w:body>
</w:document>'''.encode("utf-8")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        z.writestr("_rels/.rels", RELS_XML)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS_XML)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/numbering.xml", NUMBERING_XML)
        z.writestr("word/document.xml", document_xml)
    buf.seek(0)
    return buf.getvalue()

# ---------------------------
# UI Controls
# ---------------------------
left, right = st.columns([2,1])
with left:
    account = st.text_input("Account name (for document title)", value="")
    meet_date = st.text_input("History date label (e.g., 'June 2025')",
                              value=datetime.now().strftime("%B %Y"))
with right:
    model_name = st.text_input("OpenAI model (optional)", value="gpt-4o-mini")
    default_limit = st.number_input("Default summary word limit", 10, 200, DEFAULT_WORD_LIMIT, 5)

with st.expander("Per-section word limits (optional)"):
    word_limits: Dict[str, int] = {}
    for k, label in SECTION_ORDER:
        word_limits[k] = st.number_input(f"{label} (≤ words)", 10, 200, default_limit, 5, key=f"wl_{k}")

st.markdown("**LLM (optional)** — paste your key to improve extraction & summaries. Leave blank to use the offline heuristic.")
openai_key = st.text_input("OpenAI API Key", type="password", value=os.getenv("OPENAI_API_KEY",""))
if openai_key: st.session_state["openai_key"] = openai_key

st.markdown("**Upload Zoom transcript/summary**")
f_transcript = st.file_uploader("Transcript (.txt, .vtt, .srt, .docx, .pdf)", type=["txt","vtt","srt","docx","pdf"])

st.markdown("**Upload previous CS-Tab (.docx) — optional (preserves history)**")
f_prev = st.file_uploader("Previous CS-Tab (.docx)", type=["docx"])

go = st.button("Generate Word document", type="primary", disabled=not f_transcript)

# ---------------------------
# Main flow
# ---------------------------
if go:
    transcript_text = read_uploaded_text(f_transcript)

    # Recover previous state to preserve history
    prev = {k: {"sentiment":"", "summary":"", "history_text":""} for k,_ in SECTION_ORDER}
    if f_prev:
        prev = parse_existing_cstab_docx(f_prev)

    client = get_openai_client()
    # 1) Extract NEW entries + sentiments
    new_map = llm_extract_new_entries(client, transcript_text, model=model_name) if client else heuristic_extract_new_entries(transcript_text)

    # 2) Assemble full (newest-first) histories & summaries
    assembled: Dict[str, Dict[str, str]] = {}
    for key, label in SECTION_ORDER:
        prev_hist = (prev.get(key, {}) or {}).get("history_text","").strip()
        new_entry = (new_map.get(key, {}).get("new_history_entry") or "").strip()

        # Build newest-first history block
        parts = []
        if new_entry:
            parts.append(meet_date)
            parts.append(new_entry)
        if prev_hist:
            parts.append(prev_hist)
        full_history = "\n".join(parts).strip()

        # Summary
        limit = word_limits.get(key, DEFAULT_WORD_LIMIT)
        if client:
            summary_text = llm_summary_from_history(client, full_history, limit, model=model_name)
        else:
            summary_text = heuristic_summary(full_history, limit)

        # Consumption fallback: if transcript doesn't mention consumption at all, force summary string
        if key == "consumption" and not transcript_mentions_consumption(transcript_text):
            summary_text = "Consumption not discussed on call."

        # Sentiment: prefer new; if NA and previous exists, keep previous
        new_s = (new_map.get(key, {}).get("sentiment") or "NA")
        if (not new_s or new_s == "NA") and prev.get(key,{}).get("sentiment","").strip():
            new_s = prev[key]["sentiment"].strip()

        assembled[key] = {
            "sentiment": new_s if new_s else "NA",
            "summary": summary_text if summary_text else "NA",
            "history_text": full_history
        }

    # 3) Relationship sentiment mirrors Overall
    overall_s = (assembled.get("overall", {}).get("sentiment","") or "").lower()
    if overall_s:
        if "red" in overall_s:
            assembled["relationship"]["sentiment"] = "Red"
        elif "green" in overall_s:
            assembled["relationship"]["sentiment"] = "Green"
        else:
            assembled["relationship"]["sentiment"] = "Yellow"

    # 4) Build downloadable .docx
    title = f"Customer Success Tab — {account}" if account else "Customer Success Tab"
    doc_bytes = build_docx_bytes(title, assembled)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    fname = f"CS_Tab_{account or 'Account'}_{stamp}.docx"

    st.success("Document generated with bold-italics headings, true bullets, append-only histories, and summaries.")
    st.download_button("Download Word document", data=doc_bytes, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    # 5) Preview (Markdown bullets; preserve inner line breaks)
    st.divider()
    st.subheader("Preview")
    def bullets_for_preview(h: str) -> List[str]:
        items = history_to_bullets(h)
        return [i.replace("\n", "  \n  ") for i in items] or ["—"]
    for key, label in SECTION_ORDER:
        s = assembled[key]
        st.markdown(f"### ***{label}***")
        st.markdown(f"**Sentiment:** {s['sentiment']}")
        st.markdown(f"**Sentiment Summary** (≤ {word_limits.get(key, DEFAULT_WORD_LIMIT)} words)")
        st.write(s["summary"])
        st.markdown("**Sentiment History** (newest first)")
        for item in bullets_for_preview(s["history_text"]):
            st.markdown(f"- {item}")
