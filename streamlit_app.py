# app.py — Zoom → Customer Success Tab (.docx) with ZERO runtime installs
# - Upload transcript (.txt/.vtt/.srt/.docx/.pdf)
# - Optional: upload previous CS-Tab .docx to preserve history (we parse DOCX XML directly)
# - LLM (OpenAI) if key provided; else offline heuristic
# - Recompute word-limited summaries from FULL history
# - Output a .docx using a tiny internal DOCX builder (no python-docx)

import os, io, re, json, zipfile, html, xml.etree.ElementTree as ET
from datetime import datetime
from typing import Dict, List, Tuple, Optional

import streamlit as st

# ---------------------------
# App layout
# ---------------------------
st.set_page_config(page_title="Zoom → Customer Success Tab (Word)", layout="wide")
st.title("Zoom → Customer Success Tab (Word)")
st.caption("Upload a Zoom transcript (+ optional prior CS-Tab .docx). Get a refreshed .docx with new history entries and updated summaries. No external installs.")

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

# ---------------------------
# Helpers: read transcript files
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
    """Return plain text paragraphs from a .docx using only stdlib."""
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
            # handle line breaks inside runs
            for br in p.findall(".//w:br", ns):
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
    """Very light PDF text grabber: pulls /Contents streams and strips.
       Not perfect, but works for texty PDFs without external libs."""
    try:
        txt = raw.decode("latin-1", errors="ignore")
        # naive extract between BT/ET text operators
        chunks = re.findall(r"BT(.*?)ET", txt, flags=re.S)
        out = []
        for c in chunks:
            # pull text in parentheses, remove Tj/TJ ops
            pieces = re.findall(r"\((.*?)\)", c, flags=re.S)
            line = " ".join(p.replace("\\)", ")").replace("\\(", "(").replace("\\n", " ") for p in pieces)
            out.append(line)
        return "\n".join(out)
    except Exception:
        return ""

def read_uploaded_text(file) -> str:
    name = (file.name or "").lower()
    raw = file.read()
    file.seek(0)
    if name.endswith(".txt"): return read_txt_bytes(raw)
    if name.endswith(".vtt") or name.endswith(".srt"): return read_vtt_srt(raw)
    if name.endswith(".docx"): return read_docx_as_text(raw)
    if name.endswith(".pdf"): return read_pdf_as_text(raw)
    return read_txt_bytes(raw)

# ---------------------------
# Parse existing CS-Tab .docx (produced by this app or similar)
# ---------------------------
def parse_existing_cstab_docx(file) -> Dict[str, Dict[str, str]]:
    result = {k: {"sentiment": "", "summary": "", "history_text": ""} for k, _ in SECTION_ORDER}
    try:
        raw = file.read()
        file.seek(0)
        paras = extract_docx_paragraphs(raw)
        # Walk paragraphs: headings match section labels; then labeled blocks
        key, field = None, None
        labels = {label: k for k, label in SECTION_ORDER}
        for p in paras:
            t = p.strip()
            if t in labels:
                key, field = labels[t], None
                continue
            if t == "Sentiment":
                field = "sentiment";  continue
            if t == "Sentiment Summary":
                field = "summary";    continue
            if t == "Sentiment History":
                field = "history_text"; continue
            if key and field is not None:
                prev = result[key].get(field, "")
                result[key][field] = (prev + ("\n" if prev else "") + t).strip()
        return result
    except Exception:
        return result

# ---------------------------
# Offline heuristic extractor & summarizer
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
        entry = " ".join(buckets[k][:3]).strip()  # 1–3 sentences
        out[k] = {"sentiment": (senti or "NA"), "new_history_entry": entry}
    return out

def heuristic_summary(full_history: str, word_limit: int) -> str:
    if not full_history.strip(): return "NA"
    sents = re.split(r'(?<=[.!?])\s+', full_history.strip())
    text = " ".join(sents[:3]).strip()
    words = text.split()
    return " ".join(words[:word_limit]) if words else "NA"

# ---------------------------
# Optional OpenAI (no dependency at import)
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
        out[key] = {
            "sentiment": (sec.get("sentiment") or "NA").strip(),
            "new_history_entry": (sec.get("new_history_entry") or "").strip()
        }
    return out

def llm_summary_from_history(client, history: str, word_limit: int, model="gpt-4o-mini") -> str:
    if not history.strip(): return "NA"
    prompt = f"""Write a single-paragraph 'Sentiment Summary' (≤ {word_limit} words) based on the entire history below (newest first).
Be specific, concise, and reflect current state/trends/next steps. If trivial, return 'NA'.
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
    words = text.split()
    return " ".join(words[:word_limit]) if words else "NA"

# ---------------------------
# Tiny DOCX builder (no deps)
# ---------------------------
# Minimal styles + document boilerplate
CONTENT_TYPES_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""
RELS_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
DOC_RELS_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>
"""
STYLES_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:after="200"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="40"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/><w:qFormat/>
    <w:pPr><w:spacing w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/><w:qFormat/>
    <w:pPr><w:spacing w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
</w:styles>
"""

def p(style: Optional[str], text: str, bold=False) -> str:
    """Create a paragraph with optional style and bold run."""
    text = html.escape(text).replace("\n", "<w:br/>")
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

def build_docx_bytes(title: str, assembled: Dict[str, Dict[str, str]]) -> bytes:
    body_parts = [p("Title", title)]
    for key, label in SECTION_ORDER:
        s = assembled.get(key, {})
        body_parts.append(p("Heading2", label))
        body_parts.append(p(None, "Sentiment", bold=True))
        body_parts.append(p(None, (s.get("sentiment") or "NA")))
        body_parts.append(p(None, "Sentiment Summary", bold=True))
        body_parts.append(p(None, (s.get("summary") or "NA")))
        body_parts.append(p(None, "Sentiment History", bold=True))
        hist = s.get("history_text","")
        if hist.strip():
            for line in hist.splitlines():
                body_parts.append(p(None, line))
        else:
            body_parts.append(p(None, ""))

    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {''.join(body_parts)}
    <w:sectPr></w:sectPr>
  </w:body>
</w:document>'''.encode("utf-8")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        z.writestr("_rels/.rels", RELS_XML)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS_XML)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/document.xml", document_xml)
    buf.seek(0)
    return buf.getvalue()

# ---------------------------
# UI controls
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
if openai_key:
    st.session_state["openai_key"] = openai_key

st.markdown("**Upload Zoom transcript/summary**")
f_transcript = st.file_uploader("Transcript (.txt, .vtt, .srt, .docx, .pdf)", type=["txt","vtt","srt","docx","pdf"])

st.markdown("**Upload previous CS-Tab (.docx) — optional (preserves history)**")
f_prev = st.file_uploader("Previous CS-Tab (.docx)", type=["docx"])

go = st.button("Generate Word document", type="primary", disabled=not f_transcript)

# ---------------------------
# Main action
# ---------------------------
if go:
    transcript_text = read_uploaded_text(f_transcript)

    # Parse previous doc to preserve history/sentiment if present
    prev = {k: {"sentiment":"", "summary":"", "history_text":""} for k,_ in SECTION_ORDER}
    if f_prev:
        prev = parse_existing_cstab_docx(f_prev)

    client = get_openai_client()

    # 1) Extract NEW per-section history + sentiment
    if client:
        new_map = llm_extract_new_entries(client, transcript_text, model=model_name)
    else:
        st.info("No OpenAI key supplied — using offline heuristic.")
        new_map = heuristic_extract_new_entries(transcript_text)

    # 2) Assemble histories and recompute summaries
    assembled: Dict[str, Dict[str, str]] = {}
    for key, label in SECTION_ORDER:
        prev_hist = (prev.get(key, {}) or {}).get("history_text","").strip()
        new_entry = (new_map.get(key, {}).get("new_history_entry") or "").strip()

        parts = []
        if new_entry:
            parts.append(meet_date)
            parts.append(new_entry)
        if prev_hist:
            parts.append(prev_hist)
        full_history = "\n".join(parts).strip()

        limit = word_limits.get(key, default_limit)
        if client:
            summary_text = llm_summary_from_history(client, full_history, limit, model=model_name)
        else:
            summary_text = heuristic_summary(full_history, limit)

        new_s = (new_map.get(key, {}).get("sentiment") or "NA")
        if (not new_s or new_s == "NA") and prev.get(key,{}).get("sentiment","").strip():
            new_s = prev[key]["sentiment"].strip()

        assembled[key] = {
            "sentiment": new_s if new_s else "NA",
            "summary": summary_text if summary_text else "NA",
            "history_text": full_history
        }

    # 3) Build .docx
    title = f"Customer Success Tab — {account}" if account else "Customer Success Tab"
    doc_bytes = build_docx_bytes(title, assembled)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    fname = f"CS_Tab_{account or 'Account'}_{stamp}.docx"

    st.success("Document generated.")
    st.download_button("Download Word document", data=doc_bytes, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    # 4) On-page preview
    st.divider()
    st.subheader("Preview")
    for key, label in SECTION_ORDER:
        s = assembled[key]
        st.markdown(f"### {label}")
        st.markdown(f"**Sentiment:** {s['sentiment']}")
        st.markdown(f"**Sentiment Summary** (≤ {word_limits.get(key, default_limit)} words)")
        st.write(s["summary"])
        st.markdown("**Sentiment History** (newest first)")
        st.text(s["history_text"])
