# app.py
# Zoom → Customer Success Tab (.docx)
# - Upload transcript (.txt/.vtt/.srt/.pdf/.docx)
# - (Optional) upload prior CS-Tab .docx to preserve history
# - Extract per-section NEW history + sentiment (LLM or offline heuristic)
# - Recompute word-limited summaries from FULL history
# - Fill "NA" if info missing
# - Output a .docx matching: Sentiment / Sentiment Summary / Sentiment History
# Everything runs from inside Streamlit.

import sys, subprocess, importlib.util
def ensure(pkg, import_name=None):
    """Install pkg via pip if not importable."""
    name = import_name or pkg
    if importlib.util.find_spec(name) is None:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg], stdout=subprocess.DEVNULL)

# Minimal deps (auto-installs if missing)
ensure("streamlit")
ensure("pydantic")
ensure("python-docx")
# Optional parsers
ensure("pypdf")
# LLM (optional)
ensure("openai")

import os, io, re, json
from datetime import datetime
from typing import Dict, Optional, Tuple

import streamlit as st
from pydantic import BaseModel, Field
from docx import Document
from pypdf import PdfReader

# ---------------------------
# App layout
# ---------------------------
st.set_page_config(page_title="Zoom → Customer Success Tab (Word)", layout="wide")
st.title("Zoom → Customer Success Tab (Word)")
st.caption("Upload a Zoom transcript (and optional prior CS-Tab .docx). Get a refreshed .docx with new history entries and updated summaries. No overwrites—history is preserved.")

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
# Models
# ---------------------------
class SectionResult(BaseModel):
    sentiment: str = Field(default="NA", description="Green|Yellow|Red|NA")
    summary: str = Field(default="NA")
    new_history_entry: str = Field(default="")

# ---------------------------
# Read transcript text
# ---------------------------
def read_uploaded_text(file) -> str:
    name = (file.name or "").lower()
    raw = file.read()
    file.seek(0)

    if name.endswith(".txt"):
        return raw.decode("utf-8", errors="ignore")

    if name.endswith(".vtt") or name.endswith(".srt"):
        text = raw.decode("utf-8", errors="ignore")
        # Strip counters/timestamps
        lines = []
        for ln in text.splitlines():
            if re.match(r"^\d+$", ln.strip()):
                continue
            if re.search(r"\d{2}:\d{2}:\d{2}", ln) or re.search(r"\d{2}:\d{2}\.\d{3}", ln):
                continue
            lines.append(ln.strip())
        return "\n".join([l for l in lines if l.strip()])

    if name.endswith(".docx"):
        try:
            doc = Document(io.BytesIO(raw))
            return "\n".join(p.text for p in doc.paragraphs)
        except Exception:
            pass

    if name.endswith(".pdf"):
        out = []
        try:
            reader = PdfReader(io.BytesIO(raw))
            for pg in reader.pages:
                try:
                    out.append(pg.extract_text() or "")
                except Exception:
                    pass
        except Exception:
            pass
        return "\n".join(out)

    # Fallback decode
    return raw.decode("utf-8", errors="ignore")

# ---------------------------
# Parse existing CS-Tab .docx created by this app
# ---------------------------
def parse_existing_cstab_docx(file) -> Dict[str, Dict[str, str]]:
    result = {k: {"sentiment": "", "summary": "", "history_text": ""} for k, _ in SECTION_ORDER}
    try:
        doc = Document(file)
    except Exception:
        return result

    key = None
    field = None
    def norm(s): return re.sub(r"\s+", " ", (s or "").strip())

    for p in doc.paragraphs:
        t = norm(p.text)
        # Section heading?
        matched = False
        for k, label in SECTION_ORDER:
            if t == label:
                key, field, matched = k, None, True
                break
        if matched:
            continue

        if t == "Sentiment":
            field = "sentiment";  continue
        if t == "Sentiment Summary":
            field = "summary";    continue
        if t == "Sentiment History":
            field = "history_text"; continue

        if key and field is not None:
            prev = result[key].get(field, "")
            result[key][field] = (prev + ("\n" if prev else "") + p.text).strip()

    return result

# ---------------------------
# Heuristic (offline) extractor: works without API key
# ---------------------------
POSITIVE_KWS = ["happy","satisfied","good","improving","stable","resolved","green"]
NEGATIVE_KWS = ["blocked","delay","issue","problem","risk","concern","bad","red","escalat","degrad"]
YELLOW_KWS  = ["working on","needs to","pending","investigat","monitor","follow up","gap","accuracy","partial"]

def simple_sentiment(text: str) -> str:
    t = text.lower()
    pos = sum(t.count(k) for k in POSITIVE_KWS)
    neg = sum(t.count(k) for k in NEGATIVE_KWS)
    yel = sum(t.count(k) for k in YELLOW_KWS)
    if max(pos,neg,yel) == 0:
        return "NA"
    if neg > max(pos, yel): return "Red"
    if yel >= max(pos, neg): return "Yellow"
    return "Green"

SECTION_HINTS = {
    "overall":       ["overall","exec","business","roi","renew","escalation","account"],
    "relationship":  ["relationship","stakeholder","sponsor","champion","engage","trust","communication"],
    "consumption":   ["consumption","usage","adoption","tracked","milestone completeness","coverage","license"],
    "prod_eng":      ["engineering","product","bug","feature","roadmap","api","accuracy","latency"],
    "network":       ["network","carrier","forwarder","etl","integration","connectivity","api access"],
    "support":       ["support","ticket","sla","incident","case","helpdesk"],
    "implementation":["implement","onboard","go live","integration","blocked","project","timeline"],
}

def heuristic_extract_new_entries(transcript: str) -> Dict[str, SectionResult]:
    # Break transcript into sentences
    sents = re.split(r'(?<=[.!?])\s+', transcript.strip())
    buckets = {k: [] for k,_ in SECTION_ORDER}

    for s in sents:
        ls = s.lower()
        for k, hints in SECTION_HINTS.items():
            if any(h in ls for h in hints):
                buckets[k].append(s.strip())
    out = {}
    for k,_ in SECTION_ORDER:
        chunk = " ".join(buckets[k])[:800]
        senti = simple_sentiment(chunk)
        entry = " ".join(buckets[k][:3]).strip()  # 1–3 sentences
        out[k] = SectionResult(
            sentiment=senti if senti else "NA",
            summary="NA",
            new_history_entry=entry
        )
    return out

def heuristic_summary(full_history: str, word_limit: int) -> str:
    if not full_history.strip():
        return "NA"
    # naive: take first sentences and trim to limit
    sents = re.split(r'(?<=[.!?])\s+', full_history.strip())
    text = " ".join(sents[:3]).strip()
    words = text.split()
    return " ".join(words[:word_limit]) if words else "NA"

# ---------------------------
# LLM helpers (optional)
# ---------------------------
def get_openai_client():
    try:
        from openai import OpenAI
        # Prefer secrets if provided
        key = st.session_state.get("openai_key") or os.getenv("OPENAI_API_KEY")
        if not key:
            return None
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

def llm_extract_new_entries(client, transcript: str, model="gpt-4o-mini") -> Dict[str, SectionResult]:
    labels = [label for _, label in SECTION_ORDER]
    prompt = f"""Sections:\n{json.dumps(labels)}\n\nTranscript:\n\"\"\"{transcript[:120000]}\"\"\"\n
Return JSON exactly as:
{{"sections": {{ "<label>": {{"sentiment":"Green|Yellow|Red|NA","new_history_entry":"..."}} }} }}
Use the exact labels above. Keep entries concise. If nothing relevant, use sentiment='NA' and new_history_entry=''."""
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
        out[key] = SectionResult(
            sentiment=(sec.get("sentiment") or "NA").strip(),
            summary="NA",
            new_history_entry=(sec.get("new_history_entry") or "").strip()
        )
    return out

def llm_summary_from_history(client, history: str, word_limit: int, model="gpt-4o-mini") -> str:
    if not history.strip():
        return "NA"
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
    if not text:
        return "NA"
    words = text.split()
    return " ".join(words[:word_limit]) if words else "NA"

# ---------------------------
# Word builder
# ---------------------------
def add_label(doc: Document, text: str):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True

def write_section(doc: Document, label: str, sentiment: str, summary: str, history_text: str):
    doc.add_heading(label, level=2)
    add_label(doc, "Sentiment")
    doc.add_paragraph(sentiment or "NA")
    add_label(doc, "Sentiment Summary")
    doc.add_paragraph(summary or "NA")
    add_label(doc, "Sentiment History")
    if history_text.strip():
        for line in history_text.splitlines():
            doc.add_paragraph(line)
    else:
        doc.add_paragraph("")

def build_docx(account_name: str, assembled: Dict[str, Dict[str, str]]) -> bytes:
    doc = Document()
    title = f"Customer Success Tab — {account_name}" if account_name else "Customer Success Tab"
    doc.add_heading(title, level=1)
    for key, label in SECTION_ORDER:
        s = assembled.get(key, {})
        write_section(
            doc, label,
            s.get("sentiment","NA"),
            s.get("summary","NA"),
            s.get("history_text","").strip()
        )
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
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
    st.markdown("**Model & Limits**")
    model_name = st.text_input("OpenAI model", value="gpt-4o-mini")
    default_limit = st.number_input("Default summary word limit", 10, 200, DEFAULT_WORD_LIMIT, 5)

with st.expander("Per-section word limits (optional)"):
    word_limits: Dict[str, int] = {}
    for k, label in SECTION_ORDER:
        word_limits[k] = st.number_input(
            f"{label} (≤ words)", 10, 200, default_limit, 5, key=f"wl_{k}"
        )

st.markdown("**LLM (optional)**")
st.write("Use an API key here (preferred) or skip to run the offline heuristic.")
openai_key = st.text_input("OpenAI API Key", type="password", value=os.getenv("OPENAI_API_KEY",""))
if openai_key:
    st.session_state["openai_key"] = openai_key

st.markdown("**Upload Zoom transcript/summary**")
f_transcript = st.file_uploader("Transcript (.txt, .vtt, .srt, .pdf, .docx)", type=["txt","vtt","srt","pdf","docx"])

st.markdown("**Upload previous CS-Tab .docx (optional)**")
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

    # 1) Extract NEW history entries + sentiment
    if client:
        new_map = llm_extract_new_entries(client, transcript_text, model=model_name)
    else:
        st.info("No OpenAI key: using offline heuristic extractor.")
        new_map = heuristic_extract_new_entries(transcript_text)

    # 2) Assemble histories and recompute summaries
    assembled: Dict[str, Dict[str, str]] = {}
    for key, label in SECTION_ORDER:
        prev_hist = (prev.get(key, {}) or {}).get("history_text","").strip()
        new_entry = (new_map[key].new_history_entry or "").strip()

        # Build newest-first history
        parts = []
        if new_entry:
            parts.append(meet_date)
            parts.append(new_entry)
        if prev_hist:
            parts.append(prev_hist)
        full_history = "\n".join(parts).strip()

        # Summary
        limit = word_limits.get(key, default_limit)
        if client:
            summary_text = llm_summary_from_history(client, full_history, limit, model=model_name)
        else:
            summary_text = heuristic_summary(full_history, limit)

        # Sentiment: prefer new; if NA and previous exists, keep previous
        new_s = new_map[key].sentiment or "NA"
        if (not new_s or new_s == "NA") and prev.get(key,{}).get("sentiment","").strip():
            new_s = prev[key]["sentiment"].strip()

        assembled[key] = {
            "sentiment": new_s if new_s else "NA",
            "summary": summary_text if summary_text else "NA",
            "history_text": full_history
        }

    # 3) Build docx
    doc_bytes = build_docx(account, assembled)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    fname = f"CS_Tab_{account or 'Account'}_{stamp}.docx"

    st.success("Document generated.")
    st.download_button("Download Word document", data=doc_bytes, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    # 4) On-page preview
    st.divider()
    st.subheader("Preview")
    for key, label in SECTION_ORDER:
        st.markdown(f"### {label}")
        s = assembled[key]
        st.markdown(f"**Sentiment:** {s['sentiment']}")
        st.markdown(f"**Sentiment Summary** (≤ {word_limits.get(key, default_limit)} words)")
        st.write(s["summary"])
        st.markdown("**Sentiment History** (newest first)")
        st.text(s["history_text"])
