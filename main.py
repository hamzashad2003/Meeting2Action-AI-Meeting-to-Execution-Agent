"""
Meeting2Action – AI Meeting-to-Execution Agent
===============================================

Turns meeting transcripts into reviewed, executable work:
  1. Paste/upload a transcript
  2. AI (Groq, model "openai/gpt-oss-20b") extracts decisions, tasks,
     owners, deadlines, and flags ambiguous/missing info
  3. You review & edit the extracted action items in an editable table
  4. On approval, tasks are"""
Meeting2Action – AI Meeting-to-Execution Agent
===============================================

Turns meeting transcripts into reviewed, executable work:
  1. Paste/upload a transcript
  2. AI (Groq, model "openai/gpt-oss-20b") extracts decisions, tasks,
     owners, deadlines, and flags ambiguous/missing info
  3. You review & edit the extracted action items in an editable table
  4. On approval, tasks are pushed to a Google Sheet (and optionally
     emailed to owners)

Run locally:
    streamlit run main.py

Secrets / environment variables expected (see sidebar too):
    GROQ_API_KEY                - required, your Groq API key
    GOOGLE_SERVICE_ACCOUNT_JSON - optional, full JSON string of a Google
                                   service account with Sheets API access
    SMTP_HOST / SMTP_PORT / SMTP_USER / SMTP_PASSWORD - optional, for email
"""

import os
import json
import smtplib
from email.mime.text import MIMEText
from datetime import datetime

import streamlit as st
import pandas as pd
from openai import OpenAI

# ---------------------------------------------------------------------------
# Page config & theme
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Meeting2Action – AI Meeting-to-Execution Agent",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded",
)

GITHUB_REPO_URL = "https://github.com/hamzashad2003/Meeting2Action-AI-Meeting-to-Execution-Agent/tree/main"
GITHUB_REPO_LABEL = "Meeting2Action-AI-Meeting-to-Execution-Agent"

CUSTOM_CSS = """
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
    html, body, [class*="css"]  {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    :root {
        --m2a-primary: #4F46E5;
        --m2a-primary-dark: #3730A3;
        --m2a-accent: #06B6D4;
        --m2a-bg: #F8FAFC;
        --m2a-card: #FFFFFF;
        --m2a-text: #0F172A;
        --m2a-muted: #64748B;
        --m2a-success: #16A34A;
        --m2a-warning: #D97706;
        --m2a-danger: #DC2626;
        --m2a-border: #E2E8F0;
    }

    .stApp {
        background: linear-gradient(180deg, #F8FAFC 0%, #EEF2FF 100%);
    }

    /* ---- Force visible dark text on ALL normal main-area content ---- */
    /* (summaries, bullet lists, captions, write() output, etc.) */
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] li,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] span,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] strong,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] ul,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] ol,
    [data-testid="stMain"] [data-testid="stCaptionContainer"] p,
    [data-testid="stMain"] [data-testid="stText"] {
        color: var(--m2a-text) !important;
    }

    .m2a-hero {
        padding: 1.75rem 2rem;
        border-radius: 18px;
        background: linear-gradient(135deg, var(--m2a-primary) 0%, var(--m2a-accent) 100%);
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 30px rgba(79, 70, 229, 0.25);
    }
    .m2a-hero h1 {
        font-weight: 800;
        font-size: 2rem;
        margin-bottom: 0.25rem;
        color: white;
    }
    .m2a-hero p {
        font-size: 1rem;
        opacity: 0.92;
        margin: 0;
    }
    /* Re-assert white text inside the hero banner, overriding the broad rule above */
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] .m2a-hero p,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] .m2a-hero h1 {
        color: white !important;
    }

    .m2a-card {
        background: var(--m2a-card);
        border: 1px solid var(--m2a-border);
        border-radius: 14px;
        padding: 1.25rem 1.5rem;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
        margin-bottom: 1rem;
    }

    .m2a-section-title {
        font-weight: 700;
        font-size: 1.05rem;
        color: var(--m2a-text);
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    .m2a-badge {
        display: inline-block;
        padding: 0.15rem 0.6rem;
        border-radius: 999px;
        font-size: 0.75rem;
        font-weight: 600;
        margin-right: 0.4rem;
    }
    .m2a-badge-success { background: #DCFCE7; color: var(--m2a-success); }
    .m2a-badge-warning { background: #FEF3C7; color: var(--m2a-warning); }
    .m2a-badge-danger  { background: #FEE2E2; color: var(--m2a-danger); }
    .m2a-badge-neutral { background: #E2E8F0; color: var(--m2a-muted); }
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] span.m2a-badge-success { color: var(--m2a-success) !important; }
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] span.m2a-badge-warning { color: var(--m2a-warning) !important; }
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] span.m2a-badge-danger { color: var(--m2a-danger) !important; }
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] span.m2a-badge-neutral { color: var(--m2a-muted) !important; }

    .m2a-muted { color: var(--m2a-muted); font-size: 0.9rem; }
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] span.m2a-muted { color: var(--m2a-muted) !important; }

    div.stButton > button {
        border-radius: 10px;
        font-weight: 600;
        border: none;
        padding: 0.55rem 1.2rem;
    }
    div.stButton > button[kind="primary"] {
        background: var(--m2a-primary);
    }
    div.stButton > button[kind="primary"]:hover {
        background: var(--m2a-primary-dark);
    }

    /* ---- Main-area text inputs / text areas ---- */
    .stApp textarea,
    .stApp input[type="text"],
    .stApp input[type="password"],
    .stApp input[type="number"] {
        background-color: #FFFFFF !important;
        color: var(--m2a-text) !important;
        border: 1px solid var(--m2a-border) !important;
        caret-color: var(--m2a-text) !important;
    }
    .stApp textarea::placeholder,
    .stApp input::placeholder {
        color: var(--m2a-muted) !important;
        opacity: 1 !important;
    }

    /* ---- Radio button labels ---- */
    .stApp [data-testid="stRadio"] label p,
    .stApp [data-testid="stRadio"] label span {
        color: var(--m2a-text) !important;
    }

    /* ---- Selectbox / dropdowns ---- */
    .stApp [data-baseweb="select"] * {
        color: var(--m2a-text) !important;
    }
    .stApp [data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        border-color: var(--m2a-border) !important;
    }

    /* ---- File uploader ---- */
    .stApp [data-testid="stFileUploaderDropzone"] {
        background-color: #FFFFFF !important;
    }
    .stApp [data-testid="stFileUploaderDropzone"] * {
        color: var(--m2a-text) !important;
    }

    /* ---- Data editor / table text ---- */
    .stApp [data-testid="stDataFrame"] * {
        color: var(--m2a-text) !important;
    }

    /* ---- Checkbox labels ---- */
    .stApp [data-testid="stCheckbox"] label p {
        color: var(--m2a-text) !important;
    }

    /* ---- Sidebar ---- */
    section[data-testid="stSidebar"] {
        background: #0F172A;
    }
    section[data-testid="stSidebar"] * {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] input[type="text"],
    section[data-testid="stSidebar"] input[type="password"],
    section[data-testid="stSidebar"] textarea {
        background-color: #1E293B !important;
        color: #F1F5F9 !important;
        border: 1px solid #334155 !important;
        caret-color: #F1F5F9 !important;
    }
    section[data-testid="stSidebar"] input::placeholder,
    section[data-testid="stSidebar"] textarea::placeholder {
        color: #94A3B8 !important;
        opacity: 1 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
        background-color: #1E293B !important;
        border: 1px dashed #334155 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] * {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] small {
        color: #94A3B8 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {
        background-color: #334155 !important;
        color: #F1F5F9 !important;
        border: 1px solid #475569 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button:hover {
        background-color: #475569 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stCheckbox"] label,
    section[data-testid="stSidebar"] [data-testid="stCheckbox"] label p,
    section[data-testid="stSidebar"] [data-testid="stCheckbox"] label span {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] [data-baseweb="select"] > div {
        background-color: #1E293B !important;
        border-color: #334155 !important;
    }

    .m2a-repo-link {
        display: inline-block;
        margin: -0.5rem 0 1rem 0;
        font-size: 0.9rem;
    }
    .m2a-repo-link a {
        color: var(--m2a-primary-dark);
        font-weight: 600;
        text-decoration: none;
    }
    .m2a-repo-link a:hover {
        text-decoration: underline;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Groq (OpenAI-compatible) client
# ---------------------------------------------------------------------------

GROQ_BASE_URL = "https://api.groq.com/openai/v1"
DEFAULT_MODEL = "openai/gpt-oss-20b"

SYSTEM_PROMPT = """You are Meeting2Action, an expert meeting-execution assistant.
Given a raw meeting transcript, extract structured, actionable output.
Respond ONLY with a valid JSON object with exactly these keys:

{
  "summary": "2-4 sentence high-level summary of the meeting",
  "key_decisions": ["decision 1", "decision 2"],
  "action_items": [
    {
      "task": "clear description of the task",
      "owner": "person's name, or 'Unassigned' if not mentioned",
      "due_date": "date/timeframe mentioned, or 'Not specified'",
      "priority": "High, Medium, or Low",
      "ambiguous": true or false,
      "notes": "short note on why it's ambiguous or missing info, empty string if none"
    }
  ],
  "open_questions": ["question or missing info 1", "question or missing info 2"],
  "sentiment": "short phrase describing overall meeting tone"
}

Be precise. If information is missing or unclear, still include the item but
set "ambiguous": true and explain briefly in "notes". Do not output anything
outside the JSON object.
"""


def get_groq_client(api_key: str) -> OpenAI:
    return OpenAI(api_key=api_key, base_url=GROQ_BASE_URL)


def analyze_transcript(client: OpenAI, transcript: str, model: str) -> dict:
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"Analyze this meeting transcript:\n\n{transcript}"},
        ],
        temperature=0.2,
        response_format={"type": "json_object"},
    )
    content = response.choices[0].message.content
    return json.loads(content)


# ---------------------------------------------------------------------------
# Google Sheets export
# ---------------------------------------------------------------------------

def push_to_google_sheets(service_account_json: str, sheet_url_or_id: str, rows: list) -> str:
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    info = json.loads(service_account_json)
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)

    if sheet_url_or_id.startswith("http"):
        sh = gc.open_by_url(sheet_url_or_id)
    else:
        sh = gc.open_by_key(sheet_url_or_id)

    worksheet = sh.sheet1
    existing = worksheet.get_all_values()
    if not existing:
        worksheet.append_row(["Task", "Owner", "Due Date", "Priority", "Ambiguous", "Notes", "Added At"])

    for r in rows:
        worksheet.append_row([
            r.get("task", ""),
            r.get("owner", ""),
            r.get("due_date", ""),
            r.get("priority", ""),
            "Yes" if r.get("ambiguous") else "No",
            r.get("notes", ""),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ])
    return sh.url


# ---------------------------------------------------------------------------
# Email notifications
# ---------------------------------------------------------------------------

def send_email_notifications(smtp_host, smtp_port, smtp_user, smtp_password, owner_emails: dict, rows: list):
    sent, failed = [], []
    tasks_by_owner = {}
    for r in rows:
        owner = r.get("owner", "Unassigned")
        tasks_by_owner.setdefault(owner, []).append(r)

    with smtplib.SMTP(smtp_host, int(smtp_port)) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        for owner, tasks in tasks_by_owner.items():
            to_email = owner_emails.get(owner)
            if not to_email:
                failed.append(owner)
                continue
            body_lines = [f"Hi {owner},", "", "You have new action items from a recent meeting:", ""]
            for t in tasks:
                body_lines.append(f"- {t.get('task')} (Due: {t.get('due_date', 'Not specified')}, Priority: {t.get('priority', 'N/A')})")
            body_lines += ["", "— Sent automatically by Meeting2Action"]
            msg = MIMEText("\n".join(body_lines))
            msg["Subject"] = "Your Action Items from the Meeting"
            msg["From"] = smtp_user
            msg["To"] = to_email
            server.sendmail(smtp_user, [to_email], msg.as_string())
            sent.append(owner)
    return sent, failed


# ---------------------------------------------------------------------------
# Sidebar – configuration
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown("## ⚙️ Configuration")

    groq_api_key = st.text_input(
        "Groq API Key",
        value=os.getenv("GROQ_API_KEY", ""),
        type="password",
        help="Get a free key at console.groq.com",
    )
    model_name = st.text_input("Model", value=DEFAULT_MODEL)

    st.markdown("---")
    st.markdown("### 📊 Google Sheets (optional)")
    sheet_url = st.text_input("Sheet URL or ID", placeholder="https://docs.google.com/spreadsheets/d/...")
    service_account_file = st.file_uploader("Service Account JSON", type=["json"])
    service_account_json = None
    if service_account_file is not None:
        service_account_json = service_account_file.read().decode("utf-8")
    elif os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON"):
        service_account_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")

    st.markdown("---")
    st.markdown("### 📧 Email Notifications (optional)")
    enable_email = st.checkbox("Enable email notifications")
    smtp_host = st.text_input("SMTP Host", value=os.getenv("SMTP_HOST", "smtp.gmail.com")) if enable_email else None
    smtp_port = st.text_input("SMTP Port", value=os.getenv("SMTP_PORT", "587")) if enable_email else None
    smtp_user = st.text_input("SMTP Username / Email", value=os.getenv("SMTP_USER", "")) if enable_email else None
    smtp_password = st.text_input("SMTP Password / App Password", type="password", value=os.getenv("SMTP_PASSWORD", "")) if enable_email else None

# ---------------------------------------------------------------------------
# Hero header
# ---------------------------------------------------------------------------

st.markdown(
    """
    <div class="m2a-hero">
        <h1>✅ Meeting2Action</h1>
        <p>AI Meeting-to-Execution Agent — turn conversations into completed work, not just summaries.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f'<div class="m2a-repo-link">📦 <a href="{GITHUB_REPO_URL}" target="_blank">{GITHUB_REPO_LABEL}</a></div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

if "analysis" not in st.session_state:
    st.session_state.analysis = None
if "approved" not in st.session_state:
    st.session_state.approved = False

# ---------------------------------------------------------------------------
# Input section
# ---------------------------------------------------------------------------

st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
st.markdown('<div class="m2a-section-title">📝 Meeting Transcript</div>', unsafe_allow_html=True)

input_mode = st.radio("Input method", ["Paste text", "Upload .txt file"], horizontal=True, label_visibility="collapsed")

transcript_text = ""
if input_mode == "Paste text":
    transcript_text = st.text_area("Transcript", height=220, placeholder="Paste your meeting transcript or notes here...", label_visibility="collapsed")
else:
    uploaded = st.file_uploader("Upload transcript", type=["txt"], label_visibility="collapsed")
    if uploaded is not None:
        transcript_text = uploaded.read().decode("utf-8")
        st.text_area("Preview", value=transcript_text, height=200, disabled=True, label_visibility="collapsed")

analyze_clicked = st.button("🔍 Analyze Meeting", type="primary", use_container_width=False)
st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Run analysis
# ---------------------------------------------------------------------------

if analyze_clicked:
    if not groq_api_key:
        st.error("Please enter your Groq API key in the sidebar.")
    elif not transcript_text.strip():
        st.error("Please paste or upload a transcript first.")
    else:
        with st.spinner("Analyzing transcript with AI..."):
            try:
                client = get_groq_client(groq_api_key)
                st.session_state.analysis = analyze_transcript(client, transcript_text, model_name)
                st.session_state.approved = False
            except json.JSONDecodeError:
                st.error("The model did not return valid JSON. Try again or adjust the transcript.")
            except Exception as e:
                st.error(f"Error calling Groq API: {e}")

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------

analysis = st.session_state.analysis

if analysis:
    st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
    st.markdown('<div class="m2a-section-title">📋 Summary</div>', unsafe_allow_html=True)
    st.write(analysis.get("summary", "N/A"))
    st.markdown(
        f'<span class="m2a-badge m2a-badge-neutral">Tone: {analysis.get("sentiment", "N/A")}</span>',
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
        st.markdown('<div class="m2a-section-title">✔️ Key Decisions</div>', unsafe_allow_html=True)
        decisions = analysis.get("key_decisions", [])
        if decisions:
            for d in decisions:
                st.markdown(f"- {d}")
        else:
            st.markdown('<span class="m2a-muted">None recorded</span>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
        st.markdown('<div class="m2a-section-title">❓ Open Questions</div>', unsafe_allow_html=True)
        questions = analysis.get("open_questions", [])
        if questions:
            for q in questions:
                st.markdown(f"- {q}")
        else:
            st.markdown('<span class="m2a-muted">None recorded</span>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
    st.markdown('<div class="m2a-section-title">🗂️ Action Items — Review & Edit</div>', unsafe_allow_html=True)
    st.markdown('<span class="m2a-muted">Edit tasks, owners, or deadlines below before approving.</span>', unsafe_allow_html=True)

    action_items = analysis.get("action_items", [])
    df = pd.DataFrame(action_items) if action_items else pd.DataFrame(
        columns=["task", "owner", "due_date", "priority", "ambiguous", "notes"]
    )
    for col in ["task", "owner", "due_date", "priority", "ambiguous", "notes"]:
        if col not in df.columns:
            df[col] = ""

    edited_df = st.data_editor(
        df,
        column_config={
            "task": st.column_config.TextColumn("Task", width="large"),
            "owner": st.column_config.TextColumn("Owner"),
            "due_date": st.column_config.TextColumn("Due Date"),
            "priority": st.column_config.SelectboxColumn("Priority", options=["High", "Medium", "Low"]),
            "ambiguous": st.column_config.CheckboxColumn("Ambiguous?"),
            "notes": st.column_config.TextColumn("Notes", width="medium"),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="action_items_editor",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # -----------------------------------------------------------------------
    # Approval & execution
    # -----------------------------------------------------------------------

    st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
    st.markdown('<div class="m2a-section-title">🚀 Approve & Execute</div>', unsafe_allow_html=True)

    owners = sorted({str(o) for o in edited_df["owner"].tolist() if str(o).strip() and str(o) != "Unassigned"})
    owner_emails = {}
    if enable_email and owners:
        st.markdown('<span class="m2a-muted">Map owners to email addresses:</span>', unsafe_allow_html=True)
        for owner in owners:
            owner_emails[owner] = st.text_input(f"Email for {owner}", key=f"email_{owner}")

    approve_col, status_col = st.columns([1, 3])
    with approve_col:
        approve_clicked = st.button("✅ Approve & Send", type="primary")

    if approve_clicked:
        rows = edited_df.to_dict("records")
        if not rows:
            st.warning("No action items to send.")
        else:
            success_msgs = []
            if sheet_url and service_account_json:
                try:
                    sheet_link = push_to_google_sheets(service_account_json, sheet_url, rows)
                    success_msgs.append(f"Pushed {len(rows)} task(s) to Google Sheets.")
                except Exception as e:
                    st.error(f"Google Sheets error: {e}")
            elif sheet_url or service_account_json:
                st.warning("Provide both a Sheet URL/ID and a Service Account JSON to export to Google Sheets.")

            if enable_email:
                try:
                    sent, failed = send_email_notifications(
                        smtp_host, smtp_port, smtp_user, smtp_password, owner_emails, rows
                    )
                    if sent:
                        success_msgs.append(f"Emailed: {', '.join(sent)}.")
                    if failed:
                        success_msgs.append(f"No email on file for: {', '.join(failed)}.")
                except Exception as e:
                    st.error(f"Email error: {e}")

            if success_msgs:
                st.session_state.approved = True
                for msg in success_msgs:
                    st.success(msg)
            elif not (sheet_url and service_account_json) and not enable_email:
                st.info("No export destination configured — set up Google Sheets or email in the sidebar to send tasks.")

    st.markdown("</div>", unsafe_allow_html=True)

else:
    st.markdown(
        '<div class="m2a-card"><span class="m2a-muted">Paste or upload a transcript above, then click '
        '"Analyze Meeting" to get started.</span></div>',
        unsafe_allow_html=True,
    ) pushed to a Google Sheet (and optionally
     emailed to owners)

Run locally:
    streamlit run main.py

Secrets / environment variables expected (see sidebar too):
    GROQ_API_KEY                - required, your Groq API key
    GOOGLE_SERVICE_ACCOUNT_JSON - optional, full JSON string of a Google
                                   service account with Sheets API access
    SMTP_HOST / SMTP_PORT / SMTP_USER / SMTP_PASSWORD - optional, for email
"""

import os
import json
import smtplib
from email.mime.text import MIMEText
from datetime import datetime

import streamlit as st
import pandas as pd
from openai import OpenAI

# ---------------------------------------------------------------------------
# Page config & theme
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Meeting2Action – AI Meeting-to-Execution Agent",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded",
)

GITHUB_REPO_URL = "https://github.com/hamzashad2003/Meeting2Action-AI-Meeting-to-Execution-Agent/tree/main"
GITHUB_REPO_LABEL = "Meeting2Action-AI-Meeting-to-Execution-Agent"

CUSTOM_CSS = """
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
    html, body, [class*="css"]  {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    :root {
        --m2a-primary: #4F46E5;
        --m2a-primary-dark: #3730A3;
        --m2a-accent: #06B6D4;
        --m2a-bg: #F8FAFC;
        --m2a-card: #FFFFFF;
        --m2a-text: #0F172A;
        --m2a-muted: #64748B;
        --m2a-success: #16A34A;
        --m2a-warning: #D97706;
        --m2a-danger: #DC2626;
        --m2a-border: #E2E8F0;
    }

    .stApp {
        background: linear-gradient(180deg, #F8FAFC 0%, #EEF2FF 100%);
    }

    .m2a-hero {
        padding: 1.75rem 2rem;
        border-radius: 18px;
        background: linear-gradient(135deg, var(--m2a-primary) 0%, var(--m2a-accent) 100%);
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 30px rgba(79, 70, 229, 0.25);
    }
    .m2a-hero h1 {
        font-weight: 800;
        font-size: 2rem;
        margin-bottom: 0.25rem;
        color: white;
    }
    .m2a-hero p {
        font-size: 1rem;
        opacity: 0.92;
        margin: 0;
    }

    .m2a-card {
        background: var(--m2a-card);
        border: 1px solid var(--m2a-border);
        border-radius: 14px;
        padding: 1.25rem 1.5rem;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
        margin-bottom: 1rem;
    }

    .m2a-section-title {
        font-weight: 700;
        font-size: 1.05rem;
        color: var(--m2a-text);
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    .m2a-badge {
        display: inline-block;
        padding: 0.15rem 0.6rem;
        border-radius: 999px;
        font-size: 0.75rem;
        font-weight: 600;
        margin-right: 0.4rem;
    }
    .m2a-badge-success { background: #DCFCE7; color: var(--m2a-success); }
    .m2a-badge-warning { background: #FEF3C7; color: var(--m2a-warning); }
    .m2a-badge-danger  { background: #FEE2E2; color: var(--m2a-danger); }
    .m2a-badge-neutral { background: #E2E8F0; color: var(--m2a-muted); }

    .m2a-muted { color: var(--m2a-muted); font-size: 0.9rem; }

    div.stButton > button {
        border-radius: 10px;
        font-weight: 600;
        border: none;
        padding: 0.55rem 1.2rem;
    }
    div.stButton > button[kind="primary"] {
        background: var(--m2a-primary);
    }
    div.stButton > button[kind="primary"]:hover {
        background: var(--m2a-primary-dark);
    }

    /* ---- Main-area text inputs / text areas ---- */
    .stApp textarea,
    .stApp input[type="text"],
    .stApp input[type="password"],
    .stApp input[type="number"] {
        background-color: #FFFFFF !important;
        color: var(--m2a-text) !important;
        border: 1px solid var(--m2a-border) !important;
        caret-color: var(--m2a-text) !important;
    }
    .stApp textarea::placeholder,
    .stApp input::placeholder {
        color: var(--m2a-muted) !important;
        opacity: 1 !important;
    }

    /* ---- Radio button labels ---- */
    .stApp [data-testid="stRadio"] label p,
    .stApp [data-testid="stRadio"] label span {
        color: var(--m2a-text) !important;
    }

    /* ---- Selectbox / dropdowns ---- */
    .stApp [data-baseweb="select"] * {
        color: var(--m2a-text) !important;
    }
    .stApp [data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        border-color: var(--m2a-border) !important;
    }

    /* ---- File uploader ---- */
    .stApp [data-testid="stFileUploaderDropzone"] {
        background-color: #FFFFFF !important;
    }
    .stApp [data-testid="stFileUploaderDropzone"] * {
        color: var(--m2a-text) !important;
    }

    /* ---- Data editor / table text ---- */
    .stApp [data-testid="stDataFrame"] * {
        color: var(--m2a-text) !important;
    }

    /* ---- Checkbox labels ---- */
    .stApp [data-testid="stCheckbox"] label p {
        color: var(--m2a-text) !important;
    }

    /* ---- Sidebar ---- */
    section[data-testid="stSidebar"] {
        background: #0F172A;
    }
    section[data-testid="stSidebar"] * {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] input[type="text"],
    section[data-testid="stSidebar"] input[type="password"],
    section[data-testid="stSidebar"] textarea {
        background-color: #1E293B !important;
        color: #F1F5F9 !important;
        border: 1px solid #334155 !important;
        caret-color: #F1F5F9 !important;
    }
    section[data-testid="stSidebar"] input::placeholder,
    section[data-testid="stSidebar"] textarea::placeholder {
        color: #94A3B8 !important;
        opacity: 1 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
        background-color: #1E293B !important;
        border: 1px dashed #334155 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] * {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] small {
        color: #94A3B8 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {
        background-color: #334155 !important;
        color: #F1F5F9 !important;
        border: 1px solid #475569 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button:hover {
        background-color: #475569 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stCheckbox"] label,
    section[data-testid="stSidebar"] [data-testid="stCheckbox"] label p,
    section[data-testid="stSidebar"] [data-testid="stCheckbox"] label span {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] [data-baseweb="select"] > div {
        background-color: #1E293B !important;
        border-color: #334155 !important;
    }

    .m2a-repo-link {
        display: inline-block;
        margin: -0.5rem 0 1rem 0;
        font-size: 0.9rem;
    }
    .m2a-repo-link a {
        color: var(--m2a-primary-dark);
        font-weight: 600;
        text-decoration: none;
    }
    .m2a-repo-link a:hover {
        text-decoration: underline;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Groq (OpenAI-compatible) client
# ---------------------------------------------------------------------------

GROQ_BASE_URL = "https://api.groq.com/openai/v1"
DEFAULT_MODEL = "openai/gpt-oss-20b"

SYSTEM_PROMPT = """You are Meeting2Action, an expert meeting-execution assistant.
Given a raw meeting transcript, extract structured, actionable output.
Respond ONLY with a valid JSON object with exactly these keys:

{
  "summary": "2-4 sentence high-level summary of the meeting",
  "key_decisions": ["decision 1", "decision 2"],
  "action_items": [
    {
      "task": "clear description of the task",
      "owner": "person's name, or 'Unassigned' if not mentioned",
      "due_date": "date/timeframe mentioned, or 'Not specified'",
      "priority": "High, Medium, or Low",
      "ambiguous": true or false,
      "notes": "short note on why it's ambiguous or missing info, empty string if none"
    }
  ],
  "open_questions": ["question or missing info 1", "question or missing info 2"],
  "sentiment": "short phrase describing overall meeting tone"
}

Be precise. If information is missing or unclear, still include the item but
set "ambiguous": true and explain briefly in "notes". Do not output anything
outside the JSON object.
"""


def get_groq_client(api_key: str) -> OpenAI:
    return OpenAI(api_key=api_key, base_url=GROQ_BASE_URL)


def analyze_transcript(client: OpenAI, transcript: str, model: str) -> dict:
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"Analyze this meeting transcript:\n\n{transcript}"},
        ],
        temperature=0.2,
        response_format={"type": "json_object"},
    )
    content = response.choices[0].message.content
    return json.loads(content)


# ---------------------------------------------------------------------------
# Google Sheets export
# ---------------------------------------------------------------------------

def push_to_google_sheets(service_account_json: str, sheet_url_or_id: str, rows: list) -> str:
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    info = json.loads(service_account_json)
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)

    if sheet_url_or_id.startswith("http"):
        sh = gc.open_by_url(sheet_url_or_id)
    else:
        sh = gc.open_by_key(sheet_url_or_id)

    worksheet = sh.sheet1
    existing = worksheet.get_all_values()
    if not existing:
        worksheet.append_row(["Task", "Owner", "Due Date", "Priority", "Ambiguous", "Notes", "Added At"])

    for r in rows:
        worksheet.append_row([
            r.get("task", ""),
            r.get("owner", ""),
            r.get("due_date", ""),
            r.get("priority", ""),
            "Yes" if r.get("ambiguous") else "No",
            r.get("notes", ""),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ])
    return sh.url


# ---------------------------------------------------------------------------
# Email notifications
# ---------------------------------------------------------------------------

def send_email_notifications(smtp_host, smtp_port, smtp_user, smtp_password, owner_emails: dict, rows: list):
    sent, failed = [], []
    tasks_by_owner = {}
    for r in rows:
        owner = r.get("owner", "Unassigned")
        tasks_by_owner.setdefault(owner, []).append(r)

    with smtplib.SMTP(smtp_host, int(smtp_port)) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        for owner, tasks in tasks_by_owner.items():
            to_email = owner_emails.get(owner)
            if not to_email:
                failed.append(owner)
                continue
            body_lines = [f"Hi {owner},", "", "You have new action items from a recent meeting:", ""]
            for t in tasks:
                body_lines.append(f"- {t.get('task')} (Due: {t.get('due_date', 'Not specified')}, Priority: {t.get('priority', 'N/A')})")
            body_lines += ["", "— Sent automatically by Meeting2Action"]
            msg = MIMEText("\n".join(body_lines))
            msg["Subject"] = "Your Action Items from the Meeting"
            msg["From"] = smtp_user
            msg["To"] = to_email
            server.sendmail(smtp_user, [to_email], msg.as_string())
            sent.append(owner)
    return sent, failed


# ---------------------------------------------------------------------------
# Sidebar – configuration
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown("## ⚙️ Configuration")

    groq_api_key = st.text_input(
        "Groq API Key",
        value=os.getenv("GROQ_API_KEY", ""),
        type="password",
        help="Get a free key at console.groq.com",
    )
    model_name = st.text_input("Model", value=DEFAULT_MODEL)

    st.markdown("---")
    st.markdown("### 📊 Google Sheets (optional)")
    sheet_url = st.text_input("Sheet URL or ID", placeholder="https://docs.google.com/spreadsheets/d/...")
    service_account_file = st.file_uploader("Service Account JSON", type=["json"])
    service_account_json = None
    if service_account_file is not None:
        service_account_json = service_account_file.read().decode("utf-8")
    elif os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON"):
        service_account_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")

    st.markdown("---")
    st.markdown("### 📧 Email Notifications (optional)")
    enable_email = st.checkbox("Enable email notifications")
    smtp_host = st.text_input("SMTP Host", value=os.getenv("SMTP_HOST", "smtp.gmail.com")) if enable_email else None
    smtp_port = st.text_input("SMTP Port", value=os.getenv("SMTP_PORT", "587")) if enable_email else None
    smtp_user = st.text_input("SMTP Username / Email", value=os.getenv("SMTP_USER", "")) if enable_email else None
    smtp_password = st.text_input("SMTP Password / App Password", type="password", value=os.getenv("SMTP_PASSWORD", "")) if enable_email else None

# ---------------------------------------------------------------------------
# Hero header
# ---------------------------------------------------------------------------

st.markdown(
    """
    <div class="m2a-hero">
        <h1>✅ Meeting2Action</h1>
        <p>AI Meeting-to-Execution Agent — turn conversations into completed work, not just summaries.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f'<div class="m2a-repo-link">📦 <a href="{GITHUB_REPO_URL}" target="_blank">{GITHUB_REPO_LABEL}</a></div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

if "analysis" not in st.session_state:
    st.session_state.analysis = None
if "approved" not in st.session_state:
    st.session_state.approved = False

# ---------------------------------------------------------------------------
# Input section
# ---------------------------------------------------------------------------

st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
st.markdown('<div class="m2a-section-title">📝 Meeting Transcript</div>', unsafe_allow_html=True)

input_mode = st.radio("Input method", ["Paste text", "Upload .txt file"], horizontal=True, label_visibility="collapsed")

transcript_text = ""
if input_mode == "Paste text":
    transcript_text = st.text_area("Transcript", height=220, placeholder="Paste your meeting transcript or notes here...", label_visibility="collapsed")
else:
    uploaded = st.file_uploader("Upload transcript", type=["txt"], label_visibility="collapsed")
    if uploaded is not None:
        transcript_text = uploaded.read().decode("utf-8")
        st.text_area("Preview", value=transcript_text, height=200, disabled=True, label_visibility="collapsed")

analyze_clicked = st.button("🔍 Analyze Meeting", type="primary", use_container_width=False)
st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Run analysis
# ---------------------------------------------------------------------------

if analyze_clicked:
    if not groq_api_key:
        st.error("Please enter your Groq API key in the sidebar.")
    elif not transcript_text.strip():
        st.error("Please paste or upload a transcript first.")
    else:
        with st.spinner("Analyzing transcript with AI..."):
            try:
                client = get_groq_client(groq_api_key)
                st.session_state.analysis = analyze_transcript(client, transcript_text, model_name)
                st.session_state.approved = False
            except json.JSONDecodeError:
                st.error("The model did not return valid JSON. Try again or adjust the transcript.")
            except Exception as e:
                st.error(f"Error calling Groq API: {e}")

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------

analysis = st.session_state.analysis

if analysis:
    st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
    st.markdown('<div class="m2a-section-title">📋 Summary</div>', unsafe_allow_html=True)
    st.write(analysis.get("summary", "N/A"))
    st.markdown(
        f'<span class="m2a-badge m2a-badge-neutral">Tone: {analysis.get("sentiment", "N/A")}</span>',
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
        st.markdown('<div class="m2a-section-title">✔️ Key Decisions</div>', unsafe_allow_html=True)
        decisions = analysis.get("key_decisions", [])
        if decisions:
            for d in decisions:
                st.markdown(f"- {d}")
        else:
            st.markdown('<span class="m2a-muted">None recorded</span>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
        st.markdown('<div class="m2a-section-title">❓ Open Questions</div>', unsafe_allow_html=True)
        questions = analysis.get("open_questions", [])
        if questions:
            for q in questions:
                st.markdown(f"- {q}")
        else:
            st.markdown('<span class="m2a-muted">None recorded</span>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
    st.markdown('<div class="m2a-section-title">🗂️ Action Items — Review & Edit</div>', unsafe_allow_html=True)
    st.markdown('<span class="m2a-muted">Edit tasks, owners, or deadlines below before approving.</span>', unsafe_allow_html=True)

    action_items = analysis.get("action_items", [])
    df = pd.DataFrame(action_items) if action_items else pd.DataFrame(
        columns=["task", "owner", "due_date", "priority", "ambiguous", "notes"]
    )
    for col in ["task", "owner", "due_date", "priority", "ambiguous", "notes"]:
        if col not in df.columns:
            df[col] = ""

    edited_df = st.data_editor(
        df,
        column_config={
            "task": st.column_config.TextColumn("Task", width="large"),
            "owner": st.column_config.TextColumn("Owner"),
            "due_date": st.column_config.TextColumn("Due Date"),
            "priority": st.column_config.SelectboxColumn("Priority", options=["High", "Medium", "Low"]),
            "ambiguous": st.column_config.CheckboxColumn("Ambiguous?"),
            "notes": st.column_config.TextColumn("Notes", width="medium"),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="action_items_editor",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # -----------------------------------------------------------------------
    # Approval & execution
    # -----------------------------------------------------------------------

    st.markdown('<div class="m2a-card">', unsafe_allow_html=True)
    st.markdown('<div class="m2a-section-title">🚀 Approve & Execute</div>', unsafe_allow_html=True)

    owners = sorted({str(o) for o in edited_df["owner"].tolist() if str(o).strip() and str(o) != "Unassigned"})
    owner_emails = {}
    if enable_email and owners:
        st.markdown('<span class="m2a-muted">Map owners to email addresses:</span>', unsafe_allow_html=True)
        for owner in owners:
            owner_emails[owner] = st.text_input(f"Email for {owner}", key=f"email_{owner}")

    approve_col, status_col = st.columns([1, 3])
    with approve_col:
        approve_clicked = st.button("✅ Approve & Send", type="primary")

    if approve_clicked:
        rows = edited_df.to_dict("records")
        if not rows:
            st.warning("No action items to send.")
        else:
            success_msgs = []
            if sheet_url and service_account_json:
                try:
                    sheet_link = push_to_google_sheets(service_account_json, sheet_url, rows)
                    success_msgs.append(f"Pushed {len(rows)} task(s) to Google Sheets.")
                except Exception as e:
                    st.error(f"Google Sheets error: {e}")
            elif sheet_url or service_account_json:
                st.warning("Provide both a Sheet URL/ID and a Service Account JSON to export to Google Sheets.")

            if enable_email:
                try:
                    sent, failed = send_email_notifications(
                        smtp_host, smtp_port, smtp_user, smtp_password, owner_emails, rows
                    )
                    if sent:
                        success_msgs.append(f"Emailed: {', '.join(sent)}.")
                    if failed:
                        success_msgs.append(f"No email on file for: {', '.join(failed)}.")
                except Exception as e:
                    st.error(f"Email error: {e}")

            if success_msgs:
                st.session_state.approved = True
                for msg in success_msgs:
                    st.success(msg)
            elif not (sheet_url and service_account_json) and not enable_email:
                st.info("No export destination configured — set up Google Sheets or email in the sidebar to send tasks.")

    st.markdown("</div>", unsafe_allow_html=True)

else:
    st.markdown(
        '<div class="m2a-card"><span class="m2a-muted">Paste or upload a transcript above, then click '
        '"Analyze Meeting" to get started.</span></div>',
        unsafe_allow_html=True,
    )
