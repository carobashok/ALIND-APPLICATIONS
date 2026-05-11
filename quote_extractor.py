"""
Carob Technologies — Quote Request Extractor
============================================
Streamlit app that reads new Gmail emails, extracts quote fields
using Claude AI, and stores them in Supabase.

Install dependencies:
    pip install streamlit google-auth google-auth-oauthlib
                google-auth-httplib2 google-api-python-client
                anthropic supabase

Run:
    streamlit run quote_extractor.py

Secrets — create .streamlit/secrets.toml in the same folder:
    ANTHROPIC_API_KEY = "sk-ant-..."
    SUPABASE_URL      = "https://your-project.supabase.co"
    SUPABASE_SERVICE_KEY = "eyJ..."

    [gmail]
    type                    = "authorized_user"
    client_id               = "....apps.googleusercontent.com"
    client_secret           = "GOCSPX-..."
    refresh_token           = "1//..."
"""

import json
import base64
import re

import streamlit as st

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

import anthropic
from supabase import create_client, Client

# ── Constants ──────────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
]

QUOTE_KEYWORDS = ["quote", "rfq", "rate", "quotation", "price", "pricing", "proposal"]

EXTRACTION_PROMPT = """You are a quote request extractor for Carob Technologies, an AI and analytics company based in Chennai, India.

Extract information from the email below and return ONLY a valid JSON object — no explanation, no markdown, no extra text, no code fences.

Fields to extract:
- customer_name: full name of the sender (string or null)
- customer_email: email address of the original sender (string or null)
- company_name: company or organisation name if mentioned (string or null)
- phone: phone number if mentioned (string or null)
- product_description: what product or service they are asking about — summarise clearly (string)
- quantity: number of units, licences, projects, etc. (string or null)
- unit: unit type e.g. units, licences, nos, projects (string or null)
- deadline: when they need it by — use their exact words (string or null)
- location: city, state, or project location if mentioned (string or null)
- urgency_level: your assessment — "high" if they say urgent/ASAP/immediately, "low" if no timeline at all, "medium" otherwise (must be exactly one of: high, medium, low)
- needs_review: true if the email is ambiguous, incomplete, spam, or unclear — false if all key fields are clear (boolean)
- notes: any other relevant context not captured above (string or null)

Email subject: {subject}
From: {sender}
Email body:
{body}"""

STATUS_OPTIONS = ["new", "quoted", "won", "lost"]
STATUS_COLORS  = {"new": "🔵", "quoted": "🟡", "won": "🟢", "lost": "🔴"}
URGENCY_ICONS  = {"high": "🔴", "medium": "🟡", "low": "🟢"}


# ── Supabase ───────────────────────────────────────────────────────────────────

@st.cache_resource
def get_supabase() -> Client:
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)


# ── Gmail Auth ─────────────────────────────────────────────────────────────────

def get_gmail_service():
    """Build Gmail service from OAuth token stored in Streamlit secrets."""
    gmail_secret = dict(st.secrets["gmail"])
    creds = Credentials(
        token=None,
        refresh_token=gmail_secret["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=gmail_secret["client_id"],
        client_secret=gmail_secret["client_secret"],
        scopes=SCOPES,
    )
    if not creds.valid:
        creds.refresh(Request())
    return build("gmail", "v1", credentials=creds)


# ── Gmail Helpers ──────────────────────────────────────────────────────────────

def decode_body(payload: dict) -> str:
    if payload.get("mimeType") == "text/plain":
        data = payload.get("body", {}).get("data", "")
        if data:
            return base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
    for part in payload.get("parts", []):
        result = decode_body(part)
        if result:
            return result
    return ""


def is_quote_email(subject: str, body: str) -> bool:
    text = (subject + " " + body).lower()
    return any(kw in text for kw in QUOTE_KEYWORDS)


def fetch_unread_emails(service, log) -> list:
    log("📬 Checking Gmail inbox for unread emails...\n")
    result = service.users().messages().list(
        userId="me", labelIds=["INBOX", "UNREAD"], maxResults=20
    ).execute()
    messages = result.get("messages", [])
    if not messages:
        log("✅ No unread emails found.\n")
        return []
    log(f"📨 Found {len(messages)} unread email(s). Scanning...\n\n")
    emails = []
    for msg in messages:
        full    = service.users().messages().get(userId="me", messageId=msg["id"], format="full").execute()
        headers = {h["name"]: h["value"] for h in full["payload"]["headers"]}
        emails.append({
            "id":      msg["id"],
            "subject": headers.get("Subject", "(no subject)"),
            "sender":  headers.get("From", ""),
            "body":    decode_body(full["payload"]),
        })
    return emails


def mark_as_read(service, message_id: str):
    service.users().messages().modify(
        userId="me", id=message_id, body={"removeLabelIds": ["UNREAD"]}
    ).execute()


# ── Claude Extraction ──────────────────────────────────────────────────────────

def extract_quote_fields(email: dict, log) -> dict | None:
    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    prompt = EXTRACTION_PROMPT.format(
        subject=email["subject"],
        sender=email["sender"],
        body=email["body"][:4000],
    )
    try:
        message = client.messages.create(
            model="claude-opus-4-5",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = message.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?", "", raw).strip()
        raw = re.sub(r"```$", "", raw).strip()
        return json.loads(raw)
    except json.JSONDecodeError as e:
        log(f"   ⚠️  Could not parse Claude response: {e}\n")
        return None
    except Exception as e:
        log(f"   ❌ Claude API error: {e}\n")
        return None


# ── Supabase Insert ────────────────────────────────────────────────────────────

def insert_quote(supabase: Client, email: dict, fields: dict, log) -> bool:
    row = {
        "customer_name":       fields.get("customer_name"),
        "customer_email":      fields.get("customer_email"),
        "company_name":        fields.get("company_name"),
        "phone":               fields.get("phone"),
        "product_description": fields.get("product_description"),
        "quantity":            str(fields.get("quantity")) if fields.get("quantity") else None,
        "unit":                fields.get("unit"),
        "deadline":            fields.get("deadline"),
        "location":            fields.get("location"),
        "urgency_level":       fields.get("urgency_level", "medium"),
        "needs_review":        bool(fields.get("needs_review", False)),
        "notes":               fields.get("notes"),
        "raw_email_subject":   email["subject"],
        "raw_email_body":      email["body"][:5000],
        "sender_email":        email["sender"],
        "status":              "new",
    }
    try:
        supabase.table("quote_requests").insert(row).execute()
        return True
    except Exception as e:
        log(f"   ❌ Supabase insert error: {e}\n")
        return False


# ── Main Processor ─────────────────────────────────────────────────────────────

def run_extraction(log_fn):
    supabase = get_supabase()
    service  = get_gmail_service()
    emails   = fetch_unread_emails(service, log_fn)
    if not emails:
        return 0, 0

    found = skipped = saved = 0

    for email in emails:
        found += 1
        if not is_quote_email(email["subject"], email["body"]):
            log_fn(f"⏭️  Skipped (not a quote): {email['subject']}\n")
            mark_as_read(service, email["id"])
            skipped += 1
            continue

        log_fn(f"📋 Processing: {email['subject']}\n")
        log_fn(f"   From: {email['sender']}\n")

        fields = extract_quote_fields(email, log_fn)
        if not fields:
            log_fn("   ⚠️  Skipping — extraction failed.\n\n")
            continue

        log_fn(f"   Customer : {fields.get('customer_name') or '—'}\n")
        log_fn(f"   Product  : {fields.get('product_description') or '—'}\n")
        log_fn(f"   Quantity : {fields.get('quantity') or '—'} {fields.get('unit') or ''}\n")
        log_fn(f"   Deadline : {fields.get('deadline') or '—'}\n")
        log_fn(f"   Urgency  : {fields.get('urgency_level') or '—'}\n")
        log_fn(f"   Review?  : {'⚠️ Yes' if fields.get('needs_review') else '✅ No'}\n")

        ok = insert_quote(supabase, email, fields, log_fn)
        if ok:
            mark_as_read(service, email["id"])
            log_fn("   ✅ Saved to Supabase.\n\n")
            saved += 1
        else:
            log_fn("   ❌ Failed to save.\n\n")

    log_fn(f"─────────────────────────\n")
    log_fn(f"Done. {found} checked · {skipped} skipped · {saved} saved.\n")
    return found, saved


# ── Streamlit UI ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Quote Extractor — Carob Technologies",
    page_icon="📬",
    layout="wide",
)

st.title("📬 Quote Request Extractor")
st.caption("Reads Gmail · Extracts with Claude · Stores in Supabase")

tab_run, tab_quotes = st.tabs(["▶ Run extraction", "📋 Quote requests"])


# ── Tab 1: Run ─────────────────────────────────────────────────────────────────

with tab_run:
    st.markdown("Click the button to scan your Gmail inbox for new unread quote request emails.")
    st.info(
        "Only **unread** emails in your inbox are checked. "
        "Non-quote emails are skipped and marked as read automatically.",
        icon="ℹ️",
    )

    if st.button("▶ Check Gmail now", type="primary", use_container_width=True):
        output_area = st.empty()
        log_buffer  = []

        def log(text):
            log_buffer.append(text)
            output_area.code("".join(log_buffer), language=None)

        with st.spinner("Running..."):
            found, saved = run_extraction(log)

        if saved > 0:
            st.success(f"✅ {saved} quote request(s) saved. Open the Quote Requests tab to review.")
        elif found == 0:
            st.info("No unread emails found.")
        else:
            st.warning("Emails were found but none qualified as quote requests.")

    st.divider()
    with st.expander("⚙️ Detection keywords"):
        st.write(", ".join(f"`{k}`" for k in QUOTE_KEYWORDS))
        st.caption("Edit the QUOTE_KEYWORDS list in the script to customise.")


# ── Tab 2: Quote Requests ──────────────────────────────────────────────────────

with tab_quotes:
    supabase = get_supabase()

    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        status_filter  = st.selectbox("Status", ["All"] + STATUS_OPTIONS)
    with col2:
        urgency_filter = st.selectbox("Urgency", ["All", "high", "medium", "low"])
    with col3:
        st.write("")
        st.button("🔄 Refresh", use_container_width=True)

    try:
        query = supabase.table("quote_requests").select("*").order("created_at", desc=True)
        if status_filter  != "All":
            query = query.eq("status", status_filter)
        if urgency_filter != "All":
            query = query.eq("urgency_level", urgency_filter)
        rows = query.execute().data
    except Exception as e:
        st.error(f"Could not load from Supabase: {e}")
        rows = []

    if not rows:
        st.info("No quote requests found.")
    else:
        st.caption(f"{len(rows)} record(s)")

        for row in rows:
            created  = row.get("created_at", "")[:16].replace("T", " ")
            status   = row.get("status", "new")
            urgency  = row.get("urgency_level", "medium")
            review   = row.get("needs_review", False)
            s_icon   = STATUS_COLORS.get(status, "⚪")
            u_icon   = URGENCY_ICONS.get(urgency, "⚪")

            label = (
                f"{s_icon} {row.get('customer_name') or 'Unknown'}  —  "
                f"{(row.get('product_description') or '')[:55]}  "
                f"{'⚠️' if review else ''}  |  {created}"
            )

            with st.expander(label):
                left, right = st.columns(2)

                with left:
                    st.markdown("**Customer**")
                    st.write(f"Name    : {row.get('customer_name') or '—'}")
                    st.write(f"Email   : {row.get('customer_email') or '—'}")
                    st.write(f"Company : {row.get('company_name') or '—'}")
                    st.write(f"Phone   : {row.get('phone') or '—'}")
                    st.markdown("**Request**")
                    st.write(f"Product  : {row.get('product_description') or '—'}")
                    st.write(f"Quantity : {row.get('quantity') or '—'} {row.get('unit') or ''}")
                    st.write(f"Deadline : {row.get('deadline') or '—'}")
                    st.write(f"Location : {row.get('location') or '—'}")

                with right:
                    st.markdown("**AI assessment**")
                    st.write(f"Urgency      : {u_icon} {urgency}")
                    st.write(f"Needs review : {'⚠️ Yes' if review else '✅ No'}")
                    if row.get("notes"):
                        st.write(f"Notes : {row['notes']}")

                    st.markdown("**Update status**")
                    new_status = st.selectbox(
                        "Status",
                        STATUS_OPTIONS,
                        index=STATUS_OPTIONS.index(status),
                        key=f"status_{row['id']}",
                        label_visibility="collapsed",
                    )
                    if new_status != status:
                        try:
                            supabase.table("quote_requests").update(
                                {"status": new_status}
                            ).eq("id", row["id"]).execute()
                            st.success(f"Updated to **{new_status}**")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Update failed: {e}")

                with st.expander("📧 Original email"):
                    st.write(f"**Subject:** {row.get('raw_email_subject') or '—'}")
                    st.write(f"**From:** {row.get('sender_email') or '—'}")
                    st.text(row.get("raw_email_body") or "—")
