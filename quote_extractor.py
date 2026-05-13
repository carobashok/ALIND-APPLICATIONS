"""
Carob Technologies — Quote Request Extractor
============================================
Streamlit app: Fetch Gmail → Select emails → Extract with Claude → Store in Supabase

Install:
    pip install streamlit google-auth google-auth-oauthlib
                google-auth-httplib2 google-api-python-client
                anthropic supabase

Secrets (.streamlit/secrets.toml):
    ANTHROPIC_API_KEY = "sk-ant-..."

    [supabase]
    url = "https://xxx.supabase.co"
    key = "eyJ..."

    [gmail]
    client_id     = "....apps.googleusercontent.com"
    client_secret = "GOCSPX-..."
    refresh_token = "1//..."
"""

import json
import base64
import re
import os
from datetime import datetime, timezone
from pathlib import Path

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

ATTACHMENT_BASE = Path.home() / "Desktop" / "CarobQuotes"

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
- urgency_level: "high" if urgent/ASAP/immediately, "low" if no timeline, "medium" otherwise (must be: high, medium, or low)
- needs_review: true if ambiguous, incomplete, spam, or unclear (boolean)
- notes: any other relevant context (string or null)

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

@st.cache_resource
def get_gmail_service():
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
    """Recursively extract plain text body from Gmail payload."""
    mime = payload.get("mimeType", "")
    if mime == "text/plain":
        data = payload.get("body", {}).get("data", "")
        if data:
            return base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
    if mime == "text/html" and not payload.get("parts"):
        data = payload.get("body", {}).get("data", "")
        if data:
            html = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
            return re.sub(r"<[^>]+>", " ", html)
    for part in payload.get("parts", []):
        result = decode_body(part)
        if result:
            return result
    return ""


def get_attachments(payload: dict) -> list:
    """Return list of attachment metadata from Gmail payload."""
    attachments = []
    for part in payload.get("parts", []):
        filename = part.get("filename", "")
        if filename:
            attachments.append({
                "filename":    filename,
                "mime_type":   part.get("mimeType", ""),
                "attachment_id": part.get("body", {}).get("attachmentId", ""),
                "size":        part.get("body", {}).get("size", 0),
            })
        attachments.extend(get_attachments(part))
    return attachments


def download_attachment(service, message_id: str, attachment_id: str) -> bytes:
    """Download attachment bytes from Gmail."""
    result = service.users().messages().attachments().get(
        userId="me", messageId=message_id, id=attachment_id
    ).execute()
    data = result.get("data", "")
    return base64.urlsafe_b64decode(data)


def save_attachments(service, email: dict, customer_name: str) -> tuple[str, int]:
    """Save all attachments to Desktop/CarobQuotes/<folder>. Returns (folder_path, count)."""
    attachments = email.get("attachments", [])
    if not attachments:
        return "", 0

    date_str   = datetime.now().strftime("%Y-%m-%d")
    safe_name  = re.sub(r"[^\w\s-]", "", customer_name or "Unknown").strip().replace(" ", "_")
    folder_name = f"{date_str}_{safe_name}"
    folder_path = ATTACHMENT_BASE / folder_name
    folder_path.mkdir(parents=True, exist_ok=True)

    saved = 0
    for att in attachments:
        if not att["attachment_id"]:
            continue
        try:
            data = download_attachment(service, email["id"], att["attachment_id"])
            file_path = folder_path / att["filename"]
            file_path.write_bytes(data)
            saved += 1
        except Exception:
            pass

    return str(folder_path), saved


def fetch_unread_emails(service) -> list:
    """Fetch all unread emails from inbox with full details."""
    result = service.users().messages().list(
        userId="me", labelIds=["INBOX", "UNREAD"], maxResults=30
    ).execute()
    messages = result.get("messages", [])
    if not messages:
        return []

    emails = []
    for msg in messages:
        full    = service.users().messages().get(userId="me", id=msg["id"]).execute()
        headers = {h["name"]: h["value"] for h in full["payload"]["headers"]}
        body    = decode_body(full["payload"])
        atts    = get_attachments(full["payload"])

        # Format date
        date_raw = headers.get("Date", "")
        try:
            from email.utils import parsedate_to_datetime
            dt = parsedate_to_datetime(date_raw)
            date_fmt = dt.strftime("%d %b, %I:%M %p")
        except Exception:
            date_fmt = date_raw[:16]

        emails.append({
            "id":          msg["id"],
            "thread_id":   full.get("threadId", ""),
            "subject":     headers.get("Subject", "(no subject)"),
            "sender":      headers.get("From", ""),
            "date":        date_fmt,
            "body":        body,
            "preview":     body[:120].replace("\n", " ").strip(),
            "attachments": atts,
        })
    return emails


def mark_as_read(service, message_id: str):
    service.users().messages().modify(
        userId="me", id=message_id, body={"removeLabelIds": ["UNREAD"]}
    ).execute()


# ── Claude Extraction ──────────────────────────────────────────────────────────

def extract_quote_fields(email: dict) -> dict | None:
    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    prompt = EXTRACTION_PROMPT.format(
        subject=email["subject"],
        sender=email["sender"],
        body=email["body"][:4000],
    )
    try:
        message = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = message.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?", "", raw).strip()
        raw = re.sub(r"```$", "", raw).strip()
        return json.loads(raw)
    except Exception:
        return None


# ── Supabase Insert / Update ───────────────────────────────────────────────────

def upsert_quote(supabase: Client, service, email: dict, fields: dict) -> tuple[bool, str]:
    """
    Insert new quote or append to existing thread.
    Returns (success, action) where action is 'inserted' or 'updated'.
    """
    thread_id = email.get("thread_id", "")
    now       = datetime.now(timezone.utc).isoformat()

    # Build conversation entry for this message
    conv_entry = {
        "sender":    email["sender"],
        "timestamp": now,
        "subject":   email["subject"],
        "body":      email["body"][:2000],
    }

    # Check if thread already exists
    existing = None
    if thread_id:
        try:
            res = supabase.table("quote_requests").select("id, conversation_log, reply_count").eq("thread_id", thread_id).execute()
            if res.data:
                existing = res.data[0]
        except Exception:
            pass

    if existing:
        # Thread exists — append to conversation log
        conv_log   = existing.get("conversation_log") or []
        conv_log.append(conv_entry)
        reply_count = (existing.get("reply_count") or 0) + 1
        try:
            supabase.table("quote_requests").update({
                "conversation_log": conv_log,
                "reply_count":      reply_count,
                "last_reply_at":    now,
            }).eq("id", existing["id"]).execute()
            return True, "updated"
        except Exception as e:
            return False, str(e)
    else:
        # New thread — save attachments and insert
        customer_name = fields.get("customer_name") or "Unknown"
        folder_path, att_count = save_attachments(service, email, customer_name)

        row = {
            "thread_id":           thread_id or None,
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
            "conversation_log":    [conv_entry],
            "reply_count":         0,
            "last_reply_at":       now,
            "attachment_folder":   folder_path or None,
            "attachment_count":    att_count,
        }
        try:
            supabase.table("quote_requests").insert(row).execute()
            return True, "inserted"
        except Exception as e:
            return False, str(e)


# ── Streamlit UI ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Quote Extractor — Carob Technologies",
    page_icon="📬",
    layout="wide",
)

st.title("📬 Quote Request Extractor")
st.caption("Carob Technologies · Gmail → Claude → Supabase")

tab_inbox, tab_quotes = st.tabs(["📬 Inbox", "📋 Quote Requests"])


# ── Tab 1: Inbox ───────────────────────────────────────────────────────────────

with tab_inbox:

    # Session state
    if "emails"   not in st.session_state: st.session_state.emails   = []
    if "selected" not in st.session_state: st.session_state.selected = set()
    if "log"      not in st.session_state: st.session_state.log      = ""

    col_fetch, col_selall, col_clear, col_count = st.columns([2, 1, 1, 3])

    with col_fetch:
        if st.button("📬 Fetch unread emails", type="primary", use_container_width=True):
            with st.spinner("Connecting to Gmail..."):
                try:
                    service = get_gmail_service()
                    st.session_state.emails   = fetch_unread_emails(service)
                    st.session_state.selected = set()
                    st.session_state.log      = ""
                except Exception as e:
                    st.error(f"Gmail error: {e}")

    with col_selall:
        if st.button("☑ Select all", use_container_width=True):
            st.session_state.selected = {e["id"] for e in st.session_state.emails}
            st.rerun()

    with col_clear:
        if st.button("☐ Clear", use_container_width=True):
            st.session_state.selected = set()
            st.rerun()

    with col_count:
        total    = len(st.session_state.emails)
        selected = len(st.session_state.selected)
        if total:
            st.info(f"{total} unread email(s) fetched · {selected} selected", icon="📨")

    st.info(
        "Unselected emails stay **unread** in Gmail. "
        "Only selected emails are extracted and saved to Supabase.",
        icon="ℹ️",
    )

    # Email list
    if not st.session_state.emails:
        st.markdown("Click **Fetch unread emails** to load your inbox.")
    else:
        for email in st.session_state.emails:
            is_checked = email["id"] in st.session_state.selected
            has_att    = len(email.get("attachments", [])) > 0
            att_label  = f" 📎 {len(email['attachments'])}" if has_att else ""

            col_cb, col_body = st.columns([0.5, 11])

            with col_cb:
                st.write("")
                checked = st.checkbox(
                    "",
                    value=is_checked,
                    key=f"chk_{email['id']}",
                    label_visibility="collapsed",
                )
                if checked:
                    st.session_state.selected.add(email["id"])
                else:
                    st.session_state.selected.discard(email["id"])

            with col_body:
                with st.expander(
                    f"**{email['sender']}** · {email['subject']}{att_label} · *{email['date']}*"
                ):
                    st.write(email["body"] or "(no body)")
                    if has_att:
                        st.caption(f"📎 Attachments: {', '.join(a['filename'] for a in email['attachments'])}")

        st.divider()

        # Extract bar
        n_selected = len(st.session_state.selected)
        col_ex, col_status = st.columns([2, 4])

        with col_ex:
            extract_clicked = st.button(
                f"▶ Extract {n_selected} selected email(s) → Supabase",
                type="primary",
                disabled=(n_selected == 0),
                use_container_width=True,
            )

        if extract_clicked and n_selected > 0:
            supabase = get_supabase()
            service  = get_gmail_service()
            log_area = st.empty()
            log_buf  = []

            def log(text):
                log_buf.append(text)
                log_area.code("".join(log_buf), language=None)

            inserted = updated = failed = 0

            for email in st.session_state.emails:
                if email["id"] not in st.session_state.selected:
                    continue

                log(f"📋 Processing: {email['subject']}\n")
                log(f"   From: {email['sender']}\n")

                fields = extract_quote_fields(email)
                if not fields:
                    log("   ❌ Claude extraction failed — left unread.\n\n")
                    failed += 1
                    continue

                log(f"   Customer : {fields.get('customer_name') or '—'}\n")
                log(f"   Product  : {fields.get('product_description') or '—'}\n")
                log(f"   Quantity : {fields.get('quantity') or '—'} {fields.get('unit') or ''}\n")
                log(f"   Deadline : {fields.get('deadline') or '—'}\n")
                log(f"   Urgency  : {fields.get('urgency_level') or '—'}\n")
                log(f"   Review?  : {'⚠️ Yes' if fields.get('needs_review') else '✅ No'}\n")

                ok, action = upsert_quote(supabase, service, email, fields)
                if ok:
                    mark_as_read(service, email["id"])
                    if action == "inserted":
                        log("   ✅ New quote saved to Supabase.\n\n")
                        inserted += 1
                    else:
                        log("   🔄 Existing thread updated in Supabase.\n\n")
                        updated += 1
                else:
                    log(f"   ❌ Supabase error: {action} — left unread.\n\n")
                    failed += 1

            log(f"─────────────────────────────────────\n")
            log(f"Done. {inserted} new · {updated} thread updates · {failed} failed.\n")

            # Refresh inbox
            st.session_state.emails   = fetch_unread_emails(service)
            st.session_state.selected = set()

            if failed == 0:
                st.success(f"✅ {inserted} saved · {updated} updated. Switch to Quote Requests tab.")
            else:
                st.warning(f"{inserted} saved · {updated} updated · {failed} failed (still unread in Gmail).")


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
        st.button("🔄 Refresh", use_container_width=True, key="refresh_quotes")

    try:
        query = supabase.table("quote_requests").select("*").order("created_at", desc=True)
        if status_filter  != "All": query = query.eq("status", status_filter)
        if urgency_filter != "All": query = query.eq("urgency_level", urgency_filter)
        rows = query.execute().data
    except Exception as e:
        st.error(f"Could not load from Supabase: {e}")
        rows = []

    if not rows:
        st.info("No quote requests found.")
    else:
        st.caption(f"{len(rows)} record(s)")

        for row in rows:
            created     = row.get("created_at", "")[:16].replace("T", " ")
            status      = row.get("status", "new")
            urgency     = row.get("urgency_level", "medium")
            review      = row.get("needs_review", False)
            reply_count = row.get("reply_count", 0)
            s_icon      = STATUS_COLORS.get(status, "⚪")
            u_icon      = URGENCY_ICONS.get(urgency, "⚪")

            label = (
                f"{s_icon} {row.get('customer_name') or 'Unknown'}  —  "
                f"{(row.get('product_description') or '')[:50]}  "
                f"{'⚠️' if review else ''}  "
                f"{'💬 ' + str(reply_count + 1) if reply_count else ''}  "
                f"|  {created}"
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

                    if row.get("attachment_folder"):
                        st.markdown("**Attachments**")
                        st.write(f"📁 {row['attachment_folder']}")
                        st.write(f"📎 {row.get('attachment_count', 0)} file(s)")

                with right:
                    st.markdown("**AI Assessment**")
                    st.write(f"Urgency      : {u_icon} {urgency}")
                    st.write(f"Needs review : {'⚠️ Yes' if review else '✅ No'}")
                    if row.get("notes"):
                        st.write(f"Notes : {row['notes']}")
                    if row.get("last_reply_at"):
                        st.write(f"Last reply   : {row['last_reply_at'][:16].replace('T',' ')}")

                    st.markdown("**Update Status**")
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

                # Conversation thread
                conv_log = row.get("conversation_log") or []
                if conv_log:
                    st.markdown("---")
                    st.markdown(f"**💬 Conversation ({len(conv_log)} message(s))**")
                    for entry in conv_log:
                        sender    = entry.get("sender", "")
                        timestamp = entry.get("timestamp", "")[:16].replace("T", " ")
                        body      = entry.get("body", "")
                        subject   = entry.get("subject", "")
                        is_me     = "carobashok" in sender.lower()
                        align     = "🏢" if is_me else "👤"
                        name      = "You" if is_me else sender.split("<")[0].strip()

                        st.markdown(f"{align} **{name}** · *{timestamp}* · _{subject}_")
                        st.text(body[:500] + ("..." if len(body) > 500 else ""))
                        st.markdown("---")
