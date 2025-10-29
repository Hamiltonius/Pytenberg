#!/usr/bin/env python3
"""
Pytenberg — Gmail → Folders (standalone)
- OAuth to Gmail (readonly)
- Subject query (prompt or --query)
- Spam excluded, safe attachment types, size cap
- Dry-run + idempotency (gmail_ledger.jsonl)

Author: Thomas Galarneau
License: MIT
"""

import os, re, json, base64, argparse, datetime as dt
from pathlib import Path

# ---- Google API ----
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ---- Paths/Config (standalone) ----
SCRIPT_DIR  = Path(__file__).parent
OUT_ROOT    = SCRIPT_DIR / "pytenberg_output"
LOGS_DIR    = SCRIPT_DIR / "logs"
LEDGER_FILE = LOGS_DIR / "gmail_ledger.jsonl"   # idempotency (by Gmail msg id)

SAFE_EXT = {".pdf", ".jpg", ".jpeg", ".png", ".txt", ".doc", ".docx", ".xls", ".xlsx", ".csv"}
MAX_BYTES = 25 * 1024 * 1024  # 25 MB cap per attachment

# ----------------- helpers -----------------
def banner():
    print("\n" + "="*70)
    print("🛡️  SECURITY FEATURES ENABLED")
    print("="*70)
    print("• Spam excluded (query enforces -in:spam)")
    print(f"• Safe attachment types only: {', '.join(sorted(SAFE_EXT))}")
    print("• Oversized attachments blocked")
    print("• Dry-run available (--dry-run)")
    print("="*70 + "\n")

def sanitize_filename(name: str) -> str:
    if not name: return "attachment.bin"
    name = name.replace("\x00", "")
    name = re.sub(r'[<>:\"/\\|?*\n\r\t]', "_", name).strip().strip(".")
    return name or "attachment.bin"

def connect_gmail():
    """OAuth login; creates token.json on first run."""
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("\n❌ Missing 'credentials.json' (create OAuth Desktop credentials in Google Cloud).")
                print("   Place the file next to this script and rerun.\n")
                return None
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as f:
            f.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def search_gmail(service, query, limit=50):
    """Return list of message objects (id only) up to limit."""
    msgs, token = [], None
    while len(msgs) < limit:
        resp = service.users().messages().list(
            userId='me', q=query, maxResults=min(100, limit - len(msgs)),
            pageToken=token
        ).execute()
        msgs.extend(resp.get('messages', []) or [])
        token = resp.get('nextPageToken')
        if not token: break
    return msgs

def get_headers(full_msg):
    headers = full_msg.get('payload', {}).get('headers', [])
    return {h['name'].lower(): h['value'] for h in headers}

def get_raw_eml(service, msg_id) -> bytes:
    raw = service.users().messages().get(userId='me', id=msg_id, format='raw').execute()['raw']
    return base64.urlsafe_b64decode(raw.encode('utf-8'))

def iter_parts(payload):
    if not payload: return
    yield payload
    for p in payload.get('parts', []) or []:
        yield from iter_parts(p)

def load_ledger() -> set[str]:
    """Read gmail_ledger.jsonl and return set of processed Gmail IDs."""
    seen = set()
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    if LEDGER_FILE.exists():
        with LEDGER_FILE.open("r", encoding="utf-8") as f:
            for line in f:
                try:
                    rec = json.loads(line)
                    gid = rec.get("gmail_id")
                    if gid: seen.add(gid)
                except Exception:
                    pass
    return seen

def append_ledger(gmail_id: str, subject: str, out_dir: Path):
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    with LEDGER_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps({
            "ts": dt.datetime.utcnow().isoformat() + "Z",
            "gmail_id": gmail_id,
            "subject": subject[:120],
            "dir": str(out_dir)
        }, ensure_ascii=False) + "\n")

# ----------------- main -----------------
def main():
    OUT_ROOT.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    ap = argparse.ArgumentParser(description="Pytenberg — Gmail to folders (standalone)")
    ap.add_argument("--output-root", default=str(OUT_ROOT), help="Output directory root")
    ap.add_argument("--query", default=None, help='Gmail search query (e.g., \'from:nytimes.com newer_than:7d\')')
    ap.add_argument("--limit", type=int, default=50, help="Max emails to process")
    ap.add_argument("--dry-run", action="store_true", help="Preview without writing")
    args = ap.parse_args()

    print("☎️ Connecting to Gmail...")
    svc = connect_gmail()
    if not svc: return

    banner()

    if args.query:
        query_text = args.query.strip()
    else:
        query_text = input("What do you want to search for? (e.g., insurance claim): ").strip()
        if not query_text:
            print("No search term entered."); return

    # Safety: force inbox and exclude spam unless user already specified them
    q_bits = [query_text]
    if "in:" not in query_text:
        q_bits.append("in:inbox")
    if "-in:spam" not in query_text:
        q_bits.append("-in:spam")
    gmail_query = " ".join(q_bits)

    print(f"\n🔍 Searching Gmail for: {gmail_query}")
    msgs = search_gmail(svc, gmail_query, limit=args.limit)
    print(f"📧 Found {len(msgs)} emails")
    if not msgs:
        print("\nNo emails found."); return

    base_dir = Path(args.output_root) / sanitize_filename(re.sub(r"\s+", "_", query_text))
    print(f"\n📁 Saving to: {base_dir}/\n")

    seen = load_ledger()
    processed = blocked = 0

    for idx, m in enumerate(msgs, 1):
        mid = m["id"]
        if mid in seen and not args.dry_run:
            print(f"⏭️  {idx}/{len(msgs)} already processed (ledger)")
            continue

        full = svc.users().messages().get(userId='me', id=mid, format='full').execute()
        H = get_headers(full)
        subject = H.get('subject', 'No Subject')
        preview = (subject[:60] + "…") if len(subject) > 60 else subject

        # Plan artifacts
        email_dir  = base_dir / f"email_{idx:03d}"
        eml_path   = email_dir / "email.eml"
        attach_dir = email_dir / "attachments"

        # Gather allowed attachments
        allowed = []
        for part in iter_parts(full.get('payload', {})):
            fname = part.get('filename')
            body  = part.get('body', {})
            att_id = body.get('attachmentId')
            if not fname or not att_id:
                continue
            ext = os.path.splitext(fname)[1].lower()
            if ext not in SAFE_EXT:
                blocked += 1
                print(f"🚫 {idx}/{len(msgs)} {preview}  (blocked type: {ext})")
                continue
            meta = svc.users().messages().attachments().get(userId='me', messageId=mid, id=att_id).execute()
            data = meta.get('data', '')
            blob = base64.urlsafe_b64decode(data.encode('utf-8')) if data else b''
            if len(blob) > MAX_BYTES:
                blocked += 1
                print(f"🚫 {idx}/{len(msgs)} {preview}  (blocked size: {len(blob)} bytes)")
                continue
            allowed.append((fname, blob))

        if args.dry_run:
            print(f"✓ {idx}/{len(msgs)} {preview}  (attachments allowed: {len(allowed)}) [DRY]")
            continue

        # Write artifacts
        email_dir.mkdir(parents=True, exist_ok=True)
        with open(eml_path, "wb") as f:
            f.write(get_raw_eml(svc, mid))
        attach_dir.mkdir(exist_ok=True)
        for fn, blob in allowed:
            with open(attach_dir / sanitize_filename(fn), "wb") as f:
                f.write(blob)

        # Manifest
        with open(email_dir / "manifest.json", "w", encoding="utf-8") as f:
            json.dump({
                "gmail_id": mid,
                "subject": subject,
                "from": H.get('from'),
                "date": H.get('date'),
                "attachments_saved": [sanitize_filename(a[0]) for a in allowed],
                "dir": str(email_dir)
            }, f, indent=2, ensure_ascii=False)

        append_ledger(mid, subject, email_dir)
        processed += 1
        print(f"✓ {idx}/{len(msgs)} {preview}  (attachments: {len(allowed)}, blocked so far: {blocked})")

    print("\n" + "="*70)
    print("✅ COMPLETE")
    print("="*70)
    print(f"📊 Emails processed: {processed}")
    print(f"🚫 Attachments blocked: {blocked}")
    print(f"🧾 Ledger: {LEDGER_FILE}")
    print(f"📁 Output root: {base_dir.parent}")
    print("="*70 + "\n")

if __name__ == "__main__":
    main()
