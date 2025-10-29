# Pytenberg

**Email chaos → organized folders. Automatically.**

Drop Outlook `.msg` files into `drop/` — or connect to Gmail — and get one clean project folder per email.

---

## Install
```bash
pip install extract-msg google-api-python-client google-auth-httplib2 google-auth-oauthlib
mkdir -p drop out refs logs
```

## Usage
```bash
# Local .msg files
python3 pytenberg.py --source local

# Gmail + filter
python3 pytenberg.py --source gmail --subject invoice

# Gmail + sender + screenshot
python3 pytenberg.py --source gmail --from amazon.com --screenshot
```

**CLI Options**
```text
--source        local | gmail        # choose email source
--subject       text filter          # match subject keywords
--from          sender or domain     # match sender
--dry-run                            # preview actions, no writes
--screenshot                         # macOS only, saves .png
```

**Example:**
Subject: "Invoice – Dr. Pepper"
Creates: out/Invoice_Dr_Pepper/

## What You Get

Each email creates a folder with attachments + anything from `refs/`:
```
out/Invoice_Dr_Pepper/
├── contract.pdf          (from email attachment)
├── invoice.pdf           (from email attachment)
├── status_template.xlsx  (from refs/ - always included)
└── archive/
    └── email.msg
```

---

**Why refs/?**  
Drop recurring reference files here once — templates, checklists, or forms you want auto-included in every new folder.

---

## Example Commands
```bash
# Local Outlook messages
python3 pytenberg.py --source local

# Gmail (read-only, filtered by subject)
python3 pytenberg.py --source gmail --subject invoice

# Gmail (filtered by sender)
python3 pytenberg.py --source gmail --from amazon.com --screenshot
```

## Troubleshooting

| Issue                         | Fix                                                                                |
| ----------------------------- | ---------------------------------------------------------------------------------- |
| **No token / No connection?** | `rm token.json` then `python3 pytenberg.py --source gmail`                         |
| **No .msg files found?**      | Save from Outlook: `File → Save As → Outlook Message Format` and drop into `drop/` |
| **Missing extract_msg?**      | `pip install extract-msg`                                                          |
