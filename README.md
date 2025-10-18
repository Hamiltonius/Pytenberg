# Pytenberg

Email chaos → organized folders. Automatically.

Drop saved Outlook `.msg` emails into a folder — get structured project directories with attachments and reference files sorted automatically.

---

## Quick Start
```bash
# Install dependency
pip install extract-msg

# Create working folders
mkdir -p drop out refs

# Place .msg files in drop/

# Then run:
python3 pytenberg.py
```

**Result:** Each email becomes a folder in `out/` named from the subject line.

---

## How It Works

**Default behavior:** Extracts text before `-` or `:` in the subject line.

Example: `Acme Corp - Invoice` → creates `out/Acme_Corp/`

**Active Variants:** Want specific formats? Change the active variant in the script:
```python
ACTIVE_VARIANT = "invoice"      # Extracts INV-2024-001
ACTIVE_VARIANT = "project"      # Extracts PROJECT-ALPHA
ACTIVE_VARIANT = "case"         # Extracts case #12345
```

Built-in variants: `invoice`, `project`, `contract`, `case`, `order`, `aerospace_code`

Default (`whole_subject_extract`) works for most emails.

---

## What You Get

Each email creates a folder with attachments + anything from `refs/`:
```
out/Client_Acme/
├── contract.pdf          (from email attachment)
├── invoice.pdf           (from email attachment)
├── status_template.xlsx  (from refs/ - always included)
└── archive/
    └── email.msg
```

**Why refs/?** Drop recurring files here once (templates, checklists, reference docs) instead of attaching them to every email.

---

## Active Variants

| Variant | Example Subject | Folder Created |
|---------|----------------|----------------|
| Default | `Acme Corp - Review` | `Acme_Corp/` |
| `invoice` | `Invoice: INV-2024-001` | `INV-2024-001/` |
| `project` | `Project: ALPHA-2024` | `ALPHA-2024/` |
| `case` | `Case #12345: Bug` | `12345/` |
| `contract` | `Contract: MSA-2024-Q1` | `MSA-2024-Q1/` |

Change active variant in `pytenberg.py` around line 62.

---

## Troubleshooting

**No .msg files found:**
- Save emails from Outlook: File → Save As → Outlook Message Format
- Drop them in `drop/` folder

**"No match in subject":**
- Default variant is very permissive
- Or switch explicitly: `ACTIVE_VARIANT = "whole_subject_extract"`

**Missing extract_msg:**
```bash
pip install extract-msg
```

---

## License

MIT

---

Built by Thomas Galarneau 10.17.2025
