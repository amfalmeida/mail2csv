# Mail2CSV

Exports all your emails (every folder/label) to a single CSV file.  
Works with any IMAP-compatible email provider (Gmail, Outlook, Fastmail, Proton Mail, etc.).  
Uses Python's built-in `imaplib` — **no third-party packages needed**.

---

## 1 — Enable IMAP in Your Email Provider

### Gmail
1. Open Gmail → **Settings** (⚙) → **See all settings**
2. Go to **Forwarding and POP/IMAP**
3. Under *IMAP access*, select **Enable IMAP** → Save

### Outlook
1. Go to **Settings** → **Mail** → **Sync email**
2. Enable **IMAP**

### Other Providers
Look for "IMAP" or "Remote mailbox access" in your email settings.

---

## 2 — App Password (if required)

Many providers require an App Password instead of your regular password:

| Provider | How to get an App Password |
|----------|---------------------------|
| Gmail | [myaccount.google.com/security](https://myaccount.google.com/security) → 2-Step Verification → App passwords |
| Outlook | **Settings** → **Security** → **Sign in** → **App passwords** |
| Fastmail | **Settings** → **Security** → **App-specific passwords** |

---

## 3 — Configuration

Create a `.env` file in this directory:

```bash
cp .env.example .env
```

Edit `.env` with your credentials:

| Variable | Description | Required | Default |
|----------|-------------|----------|---------|
| `IMAP_SERVER` | IMAP server hostname | Yes | - |
| `IMAP_PORT` | IMAP server port | No | 993 |
| `EMAIL` | Email address | Yes | - |
| `PASSWORD` | Password or App Password | Yes | - |
| `MAILBOX` | Specific mailbox to export | No | all folders |
| `LIMIT` | Max emails per mailbox | No | unlimited |
| `SINCE` | Only emails on/after date (e.g. 01-Jan-2024) | No | all dates |
| `OUTPUT` | Output CSV filename | No | mail2csv.csv |

### Common IMAP Settings

| Provider | IMAP Server | Port |
|----------|-------------|------|
| Gmail | imap.gmail.com | 993 |
| Outlook | outlook.office365.com | 993 |
| Yahoo | imap.mail.yahoo.com | 993 |
| Fastmail | imap.fastmail.com | 993 |
| Proton Mail | 127.0.0.1 (use Proton Mail Bridge) | 1143 |

---

## 4 — Run with Docker Compose (recommended)

```bash
docker-compose up --build
```

The CSV will be saved to `./output/mail2csv.csv`.

---

## 5 — Run with Docker (manual)

### Build
```bash
docker build -t mail2csv .
```

### Run
```bash
docker run --rm \
  -v "$(pwd)/output":/output \
  --env-file .env \
  mail2csv
```

The CSV will appear in `./output/mail2csv.csv`.

### Example: only INBOX, last 200 emails, since 2024
```bash
IMAP_SERVER=imap.gmail.com EMAIL=user@gmail.com PASSWORD=xxxx MAILBOX=INBOX LIMIT=200 SINCE=01-Jan-2024 docker-compose up --build
```

---

## 6 — Run directly with Python (no Docker)

```bash
pip install -r requirements.txt
export IMAP_SERVER=imap.gmail.com
export EMAIL=you@gmail.com
export PASSWORD=your_app_password
python mail2csv.py
```

---

## 7 — Build Windows Executable

Use GitHub Actions (recommended):

1. Push this repository to GitHub
2. Go to **Actions** → **Build Windows Executable** → **Run workflow**
3. Download the artifact from the workflow run

The executable will be in `dist/mail2csv.exe`.

### Running the executable

Create a `.env` file next to the executable, then run:

```bash
set IMAP_SERVER=imap.gmail.com
set EMAIL=you@gmail.com
set PASSWORD=your_app_password
mail2csv.exe
```

---

## CSV columns

| Column | Description |
|--------|-------------|
| `mailbox` | Email folder/label |
| `message_id` | Unique Message-ID header |
| `date` | Sent date |
| `from` | Sender |
| `to` | Recipients |
| `cc` | CC recipients |
| `subject` | Subject line |
| `body` | Plain-text body (capped at 2 000 chars) |
| `attachments` | Semicolon-separated attachment filenames |

Duplicate emails that appear in multiple labels are exported **once**.
