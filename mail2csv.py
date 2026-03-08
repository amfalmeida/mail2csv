#!/usr/bin/env python3
"""
Mail2CSV - Email to CSV Exporter
Connects to email via IMAP and exports all emails to a CSV file.

Configuration (all via environment variables):
  IMAP_SERVER          - IMAP server hostname (required)
  IMAP_PORT            - IMAP server port (default: 993)
  EMAIL                - Email address (required)
  PASSWORD             - Email password or app password (required)
  MAILBOX              - Specific mailbox to export (default: all)
  LIMIT                - Max emails per mailbox (default: unlimited)
  SINCE                - Only fetch emails on/after date (e.g. 01-Jan-2024)
  OUTPUT               - Output CSV filename (default: mail2csv.csv)

Usage:
  docker-compose up
  # or
  python mail2csv.py
"""

import imaplib
import email
from email.header import decode_header
import csv
import sys
import os

# ── helpers ──────────────────────────────────────────────────────────────────

def decode_str(value):
    """Decode an encoded email header string."""
    if value is None:
        return ""
    parts = decode_header(value)
    decoded = []
    for part, enc in parts:
        if isinstance(part, bytes):
            try:
                decoded.append(part.decode(enc or "utf-8", errors="replace"))
            except Exception:
                decoded.append(part.decode("utf-8", errors="replace"))
        else:
            decoded.append(str(part))
    return " ".join(decoded).strip()


def get_body(msg):
    """Extract plain-text body from an email.Message object."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                charset = part.get_content_charset() or "utf-8"
                try:
                    body = part.get_payload(decode=True).decode(charset, errors="replace")
                except Exception:
                    body = part.get_payload(decode=True).decode("utf-8", errors="replace")
                break
    else:
        charset = msg.get_content_charset() or "utf-8"
        try:
            body = msg.get_payload(decode=True).decode(charset, errors="replace")
        except Exception:
            body = ""
    # Collapse excessive whitespace
    return " ".join(body.split())[:2000]  # cap at 2 000 chars for readability


def get_attachments(msg):
    """Return a comma-separated list of attachment filenames."""
    names = []
    for part in msg.walk():
        cd = str(part.get("Content-Disposition", ""))
        if "attachment" in cd:
            fname = part.get_filename()
            if fname:
                names.append(decode_str(fname))
    return "; ".join(names)


def quote_mailbox(name):
    """Properly quote mailbox name for IMAP protocol (RFC 3501)."""
    # Escape backslashes and quotes, then wrap in quotes if needed
    if any(c in name for c in ' \\"()[]{}%*'):
        # Escape backslashes first, then quotes
        escaped = name.replace('\\', '\\\\').replace('"', '\\"')
        return f'"{escaped}"'
    return name


def list_mailboxes(imap):
    """List all available mailboxes for debugging."""
    try:
        status, boxes = imap.list()
        if status == 'OK':
            print("\n📋 Available mailboxes on server:")
            for box in boxes:
                print(f"  {box}")
            print()
    except Exception as e:
        print(f"Could not list mailboxes: {e}")


# ── core export ──────────────────────────────────────────────────────────────

def export_mailbox(imap, mailbox_name, writer, limit, since, seen_ids):
    """Fetch emails from a single mailbox and write rows to CSV."""
    try:
        # Try using imap.select with proper quoting using imaplib's mechanism
        # Construct the quoted mailbox name manually with literal syntax
        status, _ = imap.select('"{}"'.format(mailbox_name.replace('\\', '\\\\').replace('"', '\\"')), readonly=True)
        if status != "OK":
            print(f"  ⚠  Could not open '{mailbox_name}', skipping.")
            return 0
    except Exception as e:
        print(f"  ⚠  Error selecting '{mailbox_name}': {e}")
        return 0

    search_criteria = f'(SINCE "{since}")' if since else "ALL"
    status, data = imap.search(None, search_criteria)
    if status != "OK" or not data or not data[0]:
        print(f"  No messages found in '{mailbox_name}'.")
        return 0

    ids = data[0].split()
    if limit:
        ids = ids[-limit:]  # most recent first when limited

    count = 0
    total = len(ids)
    print(f"  Fetching {total} message(s) from '{mailbox_name}' …")

    for i, msg_id in enumerate(reversed(ids), 1):
        # De-duplicate across mailboxes via Message-ID header
        status, raw = imap.fetch(msg_id, "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)])")
        if status != "OK":
            continue
        header_data = raw[0][1] if raw and raw[0] else b""
        mid = email.message_from_bytes(header_data).get("Message-ID", "")
        if mid and mid in seen_ids:
            continue
        if mid:
            seen_ids.add(mid)

        # Fetch full message
        status, raw_full = imap.fetch(msg_id, "(RFC822)")
        if status != "OK" or not raw_full or raw_full[0] is None:
            continue

        msg = email.message_from_bytes(raw_full[0][1])

        writer.writerow({
            "mailbox":     mailbox_name,
            "message_id":  mid,
            "date":        decode_str(msg.get("Date", "")),
            "from":        decode_str(msg.get("From", "")),
            "to":          decode_str(msg.get("To", "")),
            "cc":          decode_str(msg.get("Cc", "")),
            "subject":     decode_str(msg.get("Subject", "")),
            "body":        get_body(msg),
        })
        count += 1

        if i % 50 == 0 or i == total:
            print(f"    {i}/{total} processed …", end="\r")

    print()
    return count


def main():
    imap_server = os.environ.get("IMAP_SERVER")
    imap_port = os.environ.get("IMAP_PORT", "993")
    email = os.environ.get("EMAIL")
    password = os.environ.get("PASSWORD")
    mailbox = os.environ.get("MAILBOX")
    limit = os.environ.get("LIMIT")
    since = os.environ.get("SINCE")
    output = os.environ.get("OUTPUT", "mail2csv.csv")

    if imap_port:
        imap_port = int(imap_port)

    if limit:
        limit = int(limit)

    if not imap_server or not email or not password:
        print("❌ Error: IMAP_SERVER, EMAIL, and PASSWORD are required.")
        print("Set them in your .env file or environment.")
        sys.exit(1)

    print("\n📧 Configuration:")
    print(f"   Server:    {imap_server}:{imap_port}")
    print(f"   Email:     {email}")
    print(f"   Mailbox:   {mailbox or 'all'}")
    print(f"   Limit:     {limit or 'unlimited'}")
    print(f"   Since:     {since or 'all dates'}")
    print(f"   Output:    {output}")
    print()

    print(f"Connecting to {imap_server} as {email} …")
    try:
        imap = imaplib.IMAP4_SSL(imap_server, imap_port)
        imap.login(email, password)
    except imaplib.IMAP4.error as e:
        print(f"\n❌ Login failed: {e}")
        print("\nTips:")
        print("  • Make sure IMAP is enabled in your email settings.")
        print("  • Use an App Password if your provider requires it.")
        sys.exit(1)

    print("✅ Connected!\n")

    # Debug: List all available mailboxes
    list_mailboxes(imap)

    fieldnames = ["mailbox", "message_id", "date", "from", "to", "cc", "subject", "body"]

    with open(output, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        seen_ids = set()

        if mailbox:
            mailboxes = [mailbox]
        else:
            status, boxes = imap.list()
            mailboxes = []
            for box in boxes:
                parts = box.decode().split('"/"')
                name = parts[-1].strip().strip('"')
                mailboxes.append(name)

        total_exported = 0
        for mb in mailboxes:
            print(f"📁 Mailbox: {mb}")
            n = export_mailbox(imap, mb, writer, limit, since, seen_ids)
            total_exported += n
            print(f"  ✓ {n} email(s) exported.\n")

    imap.logout()
    print(f"🎉 Done! {total_exported} unique email(s) exported to '{output}'.")


if __name__ == "__main__":
    main()
