#!/usr/bin/env python3
"""
run_daily.py — End-to-end Ventas daily dashboard pipeline.

Steps each run:
  1. Find the LATEST "VENTAS REPORT" email and download its CSV attachments.
  2. Generate the dark-themed dashboard PPTX (ventas_slide_generator.py, CSV mode).
  3. (optional) Upload the PPTX to Google Drive.
  4. Email the PPTX (as an attachment) with the slide shown INLINE in the body.

Single OAuth token (token_ventas.json) covers Gmail read, Gmail send, and Drive.
Reuses the existing credentials.json (OAuth Desktop client).

ONE-TIME SETUP (on a machine with a browser, e.g. your Windows box):
    1. In Google Cloud Console, enable Gmail API + Drive API on the project.
    2. Put the OAuth Desktop credentials.json in this folder.
    3. Run once:   python run_daily.py --auth
       (browser opens -> sign in -> approve). Writes token_ventas.json.

INLINE SLIDE IMAGE:
    The slide is rendered to PNG and embedded in the email body. On Windows this
    uses PowerPoint (pip install comtypes); otherwise LibreOffice (soffice) is
    used if present. If neither is available, the email is still sent with the
    .pptx attached, just without the inline preview.

DAILY USAGE (what the 10 AM job runs):
    python run_daily.py --to raj.mohan@6dtech.co.in

OPTIONS:
    --auth            Do the one-time sign-in only.
    --to A,B          Comma-separated recipients (default raj.mohan@6dtech.co.in).
    --no-drive        Skip the Google Drive upload.
    --no-image        Skip the inline slide image (attachment only).
    --date YYYY-MM-DD Override report date (default: from the email).
    --query Q         Gmail search query for the report email.
"""

import os, sys, re, base64, argparse, subprocess, shutil
from pathlib import Path
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders

HERE       = Path(__file__).parent
GOOGLE_HOSTS = ["oauth2.googleapis.com", "gmail.googleapis.com", "www.googleapis.com"]
SCOPES     = ["https://www.googleapis.com/auth/gmail.readonly",
              "https://www.googleapis.com/auth/gmail.send",
              "https://www.googleapis.com/auth/drive.file"]
CREDS_FILE = HERE / "credentials.json"
TOKEN_FILE = HERE / "token_ventas.json"
DRIVE_FOLDER_ID = "1yCgVqztLDc79kACnyA4y5WYQIGOEuEaj"
DEFAULT_TO = "raj.mohan@6dtech.co.in"
DEFAULT_QUERY = 'subject:"VENTAS REPORT" has:attachment'
GENERATOR  = HERE / "ventas_slide_generator.py"
WORKDIR    = Path(os.environ.get("VENTAS_WORKDIR", str(HERE)))   # writable dir (Cloud Run: /tmp)
ENV_TO     = os.environ.get("MAIL_TO",  "")
ENV_CC     = os.environ.get("MAIL_CC",  "")
ENV_BCC    = os.environ.get("MAIL_BCC", "")
MIME_PPTX  = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
INLINE_CID  = "ventas_slide"
INLINE_CID1 = "ventas_slide1"
INLINE_CID2 = "ventas_slide2"


def check_connectivity(timeout=8):
    """Verify Google's API endpoints are reachable before doing real work.

    Returns (ok: bool, detail: str). Any HTTP response (even 4xx) counts as
    reachable; only connection/proxy/DNS failures count as unreachable.
    """
    import urllib.request, urllib.error
    last = "no hosts checked"
    for host in GOOGLE_HOSTS:
        try:
            urllib.request.urlopen("https://%s/" % host, timeout=timeout)
            return True, host
        except urllib.error.HTTPError:
            return True, host
        except Exception as e:
            last = "%s: %s" % (host, e)
            continue
    return False, last


def _looks_like_network(exc):
    s = (str(exc) + " " + repr(exc)).lower()
    return any(k in s for k in ("proxy", "tunnel", "max retries", "connection",
                                "timed out", "timeout", "getaddrinfo",
                                "name or service", "transporterror", "403 forbidden"))


def _creds_from_env():
    """Unattended credentials from environment (Cloud Run / Secret Manager).

    Priority:
      1. Service account + domain-wide delegation:
         SA_KEY_JSON (service-account key JSON) + GMAIL_DELEGATED_USER (mailbox to impersonate)
      2. User OAuth token JSON in TOKEN_JSON (content of token_ventas.json)
    Returns Credentials or None (fall back to local-file flow).
    """
    import json
    sa_key = os.environ.get("SA_KEY_JSON")
    subject = os.environ.get("GMAIL_DELEGATED_USER")
    if sa_key and subject:
        from google.oauth2 import service_account
        info = json.loads(sa_key)
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=SCOPES).with_subject(subject)
        print("Auth: service account + domain-wide delegation as", subject)
        return creds
    token_json = os.environ.get("TOKEN_JSON")
    if token_json:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        creds = Credentials.from_authorized_user_info(json.loads(token_json), SCOPES)
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        print("Auth: user OAuth token from environment")
        return creds
    return None


def get_creds():
    env_creds = _creds_from_env()
    if env_creds is not None:
        return env_creds
    try:
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
    except ImportError:
        print("ERROR: pip install google-auth-oauthlib google-api-python-client")
        sys.exit(1)
    creds = None
    if TOKEN_FILE.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
        except (ValueError, KeyError) as e:
            print("ERROR: token_ventas.json is unreadable/corrupted (%s)." % e)
            print("Fix: re-run one-time sign-in ->  python run_daily.py --auth")
            sys.exit(4)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                if _looks_like_network(e):
                    print("ERROR: Google is unreachable - cannot refresh token.")
                    print("Detail:", e)
                    print("This usually means no internet / blocked proxy on the "
                          "machine running the job. Run it where Google APIs are "
                          "reachable (e.g. your Windows box).")
                    sys.exit(5)
                raise
        else:
            if not CREDS_FILE.exists():
                print(f"ERROR: {CREDS_FILE} not found.")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())
        print(f"Token saved to {TOKEN_FILE}")
    return creds


def _svc(name, ver, creds):
    from googleapiclient.discovery import build
    return build(name, ver, credentials=creds)


# -- render slide 1 of the deck to a PNG (for the inline email image) --
def slide_to_png(pptx_path, out_png):
    """Return Path to a PNG of slide 1, or None if rendering isn't possible.
    Tries PowerPoint (Windows, via comtypes) first, then LibreOffice."""
    pptx_path = Path(pptx_path); out_png = Path(out_png)
    # 1) PowerPoint COM automation (Windows)
    try:
        import comtypes.client
        ppt = comtypes.client.CreateObject("PowerPoint.Application")
        try:
            pres = ppt.Presentations.Open(str(pptx_path), WithWindow=False)
            pres.Slides(1).Export(str(out_png), "PNG", 1920, 1080)
            pres.Close()
        finally:
            ppt.Quit()
        if out_png.exists():
            print("Slide image rendered via PowerPoint.")
            return out_png
    except Exception as e:
        print(f"(PowerPoint render unavailable: {e})")
    # 2) LibreOffice headless (CI/Linux). Prefer PDF -> PNG for a sharp image.
    try:
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            outdir = out_png.parent
            prof = "-env:UserInstallation=file://" + str(outdir / ".lo_profile")
            pdftoppm = shutil.which("pdftoppm")
            if pdftoppm:
                subprocess.check_call([soffice, prof, "--headless", "--convert-to", "pdf",
                                       "--outdir", str(outdir), str(pptx_path)])
                pdf = outdir / (pptx_path.stem + ".pdf")
                if pdf.exists():
                    subprocess.check_call([pdftoppm, "-png", "-r", "150", "-singlefile",
                                           str(pdf), str(out_png.with_suffix(""))])
                    if out_png.exists():
                        print("Slide image rendered via LibreOffice + pdftoppm.")
                        return out_png
            # Fallback: direct PNG export
            subprocess.check_call([soffice, prof, "--headless", "--convert-to", "png",
                                   "--outdir", str(outdir), str(pptx_path)])
            produced = outdir / (pptx_path.stem + ".png")
            if produced.exists():
                if produced.resolve() != out_png.resolve():
                    produced.replace(out_png)
                print("Slide image rendered via LibreOffice.")
                return out_png
        else:
            print("(LibreOffice not found on PATH)")
    except Exception as e:
        print(f"(LibreOffice render unavailable: {e})")
    print("WARN: could not render slide image; sending without inline preview.")
    return None


# -- 1. fetch latest CSVs --
def _report_date(msg):
    def walk(part):
        if part.get("body", {}).get("data"):
            try:
                txt = base64.urlsafe_b64decode(part["body"]["data"]).decode("utf-8", "ignore")
                m = re.search(r"Report Date\s*:\s*(?:</b>)?\s*(\d{4}-\d{2}-\d{2})", txt)
                if m: return m.group(1)
            except Exception: pass
        for p in part.get("parts", []) or []:
            r = walk(p)
            if r: return r
        return None
    rd = walk(msg["payload"])
    if rd: return rd
    ts = int(msg.get("internalDate", "0")) / 1000
    return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")


def _iter_csv_parts(part):
    if part.get("filename", "").lower().endswith(".csv"):
        yield part
    for p in part.get("parts", []) or []:
        yield from _iter_csv_parts(p)


def fetch_latest(gmail, query, outbase="csv"):
    resp = gmail.users().messages().list(userId="me", q=query, maxResults=5).execute()
    msgs = resp.get("messages", [])
    if not msgs:
        print(f"No messages match: {query}"); sys.exit(2)
    msg = gmail.users().messages().get(userId="me", id=msgs[0]["id"], format="full").execute()
    rd = _report_date(msg)
    dest = WORKDIR / outbase / rd
    dest.mkdir(parents=True, exist_ok=True)
    saved = []
    for part in _iter_csv_parts(msg["payload"]):
        aid = part["body"].get("attachmentId")
        if not aid: continue
        att = gmail.users().messages().attachments().get(
            userId="me", messageId=msgs[0]["id"], id=aid).execute()
        fp = dest / part["filename"]
        fp.write_bytes(base64.urlsafe_b64decode(att["data"]))
        saved.append(fp)
    if not saved:
        print("No .csv attachments on the latest message."); sys.exit(3)
    print(f"Report date {rd}: downloaded {len(saved)} CSV(s) -> {dest}")
    return dest, rd, saved


def _pick(paths, key):
    for p in paths:
        if key in p.name.upper(): return str(p)
    return None


# -- 2. generate --
def generate(csv_paths, report_date, out_path, png_path=None, table_png_path=None):
    pri = _pick(csv_paths, "PRIMARY"); pos = _pick(csv_paths, "POS")
    stk = _pick(csv_paths, "STOCK") or _pick(csv_paths, "TRANSFER")
    args = [sys.executable, str(GENERATOR), "--csv-file"]
    # pass EVERY downloaded CSV so all summary sections are included; keep the
    # three core ones first so slide-1 classification is unambiguous.
    core=[p for p in (pri, pos, stk) if p]
    extra=[str(p) for p in csv_paths if str(p) not in core]
    args += core + extra
    args += ["--date", report_date, "--output", str(out_path)]
    if png_path:
        args += ["--png", str(png_path)]
    if table_png_path:
        args += ["--table-png", str(table_png_path)]
    print("Generating deck:", " ".join(args))
    subprocess.check_call(args, cwd=str(HERE))
    return out_path


# -- 3. drive upload --
def upload_drive(creds, path, folder_id=DRIVE_FOLDER_ID):
    from googleapiclient.http import MediaFileUpload
    drive = _svc("drive", "v3", creds)
    meta = {"name": Path(path).name, "parents": [folder_id]}
    media = MediaFileUpload(str(path), mimetype=MIME_PPTX, resumable=True)
    f = drive.files().create(body=meta, media_body=media,
                             fields="id,webViewLink").execute()
    link = f.get("webViewLink", f"https://drive.google.com/file/d/{f['id']}/view")
    print("Uploaded to Drive:", link)
    return link


# -- 4. send email (inline slide image + pptx attachment) --
def send_email(gmail, to_list, subject, html, attach_path,
               cc_list=None, bcc_list=None, inline_imgs=None):
    outer = MIMEMultipart("mixed")
    outer["to"] = ", ".join(to_list)
    if cc_list:
        outer["cc"] = ", ".join(cc_list)
    if bcc_list:
        outer["bcc"] = ", ".join(bcc_list)
    outer["subject"] = subject

    related = MIMEMultipart("related")
    related.attach(MIMEText(html, "html"))
    for cid, png in (inline_imgs or []):
        if png and Path(png).exists():
            img = MIMEImage(Path(png).read_bytes(), _subtype="png")
            img.add_header("Content-ID", f"<{cid}>")
            img.add_header("Content-Disposition", "inline", filename=Path(png).name)
            related.attach(img)
    outer.attach(related)

    part = MIMEBase(*MIME_PPTX.split("/"))
    part.set_payload(Path(attach_path).read_bytes())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment", filename=Path(attach_path).name)
    outer.attach(part)

    raw = base64.urlsafe_b64encode(outer.as_bytes()).decode()
    sent = gmail.users().messages().send(userId="me", body={"raw": raw}).execute()
    _all = to_list + (cc_list or []) + (bcc_list or [])
    print(f"Email sent to {', '.join(_all)} (id={sent.get('id')})")
    return sent.get("id")


def _email_html(report_date, drive_link=None, images=None):
    link = f'<p>Drive copy: <a href="{drive_link}">{drive_link}</a></p>' if drive_link else ""
    blocks = ""
    for cid, label in (images or []):
        blocks += (f'<p style="margin:16px 0 4px;font-weight:bold;color:#0E518C;font-size:14px;">{label}</p>'
                   f'<p><img src="cid:{cid}" alt="{label}" '
                   f'style="width:100%;max-width:1000px;border:1px solid #ddd;border-radius:4px;"></p>')
    return f"""\
<div style="font-family:Arial,sans-serif;color:#0A2A57;">
  <h2 style="color:#0E518C;margin-bottom:4px;">Ventas Daily Dashboard</h2>
  <p style="margin:2px 0;color:#444;">Airtel Congo (CG) &nbsp;|&nbsp; Report date {report_date} &nbsp;|&nbsp; Last 7 days (excl. report day)</p>
  <p>Please find the Ventas USDM 2.0 daily report below. The editable deck is attached.</p>
  {blocks}
  {link}
  <p style="color:#888;font-size:12px;">Generated automatically by Ventas USDM 2.0.</p>
</div>"""


def main():
    ap = argparse.ArgumentParser(description="Ventas daily dashboard pipeline")
    ap.add_argument("--auth", action="store_true")
    ap.add_argument("--to", default=ENV_TO or DEFAULT_TO)
    ap.add_argument("--cc", default=ENV_CC)
    ap.add_argument("--bcc", default=ENV_BCC)
    ap.add_argument("--no-drive", action="store_true")
    ap.add_argument("--no-image", action="store_true")
    ap.add_argument("--date", default="")
    ap.add_argument("--query", default=DEFAULT_QUERY)
    args = ap.parse_args()

    ok, detail = check_connectivity()
    if not ok:
        print("ERROR: Google API endpoints are UNREACHABLE from this machine.")
        print("Last attempt -> %s" % detail)
        print("The Ventas job needs internet access to Gmail/Drive APIs. Run it "
              "on a machine/network where Google is reachable (e.g. your Windows "
              "box via run_daily.bat), not inside a restricted/sandboxed runner.")
        sys.exit(6)
    print("Connectivity OK (%s reachable)." % detail)

    creds = get_creds()
    if args.auth:
        print("Auth OK. token_ventas.json is ready."); return

    gmail = _svc("gmail", "v1", creds)
    csv_dir, rd, csvs = fetch_latest(gmail, args.query)
    report_date = args.date or rd
    out = WORKDIR / f"Ventas_Dashboard_{report_date}.pptx"
    png1 = WORKDIR / f"Ventas_Dashboard_{report_date}_slide1.png"   # dashboard
    png2 = WORKDIR / f"Ventas_Dashboard_{report_date}_slide2.png"   # trend table
    generate(csvs, report_date, out,
             png_path=(None if args.no_image else png1),
             table_png_path=(None if args.no_image else png2))

    images = []        # (cid, label, path) for inline embedding
    if not args.no_image:
        if not png1.exists():
            slide_to_png(out, png1)   # fallback to PowerPoint/LibreOffice for slide 1
        if png1.exists(): images.append((INLINE_CID1, "Daily Dashboard", png1))
        if png2.exists(): images.append((INLINE_CID2, "Weekly Transaction Trend", png2))
        print("Inline images:", [str(p) for _,_,p in images])

    drive_link = None
    if not args.no_drive:
        try:
            drive_link = upload_drive(creds, out)
        except Exception as e:
            print(f"WARN: Drive upload failed ({e}); continuing to email.")

    to_list  = [a.strip() for a in args.to.split(",")  if a.strip()]
    cc_list  = [a.strip() for a in args.cc.split(",")  if a.strip()]
    bcc_list = [a.strip() for a in args.bcc.split(",") if a.strip()]
    subject = f"Ventas Daily Dashboard | Airtel Congo (CG) | {report_date} (Last 7 Days)"
    html = _email_html(report_date, drive_link, images=[(c,l) for c,l,_ in images])
    send_email(gmail, to_list, subject, html, out, cc_list, bcc_list,
               inline_imgs=[(c,p) for c,_,p in images])
    print("DONE.")


if __name__ == "__main__":
    main()
