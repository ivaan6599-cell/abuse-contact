"""
VS Code mein bas ye file kholo, neeche SETTINGS bharo, STEP badlo, aur ▶ Run dabao.
Pehli baar terminal mein ek hi baar:  pip install pandas openpyxl dnspython
"""
import os, argparse
from pathlib import Path

# ======================= SETTINGS (sirf yahan badlo) =======================
GMAIL         = "210490131057@gmail.com"
APP_PASSWORD  = "nleb yydp ynuo xflz"   # 16-letter code yahin paste karo (spaces chalenge)
YOUR_NAME     = "ivaan | mail test"
ORG_NAME      = "ivaan"
EXCEL_PATH    = "abuse_contacts.xlsx"          # isi folder mein rakho
SHEET_NAME    = "All Hosts"
EMAIL_COLUMN  = "Abuse Email"

# Pehle STEP = "test" chalao. Sab theek ho to "prep" -> "send" -> "collect" -> "report"
STEP          = "test"
# STEP options:
#   "login"    -> sirf Gmail login check karta hai (koi mail nahi). Pehle isse chalao.
#   "test"     -> sirf TEST_EMAILS ko mail bhejta hai (alag state file, real list ko nahi chhuta)
#   "test_collect" / "test_report"  -> test ka bounce padho / test ka Excel
#   "prep"     -> Excel se emails nikalo + domain check (koi mail nahi)
#   "dry"      -> preview, koi mail nahi jaati
#   "send"     -> SEND_LIMIT utni mails bhejo
#   "collect"  -> Gmail inbox se bounces padho (bhejne ke 1-3 din baad)
#   "report"   -> final Excel banao

SEND_LIMIT    = 50          # ek run mein max mails
DELAY_SEC     = 8           # do mails ke beech gap
TEST_EMAILS   = ["210490131057@gmail.com", "zz-no-such-user-84713@gmail.com"]   # ek real, ek nakli
# ===========================================================================

# --- sanity check (galat settings par saaf error) ---
APP_PASSWORD = APP_PASSWORD.replace(" ", "").strip()
if "@" not in GMAIL or not GMAIL.endswith("@gmail.com") or GMAIL.count("@") != 1:
    raise SystemExit(f"GMAIL galat lag raha hai: {GMAIL!r}")
if len(APP_PASSWORD) != 16 or not APP_PASSWORD.isalpha():
    raise SystemExit("APP_PASSWORD 16 letters ka hona chahiye (sirf a-z, spaces ke bina). Naya App Password banake daalo.")

os.environ.update({
    "SMTP_HOST": "smtp.gmail.com", "SMTP_PORT": "465", "SMTP_USER": GMAIL, "SMTP_PASS": APP_PASSWORD,
    "FROM_ADDR": GMAIL, "FROM_NAME": YOUR_NAME, "ORG_NAME": ORG_NAME, "REPLY_TO": GMAIL,
    "IMAP_HOST": "imap.gmail.com", "IMAP_PORT": "993", "IMAP_USER": GMAIL, "IMAP_PASS": APP_PASSWORD,
    "IMAP_FOLDER": "INBOX",
})
import abuse_verify as av
HERE = Path(__file__).parent
TEMPLATE = str(HERE / "template.txt")


def make_test_state():
    import dns.resolver, uuid, pandas as pd
    res = dns.resolver.Resolver(); res.lifetime = res.timeout = 6
    rows = []
    for e in TEST_EMAILS:
        d = e.split("@")[-1]
        r = {c: "" for c in av.COLS}
        r.update(email=e.lower(), domain=d, syntax_ok=str(bool(av.EMAIL_RE.match(e))),
                 mx_status=av.dns_check(d, res), token=av.TOKEN_PREFIX + uuid.uuid4().hex[:12])
        rows.append(r)
    import pandas as pd
    pd.DataFrame(rows)[av.COLS].to_csv(av.STATE, index=False)
    print("test state ready:", TEST_EMAILS)


def check_login():
    """Gmail login pehle 465 (SSL) par, fail ho to 587 par try karta hai. Saaf message deta hai."""
    import smtplib, ssl
    ctx = ssl.create_default_context()
    last = ""
    for port in (465, 587):
        try:
            if port == 465:
                s = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ctx, timeout=30)
            else:
                s = smtplib.SMTP("smtp.gmail.com", 587, timeout=30); s.starttls(context=ctx)
            s.login(GMAIL, APP_PASSWORD); s.quit()
            os.environ["SMTP_PORT"] = str(port)
            print(f"Gmail login OK (port {port})")
            return
        except smtplib.SMTPAuthenticationError as e:
            raise SystemExit("LOGIN FAIL: App Password ya Gmail address galat hai (535). "
                             "myaccount.google.com/apppasswords pe naya banao aur dobara paste karo.")
        except Exception as e:
            last = f"port {port}: {type(e).__name__}: {e}"
            print("try fail ->", last)
    raise SystemExit("LOGIN FAIL: Gmail tak pahunch nahi paa rahe (network/antivirus block). "
                     "Mobile hotspot se try karo, ya antivirus ka email-scan band karo.\nLast error: " + last)


if STEP in ("login", "test", "send"):
    check_login()

if STEP.startswith("test"):
    av.STATE = HERE / "state_test.csv"
    out = str(HERE / "test_result.xlsx")
else:
    out = str(HERE / "abuse_verification_result.xlsx")

if STEP == "login":
    pass
elif STEP == "test":
    make_test_state()
    av.cmd_send(argparse.Namespace(template=TEMPLATE, limit=10, delay=3, dry_run=False))
    print("\nAb 2-3 minute ruko, phir STEP = 'test_collect', phir 'test_report'.")
elif STEP in ("test_collect", "collect"):
    av.cmd_collect(argparse.Namespace(days=5))
elif STEP in ("test_report", "report"):
    av.cmd_report(argparse.Namespace(out=out))
elif STEP == "prep":
    av.cmd_prep(argparse.Namespace(input=str(HERE / EXCEL_PATH), sheet=SHEET_NAME, column=EMAIL_COLUMN, workers=40, dns=None))
elif STEP == "dry":
    av.cmd_send(argparse.Namespace(template=TEMPLATE, limit=SEND_LIMIT, delay=DELAY_SEC, dry_run=True))
elif STEP == "send":
    av.cmd_send(argparse.Namespace(template=TEMPLATE, limit=SEND_LIMIT, delay=DELAY_SEC, dry_run=False))
else:
    raise SystemExit(f"unknown STEP: {STEP}")