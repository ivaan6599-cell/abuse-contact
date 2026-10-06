#!/usr/bin/env python3
"""
abuse_verify.py - verify abuse-contact emails via DNS + real send + bounce (DSN) collection.

Pipeline (state lives in state.csv, every step is resumable):
  prep     xlsx -> clean/dedupe/split emails, syntax + MX/A check      (no mail sent)
  send     send template mail one-by-one, throttled                    (--dry-run by default OFF, see flags)
  collect  read IMAP mailbox, parse bounces (DSN), update state
  report   write final xlsx (Send Mail Successfully / Bounce + extras)

Config: env vars or a .env file next to this script (see .env.example).
"""
import argparse, concurrent.futures as cf, email, imaplib, os, random, re, smtplib, ssl, sys, time, uuid
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from email.utils import formatdate, make_msgid, formataddr
from pathlib import Path

import pandas as pd

HERE = Path(__file__).parent
STATE = HERE / "state.csv"
TOKEN_PREFIX = "abv-"
TOKEN_RE = re.compile(r"abv-([0-9a-f]{12})")
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-']+@(?:[A-Za-z0-9\-]+\.)+[A-Za-z]{2,}$")
STATUS_RE = re.compile(r"\b([245])\.(\d{1,3})\.(\d{1,3})\b")
SOFT_WAIT_HOURS = 72

COLS = ["email", "domain", "syntax_ok", "mx_status", "token", "sent_at", "send_ok",
        "send_error", "bounce", "bounce_type", "bounce_code", "bounce_reason", "bounce_at"]


# ---------------------------------------------------------------- config
def load_env():
    p = HERE / ".env"
    if p.exists():
        for line in p.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def cfg(k, default=None, required=False):
    v = os.environ.get(k, default)
    if required and not v:
        sys.exit(f"missing config: {k}")
    return v


def load_state():
    if not STATE.exists():
        sys.exit("state.csv not found - run `prep` first")
    return pd.read_csv(STATE, dtype=str, keep_default_na=False)


def save_state(df):
    tmp = STATE.with_suffix(".tmp")
    df[COLS].to_csv(tmp, index=False)
    tmp.replace(STATE)


# ---------------------------------------------------------------- prep
def extract_emails(xlsx, sheet, col):
    df = pd.read_excel(xlsx, sheet_name=sheet, dtype=str)
    cols = [c for c in df.columns if c.strip().lower() in (col.lower(), "add new mail")]
    if not cols:
        sys.exit(f"column '{col}' not in sheet '{sheet}'. Columns: {list(df.columns)}")
    out = []
    for c in cols:
        for cell in df[c].dropna():
            for part in re.split(r"[,;\s/|]+", str(cell)):
                part = part.strip().strip("<>()[]\"'").lower()
                if "@" in part:
                    out.append(part)
    return sorted(set(out))


def dns_check(domain, resolver):
    import dns.exception, dns.resolver
    try:
        ans = resolver.resolve(domain, "MX")
        hosts = [str(r.exchange).rstrip(".") for r in ans]
        if hosts == [""] or all(h == "" for h in hosts):  # null MX (RFC 7505)
            return "null_mx"
        return "mx"
    except dns.resolver.NXDOMAIN:
        return "nxdomain"
    except dns.resolver.NoAnswer:
        try:
            resolver.resolve(domain, "A")
            return "a_only"
        except dns.resolver.NXDOMAIN:
            return "nxdomain"
        except Exception:
            return "no_mx"
    except (dns.exception.Timeout, dns.resolver.NoNameservers):
        return "dns_error"
    except Exception:
        return "dns_error"


def cmd_prep(a):
    import dns.resolver
    emails = extract_emails(a.input, a.sheet, a.column)
    print(f"unique emails extracted: {len(emails)}")
    res = dns.resolver.Resolver()
    res.lifetime = res.timeout = 6
    if a.dns:
        res.nameservers = a.dns.split(",")

    rows = {e: {"email": e, "domain": e.split("@")[-1], "syntax_ok": str(bool(EMAIL_RE.match(e)))} for e in emails}
    domains = sorted({r["domain"] for r in rows.values() if r["syntax_ok"] == "True"})
    print(f"unique domains to resolve: {len(domains)}")
    with cf.ThreadPoolExecutor(max_workers=a.workers) as ex:
        dmap = dict(zip(domains, ex.map(lambda d: dns_check(d, res), domains)))
    # retry transient dns errors once, serially
    for d, s in list(dmap.items()):
        if s == "dns_error":
            time.sleep(1)
            dmap[d] = dns_check(d, res)

    old = pd.read_csv(STATE, dtype=str, keep_default_na=False).set_index("email") if STATE.exists() else None
    out = []
    for e, r in rows.items():
        r["mx_status"] = dmap.get(r["domain"], "bad_syntax")
        r["token"] = TOKEN_PREFIX + uuid.uuid4().hex[:12]
        for c in COLS:
            r.setdefault(c, "")
        if old is not None and e in old.index:          # keep progress on re-run
            for c in COLS[4:]:
                r[c] = old.loc[e, c]
        out.append(r)
    df = pd.DataFrame(out)[COLS]
    save_state(df)
    print(df["mx_status"].value_counts().to_string())
    print(f"-> {STATE}")


# ---------------------------------------------------------------- send
def render_template(path, email_addr):
    raw = Path(path).read_text(encoding="utf-8")
    m = re.match(r"Subject:\s*(.+?)\r?\n\r?\n(.*)", raw, re.S)
    if not m:
        sys.exit("template must start with 'Subject: ...' then a blank line, then the body")
    subj, body = m.group(1).strip(), m.group(2)
    rep = {"{{email}}": email_addr, "{{domain}}": email_addr.split("@")[-1],
           "{{sender_name}}": cfg("FROM_NAME", ""), "{{org}}": cfg("ORG_NAME", "")}
    for k, v in rep.items():
        subj, body = subj.replace(k, v), body.replace(k, v)
    return subj, body


def build_msg(to, token, subj, body):
    from_addr = cfg("FROM_ADDR", required=True)
    m = EmailMessage()
    m["From"] = formataddr((cfg("FROM_NAME", ""), from_addr))
    m["To"] = to
    m["Subject"] = subj
    m["Date"] = formatdate(localtime=True)
    m["Message-ID"] = make_msgid(idstring=token, domain=from_addr.split("@")[-1])
    m["X-Tracking-ID"] = token
    if cfg("REPLY_TO"):
        m["Reply-To"] = cfg("REPLY_TO")
    m.set_content(f"{body}\n\n--\nRef: {token}")   # token also in body so DSNs quoting it can be matched
    return m


def cmd_send(a):
    df = load_state()
    ok_mx = df["mx_status"].isin(["mx", "a_only"]) & (df["syntax_ok"] == "True")
    todo = df[ok_mx & (df["sent_at"] == "")]
    if a.limit:
        todo = todo.head(a.limit)
    print(f"to send now: {len(todo)}  (already sent: {(df['sent_at'] != '').sum()})")
    if a.dry_run:
        for _, r in todo.head(3).iterrows():
            s, b = render_template(a.template, r["email"])
            print("---", r["email"], "|", s, "\n", b[:300])
        print("dry-run: nothing sent")
        return

    env_from = cfg("BOUNCE_ADDR") or cfg("FROM_ADDR", required=True)
    ctx = ssl.create_default_context()
    port = int(cfg("SMTP_PORT", "587"))

    def connect():
        s = smtplib.SMTP_SSL(cfg("SMTP_HOST", required=True), port, context=ctx, timeout=30) if port == 465 \
            else smtplib.SMTP(cfg("SMTP_HOST", required=True), port, timeout=30)
        if port != 465:
            s.starttls(context=ctx)
        s.login(cfg("SMTP_USER", required=True), cfg("SMTP_PASS", required=True))
        return s

    smtp = connect()
    sent = 0
    for idx, r in todo.iterrows():
        subj, body = render_template(a.template, r["email"])
        msg = build_msg(r["email"], r["token"], subj, body)
        i = df.index[df["email"] == r["email"]][0]
        try:
            smtp.send_message(msg, from_addr=env_from, to_addrs=[r["email"]])
            df.loc[i, ["send_ok", "sent_at"]] = ["True", datetime.now(timezone.utc).isoformat(timespec="seconds")]
            sent += 1
            print(f"[{sent}/{len(todo)}] sent  {r['email']}")
        except smtplib.SMTPRecipientsRefused as e:       # synchronous reject = immediate hard/soft bounce
            code, reason = list(e.recipients.values())[0]
            reason = reason.decode(errors="replace") if isinstance(reason, bytes) else str(reason)
            df.loc[i, ["send_ok", "sent_at", "bounce", "bounce_type", "bounce_code", "bounce_reason", "bounce_at"]] = [
                "False", datetime.now(timezone.utc).isoformat(timespec="seconds"), "True",
                "hard" if 500 <= code < 600 else "soft", str(code), reason[:300],
                datetime.now(timezone.utc).isoformat(timespec="seconds")]
            print(f"[reject] {r['email']} {code} {reason[:80]}")
        except (smtplib.SMTPServerDisconnected, smtplib.SMTPSenderRefused, smtplib.SMTPDataError) as e:
            df.loc[i, ["send_ok", "send_error"]] = ["False", f"{type(e).__name__}: {e}"[:300]]
            print(f"[error] {r['email']} {e}")
            if isinstance(e, smtplib.SMTPSenderRefused) or "quota" in str(e).lower() or "limit" in str(e).lower():
                save_state(df)
                sys.exit("sender refused / rate-limited - stopping. Lower --limit or wait.")
            try: smtp.close()
            except Exception: pass
            smtp = connect()
        except Exception as e:
            df.loc[i, ["send_ok", "send_error"]] = ["False", f"{type(e).__name__}: {e}"[:300]]
            print(f"[error] {r['email']} {e}")
        save_state(df)
        time.sleep(a.delay + random.uniform(0, a.delay * 0.4))
    try: smtp.quit()
    except Exception: pass


# ---------------------------------------------------------------- collect
def parse_bounce(msg):
    """Return dict(token, rcpt, code, reason) if msg looks like a bounce, else None."""
    sender = (msg.get("From", "") + msg.get("Return-Path", "")).lower()
    ctype = msg.get_content_type()
    is_dsn = ctype == "multipart/report" or "delivery-status" in str(msg.get("Content-Type", "")).lower()
    looks = is_dsn or any(k in sender for k in ("mailer-daemon", "postmaster", "mail delivery"))
    subj = str(msg.get("Subject", "")).lower()
    looks = looks or any(k in subj for k in ("undeliver", "delivery status", "returned mail", "failure notice", "delivery has failed"))
    if not looks:
        return None

    text_parts, rcpt, code, reason = [], None, None, None
    for part in msg.walk():
        pt = part.get_content_type()
        if pt == "message/delivery-status":
            for blk in part.get_payload():           # list of header-blocks
                fr = blk.get("Final-Recipient") or blk.get("Original-Recipient")
                if fr and not rcpt:
                    rcpt = fr.split(";")[-1].strip().strip("<>").lower()
                st = blk.get("Status")
                if st and not code:
                    code = st.strip()
                dc = blk.get("Diagnostic-Code")
                if dc and not reason:
                    reason = re.sub(r"\s+", " ", dc.split(";", 1)[-1]).strip()
        elif pt in ("text/plain", "message/rfc822", "text/rfc822-headers", "text/html"):
            try:
                pl = part.get_payload(decode=True)
                text_parts.append(pl.decode("utf-8", "replace") if pl else str(part.get_payload()))
            except Exception:
                text_parts.append(str(part.get_payload()))
    blob = "\n".join(text_parts) + "\n" + msg.as_string()
    tm = TOKEN_RE.search(blob)
    if not code:
        sm = STATUS_RE.search("\n".join(text_parts))
        code = sm.group(0) if sm else None
    if not reason:
        reason = re.sub(r"\s+", " ", "\n".join(text_parts))[:300].strip()
    if not code or not code[0] in "245":
        code = code or ""
    if code.startswith("2"):                          # success DSN (delivered/relayed) - not a bounce
        return None
    return {"token": tm.group(0) if tm else None, "rcpt": rcpt, "code": code, "reason": (reason or "")[:300]}


def cmd_collect(a):
    df = load_state()
    by_token = {t: i for i, t in zip(df.index, df["token"])}
    by_email = {e: i for i, e in zip(df.index, df["email"])}
    M = imaplib.IMAP4_SSL(cfg("IMAP_HOST", required=True), int(cfg("IMAP_PORT", "993")))
    M.login(cfg("IMAP_USER", required=True), cfg("IMAP_PASS", required=True))
    M.select(cfg("IMAP_FOLDER", "INBOX"), readonly=True)
    since = (datetime.now() - timedelta(days=a.days)).strftime("%d-%b-%Y")
    _, data = M.search(None, "SINCE", since)
    ids = data[0].split()
    print(f"scanning {len(ids)} messages since {since}")
    hits = unmatched = 0
    for mid in ids:
        _, md = M.fetch(mid, "(RFC822)")
        msg = email.message_from_bytes(md[0][1])
        b = parse_bounce(msg)
        if not b:
            continue
        i = by_token.get(b["token"]) if b["token"] else None
        if i is None and b["rcpt"]:
            i = by_email.get(b["rcpt"])
        if i is None:
            unmatched += 1
            continue
        if df.loc[i, "sent_at"] == "":
            continue
        btype = "hard" if b["code"].startswith("5") else "soft" if b["code"].startswith("4") else "unknown"
        df.loc[i, ["bounce", "bounce_type", "bounce_code", "bounce_reason", "bounce_at"]] = [
            "True", btype, b["code"], b["reason"], datetime.now(timezone.utc).isoformat(timespec="seconds")]
        hits += 1
    M.logout()
    save_state(df)
    print(f"bounces matched: {hits} | bounce-looking but unmatched: {unmatched}")


# ---------------------------------------------------------------- report
def cmd_report(a):
    df = load_state()
    now = datetime.now(timezone.utc)

    def status(r):
        if r["syntax_ok"] != "True": return "invalid_syntax"
        if r["mx_status"] in ("nxdomain", "no_mx", "null_mx"): return "domain_cannot_receive_mail"
        if r["mx_status"] == "dns_error": return "dns_error_retry"
        if r["bounce"] == "True": return f"{r['bounce_type']}_bounce"
        if r["send_ok"] == "True":
            age = now - datetime.fromisoformat(r["sent_at"])
            return "delivered_no_bounce" if age >= timedelta(hours=SOFT_WAIT_HOURS) else "sent_waiting_for_bounce_window"
        if r["send_ok"] == "False": return "send_failed"
        return "not_sent_yet"

    out = pd.DataFrame({
        "Email": df["email"], "Domain": df["domain"], "MX Status": df["mx_status"],
        "Send Mail Successfully": df["send_ok"].map(lambda x: "Yes" if x == "True" else "No"),
        "Bounce": df["bounce"].map(lambda x: "Yes" if x == "True" else "No"),
        "Bounce Type": df["bounce_type"], "Bounce Code": df["bounce_code"],
        "Bounce Reason": df["bounce_reason"], "Sent At (UTC)": df["sent_at"],
        "Final Status": df.apply(status, axis=1),
    })
    with pd.ExcelWriter(a.out, engine="openpyxl") as w:
        out.to_excel(w, sheet_name="Results", index=False)
        out["Final Status"].value_counts().rename_axis("Final Status").reset_index(name="Count").to_excel(w, sheet_name="Summary", index=False)
        ws = w.sheets["Results"]
        ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
        for col, wd in zip("ABCDEFGHIJ", (36, 26, 12, 14, 9, 11, 11, 60, 22, 30)):
            ws.column_dimensions[col].width = wd
    print(out["Final Status"].value_counts().to_string())
    print(f"-> {a.out}")


# ---------------------------------------------------------------- main
def main():
    load_env()
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = p.add_subparsers(dest="cmd", required=True)

    s = sp.add_parser("prep"); s.add_argument("input"); s.add_argument("--sheet", default="All Hosts")
    s.add_argument("--column", default="Abuse Email"); s.add_argument("--workers", type=int, default=40)
    s.add_argument("--dns", help="comma-separated resolvers, e.g. 1.1.1.1,8.8.8.8"); s.set_defaults(f=cmd_prep)

    s = sp.add_parser("send"); s.add_argument("--template", default=str(HERE / "template.txt"))
    s.add_argument("--limit", type=int, default=50, help="max mails this run (default 50; keep daily volume low)")
    s.add_argument("--delay", type=float, default=8, help="base seconds between mails (+0-40%% jitter)")
    s.add_argument("--dry-run", action="store_true"); s.set_defaults(f=cmd_send)

    s = sp.add_parser("collect"); s.add_argument("--days", type=int, default=5); s.set_defaults(f=cmd_collect)

    s = sp.add_parser("report"); s.add_argument("--out", default="abuse_verification_result.xlsx"); s.set_defaults(f=cmd_report)

    a = p.parse_args(); a.f(a)


if __name__ == "__main__":
    main()
