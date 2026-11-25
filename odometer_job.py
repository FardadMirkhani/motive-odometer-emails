"""Cloud version (GitHub Actions) of the Mon/Wed/Fri shop odometer emails.

Each shop gets ONLY the trucks that the Ryder Truck Mapping file (read fresh from Dropbox
on every run) assigns to that shop.

Environment (GitHub secrets):
  MOTIVE_API_KEY
  DROPBOX_APP_KEY, DROPBOX_REFRESH_TOKEN
  RESEND_API_KEY   (email via Resend; preferred)
    or MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET (email via Microsoft Graph)
  RECIPIENTS_JSON  (sender + addresses, kept out of the repo), e.g.
    {"sender": "...", "cc": ["..."],
     "shops": {"Jack's Diesel": {"to": ["..."]}, "Ryder Fontana": {"to": ["..."], "cc": []}}}
Usage:
  python odometer_job.py --mode dry-run                 # build everything, send nothing
  python odometer_job.py --mode test --test-to x@y.com  # every shop email to x only, no CC
  python odometer_job.py --mode send                    # real send
"""
import argparse
import html
import io
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
import warnings
from datetime import datetime
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

warnings.filterwarnings("ignore", module="openpyxl")
HERE = os.path.dirname(os.path.abspath(__file__))
PACIFIC = ZoneInfo("America/Los_Angeles")


# ---------- HTTP helpers ----------
def http(url, method="GET", headers=None, data=None, form=None):
    body = None
    headers = dict(headers or {})
    if form is not None:
        body = urllib.parse.urlencode(form).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    elif data is not None:
        body = data if isinstance(data, bytes) else json.dumps(data).encode()
        headers.setdefault("Content-Type", "application/json")
    req = urllib.request.Request(url, data=body, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=90) as r:
            return r.read()
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"{method} {url.split('?')[0]} -> {e.code}: {e.read()[:500]!r}") from None


def jget(url, headers):
    return json.loads(http(url, headers=headers))


# ---------- Motive ----------
def motive_paged(path, key, list_key):
    out, page = [], 1
    while True:
        d = jget(f"https://api.gomotive.com{path}?per_page=100&page_no={page}", {"X-Api-Key": key})
        items = d.get(list_key) or []
        out.extend(items)
        if not items or len(out) >= (d.get("pagination") or {}).get("total", len(out)):
            return out
        page += 1


def motive_rows(key, group_name):
    groups = [g["group"] for g in motive_paged("/v1/groups", key, "groups")]
    match = [g for g in groups if g["name"].strip().lower() == group_name.strip().lower()]
    if not match:
        raise RuntimeError(f"Motive group '{group_name}' not found; have: {[g['name'] for g in groups]}")
    gid = match[0]["id"]
    vehicles = [v["vehicle"] for v in motive_paged("/v1/vehicles", key, "vehicles")]
    members = [v for v in vehicles if gid in (v.get("group_ids") or [])]
    locs = {i["vehicle"]["id"]: (i["vehicle"].get("current_location") or {})
            for i in motive_paged("/v2/vehicle_locations", key, "vehicles")}
    rows = []
    for v in members:
        loc = locs.get(v["id"], {})
        odo = loc.get("true_odometer")
        if odo is None:
            odo = loc.get("odometer")  # miles for this account (metric_units=false)
        ts = loc.get("located_at")
        when = datetime.fromisoformat(ts.replace("Z", "+00:00")).astimezone(PACIFIC) if ts else None
        rows.append({"number": str(v.get("number") or ""), "year": v.get("year") or "",
                     "vin": v.get("vin") or "", "odometer": round(odo) if odo is not None else None,
                     "when": when})
    return rows


# ---------- Dropbox mapping file ----------
def dropbox_mapping(app_key, refresh_token, path, shops):
    tok = json.loads(http("https://api.dropboxapi.com/oauth2/token", "POST", form={
        "grant_type": "refresh_token", "refresh_token": refresh_token, "client_id": app_key}))["access_token"]
    auth = {"Authorization": f"Bearer {tok}"}
    # Team Dropbox: address paths from the team root namespace.
    acct = json.loads(http("https://api.dropboxapi.com/2/users/get_current_account", "POST", auth, data=b"null"))
    root_ns = acct["root_info"]["root_namespace_id"]
    headers = dict(auth)
    headers["Dropbox-API-Arg"] = json.dumps({"path": path})
    headers["Dropbox-API-Path-Root"] = json.dumps({".tag": "root", "root": root_ns})
    headers["Content-Type"] = "application/octet-stream"
    content = http("https://content.dropboxapi.com/2/files/download", "POST", headers, data=b"")
    wb = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    mapping = {}
    for row in wb["Sheet1"].iter_rows(values_only=True):
        if len(row) < 4 or row[1] is None:
            continue
        num = str(row[1]).strip()
        if num.endswith(".0"):
            num = num[:-2]
        if not num.isdigit():
            continue
        loc = str(row[3] or "").strip().lower().replace("’", "'")
        hits = [s["label"] for s in shops if s["match"] in loc]
        mapping[num] = hits[0] if len(hits) == 1 else None
    return mapping


# ---------- Email (Microsoft Graph) ----------
def html_table(rows):
    td = "border:1px solid #bbb;padding:4px 10px;font-family:Calibri,Arial,sans-serif;font-size:11pt;"
    heads = ["Vehicle Number", "Year", "VIN", "Odometer (mi)", "Reading Time (Pacific)"]
    out = ['<table style="border-collapse:collapse;"><tr>']
    out += [f'<th style="{td}background:#1f4e79;color:#fff;text-align:left;">{h}</th>' for h in heads]
    out.append("</tr>")
    for i, r in enumerate(rows):
        cells = [(r["number"], "left"), (str(r["year"]), "left"), (r["vin"], "left"),
                 (f"{r['odometer']:,}" if r["odometer"] is not None else "", "right"),
                 (r["when"].strftime("%m/%d/%Y %I:%M %p") if r["when"] else "", "left")]
        bg = "#f2f2f2" if i % 2 else "#ffffff"
        out.append(f'<tr style="background:{bg};">' + "".join(
            f'<td style="{td}text-align:{a};">{html.escape(v)}</td>' for v, a in cells) + "</tr>")
    out.append("</table>")
    return "\n".join(out)


def graph_send(sender, to, cc, subject, body_html):
    tok = json.loads(http(
        f"https://login.microsoftonline.com/{os.environ['MS_TENANT_ID']}/oauth2/v2.0/token", "POST", form={
            "client_id": os.environ["MS_CLIENT_ID"], "client_secret": os.environ["MS_CLIENT_SECRET"],
            "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"}))["access_token"]
    msg = {"message": {
        "subject": subject,
        "body": {"contentType": "HTML", "content": body_html},
        "toRecipients": [{"emailAddress": {"address": a}} for a in to],
        "ccRecipients": [{"emailAddress": {"address": a}} for a in cc],
    }, "saveToSentItems": True}
    http(f"https://graph.microsoft.com/v1.0/users/{urllib.parse.quote(sender)}/sendMail", "POST",
         {"Authorization": f"Bearer {tok}"}, data=msg)


def resend_send(sender, to, cc, bcc, subject, body_html):
    msg = {"from": f"Seaside Logistics <{sender}>", "to": to, "subject": subject, "html": body_html,
           "reply_to": sender}
    if cc:
        msg["cc"] = cc
    if bcc:
        msg["bcc"] = bcc
    http("https://api.resend.com/emails", "POST",
         {"Authorization": f"Bearer {os.environ['RESEND_API_KEY']}", "User-Agent": "odometer-job/1.0"}, data=msg)


def send_email(sender, to, cc, bcc, subject, body_html):
    if os.environ.get("RESEND_API_KEY"):
        resend_send(sender, to, cc, bcc, subject, body_html)
    else:
        graph_send(sender, to, cc, subject, body_html)  # Graph saves to Sent Items itself


# ---------- main ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["dry-run", "test", "send"], default="dry-run")
    ap.add_argument("--test-to")
    ap.add_argument("--shop", help="limit to one shop label")
    ap.add_argument("--schedule", default="", help="cron string that triggered the run (DST guard)")
    args = ap.parse_args()

    with open(os.path.join(HERE, "config.json"), encoding="utf-8") as f:
        cfg = json.load(f)
    rcp = json.loads(os.environ["RECIPIENTS_JSON"])
    cfg["sender"], cfg["cc"] = rcp["sender"], rcp.get("cc", [])
    for shop in cfg["shops"]:
        shop.update(rcp.get("shops", {}).get(shop["label"], {}))
    now = datetime.now(PACIFIC)

    if args.schedule:  # two UTC crons exist; only the one matching current DST state proceeds
        is_dst = bool(now.dst())
        want = cfg["cron_dst"] if is_dst else cfg["cron_std"]
        if args.schedule.strip() != want:
            print(f"Skip: schedule '{args.schedule}' is not the active one ({'PDT' if is_dst else 'PST'}).")
            return
    if args.mode == "send" and now.strftime("%A") not in cfg["send_days"]:
        print(f"Skip: {now:%A} is not a send day.")
        return
    if args.mode == "test" and not args.test_to:
        sys.exit("--test-to required in test mode")

    rows = motive_rows(os.environ["MOTIVE_API_KEY"], cfg["motive_group"])
    mapping = dropbox_mapping(os.environ["DROPBOX_APP_KEY"], os.environ["DROPBOX_REFRESH_TOKEN"],
                              cfg["mapping_dropbox_path"], cfg["shops"])
    today = now.strftime("%Y-%m-%d")
    unassigned = sorted(r["number"] for r in rows if not mapping.get(r["number"]))
    if unassigned:
        print(f"WARNING: no shop in mapping file (sent to nobody): {unassigned}")

    def nkey(r):
        return (0, int(r["number"])) if r["number"].isdigit() else (1, r["number"])

    for shop in cfg["shops"]:
        if args.shop and shop["label"] != args.shop:
            continue
        shop_rows = sorted((r for r in rows if mapping.get(r["number"]) == shop["label"]), key=nkey)
        if any(mapping.get(r["number"]) != shop["label"] for r in shop_rows):
            sys.exit(f"Safety check failed for {shop['label']}; nothing sent for this shop.")
        if not shop.get("to"):
            print(f"{shop['label']}: {len(shop_rows)} trucks (no recipients configured, not emailed)")
            continue
        p = 'style="font-family:Calibri,Arial,sans-serif;font-size:11pt;"'
        body = (f"<html><body><p {p}>Hello,</p><p {p}>Below are the latest odometer readings for our "
                f"trucks serviced at your shop, as of {today}.</p>{html_table(shop_rows)}"
                f"<p {p}>Thank you,<br>Seaside Logistics</p></body></html>")
        cc = shop.get("cc", cfg["cc"])
        subject = f"Odometer Readings - {today}"
        to = shop["to"]
        if args.mode == "test":
            subject = f"[TEST - real email goes to {'; '.join(to)} cc {'; '.join(cc)}] {subject}"
            to, cc = [args.test_to], []
        nums = ", ".join(r["number"] for r in shop_rows)
        if args.mode == "dry-run":
            print(f"DRY RUN {shop['label']}: to={to} cc={cc} trucks({len(shop_rows)})={nums}")
            continue
        bcc = [] if args.mode == "test" else [cfg["sender"]]  # keep a copy in the Administration mailbox
        send_email(cfg["sender"], to, cc, bcc, subject, body)
        print(f"SENT {shop['label']}: to={to} cc={cc} trucks({len(shop_rows)})={nums}")


if __name__ == "__main__":
    main()
