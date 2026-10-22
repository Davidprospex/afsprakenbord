#!/usr/bin/env python3
"""Afsprakenbord: telt nieuwe afspraakbevestigingen uit de mailbox van socialassistant@prospex.nl.

Draait in GitHub Actions (elke 5 minuten), zonder Claude en zonder laptop.
Leest via Microsoft Graph (app-registratie, alleen-lezen, beperkt tot deze ene mailbox),
schrijft data.json en bouwt index.html. Slaat alleen bedrijfsnaam/klant, profiel en Ctpid op:
geen contactpersonen.

Omgeving: MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MAILBOX (standaard socialassistant@prospex.nl).
Optioneel: FULL=1 om alles vanaf START opnieuw op te bouwen.
"""
import datetime as dt
import hashlib
import html
import json
import os
import re
import sys
import urllib.parse
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
START = "2026-01-01"  # begin van het bord
LOOKBACK_DAYS = 3     # normale run: alleen recente mail ophalen en samenvoegen

# Klantnamen zoals ze op het bord moeten staan.
KANON = ["Prospex", "Marcus", "Mfactor", "Lyreco", "Prostrive", "Nowy Styl", "Personeelsadviseur", "Novar",
         "Hovingh", "Riff", "Thorlux", "OnItNow", "RBBLS", "SGS", "QRC", "Solite", "Touch", "Verdi", "Axelio"]
ALIASES = {
    "prospex marketing": "Prospex",
    "touch incentive": "Touch",
    "onitnow": "OnItNow",
    "de virtuele personeelsadviseur": "Personeelsadviseur",
    "hovingh & partners": "Hovingh",
    "qrc security": "QRC",
    "nowystyl": "Nowy Styl",
    "m factor": "Mfactor",
}
KANON_L = {k.lower(): k for k in KANON}

SUBJECT_RE = re.compile(r"^\s*(Afspraakbevestiging|Appointment confirmation)\b", re.I)
STAMP_RE = re.compile(r"(\d{2})-(\d{2})-(\d{4})\s+(\d{2}):(\d{2})(?::\d{2})?")


def body_text(html_body):
    t = re.sub(r"(?i)<br\s*/?>|</(p|div|tr|td)>|<hr\s*/?>", "\n", html_body or "")
    t = re.sub(r"<[^>]+>", "", t)
    return html.unescape(t)


def field(text, *names):
    for n in names:
        m = re.search(r"^\s*" + re.escape(n) + r"\s*:\s*(.*)$", text, re.M | re.I)
        if m:
            return m.group(1).strip()
    return ""


def klant_uit_campagne(campagne):
    deel = campagne.split(" - ")[0].strip().lstrip("[").strip()
    deel = re.sub(r"\s*\b20\d\d\b\s*$", "", deel).strip().rstrip("-").strip()
    l = deel.lower()
    return ALIASES.get(l) or KANON_L.get(l) or deel


def profiel_uit_project(project, afspraak_voor):
    m = re.search(r"(?:Connecties|Connections|connecties)\s*-?\s*(.+)$", project)
    if m and m.group(1).strip():
        return m.group(1).split(" - ")[0].strip()
    return afspraak_voor


def gemaakt_op(text, fallback_iso):
    i = re.search(r"(OVERIGE INFORMATIE|OTHER INFORMATION)", text, re.I)
    m = STAMP_RE.search(text, i.end()) if i else None
    if m:
        d, mo, y, h, mi = m.groups()
        return dt.datetime(int(y), int(mo), int(d), int(h), int(mi))
    # Graph geeft UTC; zet om naar Nederlandse tijd (CET/CEST benadering via zoneinfo).
    from zoneinfo import ZoneInfo
    u = dt.datetime.fromisoformat(fallback_iso.replace("Z", "+00:00"))
    return u.astimezone(ZoneInfo("Europe/Amsterdam")).replace(tzinfo=None)


def parse(msg):
    """Geeft een afspraak-dict terug, of None als het geen nieuwe afspraak is."""
    if not SUBJECT_RE.match(msg.get("subject") or ""):
        return None  # Re:, FW:, informatieverzoeken, overige mail
    text = body_text((msg.get("body") or {}).get("content", ""))
    code = field(text, "Resultaatcode", "Resultcode")
    if not re.match(r"(?i)^(afspraak|appointment)", code):
        return None
    campagne = field(text, "Campagne", "Campaign")
    project = field(text, "Project")
    if not campagne:
        return None
    if re.search(r"\btest\b|test-", project + " " + (msg.get("subject") or ""), re.I):
        return None  # testbevestigingen
    gemaakt = gemaakt_op(text, msg["receivedDateTime"])
    dag = gemaakt.date()
    if dag.weekday() >= 5:  # weekend telt op de maandag erna
        dag += dt.timedelta(days=7 - dag.weekday())
    return {
        "id": msg.get("internetMessageId") or msg["id"],
        "datum": dag.isoformat(),
        "tijd": gemaakt.strftime("%H:%M"),
        "klant": klant_uit_campagne(campagne),
        "profiel": profiel_uit_project(project, field(text, "Afspraak voor", "Appointment for")),
        "nummer": field(text, "Ctpid"),
        "status": "telt",
        # vingerafdruk van het onderwerp (bevat de prospectnaam, die slaan we niet leesbaar op)
        "onderwerp": hashlib.sha1(re.sub(r"\s+", " ", msg.get("subject") or "").strip().lower().encode()).hexdigest()[:12],
    }


def token():
    data = urllib.parse.urlencode({
        "client_id": os.environ["MS_CLIENT_ID"],
        "client_secret": os.environ["MS_CLIENT_SECRET"],
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }).encode()
    url = f"https://login.microsoftonline.com/{os.environ['MS_TENANT_ID']}/oauth2/v2.0/token"
    with urllib.request.urlopen(urllib.request.Request(url, data=data)) as r:
        return json.load(r)["access_token"]


def fetch(since_iso):
    mailbox = os.environ.get("MAILBOX", "socialassistant@prospex.nl")
    q = urllib.parse.urlencode({
        "$filter": f"receivedDateTime ge {since_iso}",
        "$select": "id,subject,body,receivedDateTime,internetMessageId",
        "$orderby": "receivedDateTime asc",
        "$top": "100",
    })
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/inbox/messages?{q}"
    headers = {"Authorization": f"Bearer {token()}", "Prefer": 'outlook.body-content-type="html"'}
    while url:
        with urllib.request.urlopen(urllib.request.Request(url, headers=headers)) as r:
            page = json.load(r)
        yield from page.get("value", [])
        url = page.get("@odata.nextLink")


def main():
    path = os.path.join(ROOT, "data.json")
    data = {"bron": "mail", "dagen": {}}
    if os.path.exists(path):
        old = json.load(open(path))
        if old.get("bron") == "mail" and os.environ.get("FULL") != "1":
            data = old
    full = data["dagen"] == {}
    since = START if full else (dt.date.today() - dt.timedelta(days=LOOKBACK_DAYS)).isoformat()

    alle = [dict(a, datum=d) for d, dag in data["dagen"].items() for a in dag]
    ids = {a["id"] for a in alle}
    sleutels = {(a["klant"], a["nummer"]) for a in alle if a["nummer"]}
    gezien = {}
    for a in alle:
        if a.get("onderwerp"):
            gezien[a["onderwerp"]] = max(gezien.get(a["onderwerp"], a["datum"]), a["datum"])
    nieuw = 0
    telling = {"opgehaald": 0, "geen bevestiging": 0, "al op bord": 0, "zelfde ctpid": 0, "zelfde onderwerp": 0}
    for msg in fetch(since + "T00:00:00Z"):
        telling["opgehaald"] += 1
        a = parse(msg)
        if not a:
            telling["geen bevestiging"] += 1
            subj = msg.get("subject") or ""
            if SUBJECT_RE.match(subj) and telling.setdefault("bevestiging afgekeurd", 0) < 3:
                telling["bevestiging afgekeurd"] += 1
                body = (msg.get("body") or {})
                text = body_text(body.get("content", ""))
                print("DEBUG", {"onderwerp_ok": bool(SUBJECT_RE.match(subj)), "begint_met": subj[:22],
                                "body_type": body.get("contentType"), "body_lengte": len(body.get("content", "")),
                                "resultaatcode": field(text, "Resultaatcode", "Resultcode")[:30],
                                "campagne_gevonden": bool(field(text, "Campagne", "Campaign")),
                                "regels": text.count("\n"),
                                "rc_regel": [l[:40] for l in text.splitlines() if "esult" in l][:2],
                                "ruw_rond_rc": body.get("content", "")[body.get("content", "").find("esultaat")-60:body.get("content", "").find("esultaat")+40] if "esultaat" in body.get("content", "") else "geen"})
            continue
        if a["id"] in ids or a["datum"] < START:
            telling["al op bord"] += 1
            continue
        if a["nummer"] and (a["klant"], a["nummer"]) in sleutels:
            telling["zelfde ctpid"] += 1
            continue  # zelfde Ctpid voor dezelfde klant: herplanning of dubbele mail
        vorige = gezien.get(a["onderwerp"])
        if vorige and (dt.date.fromisoformat(a["datum"]) - dt.date.fromisoformat(vorige)).days <= 14:
            telling["zelfde onderwerp"] += 1
            continue  # zelfde bevestiging opnieuw verstuurd
        gezien[a["onderwerp"]] = a["datum"]
        dag = a.pop("datum")
        data["dagen"].setdefault(dag, []).append(a)
        ids.add(a["id"])
        if a["nummer"]:
            sleutels.add((a["klant"], a["nummer"]))
        nieuw += 1
    for dag in data["dagen"].values():
        dag.sort(key=lambda a: a["tijd"])
    data["dagen"] = dict(sorted(data["dagen"].items()))

    print(f"Sinds {since}: {telling}, nieuw: {nieuw}")
    changed = nieuw > 0 or full or "laatsteRun" not in data
    if not changed:
        print("Geen nieuwe afspraken.")
        return
    from zoneinfo import ZoneInfo
    data["laatsteRun"] = dt.datetime.now(ZoneInfo("Europe/Amsterdam")).strftime("%Y-%m-%dT%H:%M")
    json.dump(data, open(path, "w"), ensure_ascii=False, indent=1)
    build(data)
    print(f"{nieuw} nieuwe afspraken; totaal {sum(len(v) for v in data['dagen'].values())}; gewijzigd={changed}")


def build(data):
    css = open(os.path.join(ROOT, "prospex-design-system.css")).read()
    tpl = open(os.path.join(ROOT, "template.html")).read()
    blob = json.dumps({"laatsteRun": data["laatsteRun"], "dagen": data["dagen"]}, ensure_ascii=False).replace("</", "<\\/")
    page = tpl.replace("/*DS*/", css, 1).replace("/*DATA*/", blob, 1)
    i = page.index('<div class="prospex prospex-dark">')
    page = ('<!doctype html><html lang="nl"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
            '<meta name="robots" content="noindex,nofollow">'
            '<meta http-equiv="refresh" content="300">' + page[:i] + '</head><body>' + page[i:] + '</body></html>')
    open(os.path.join(ROOT, "index.html"), "w").write(page)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        # Lokale test: python3 update.py --test bericht.json [...]
        for f in sys.argv[2:]:
            print(json.dumps(parse(json.load(open(f))), ensure_ascii=False))
    else:
        main()
