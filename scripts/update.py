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
import urllib.error
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
    "verdi groep": "Verdi",
}
KANON_L = {k.lower(): k for k in KANON}

SUBJECT_RE = re.compile(r"^[\s\u200b\u200e\u200f\ufeff]*(?:(?:FW|FWD|Fwd|Doorgestuurd)\s*:\s*)*(Afspraakbevestiging|Appointment confirmation)\b", re.I)
STAMP_RE = re.compile(r"(\d{2})-(\d{2})-(\d{4})\s+(\d{2}):(\d{2})(?::\d{2})?")


def body_text(html_body):
    # Regeleinde bij elke br/hr en elk blok-element, ook als er attributen in de tag staan (zoals bij Microsoft Graph).
    t = re.sub(r"(?i)<br[^>]*>|<hr[^>]*>|</?(p|div|tr|td|table|tbody)[^>]*>", "\n", html_body or "")
    t = re.sub(r"<[^>]+>", "", t)
    return html.unescape(t).replace("\xa0", " ")


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
    if re.search(r"(?i)opnieuw|herplan|verzet|verplaatst|resched", code):
        return None  # herplanning (bijv. "Afspraak opnieuw ingepland") telt niet als nieuwe afspraak
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
        "onderwerp": hashlib.sha1(re.sub(r"^(?:\s*(?:fw|fwd|doorgestuurd)\s*:\s*)+", "", re.sub(r"\s+", " ", msg.get("subject") or "").strip().lower()).encode()).hexdigest()[:12],
        # Ctpid is alleen uniek binnen een project: project (als vingerafdruk) hoort bij de sleutel
        "proj": hashlib.sha1(project.strip().lower().encode()).hexdigest()[:8],
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


MAPPEN = ["inbox", "deleteditems", "archive"]  # bevestigingen staan niet alleen in de Postvak IN


def _pages(url, headers):
    while url:
        with urllib.request.urlopen(urllib.request.Request(url, headers=headers)) as r:
            page = json.load(r)
        yield from page.get("value", [])
        url = page.get("@odata.nextLink")


def fetch(since_iso):
    """Alle bevestigingen sinds since_iso uit alle relevante mappen, van elke afzender, oudste eerst."""
    mailbox = os.environ.get("MAILBOX", "socialassistant@prospex.nl")
    headers = {"Authorization": f"Bearer {token()}", "Prefer": 'outlook.body-content-type="html"'}
    alles = []
    for map_ in MAPPEN:
        basis = {"$select": "id,subject,body,receivedDateTime,internetMessageId", "$top": "100"}
        # Geen onderwerpfilter op de server: doorgestuurde bevestigingen beginnen anders. We kiezen zelf (SUBJECT_RE).
        filters = [f"receivedDateTime ge {since_iso}"]
        for f in filters:
            q = urllib.parse.urlencode({**basis, "$filter": f})
            url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{map_}/messages?{q}"
            try:
                berichten = list(_pages(url, headers))
                break
            except urllib.error.HTTPError as e:
                if e.code == 404:  # map bestaat niet
                    berichten = []
                    break
                berichten = None
        print(f"Map {map_}: {len(berichten or [])} berichten")
        alles += berichten or []
    alles.sort(key=lambda m: m.get("receivedDateTime", ""))
    return alles


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
    # Ctpid per klant met laatste datum: zelfde Ctpid binnen 30 dagen = herplanning of al geteld (ook via WhatsApp).
    sleutels = {}
    for a in alle:
        if a["nummer"]:
            k = (a["klant"], a.get("proj", ""), a["nummer"])
            sleutels[k] = max(sleutels.get(k, a["datum"]), a["datum"])
    gezien = {}
    for a in alle:
        if a.get("onderwerp"):
            gezien[a["onderwerp"]] = max(gezien.get(a["onderwerp"], a["datum"]), a["datum"])
    nieuw = 0
    telling = {"opgehaald": 0, "geen bevestiging": 0, "al op bord": 0, "zelfde ctpid": 0, "zelfde onderwerp": 0}
    debug_klant = os.environ.get("DEBUG_KLANT", "lyreco").lower()  # bijv. DEBUG_KLANT=lyreco voor diagnose
    def dbg(msg, uitkomst, a=None):
        subj = (msg.get("subject") or "").lower()
        if debug_klant and debug_klant in subj and uitkomst != "TELT" and "undeliverable" not in subj:
            text = body_text((msg.get("body") or {}).get("content", ""))
            print("DEBUG", msg.get("receivedDateTime", "")[:16], "ctpid", field(text, "Ctpid") or "-",
                  "| code:", (field(text, "Resultaatcode", "Resultcode") or "-")[:30], "| campagne:", field(text, "Campagne", "Campaign")[:30],
                  "| onderwerp_start:", subj[:26], "| =>", uitkomst, (a or {}).get("datum", ""))
    for msg in fetch(since + "T00:00:00Z"):
        telling["opgehaald"] += 1
        a = parse(msg)
        if not a:
            dbg(msg, "GEEN BEVESTIGING")
        if not a:
            telling["geen bevestiging"] += 1
            continue
        if a["id"] in ids or a["datum"] < START:
            dbg(msg, "AL OP BORD", a)
            telling["al op bord"] += 1
            continue
        vorig = sleutels.get((a["klant"], a.get("proj", ""), a["nummer"])) if a["nummer"] else None
        if vorig and abs((dt.date.fromisoformat(a["datum"]) - dt.date.fromisoformat(vorig)).days) <= 30:
            dbg(msg, "ZELFDE CTPID (binnen 30 dagen)", a)
            telling["zelfde ctpid"] += 1
            continue  # zelfde Ctpid voor dezelfde klant: herplanning of dubbele mail
        vorige = gezien.get(a["onderwerp"])
        if vorige and (dt.date.fromisoformat(a["datum"]) - dt.date.fromisoformat(vorige)).days <= 14:
            dbg(msg, "ZELFDE ONDERWERP (binnen 14 dagen)", a)
            telling["zelfde onderwerp"] += 1
            continue  # zelfde bevestiging opnieuw verstuurd
        gezien[a["onderwerp"]] = a["datum"]
        dbg(msg, "TELT", a)
        dag = a.pop("datum")
        data["dagen"].setdefault(dag, []).append(a)
        ids.add(a["id"])
        if a["nummer"]:
            sleutels[(a["klant"], a.get("proj", ""), a["nummer"])] = dag
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



# ---------- Live bord: plusjes van ntfy.sh vastleggen in live.json ----------
LIVE_TOPIC = "prospex-afsprakenbord-ffb9535c0211e74c8e65acc0"


def update_live():
    """Leest de +/- van het Live-tabblad (ntfy.sh bewaart ze 12 uur) en legt de stand vast in live.json.
    Zo ziet elk apparaat bij openen de juiste stand, ook als er geen bord openstaat."""
    from zoneinfo import ZoneInfo
    nu = dt.datetime.now(ZoneInfo("Europe/Amsterdam"))
    week = (nu.date() - dt.timedelta(days=nu.weekday())).isoformat()
    path = os.path.join(ROOT, "live.json")
    live = json.load(open(path)) if os.path.exists(path) else {}
    if live.get("week") != week:
        live = {"week": week, "tel": {}, "dag": {}, "eids": [], "sinds": 0}
    sinds = max(int(live.get("sinds", 0)), int(nu.timestamp()) - 12 * 3600)
    url = f"https://ntfy.sh/{LIVE_TOPIC}/json?poll=1&since={sinds}"
    try:
        with urllib.request.urlopen(url, timeout=20) as r:
            regels = r.read().decode().splitlines()
    except Exception as e:
        print("Live: ntfy niet bereikbaar:", e)
        return False
    eids, gewijzigd = set(live["eids"]), False
    klant_events = []
    for regel in regels:
        try:
            m = json.loads(regel)
            live["sinds"] = max(live["sinds"], int(m.get("time", 0)))
            if m.get("event") != "message":
                continue
            b = json.loads(m.get("message", ""))
        except Exception:
            continue
        if b.get("type") == "klant":
            klant_events.append(b)
            continue
        if b.get("type") != "delta" or b.get("w") != week or not b.get("eid") or b["eid"] in eids:
            continue
        k, d = str(b.get("k", ""))[:60], int(b.get("d", 0))
        if d not in (1, -1) or not k or live["tel"].get(k, 0) + d < 0:
            continue
        dag = dt.datetime.fromtimestamp(int(m["time"]), ZoneInfo("Europe/Amsterdam")).date().isoformat()
        live["tel"][k] = live["tel"].get(k, 0) + d
        live["dag"][dag] = max(0, live["dag"].get(dag, 0) + d)
        eids.add(b["eid"])
        live["eids"].append(b["eid"])
        gewijzigd = True
    update_klanten(klant_events)
    live["eids"] = live["eids"][-500:]
    if gewijzigd or not os.path.exists(path):
        json.dump(live, open(path, "w"), ensure_ascii=False)
        print("Live: stand bijgewerkt", live["tel"])
    return gewijzigd


def update_klanten(events):
    """Klant toevoegen/verwijderen door een admin vastleggen in klanten.json (blijft staan, geen weekreset)."""
    path = os.path.join(ROOT, "klanten.json")
    kl = json.load(open(path)) if os.path.exists(path) else {"toegevoegd": {}, "verwijderd": {}, "eids": []}
    seen, gewijzigd = set(kl["eids"]), False
    for b in events:
        k = str(b.get("k", "")).strip()[:60]
        eid, actie, week = b.get("eid"), b.get("actie"), b.get("week")
        if not k or not eid or eid in seen or actie not in ("toevoegen", "verwijderen", "terugzetten"):
            continue
        if actie == "toevoegen":
            kl["toegevoegd"][k] = week
            kl["verwijderd"].pop(k, None)
        elif actie == "verwijderen":
            kl["verwijderd"][k] = week
            kl["toegevoegd"].pop(k, None)
        else:
            kl["verwijderd"].pop(k, None)
        kl["eids"].append(eid)
        seen.add(eid)
        gewijzigd = True
    kl["eids"] = kl["eids"][-300:]
    if gewijzigd or not os.path.exists(path):
        json.dump(kl, open(path, "w"), ensure_ascii=False)
        print("Klanten bijgewerkt:", {"toegevoegd": kl["toegevoegd"], "verwijderd": kl["verwijderd"]})

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        # Lokale test: python3 update.py --test bericht.json [...]
        for f in sys.argv[2:]:
            print(json.dumps(parse(json.load(open(f))), ensure_ascii=False))
    else:
        update_live()
        main()
