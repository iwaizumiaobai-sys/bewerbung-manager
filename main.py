"""Bewerbungs-Manager — alles in einer Datei.

Start lokal:  uvicorn main:app --reload
Auf Railway:  laeuft ueber das Procfile automatisch.
"""

import email
import hashlib
import io
import html
import json
import math
import os
import re
import secrets
import sqlite3
import threading
import time
from datetime import datetime, timedelta, timezone
from email.header import decode_header
from urllib.parse import parse_qs, quote

from fastapi import FastAPI, File, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, StreamingResponse
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

# ---------------------------------------------------------------- Einstellungen

FIRMA = "Places to Be"
FARBE = "#1B3A8C"          # Hauptfarbe aus dem Logo. Kopfbereich und Knoepfe.
FARBE_HELL = "#E9EFFB"     # Sehr helle Variante fuer Flaechen.
AKZENT = "#D42B27"         # Das Rot aus dem Logo, sparsam eingesetzt.


def abdunkeln(hexwert, anteil=0.35):
    """Macht eine Farbe dunkler, fuer den Verlauf im Kopfbereich."""
    hexwert = hexwert.lstrip("#")
    r, g, b = (int(hexwert[i:i + 2], 16) for i in (0, 2, 4))
    return "#%02x%02x%02x" % tuple(int(k * (1 - anteil)) for k in (r, g, b))


def kuerzel_aus(name):
    """Bildet ein Kuerzel, falls kein Logo hinterlegt ist."""
    teile = [w for w in re.split(r"\s+", name) if w]
    return "".join(w[0] for w in teile[:3]).upper() or "?"


AUFBEWAHRUNG_TAGE = 3

DB = os.environ.get("DB_PATH", "/tmp/bewerbungen.db")
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
MODELL = os.environ.get("AI_MODEL", "claude-sonnet-4-6")

# --- Schutz gegen Missbrauch und ausufernde Kosten ---
PASSWORT = os.environ.get("PASSWORT", "")
def zahl_lesen(name, vorgabe, kleinste, groesste):
    """Liest eine Zahl aus einer Variable. Bei Tippfehlern gilt die Vorgabe."""
    try:
        wert = float(os.environ.get(name, "").strip().replace(",", "."))
    except ValueError:
        return vorgabe
    if not math.isfinite(wert):
        return vorgabe
    return max(kleinste, min(groesste, wert))


TAGESLIMIT = int(zahl_lesen("TAGESLIMIT", 150, 1, 100000))   # Bewertungen pro Tag
MAX_MB = zahl_lesen("MAX_MB", 8, 0.5, 50)                     # groesste erlaubte Datei

# Wert, den der Zutritts-Cookie tragen muss. Aendert sich mit dem Passwort.
MARKE = hashlib.sha256(("ptb-zutritt-" + PASSWORT).encode()).hexdigest()[:40]

# Bewertungskriterien: Schluessel, Anzeigename, Gewicht, Erklaerung fuer die KI.
# Ueber die Variable KRITERIEN aenderbar, eine Zeile je Kriterium:
#   kuerzel|Anzeigename|Gewicht|Was gemeint ist
# Bei leerer oder fehlerhafter Variable gilt diese Vorgabe.
KRITERIEN_VORGABE = [
    ("resilienz", "Umgang mit Ablehnung", 25,
     "Frustrationstoleranz, Umgang mit Absagen und Gegenwind"),
    ("reise", "Reisebereitschaft", 20,
     "zeitliche und raeumliche Flexibilitaet, Bereitschaft zu reisen"),
    ("extro", "Extrovertiertheit", 20,
     "Offenheit, Kommunikationsfreude, Zugehen auf Fremde"),
    ("erfahrung", "Vorerfahrung", 15,
     "Vertrieb, Promotion, Standarbeit, Kundenkontakt"),
    ("motivation", "Motivation", 15,
     "Warum gerade Fundraising, Bezug zu gemeinnuetziger Arbeit"),
    ("sprache", "Sprachkenntnisse", 5,
     "Ausdruck und erwaehnte Sprachkenntnisse"),
]


# Diese Namen nutzt die App selbst, sie duerfen kein Kriterium heissen.
RESERVIERT = {"fehler", "erstattet", "fuehrerschein", "zusammenfassung", "verfuegbarkeit"}


def kriterien_lesen():
    """Liest die Kriterien aus der Variable. Faellt auf die Vorgabe zurueck."""
    roh = os.environ.get("KRITERIEN", "").strip()
    if not roh:
        return KRITERIEN_VORGABE

    gelesen = []
    for zeile in roh.splitlines():
        zeile = zeile.strip()
        if not zeile or zeile.startswith("#"):
            continue
        teile = [t.strip() for t in zeile.split("|")]
        if len(teile) < 3:
            continue
        schluessel = re.sub(r"[^a-z0-9_]", "", teile[0].lower())
        if not schluessel or schluessel in RESERVIERT:
            continue
        try:
            zahl = float(teile[2].replace(",", "."))
        except ValueError:
            continue
        if not math.isfinite(zahl):
            continue
        gewicht = max(1, min(100, int(zahl)))
        titel = teile[1] or schluessel
        erklaerung = teile[3] if len(teile) > 3 else titel
        gelesen.append((schluessel, titel, gewicht, erklaerung))

    # Doppelte Schluessel entfernen, Reihenfolge behalten
    gesehen, sauber = set(), []
    for eintrag in gelesen:
        if eintrag[0] in gesehen:
            continue
        gesehen.add(eintrag[0])
        sauber.append(eintrag)

    if len(sauber) < 2:
        return KRITERIEN_VORGABE
    return sauber[:10]


KRITERIEN = kriterien_lesen()
GEWICHT_SUMME = sum(g for _, _, g, _ in KRITERIEN) or 1

STATUS_FARBEN = {
    "gruen": ("C6EFCE", "Einladen"),
    "gelb": ("FFEB9C", "Prüfen"),
    "grau": ("E7E6E6", "Angaben fehlen"),
    "rot": ("FFC7CE", "Unpassend"),
    "fehler": ("F8CBAD", "Bewertung fehlgeschlagen"),
}

# ---------------------------------------------------------------- Datenbank


def conn():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    return c


def init():
    with conn() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS bewerbungen (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT, absender TEXT, betreff TEXT, text TEXT,
                status TEXT, gesamt REAL, scores TEXT,
                zusammenfassung TEXT, fuehrerschein TEXT,
                angelegt TEXT
            )
        """)
        # Nachtraeglich ergaenzte Spalten: still uebergehen, wenn sie schon da sind.
        for spalte in ("verfuegbarkeit TEXT",):
            try:
                c.execute(f"ALTER TABLE bewerbungen ADD COLUMN {spalte}")
            except sqlite3.OperationalError:
                pass
        c.execute("""
            CREATE TABLE IF NOT EXISTS tageszaehler (
                tag TEXT PRIMARY KEY,
                anzahl INTEGER NOT NULL DEFAULT 0
            )
        """)


def heute():
    """Der heutige Tag in deutscher Zeit. Das Tageslimit springt um Mitternacht zurueck."""
    return berlin_zeit(datetime.now(timezone.utc)).strftime("%Y-%m-%d")


def verbrauch_heute():
    """Wie viele Bewertungen heute schon liefen."""
    with conn() as c:
        reihe = c.execute("SELECT anzahl FROM tageszaehler WHERE tag=?",
                          (heute(),)).fetchone()
    return reihe["anzahl"] if reihe else 0


def kontingent_buchen():
    """Bucht eine Bewertung. Gibt False zurueck, wenn das Tageslimit erreicht ist.

    Zaehlt auch dann hoch, wenn die Bewerbung spaeter geloescht wird, denn der
    API-Aufruf hat bereits Geld gekostet.
    """
    with conn() as c:
        c.execute("INSERT OR IGNORE INTO tageszaehler (tag, anzahl) VALUES (?, 0)",
                  (heute(),))
        zeiger = c.execute(
            "UPDATE tageszaehler SET anzahl = anzahl + 1 "
            "WHERE tag = ? AND anzahl < ?", (heute(), TAGESLIMIT))
        return zeiger.rowcount > 0


def kontingent_zurueck():
    """Gibt eine Buchung zurueck, wenn der KI-Aufruf gar nicht erst geklappt hat."""
    with conn() as c:
        c.execute("UPDATE tageszaehler SET anzahl = MAX(0, anzahl - 1) WHERE tag = ?",
                  (heute(),))


def aufraeumen():
    grenze =(datetime.now(timezone.utc) - timedelta(days=AUFBEWAHRUNG_TAGE)).isoformat()
    with conn() as c:
        c.execute("DELETE FROM bewerbungen WHERE angelegt < ?", (grenze,))


# ---------------------------------------------------------------- Mail lesen


STEUERZEICHEN = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def sauber(wert, laenge=None):
    """Entfernt unsichtbare Steuerzeichen (die z.B. die Excel-Ausgabe stoeren) und kuerzt."""
    wert = STEUERZEICHEN.sub("", str(wert or "")).strip()
    return wert[:laenge] if laenge else wert


def kopf(wert):
    if not wert:
        return ""
    teile = []
    try:
        stuecke = decode_header(wert)
    except Exception:
        return sauber(str(wert))
    for stueck, kodierung in stuecke:
        if isinstance(stueck, bytes):
            try:
                teile.append(stueck.decode(kodierung or "utf-8", errors="replace"))
            except LookupError:                      # unbekannter Zeichensatz
                teile.append(stueck.decode("utf-8", errors="replace"))
        else:
            teile.append(stueck)
    return sauber("".join(teile))


def html_zu_text(roh):
    roh = re.sub(r"(?is)<(script|style).*?</\1>", " ", roh)
    roh = re.sub(r"(?i)<br\s*/?>", "\n", roh)
    roh = re.sub(r"(?i)</p>", "\n\n", roh)
    roh = re.sub(r"<[^>]+>", " ", roh)
    roh = html.unescape(roh).replace("\xa0", " ")
    return re.sub(r"\n{3,}", "\n\n", re.sub(r"[ \t]{2,}", " ", roh)).strip()


def pdf_zu_text(rohdaten):
    """Liest den Text aus einer PDF. Gibt leeren String zurueck, wenn es nicht geht."""
    try:
        from pypdf import PdfReader

        leser = PdfReader(io.BytesIO(rohdaten))
        seiten = [(s.extract_text() or "") for s in leser.pages[:15]]
        return re.sub(r"\n{3,}", "\n\n", "\n".join(seiten)).strip()
    except Exception:
        return ""


def text_entschluesseln(rohdaten, zeichensatz=None):
    """Wandelt Bytes in Text. Probiert den angegebenen Zeichensatz, dann UTF-8,
    dann Windows-1252 (typisch fuer aeltere Word- und Outlook-Dateien)."""
    for versuch in (zeichensatz, "utf-8", "cp1252"):
        if not versuch:
            continue
        try:
            return rohdaten.decode(versuch)
        except (LookupError, UnicodeDecodeError):
            continue
    return rohdaten.decode("utf-8", errors="replace")


def eml_lesen(rohdaten):
    msg = email.message_from_bytes(rohdaten)
    absender = kopf(msg.get("From"))
    betreff = kopf(msg.get("Subject"))

    text, html_teil, anhaenge = "", "", []
    if msg.is_multipart():
        for teil in msg.walk():
            if teil.get_content_maintype() == "multipart":
                continue
            dateiname = kopf(teil.get_filename() or "")
            ist_anhang = ("attachment" in str(teil.get("Content-Disposition", ""))
                          or dateiname)
            if ist_anhang:
                if dateiname.lower().endswith(".pdf"):
                    inhalt = teil.get_payload(decode=True)
                    if inhalt:
                        gelesen = pdf_zu_text(inhalt)
                        if gelesen:
                            anhaenge.append(f"[Anhang {dateiname}]\n{gelesen}")
                continue
            try:
                inhalt = teil.get_payload(decode=True)
                if inhalt is None:
                    continue
                entpackt = text_entschluesseln(inhalt, teil.get_content_charset())
            except Exception:
                continue
            if teil.get_content_type() == "text/plain" and not text:
                text = entpackt
            elif teil.get_content_type() == "text/html" and not html_teil:
                html_teil = entpackt
    else:
        inhalt = msg.get_payload(decode=True) or b""
        entpackt = text_entschluesseln(inhalt, msg.get_content_charset())
        if msg.get_content_type() == "text/html":
            html_teil = entpackt
        else:
            text = entpackt

    koerper = text.strip() or html_zu_text(html_teil)
    if anhaenge:
        koerper = (koerper + "\n\n" + "\n\n".join(anhaenge)).strip()

    name = absender
    if "<" in absender:
        name = absender.split("<")[0].strip().strip('"')
    if not name:
        name = absender

    return {"name": name or "Unbekannt", "absender": absender,
            "betreff": betreff or "(ohne Betreff)", "text": koerper}


# ---------------------------------------------------------------- Bewertung

def mail_raten(text):
    # Begrenzte Laengen, damit lange Zeichenketten ohne Leerzeichen nicht ewig dauern.
    treffer = re.search(r"[\w.+-]{1,64}@[\w-]{1,63}(?:\.[\w-]{1,63}){1,8}", text[:20000])
    return treffer.group(0).rstrip(".,;:") if treffer else ""


GRUSS = (r"(?:(?:Viele|Beste|Freundliche|Liebe|Herzliche|Sonnige|Schöne)\s+Gr(?:ü|ue|u)(?:ß|ss)en?"
         r"|Mit\s+(?:freundlichen|besten|lieben|herzlichen)\s+Gr(?:ü|ue|u)(?:ß|ss)en"
         r"|Mit\s+freundlichem\s+Gru(?:ß|ss)|MfG|LG|VG|Best regards|Kind regards|Regards)")
NAME = r"([A-ZÄÖÜ][\wäöüß'-]+(?:[ \t]+[A-ZÄÖÜ][\wäöüß'-]+){0,2})"   # nur innerhalb einer Zeile


def name_raten(text, dateiname=""):
    """Sucht den Namen: erst in der Grussformel am Ende, dann im Text, dann im Dateinamen."""
    schluss = text[-500:]
    treffer = None
    for treffer in re.finditer(r"(?<!\w)" + GRUSS + r",?[ \t]*\r?\n+\s*" + NAME, schluss):
        pass                                  # die letzte Grussformel zaehlt
    if treffer:
        return treffer.group(1).strip()

    treffer = re.search(r"(?:mein Name ist|Ich hei[sß]e)\s+"
                        + NAME, text[:1500])
    if treffer:
        return treffer.group(1).strip()

    roh = re.sub(r"\.[^.]+$", "", dateiname)
    roh = re.sub(r"(?i)bewerbung|lebenslauf|anschreiben|unterlagen|dokument|document"
                 r"|scan|datei|image|img|cv|final|neu", " ", roh)
    roh = re.sub(r"[_\-.]+", " ", roh).strip()
    woerter = [w for w in roh.split() if len(w) > 1 and not w.isdigit()][:3]
    return " ".join(w.capitalize() for w in woerter) or "Unbekannt"


ROLLE = os.environ.get(
    "ROLLE",
    "Du hilfst einer Fundraising-Agentur bei der Vorsortierung von Bewerbungen\n"
    "fuer die Taetigkeit als Werber an Infostaenden (Mitgliederwerbung fuer\n"
    "gemeinnuetzige Organisationen)."
)


def prompt_bauen():
    """Baut die Anweisung an die KI aus den aktuellen Kriterien."""
    punkte = "\n".join(f"- {k}: {erklaerung}" for k, _, _, erklaerung in KRITERIEN)
    felder = ", ".join(f'"{k}": 0' for k, _, _, _ in KRITERIEN)
    return (
        ROLLE + "\n\n"
        "Bewerte den folgenden Bewerbungstext auf einer Skala von 1 bis 10\n"
        "je Kriterium:\n\n"
        + punkte + "\n\n"
        "Wenn zu einem Kriterium nichts im Text steht, gib 3 und erwaehne die Luecke.\n\n"
        "Erfasse ausserdem ein paar Eckdaten. Das wird nicht benotet, sondern als\n"
        "Fakten gesammelt. Nimm nur, was wirklich im Text steht, und erfinde nichts.\n"
        "Steht zu einem Feld nichts im Text, schreibe genau \"unbekannt\".\n\n"
        "- zeitraum: ab wann und wie lange die Person kann, in einem kurzen Ausdruck.\n"
        "  Beispiele: \"sofort, 3 Monate\", \"ab 15.10., bis 31.1.\", \"ab Februar, unbefristet\"\n"
        "- beschaeftigung: was die Person aktuell macht.\n"
        "  Beispiele: \"Student (BWL, 3. Semester)\", \"Schueler\", \"Angestellt in Vollzeit\",\n"
        "  \"Arbeitssuchend\", \"Azubi im 2. Lehrjahr\", \"Selbststaendig\"\n"
        "- sprachen: alle Sprachen, die die Person nennt, jeweils mit dem Niveau, das\n"
        "  im Text steht. Muttersprache zuerst, mehrere durch Komma trennen.\n"
        "  Beispiele: \"Deutsch (Muttersprache), Englisch (fliessend), Tuerkisch (Grundkenntnisse)\",\n"
        "  \"Deutsch (Muttersprache), Englisch (B2), Franzoesisch (Schulniveau)\".\n"
        "  Schaetze das Niveau nicht. Nennt der Text eine Sprache ohne Niveau, schreibe\n"
        "  \"Spanisch (Niveau nicht genannt)\". Nennt der Text gar keine Sprachen,\n"
        "  schreibe \"unbekannt\".\n\n"
        "Antworte ausschliesslich mit JSON, ohne Vorrede und ohne Codebloecke:\n"
        "{" + felder + ', "fuehrerschein": "ja|nein|unbekannt", '
        '"verfuegbarkeit": {"zeitraum": "", "beschaeftigung": "", "sprachen": ""}, '
        '"zusammenfassung": "ein bis zwei Saetze"}\n\n'
        "Hier ist die Bewerbung:\n\n"
    )


PROMPT = prompt_bauen()


VERFUEG_FELDER = [
    ("zeitraum", "Zeitliche Verfügbarkeit"),
    ("beschaeftigung", "Aktuelle Beschäftigung"),
    ("sprachen", "Sprachen"),
]
# Laengenbegrenzung je Feld. Sprachen mit Niveau brauchen etwas mehr Platz.
VERFUEG_MAXLAENGE = {"zeitraum": 70, "beschaeftigung": 70, "sprachen": 140}
# Werte, die nur "nichts bekannt" ausdruecken. Sie werden in der Tabelle grau gesetzt.
LEERE_WERTE = ("unbekannt", "keine genannt", "keine angabe", "")


def leere_verfuegbarkeit():
    """Alle Verfuegbarkeitsfelder auf unbekannt."""
    return {schluessel: "unbekannt" for schluessel, _ in VERFUEG_FELDER}


def verfuegbarkeit_der_reihe(reihe):
    """Liest die gespeicherte Verfuegbarkeit, auch bei alten Eintraegen.

    Aeltere Eintraege kennen noch die Felder ab und dauer. Sie werden im neuen
    Feld zeitraum zusammengefuehrt, damit nichts verloren geht.
    """
    try:
        roh = reihe["verfuegbarkeit"]
    except (IndexError, KeyError):
        return leere_verfuegbarkeit()
    if not roh:
        return leere_verfuegbarkeit()
    try:
        geladen = json.loads(roh)
    except (ValueError, TypeError):
        return leere_verfuegbarkeit()
    grund = leere_verfuegbarkeit()
    if not isinstance(geladen, dict):
        return grund

    def lesen(schluessel):
        wert = str(geladen.get(schluessel, "")).strip()
        return "" if wert.lower() in LEERE_WERTE else wert

    for schluessel, _ in VERFUEG_FELDER:
        wert = str(geladen.get(schluessel, "")).strip()
        if wert:
            grund[schluessel] = wert
    if "zeitraum" not in geladen:
        alt = ", ".join(t for t in (lesen("ab"), lesen("dauer")) if t)
        if alt:
            grund["zeitraum"] = alt
    return grund


def verfuegbarkeit_aus(daten):
    """Holt die Verfuegbarkeit aus der KI-Antwort und raeumt sie auf."""
    roh = daten.get("verfuegbarkeit")
    if not isinstance(roh, dict):
        return leere_verfuegbarkeit()
    sauber = {}
    for schluessel, _ in VERFUEG_FELDER:
        grenze = VERFUEG_MAXLAENGE.get(schluessel, 70)
        wert = roh.get(schluessel)
        wert = "" if wert is None else str(wert).strip()[:grenze]
        sauber[schluessel] = wert or "unbekannt"
    return sauber


def fehler_text(fehler):
    """Macht aus einem Fehler des KI-Aufrufs einen verstaendlichen deutschen Hinweis."""
    art = type(fehler).__name__
    roh = str(fehler)
    klein = roh.lower()
    if "credit balance" in klein or "billing" in klein:
        return "Das Guthaben bei Anthropic ist aufgebraucht. Bitte in der Konsole aufladen."
    if art == "AuthenticationError" or "invalid x-api-key" in klein:
        return "Der API-Schluessel wird nicht akzeptiert. ANTHROPIC_API_KEY bei Railway pruefen."
    if art == "NotFoundError" or "model" in klein and "not found" in klein:
        return f"Das Modell \"{MODELL}\" wurde nicht gefunden. AI_MODEL bei Railway pruefen."
    if art in ("RateLimitError", "InternalServerError", "APIConnectionError",
               "APITimeoutError") or "overloaded" in klein:
        return "Die KI ist gerade ueberlastet oder nicht erreichbar. Bitte spaeter erneut versuchen."
    return f"{art}: {roh}"[:220]


def bewerten(betreff, text):
    """Bewertet eine Bewerbung. Bei Fehlern kommt ein Ergebnis mit Schluessel "fehler"."""
    if not API_KEY:
        return {k: 5 for k, _, _, _ in KRITERIEN} | {
            "fuehrerschein": "unbekannt",
            "verfuegbarkeit": leere_verfuegbarkeit(),
            "zusammenfassung": "Keine KI-Bewertung aktiv (ANTHROPIC_API_KEY fehlt).",
        }

    def misslungen(hinweis, erstattet):
        # erstattet: Der Aufruf selbst ist gescheitert, es ist also nichts angefallen.
        return {k: 3 for k, _, _, _ in KRITERIEN} | {
            "fuehrerschein": "unbekannt",
            "verfuegbarkeit": leere_verfuegbarkeit(),
            "zusammenfassung": f"Bewertung fehlgeschlagen. {hinweis}",
            "fehler": hinweis, "erstattet": erstattet,
        }

    try:
        import anthropic

        klient = anthropic.Anthropic(api_key=API_KEY)
        antwort = klient.messages.create(
            model=MODELL,
            max_tokens=900,
            messages=[{"role": "user",
                       "content": PROMPT + f"Betreff: {betreff}\n\nText:\n{text[:12000]}"}],
        )
    except Exception as fehler:
        return misslungen(fehler_text(fehler), erstattet=True)

    try:
        roh = "".join(b.text for b in antwort.content if b.type == "text")
        roh = re.sub(r"```(?:json)?|```", "", roh).strip()
        daten = json.loads(roh)
        if not isinstance(daten, dict):
            raise ValueError("keine JSON-Antwort")
    except Exception:
        # Die KI hat geantwortet und Geld gekostet, die Antwort war nur nicht lesbar.
        return misslungen("Die KI-Antwort war nicht lesbar. Bitte erneut hochladen.",
                          erstattet=False)

    ergebnis = {}
    for schluessel, _, _, _ in KRITERIEN:
        try:
            ergebnis[schluessel] = max(1, min(10, int(round(float(daten.get(schluessel, 3))))))
        except Exception:
            ergebnis[schluessel] = 3
    fuehrer = str(daten.get("fuehrerschein") or "").strip().lower()
    ergebnis["fuehrerschein"] = fuehrer if fuehrer in ("ja", "nein") else "unbekannt"
    ergebnis["verfuegbarkeit"] = verfuegbarkeit_aus(daten)
    ergebnis["zusammenfassung"] = str(daten.get("zusammenfassung") or "")[:500]
    return ergebnis


def gesamtnote(scores):
    summe = sum(scores[k] * g for k, _, g, _ in KRITERIEN)
    return round(summe / GEWICHT_SUMME, 1)


def status_aus(note, scores):
    fehlend = sum(1 for k, _, _, _ in KRITERIEN if scores[k] == 3)
    if fehlend >= 3:
        return "grau"
    if note >= 7.5:
        return "gruen"
    if note >= 5:
        return "gelb"
    return "rot"


class LimitErreicht(Exception):
    """Das Tageskontingent ist aufgebraucht."""


def verarbeiten(daten):
    daten = dict(daten)
    daten["name"] = sauber(daten.get("name"), 120) or "Unbekannt"
    daten["absender"] = sauber(daten.get("absender"), 200)
    daten["betreff"] = sauber(daten.get("betreff"), 300) or "(ohne Betreff)"
    daten["text"] = STEUERZEICHEN.sub("", daten.get("text") or "")
    if API_KEY and not kontingent_buchen():
        raise LimitErreicht(
            f"Tageslimit von {TAGESLIMIT} Bewertungen erreicht. "
            "Morgen geht es weiter, oder das Limit bei Railway hochsetzen.")
    scores = bewerten(daten["betreff"], daten["text"])
    if scores.get("fehler"):
        # Fehlgeschlagene Bewertung: klar markieren statt Dreier vorzutaeuschen.
        if scores.get("erstattet") and API_KEY:
            kontingent_zurueck()
        note, status, nur_scores = 0, "fehler", {}
    else:
        note = gesamtnote(scores)
        status = status_aus(note, scores)
        nur_scores = {k: scores[k] for k, _, _, _ in KRITERIEN}

    verfuegbar = scores.get("verfuegbarkeit") or leere_verfuegbarkeit()

    with conn() as c:
        zeiger = c.execute(
            """INSERT INTO bewerbungen
               (name, absender, betreff, text, status, gesamt, scores,
                zusammenfassung, fuehrerschein, verfuegbarkeit, angelegt)
               VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (daten["name"], daten["absender"], daten["betreff"], daten["text"],
             status, note, json.dumps(nur_scores), scores["zusammenfassung"],
             scores["fuehrerschein"], json.dumps(verfuegbar),
             datetime.now(timezone.utc).isoformat()),
        )
        neue_id = zeiger.lastrowid

    return {"id": neue_id, "name": daten["name"], "betreff": daten["betreff"],
            "status": status, "gesamt": note, "scores": nur_scores,
            "fuehrerschein": scores["fuehrerschein"],
            "verfuegbarkeit": verfuegbar,
            "zusammenfassung": scores["zusammenfassung"]}


# ---------------------------------------------------------------- Web

app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)
init()


def regelmaessig_aufraeumen():
    """Loescht alte Bewerbungen stuendlich, auch wenn niemand die Seite oeffnet."""
    while True:
        try:
            aufraeumen()
        except Exception:
            pass
        time.sleep(3600)


threading.Thread(target=regelmaessig_aufraeumen, daemon=True).start()

# --- Zutritt ---------------------------------------------------------------
# Ein gemeinsames Passwort fuers Team. Wer es kennt, bekommt einen Cookie.
# Ohne gesetztes Passwort bleibt die Seite offen und warnt sichtbar davor.

FEHLVERSUCHE = {}          # Adresse -> [Anzahl, Zeitpunkt des ersten Versuchs]
SPERRE_AB = 8              # so viele Fehlversuche
SPERRE_DAUER = 15 * 60     # dann so lange Pause, in Sekunden
OFFEN = ("/login", "/static/logo.png")


def gesperrt(adresse):
    eintrag = FEHLVERSUCHE.get(adresse)
    if not eintrag:
        return False
    anzahl, seit = eintrag
    if time.time() - seit > SPERRE_DAUER:
        FEHLVERSUCHE.pop(adresse, None)
        return False
    return anzahl >= SPERRE_AB


def fehlversuch(adresse):
    anzahl, seit = FEHLVERSUCHE.get(adresse, (0, time.time()))
    if time.time() - seit > SPERRE_DAUER:
        anzahl, seit = 0, time.time()
    FEHLVERSUCHE[adresse] = (anzahl + 1, seit)
    if len(FEHLVERSUCHE) > 500:       # Speicher nicht volllaufen lassen
        jetzt = time.time()
        for schluessel in [a for a, (_, t) in FEHLVERSUCHE.items()
                           if jetzt - t > SPERRE_DAUER]:
            FEHLVERSUCHE.pop(schluessel, None)


async def groesse_pruefen(request, call_next):
    """Weist zu grosse Uploads ab, bevor sie eingelesen und zwischengespeichert werden."""
    if request.method == "POST":
        try:
            laenge = int(request.headers.get("content-length", "0"))
        except ValueError:
            laenge = 0
        if laenge > MAX_MB * 1024 * 1024 + 64 * 1024:     # etwas Luft fuer Formularkopf
            return JSONResponse({"fehler": f"Datei groesser als {MAX_MB:g} MB. "
                                           "Bitte kleiner speichern oder den Text einfuegen."}, 413)
    return await call_next(request)


@app.middleware("http")
async def tuersteher(request: Request, call_next):
    """Laesst nur durch, wer angemeldet ist. Gilt fuer jede Route."""
    if request.url.path in OFFEN:
        return await call_next(request)
    if not PASSWORT:
        return await groesse_pruefen(request, call_next)
    if secrets.compare_digest(request.cookies.get("zutritt", "").encode(), MARKE.encode()):
        return await groesse_pruefen(request, call_next)
    if request.method == "GET" and not request.url.path.startswith("/api/"):
        ziel = request.url.path + (f"?{request.url.query}" if request.url.query else "")
        if ziel == "/":
            return RedirectResponse("/login", status_code=303)
        return RedirectResponse(f"/login?weiter={quote(ziel, safe='')}", status_code=303)
    return JSONResponse({"fehler": "Nicht angemeldet. Seite neu laden."}, 401)


LOGIN_MELDUNGEN = {
    "passwort": "Passwort stimmt nicht.",
    "gesperrt": "Zu viele Versuche. Bitte 15 Minuten warten.",
}


def sicheres_ziel(ziel):
    """Nur Pfade dieser App sind als Ruecksprung erlaubt, keine fremden Seiten."""
    if (ziel and re.fullmatch(r"/[A-Za-z0-9/_\-.?=&%]*", ziel)
            and not ziel.startswith("//") and not ziel.startswith("/login")):
        return ziel
    return "/"


@app.get("/login", response_class=HTMLResponse)
def login_seite(fehler: str = "", weiter: str = ""):
    if not PASSWORT:
        return RedirectResponse("/", status_code=303)
    text = LOGIN_MELDUNGEN.get(fehler, "")
    meldung = f'<p class="warn">{text}</p>' if text else ""
    meldung += (f'<input type="hidden" name="weiter" value="{html.escape(sicheres_ziel(weiter))}">')
    return (LOGIN.replace("{{FARBE_DUNKEL}}", abdunkeln(FARBE))
                 .replace("{{FARBE_HELL}}", FARBE_HELL)
                 .replace("{{AKZENT}}", AKZENT)
                 .replace("{{FARBE}}", FARBE)
                 .replace("{{KUERZEL}}", kuerzel_aus(FIRMA))
                 .replace("{{MELDUNG}}", meldung))


def absender_adresse(request):
    """Adresse des Besuchers. Hinter dem Railway-Proxy steht sie im letzten Eintrag
    von X-Forwarded-For, den der Proxy selbst anhaengt (vorne kann jeder etwas faelschen)."""
    weiter = request.headers.get("x-forwarded-for", "")
    if weiter.strip():
        return weiter.split(",")[-1].strip()
    return request.client.host if request.client else "unbekannt"


@app.post("/login")
async def anmelden(request: Request):
    # Formular selbst lesen und nach wenigen KB abbrechen: Diese Seite ist ohne
    # Anmeldung erreichbar und darf keine grossen Datenmengen annehmen.
    roh = b""
    async for stueck in request.stream():
        roh += stueck
        if len(roh) > 4096:
            return JSONResponse({"fehler": "Anfrage zu gross."}, 413)
    felder = parse_qs(roh.decode("utf-8", errors="replace"))
    passwort = (felder.get("passwort") or [""])[0]
    weiter = (felder.get("weiter") or [""])[0]
    adresse = absender_adresse(request)
    ziel = sicheres_ziel(weiter)
    zurueck = f"&weiter={quote(ziel, safe='')}" if ziel != "/" else ""
    if gesperrt(adresse):
        return RedirectResponse(
            "/login?fehler=gesperrt" + zurueck,
            status_code=303)
    if not secrets.compare_digest(passwort.encode(), PASSWORT.encode()):
        fehlversuch(adresse)
        return RedirectResponse("/login?fehler=passwort" + zurueck,
                                status_code=303)
    FEHLVERSUCHE.pop(adresse, None)
    antwort = RedirectResponse(ziel, status_code=303)
    antwort.set_cookie("zutritt", MARKE, max_age=30 * 24 * 3600,
                       httponly=True, samesite="lax", secure=True)
    return antwort


@app.get("/", response_class=HTMLResponse)
def seite():
    aufraeumen()
    warnung = ""
    if not PASSWORT:
        warnung = ('<div class="warnbalken"><b>Diese Seite ist ungeschützt.</b>'
                   'Jeder, der die Adresse kennt, kann Bewerbungen hochladen und '
                   'damit Kosten verursachen. Bei Railway eine Variable PASSWORT '
                   'anlegen und neu starten.</div>')
    return (SEITE.replace("{{FIRMA}}", FIRMA)
                 .replace("{{KUERZEL}}", kuerzel_aus(FIRMA))
                 .replace("{{FARBE_DUNKEL}}", abdunkeln(FARBE))
                 .replace("{{FARBE_HELL}}", FARBE_HELL)
                 .replace("{{AKZENT}}", AKZENT)
                 .replace("{{FARBE}}", FARBE)
                 .replace("{{WARNUNG}}", warnung)
                 .replace("{{TAGE}}", str(AUFBEWAHRUNG_TAGE)))


LOGO_URL = os.environ.get("LOGO_URL", "")   # Falls kein logo.png im Repo liegt.


@app.get("/static/logo.png")
def logo():
    from fastapi.responses import FileResponse, RedirectResponse

    if os.path.exists("logo.png"):
        return FileResponse("logo.png")
    if LOGO_URL:
        return RedirectResponse(LOGO_URL)
    return JSONResponse({"fehler": "kein Logo hinterlegt"}, 404)


def datei_auswerten(rohdaten, dateiname, endung):
    """Liest Name, Absender und Text aus einer Datei. Gibt (daten, fehlertext) zurueck."""
    if endung == "eml":
        try:
            return eml_lesen(rohdaten), ""
        except Exception as fehler:
            return None, f"Datei nicht lesbar: {sauber(fehler, 150)}"
    if endung == "pdf":
        text = pdf_zu_text(rohdaten)
        if not text:
            return None, ("Aus dieser PDF laesst sich kein Text lesen. "
                          "Vermutlich ein Scan — bitte den Text von Hand einfuegen.")
    elif endung in ("txt", "text", "md", "rtf"):
        text = text_entschluesseln(rohdaten)
    else:
        return None, "Moegliche Formate: .eml, .pdf, .txt"
    return {"name": name_raten(text, dateiname), "absender": mail_raten(text),
            "betreff": f"Bewerbung ({dateiname})", "text": text}, ""


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    # Blockweise lesen und abbrechen, sobald die Grenze ueberschritten ist.
    # So landet eine riesige Datei nie vollstaendig im Speicher.
    grenze = int(MAX_MB * 1024 * 1024)
    stuecke, gelesen = [], 0
    while True:
        stueck = await file.read(256 * 1024)
        if not stueck:
            break
        gelesen += len(stueck)
        if gelesen > grenze:
            return JSONResponse(
                {"fehler": f"Datei groesser als {MAX_MB:g} MB. "
                           "Bitte kleiner speichern oder den Text einfuegen."}, 413)
        stuecke.append(stueck)
    rohdaten = b"".join(stuecke)

    dateiname = file.filename or ""
    endung = dateiname.lower().rsplit(".", 1)[-1] if "." in dateiname else ""

    daten, fehlertext = await run_in_threadpool(datei_auswerten, rohdaten, dateiname, endung)
    if fehlertext:
        return JSONResponse({"fehler": fehlertext}, 400)

    if not daten["text"].strip():
        return JSONResponse({"fehler": "Die Datei enthaelt keinen Text."}, 400)
    try:
        return await run_in_threadpool(verarbeiten, daten)
    except LimitErreicht as grenze:
        return JSONResponse({"fehler": str(grenze)}, 429)


@app.post("/api/text")
async def per_hand(nutzlast: dict):
    def feld(name, vorgabe="", laenge=200):
        wert = nutzlast.get(name)
        return (str(wert).strip() if wert is not None else "")[:laenge] or vorgabe

    text = feld("text", laenge=10 ** 7)
    if not text:
        return JSONResponse({"fehler": "Kein Text angegeben."}, 400)
    if len(text) > MAX_MB * 1024 * 1024:
        return JSONResponse({"fehler": f"Text groesser als {MAX_MB:g} MB."}, 413)
    try:
        return await run_in_threadpool(verarbeiten, {
            "name": feld("name", "Unbekannt"),
            "absender": feld("absender"),
            "betreff": feld("betreff", "Bewerbung"),
            "text": text,
        })
    except LimitErreicht as grenze:
        return JSONResponse({"fehler": str(grenze)}, 429)


@app.get("/api/liste")
def liste():
    """Alle gespeicherten Bewerbungen, beste Note zuerst."""
    aufraeumen()
    with conn() as c:
        reihen = c.execute(
            "SELECT * FROM bewerbungen "
            "ORDER BY (status = 'fehler') DESC, gesamt DESC, id DESC").fetchall()

    eintraege = []
    for reihe in reihen:
        try:
            scores = json.loads(reihe["scores"])
        except (ValueError, TypeError):
            scores = {}
        eintraege.append({
            "id": reihe["id"],
            "name": reihe["name"],
            "absender": reihe["absender"],
            "betreff": reihe["betreff"],
            "status": reihe["status"],
            "gesamt": reihe["gesamt"],
            "scores": scores,
            "verfuegbarkeit": verfuegbarkeit_der_reihe(reihe),
            "fuehrerschein": reihe["fuehrerschein"],
            "zusammenfassung": reihe["zusammenfassung"],
        })
    return {"anzahl": len(eintraege), "eintraege": eintraege,
            "verbraucht": verbrauch_heute(), "limit": TAGESLIMIT}


@app.post("/api/loeschen/{eintrag_id}")
def loeschen(eintrag_id: int):
    """Entfernt eine einzelne Bewerbung, etwa einen Fehlversuch."""
    with conn() as c:
        c.execute("DELETE FROM bewerbungen WHERE id=?", (eintrag_id,))
    return {"geloescht": eintrag_id}


@app.get("/mail/{eintrag_id}", response_class=HTMLResponse)
def mail_ansehen(eintrag_id: int):
    aufraeumen()
    with conn() as c:
        reihe = c.execute("SELECT * FROM bewerbungen WHERE id=?", (eintrag_id,)).fetchone()
    if reihe is None:
        return HTMLResponse(
            f"<body style='font-family:system-ui;padding:40px;max-width:600px;margin:auto'>"
            f"<h2>Nicht mehr vorhanden</h2><p>Bewerbungen werden nach "
            f"{AUFBEWAHRUNG_TAGE} Tagen automatisch geloescht.</p></body>", 404)
    sicher = html.escape(reihe["text"] or "", quote=False)
    note = ("Bewertung fehlgeschlagen" if reihe["status"] == "fehler"
            else f"Gesamtnote {reihe['gesamt']}/10")
    return HTMLResponse(f"""<!doctype html><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<body style="font-family:system-ui,-apple-system,sans-serif;max-width:760px;margin:0 auto;padding:28px;color:#1c1c1a">
<div style="background:{FARBE};color:#fff;padding:16px 20px;border-radius:12px">
<div style="font-size:13px;opacity:.85">{html.escape(reihe['betreff'] or '')}</div>
<div style="font-size:20px;font-weight:600;margin-top:2px">{html.escape(reihe['name'] or '')}</div></div>
<p style="color:#6b6963;font-size:13px;margin:14px 0 4px">{html.escape(reihe['absender'] or '')}</p>
<p style="color:#6b6963;font-size:13px;margin:0 0 18px">{note}</p>
<pre style="white-space:pre-wrap;font:inherit;line-height:1.65;background:#faf9f7;border:1px solid #e6e3dd;border-radius:12px;padding:18px">{sicher}</pre>
</body>""")


# ---------------------------------------------------------------- Excel

def berlin_zeit(zeitpunkt):
    """Rechnet eine UTC-Zeit in deutsche Ortszeit um (Sommerzeit nach EU-Regel)."""
    def letzter_sonntag(monat):
        tag = datetime(zeitpunkt.year, monat, 31, 1, 0, tzinfo=timezone.utc)
        while tag.weekday() != 6:
            tag -= timedelta(days=1)
        return tag

    sommer = letzter_sonntag(3) <= zeitpunkt < letzter_sonntag(10)
    return (zeitpunkt + timedelta(hours=2 if sommer else 1)).replace(tzinfo=None, microsecond=0)


def eingang_als_datum(iso):
    """Wandelt den gespeicherten UTC-Zeitstempel in deutsche Ortszeit um."""
    try:
        d = datetime.fromisoformat(iso)
    except (TypeError, ValueError):
        return None
    if d.tzinfo is None:
        d = d.replace(tzinfo=timezone.utc)
    return berlin_zeit(d.astimezone(timezone.utc))


def zeilen_noetig(text, breite):
    """Schaetzt, wie viele Zeilen ein Text in einer Spalte dieser Breite braucht."""
    if not text:
        return 1
    zeichen = max(8, int(breite * 1.1))
    return sum(max(1, math.ceil(len(teil) / zeichen)) for teil in str(text).split("\n"))


VERFUEG_BREITEN = {"zeitraum": 24, "beschaeftigung": 26, "sprachen": 36}


def tabelle_bauen(reihen, basis):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bewerbungen"
    ws.sheet_view.showGridLines = False

    schrift = "Calibri"
    linie = Side(style="thin", color="D9DEE3")
    rand = Border(left=linie, right=linie, top=linie, bottom=linie)

    anzahl_krit = len(KRITERIEN)
    anzahl_verf = len(VERFUEG_FELDER)
    kopfzeile = (["Status", "Name", "Absender", "Eingang", "Gesamt (1-10)"]
                 + [t for _, t, _, _ in KRITERIEN]
                 + [t for _, t in VERFUEG_FELDER]
                 + ["Führerschein", "Zusammenfassung", "Einladung", "Mail"])

    # Spaltenbreiten richten sich nach den Titeln, nicht nach festen Positionen.
    def titelbreite(titel, mindest):
        laengstes = max((len(w) for w in str(titel).split()), default=0)
        return max(mindest, laengstes + 3)

    breiten = [15, 22, 28, 17, 11]
    breiten += [titelbreite(t, 12) for _, t, _, _ in KRITERIEN]
    breiten += [VERFUEG_BREITEN.get(k, 22) for k, _ in VERFUEG_FELDER]
    breiten += [14, 62, 12, 10]

    spalte_krit_von = 6
    spalte_krit_bis = 5 + anzahl_krit
    spalte_verf_von = spalte_krit_bis + 1
    spalte_verf_bis = spalte_krit_bis + anzahl_verf
    spalte_fuehrer = spalte_verf_bis + 1
    spalte_zusammen = spalte_fuehrer + 1
    spalte_einladung = spalte_zusammen + 1
    spalte_mail = spalte_einladung + 1

    for spalte, titel in enumerate(kopfzeile, 1):
        zelle = ws.cell(row=1, column=spalte, value=titel)
        zelle.fill = PatternFill("solid", start_color="2F4F4F")
        zelle.font = Font(name=schrift, bold=True, color="FFFFFF", size=10)
        zelle.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        zelle.border = rand
        ws.column_dimensions[get_column_letter(spalte)].width = breiten[spalte - 1]
    ws.row_dimensions[1].height = 34

    for nr, reihe in enumerate(reihen, start=2):
        scores = json.loads(reihe["scores"])
        verfuegbar = verfuegbarkeit_der_reihe(reihe)
        farbe, beschriftung = STATUS_FARBEN.get(reihe["status"], ("FFFFFF", ""))
        hintergrund = "F6F8FA" if nr % 2 == 1 else "FFFFFF"
        eingang = eingang_als_datum(reihe["angelegt"])

        fehler = reihe["status"] == "fehler"
        werte = ([beschriftung, reihe["name"], reihe["absender"], eingang,
                  None if fehler else reihe["gesamt"]]
                 + [scores.get(k, "") for k, _, _, _ in KRITERIEN]
                 + [verfuegbar.get(k, "unbekannt") for k, _ in VERFUEG_FELDER]
                 + [reihe["fuehrerschein"], reihe["zusammenfassung"], "", ""])

        for spalte, wert in enumerate(werte, 1):
            if isinstance(wert, str):
                wert = sauber(wert, 32000)
            zelle = ws.cell(row=nr, column=spalte, value=wert)
            if isinstance(wert, str) and wert.startswith("="):
                zelle.data_type = "s"   # Fremder Text darf nie als Formel laufen.
            zelle.border = rand
            zelle.fill = PatternFill("solid", start_color=hintergrund)
            zelle.font = Font(name=schrift, size=10, color="1F2933")
            zelle.alignment = Alignment(vertical="center", wrap_text=True)

        status = ws.cell(row=nr, column=1)
        status.fill = PatternFill("solid", start_color=farbe)
        status.font = Font(name=schrift, size=10, bold=True, color="1F2933")
        status.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        gesamt = ws.cell(row=nr, column=5)
        gesamt.fill = PatternFill("solid", start_color=farbe)
        gesamt.font = Font(name=schrift, size=12, bold=True, color="1F2933")
        gesamt.number_format = "0.0"
        gesamt.alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=nr, column=4).number_format = "DD.MM.YYYY HH:MM"
        ws.cell(row=nr, column=4).alignment = Alignment(horizontal="center", vertical="center")

        for spalte in range(spalte_krit_von, spalte_krit_bis + 1):
            zelle = ws.cell(row=nr, column=spalte)
            zelle.alignment = Alignment(horizontal="center", vertical="center")
            zelle.font = Font(name=schrift, size=11, bold=True, color="1F2933")
            zelle.number_format = "0"

        # "unbekannt" tritt in den Hintergrund, damit echte Angaben auffallen.
        for spalte in range(spalte_verf_von, spalte_verf_bis + 1):
            zelle = ws.cell(row=nr, column=spalte)
            if str(zelle.value).strip().lower() in LEERE_WERTE:
                zelle.font = Font(name=schrift, size=10, italic=True, color="9AA0A6")

        fuehrer = ws.cell(row=nr, column=spalte_fuehrer)
        fuehrer.alignment = Alignment(horizontal="center", vertical="center")
        wert_f = str(fuehrer.value or "").strip().lower()
        if wert_f in LEERE_WERTE:
            fuehrer.font = Font(name=schrift, size=10, italic=True, color="9AA0A6")
        elif wert_f == "nein":
            fuehrer.font = Font(name=schrift, size=10, bold=True, color="B42318")
        ws.cell(row=nr, column=spalte_zusammen).alignment = Alignment(
            vertical="center", wrap_text=True)

        link = ws.cell(row=nr, column=spalte_mail)
        link.alignment = Alignment(horizontal="center", vertical="center")
        if basis:
            link.value = "Öffnen"
            link.hyperlink = f"{basis}/mail/{reihe['id']}"
            link.font = Font(name=schrift, size=10, color="0563C1", underline="single")

        zeilen = max(
            zeilen_noetig(reihe["zusammenfassung"], breiten[spalte_zusammen - 1]),
            zeilen_noetig(verfuegbar.get("sprachen", ""),
                          breiten[spalte_verf_bis - 1]),
            2,
        )
        ws.row_dimensions[nr].height = min(150, 13.5 * zeilen + 8)

    letzte = len(reihen) + 1
    if reihen:
        # Farbskala von Rot (1) ueber Gelb zu Gruen (10), gleich fuer alle Kriterien.
        bereich = (f"{get_column_letter(spalte_krit_von)}2:"
                   f"{get_column_letter(spalte_krit_bis)}{letzte}")
        ws.conditional_formatting.add(bereich, ColorScaleRule(
            start_type="num", start_value=1, start_color="F8A5A5",
            mid_type="num", mid_value=5.5, mid_color="FFE699",
            end_type="num", end_value=10, end_color="8FD19E"))

        # Auswahlliste fuer die Spalte "Einladung", damit man sie schnell pflegen kann.
        auswahl = DataValidation(type="list", formula1='"Ja,Nein,Offen"', allow_blank=True)
        ws.add_data_validation(auswahl)
        spalte_e = get_column_letter(spalte_einladung)
        auswahl.add(f"{spalte_e}2:{spalte_e}{letzte}")
        for nr in range(2, letzte + 1):
            ws.cell(row=nr, column=spalte_einladung).alignment = Alignment(
                horizontal="center", vertical="center")

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(kopfzeile))}{max(1, letzte)}"

    hinweis = letzte + 2
    ws.cell(row=hinweis, column=1,
            value=f"Mail-Links sind {AUFBEWAHRUNG_TAGE} Tage gültig. "
                  f"Stand: {berlin_zeit(datetime.now(timezone.utc)).strftime('%d.%m.%Y %H:%M')}"
            ).font = Font(name=schrift, size=9, italic=True, color="888888")

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = "1:1"
    return wb


@app.get("/api/excel")
def excel(request: Request):
    aufraeumen()
    with conn() as c:
        reihen = c.execute(
            "SELECT * FROM bewerbungen "
            "ORDER BY (status = 'fehler') DESC, gesamt DESC, id DESC").fetchall()

    # Adresse fuer die Mail-Links: Einstellung oder die Adresse, unter der die App gerade laeuft.
    basis = os.environ.get("BASIS_URL", "").rstrip("/")
    if not basis:
        host = request.headers.get("x-forwarded-host") or request.headers.get("host", "")
        basis = f"https://{host}" if host else ""

    wb = tabelle_bauen(reihen, basis)
    puffer = io.BytesIO()
    wb.save(puffer)
    puffer.seek(0)
    return StreamingResponse(
        puffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":
                 f'attachment; filename="bewerbungen_{heute()}.xlsx"'},
    )


# ---------------------------------------------------------------- Oberflaeche

LOGIN = """<!doctype html><html lang="de"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Anmelden</title><style>
*{box-sizing:border-box}
body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;
padding:20px;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",system-ui,sans-serif;
color:#152238;
background:
radial-gradient(900px 520px at 88% 96%,rgba(120,170,235,.20),transparent 62%),
radial-gradient(700px 420px at 4% 4%,rgba(150,195,245,.16),transparent 60%),
linear-gradient(170deg,#f4f8fe 0%,#e9f1fc 100%)}
.box{background:#fff;border-radius:20px;padding:34px 30px;width:100%;max-width:380px;
box-shadow:0 10px 36px rgba(30,70,150,.14);text-align:center}
.zeichen{width:68px;height:68px;margin:0 auto 16px;border-radius:50%;
background:linear-gradient(140deg,{{FARBE}},{{FARBE_DUNKEL}});color:#fff;
display:flex;align-items:center;justify-content:center;font-size:21px;font-weight:700}
h1{margin:0 0 4px;font-size:21px;font-weight:700;letter-spacing:-.02em;color:#12224a}
p.sub{margin:0 0 22px;font-size:13.5px;color:#7e8ca6}
input{width:100%;padding:13px 15px;border:1px solid #d5e0f2;border-radius:12px;
font-size:16px;font-family:inherit;background:#fbfcff;color:#152238;text-align:center}
input:focus{outline:0;border-color:{{FARBE}};background:#fff}
button{margin-top:12px;width:100%;background:linear-gradient(100deg,{{FARBE}},#2f6ae0);
color:#fff;border:0;border-radius:12px;padding:14px;font-size:15.5px;font-weight:600;
font-family:inherit;cursor:pointer;box-shadow:0 7px 20px rgba(30,80,190,.30)}
button:active{transform:translateY(1px)}
.warn{margin:0 0 14px;background:#fdeceb;color:{{AKZENT}};border-radius:11px;
padding:11px 14px;font-size:13.5px;font-weight:500}
</style></head><body>
<form class="box" method="post" action="/login">
<div class="zeichen">{{KUERZEL}}</div>
<h1>Bewerbungen</h1>
<p class="sub">Bitte Passwort eingeben</p>
{{MELDUNG}}
<input type="password" name="passwort" placeholder="Passwort" autofocus
autocomplete="current-password">
<button type="submit">Anmelden</button>
</form></body></html>"""


SEITE = """<!doctype html><html lang="de"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Bewerbungen — {{FIRMA}}</title><style>
*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",system-ui,sans-serif;
color:#152238;line-height:1.6;min-height:100vh;
background:
radial-gradient(900px 520px at 88% 96%,rgba(120,170,235,.20),transparent 62%),
radial-gradient(700px 420px at 4% 4%,rgba(150,195,245,.16),transparent 60%),
linear-gradient(170deg,#f4f8fe 0%,#e9f1fc 100%)}
.huelle{max-width:860px;margin:0 auto;padding:18px 16px 40px}

.marke{display:flex;align-items:center;gap:14px;padding:6px 2px 2px}
.marke img{height:66px;width:auto;flex:0 0 auto}
.marke .kuerzel{height:66px;width:66px;flex:0 0 66px;border-radius:50%;
background:linear-gradient(140deg,{{FARBE}},{{FARBE_DUNKEL}});color:#fff;
display:flex;align-items:center;justify-content:center;font-size:20px;font-weight:700}
.marke h1{margin:0;font-size:24px;font-weight:700;letter-spacing:-.02em;color:#12224a}
.marke h1 i{font-style:normal;color:{{AKZENT}}}
.marke p{margin:0;font-size:11px;letter-spacing:.13em;text-transform:uppercase;
color:#8fa2c2;font-weight:500}

.band{margin-top:18px;border-radius:18px;padding:22px 24px;color:#fff;position:relative;
overflow:hidden;display:flex;align-items:center;gap:16px;
background:linear-gradient(115deg,{{FARBE_DUNKEL}} 0%,{{FARBE}} 58%,#2a5fd0 100%);
box-shadow:0 10px 30px rgba(20,50,120,.26)}
.band:before{content:"";position:absolute;right:-90px;top:-70px;width:280px;height:280px;
border-radius:50%;background:radial-gradient(circle at 34% 34%,
rgba(120,190,255,.55),rgba(60,120,220,.20) 52%,transparent 72%)}
.band .ikon{width:52px;height:52px;flex:0 0 52px;border-radius:14px;
background:rgba(255,255,255,.16);display:flex;align-items:center;justify-content:center;
font-size:24px;position:relative;z-index:1}
.band .txt{position:relative;z-index:1;min-width:0}
.warnbalken{margin-top:16px;background:#fdeceb;border:1px solid #f3cfcd;
border-radius:14px;padding:13px 16px;font-size:13.5px;color:{{AKZENT}};
font-weight:500;line-height:1.5}
.warnbalken b{display:block;font-weight:700;margin-bottom:2px}
.band h2{margin:0;font-size:19px;font-weight:700;letter-spacing:-.015em}
.band p{margin:2px 0 0;font-size:13.5px;opacity:.86}

.zone{margin-top:16px;background:#fff;border-radius:20px;padding:12px;
box-shadow:0 6px 26px rgba(30,70,150,.10);transition:.18s ease}
.zone .innen{border:2px dashed #b9cdea;border-radius:15px;padding:30px 18px;
text-align:center;transition:.18s ease}
.zone.aktiv{transform:scale(1.012)}
.zone.aktiv .innen{border-color:{{FARBE}};background:{{FARBE_HELL}}}
.zone .sym{width:60px;height:60px;margin:0 auto;border-radius:50%;
background:radial-gradient(circle at 50% 38%,#fff,{{FARBE_HELL}});
border:1px solid #dbe7f8;display:flex;align-items:center;justify-content:center;
font-size:26px;box-shadow:0 3px 14px rgba(40,90,180,.13)}
.zone h3{margin:12px 0 2px;font-size:20px;font-weight:700;letter-spacing:-.02em;color:#12224a}
.zone .unter{margin:0;color:#7e8ca6;font-size:14px}
.marken{display:flex;gap:8px;justify-content:center;margin-top:12px;flex-wrap:wrap}
.marken span{font-size:12px;font-weight:700;letter-spacing:.5px;color:{{FARBE}};
background:{{FARBE_HELL}};border:1px solid #dbe7f8;border-radius:9px;padding:5px 15px}

button{font-family:inherit;cursor:pointer;border:0;font-weight:600}
.haupt{margin-top:16px;background:linear-gradient(100deg,{{FARBE}},#2f6ae0);
color:#fff;border-radius:13px;padding:15px 26px;font-size:15.5px;width:100%;max-width:330px;
display:inline-flex;align-items:center;justify-content:center;gap:10px;
box-shadow:0 7px 20px rgba(30,80,190,.30);transition:.15s}
.haupt:active{transform:translateY(1px)}
.haupt:disabled{opacity:.55;box-shadow:none}
.zweit{width:100%;background:#fff;color:{{FARBE}};border:1.6px solid {{FARBE}};
border-radius:13px;padding:13px 20px;font-size:15px;
display:inline-flex;align-items:center;justify-content:center;gap:9px}
.zweit:active{background:{{FARBE_HELL}}}

.trenner{display:flex;align-items:center;gap:12px;margin:20px 0 13px;
color:#8fa2c2;font-size:13px;font-weight:600}
.trenner:before,.trenner:after{content:"";flex:1;height:1px;background:#d6e2f3}

.karte{background:#fff;border-radius:20px;padding:18px 20px;margin-top:16px;
box-shadow:0 6px 26px rgba(30,70,150,.10)}
.karte.aus{display:none}
input,textarea{width:100%;padding:12px 14px;border:1px solid #d5e0f2;border-radius:11px;
font-size:15px;font-family:inherit;margin-bottom:10px;background:#fbfcff;color:#152238}
input:focus,textarea:focus{outline:0;border-color:{{FARBE}};background:#fff}
textarea{min-height:150px;resize:vertical}

.fortschritt{margin-top:16px;background:#fff;border-radius:16px;padding:15px 18px;
box-shadow:0 6px 26px rgba(30,70,150,.10);display:none}
.fortschritt.an{display:block}
.fortschritt p{margin:0 0 9px;font-size:14px;font-weight:600;color:{{FARBE}}}
.fortschritt .spur{height:8px;background:#eaeef6;border-radius:5px;overflow:hidden}
.fortschritt .spur i{display:block;height:100%;width:0;border-radius:5px;
background:linear-gradient(90deg,{{FARBE}},#4f8fe8);transition:width .3s ease}

.fehler{background:#fdeceb;color:{{AKZENT}};border-radius:12px;padding:12px 15px;
font-size:14px;margin-top:12px;font-weight:500}

.tabkopf{display:flex;align-items:center;justify-content:space-between;gap:12px;
flex-wrap:wrap;margin-bottom:4px}
.tabkopf h2{margin:0;font-size:18px;font-weight:700;color:#12224a;letter-spacing:-.01em}
.tabkopf .zahl{font-size:13px;color:#7e8ca6;font-weight:500}
.tabkopf .xls{background:linear-gradient(100deg,{{FARBE}},#2f6ae0);color:#fff;
border-radius:11px;padding:11px 18px;font-size:14px;display:inline-flex;
align-items:center;gap:8px;box-shadow:0 5px 15px rgba(30,80,190,.26)}

.leer{text-align:center;padding:26px 10px;color:#9aa9c2;font-size:14px}

.reihe{border-top:1px solid #eef3fb}
.reihe:first-of-type{border-top:0}
.kopfzeile{display:flex;align-items:center;gap:12px;padding:13px 2px;cursor:pointer}
.ampel{width:10px;height:38px;border-radius:5px;flex:0 0 10px}
.wer{flex:1;min-width:0}
.wer b{display:block;font-size:15.5px;font-weight:600;color:#12224a;
white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.wer small{display:block;font-size:12.5px;color:#7e8ca6;
white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.wert{text-align:right;flex:0 0 auto}
.wert b{font-size:20px;font-weight:700;color:#12224a;line-height:1.1}
.wert small{display:block;font-size:11px;color:#9aa9c2}
.pfeilchen{flex:0 0 14px;color:#b6c2d6;font-size:13px;transition:transform .2s}
.reihe.offen .pfeilchen{transform:rotate(90deg)}

.detail{display:none;padding:2px 2px 18px 24px}
.reihe.offen .detail{display:block}
.zeile{margin-bottom:10px}
.lab{display:flex;justify-content:space-between;font-size:13px;margin-bottom:4px}
.lab span{color:#7e8ca6}
.lab em{font-style:normal;color:#a7b5cc;font-size:11px;margin-left:5px}
.lab b{font-weight:700;color:#12224a}
.spur{height:8px;background:#eef3fb;border-radius:5px;overflow:hidden}
.spur i{display:block;height:100%;border-radius:5px}
.blocktitel{margin:15px 0 8px;font-size:12px;font-weight:700;letter-spacing:.09em;
text-transform:uppercase;color:#8fa2c2}
.verf{list-style:none;margin:0;padding:0;display:grid;gap:6px}
.verf li{display:grid;grid-template-columns:minmax(120px,175px) 1fr;align-items:baseline;
gap:12px;font-size:14px}
.verf li b{font-weight:500;color:#7e8ca6;font-size:13px}
.verf li span{color:#12224a;font-weight:500}
.verf li.offen span{color:#9aa9c2;font-weight:400;font-style:italic}
.fazit{font-size:14px;color:#48566e;margin:14px 0 0}
.fuss{font-size:13px;color:#7e8ca6;margin:5px 0 0}
.knoepfe{display:flex;gap:9px;margin-top:14px;flex-wrap:wrap}
.knoepfe a,.knoepfe button{font-size:13.5px;font-weight:600;border-radius:10px;
padding:10px 16px;text-decoration:none;display:inline-flex;align-items:center;gap:7px}
.knoepfe a{background:{{FARBE_HELL}};color:{{FARBE}};border:1px solid #d5e0f2}
.knoepfe button{background:#fff;color:{{AKZENT}};border:1px solid #f0d4d3}

.sicher{display:flex;align-items:center;justify-content:center;gap:9px;margin:26px 0 0}
.sicher .schild{font-size:19px}
.sicher p{margin:0;font-size:13px;color:#5b6b86}
.sicher p b{color:#12224a}
.sicher p small{display:block;font-size:11.5px;color:#93a3bf}
</style></head><body><div class="huelle">

<div class="marke">
<img src="/static/logo.png" alt="" onerror="this.outerHTML='<div class=\\'kuerzel\\'>{{KUERZEL}}</div>'">
<div><h1>Places <i>to</i> Be</h1><p>Bewerbungen sichten</p></div>
</div>

{{WARNUNG}}

<div class="band">
<div class="ikon">📄</div>
<div class="txt"><h2>Bewerbungen verarbeiten</h2>
<p>Unterlagen hochladen, Liste füllt sich von selbst.</p></div>
</div>

<div class="zone" id="zone"><div class="innen">
<div class="sym">☁️</div>
<h3>Dateien hier ablegen</h3>
<p class="unter">mehrere auf einmal möglich</p>
<div class="marken"><span>EML</span><span>PDF</span><span>TXT</span></div>
<button class="haupt" id="waehlen" onclick="datei.click()">Dateien auswählen <span>→</span></button>
<input type="file" id="datei" accept=".eml,.pdf,.txt,.md,.rtf" multiple hidden>
</div></div>

<div class="trenner">oder</div>
<button class="zweit" onclick="handForm()">⌨️ Text von Hand eingeben</button>

<div class="karte aus" id="hand">
<input id="h-name" placeholder="Name">
<input id="h-mail" placeholder="E-Mail-Adresse">
<input id="h-betreff" placeholder="Betreff">
<textarea id="h-text" placeholder="Bewerbungstext einfügen"></textarea>
<button class="haupt" style="max-width:none" onclick="sendeText()">Bewerten <span>→</span></button>
</div>

<div class="fortschritt" id="fortschritt">
<p id="f-text">Wird ausgewertet …</p>
<div class="spur"><i id="f-balken"></i></div>
</div>

<div id="fehler"></div>

<div class="karte" id="tabelle">
<div class="tabkopf">
<div><h2>Bewertungen</h2><span class="zahl" id="t-zahl">noch keine</span></div>
<button class="xls" onclick="location.href='/api/excel'">⬇ Excel</button>
</div>
<div id="t-inhalt"><p class="leer">Noch nichts verarbeitet. Lad oben eine Bewerbung hoch.</p></div>
</div>

<div class="sicher"><span class="schild">🛡️</span>
<p><b>Die Entscheidung trifft das Team.</b>
<small>Bewerbungen werden nach {{TAGE}} Tagen automatisch gelöscht.</small></p></div>
</div>

<script>
const zone=document.getElementById('zone'),datei=document.getElementById('datei');
const KRIT=[{{KRITLISTE}}];
const VERF=[{{VERFLISTE}}];
const PILL={gruen:['#d9efdc','#2a6b36','Einladen','#4f8a3d'],
gelb:['#faeecb','#8a6410','Prüfen','#e0a423'],
grau:['#ebe9e3','#5f5c56','Angaben fehlen','#a8b0bf'],
rot:['#f8dcd9','#9c3229','Unpassend','#c4483f'],
fehler:['#fde8d4','#9a4a0b','Fehlgeschlagen','#e07b24']};
const offeneZeilen=new Set();

['dragenter','dragover'].forEach(e=>zone.addEventListener(e,v=>{
v.preventDefault();zone.classList.add('aktiv')}));
['dragleave','drop'].forEach(e=>zone.addEventListener(e,v=>{
v.preventDefault();zone.classList.remove('aktiv')}));
zone.addEventListener('drop',v=>{if(v.dataTransfer.files.length)stapel(v.dataTransfer.files)});
datei.addEventListener('change',()=>{if(datei.files.length)stapel(datei.files)});

function handForm(){document.getElementById('hand').classList.toggle('aus')}
function fehler(t){const ort=document.getElementById('fehler');ort.innerHTML='';
if(!t)return;const box=document.createElement('div');box.className='fehler';
box.style.whiteSpace='pre-line';box.textContent=t;ort.appendChild(box)}
async function antwortLesen(a){
if(a.status===401){location.href='/login';throw new Error('Nicht angemeldet.')}
try{return await a.json()}
catch(e){return{fehler:'Server antwortet nicht richtig (Code '+a.status+'). Bitte erneut versuchen.'}}}
function farbe(n){return n>=7?'#4f8a3d':n>=4?'#c98a1e':'#c4483f'}

function fortschritt(an,text,anteil){
const box=document.getElementById('fortschritt');
box.classList.toggle('an',an);
if(text)document.getElementById('f-text').textContent=text;
document.getElementById('f-balken').style.width=(anteil||0)+'%';
document.getElementById('waehlen').disabled=an}

async function stapel(dateien){
fehler('');const liste=Array.from(dateien),probleme=[];
for(let i=0;i<liste.length;i++){
const f=liste[i];
fortschritt(true,liste.length>1?('Verarbeite '+(i+1)+' von '+liste.length+': '+f.name)
:'Wird ausgewertet …',Math.round(i/liste.length*100));
try{const fd=new FormData();fd.append('file',f);
const a=await fetch('/api/upload',{method:'POST',body:fd});
const d=await antwortLesen(a);
if(!a.ok)throw new Error(d.fehler||'Fehler');
}catch(e){probleme.push(f.name+': '+e.message)}}
fortschritt(false);datei.value='';
if(probleme.length)fehler('Nicht verarbeitet\\n'+probleme.join('\\n'));
await ladeListe()}

async function sendeText(){
const text=document.getElementById('h-text').value.trim();
if(!text){fehler('Bitte den Bewerbungstext einfügen.');return}
fehler('');fortschritt(true,'Wird ausgewertet …',40);
try{const a=await fetch('/api/text',{method:'POST',
headers:{'Content-Type':'application/json'},body:JSON.stringify({
name:document.getElementById('h-name').value,
absender:document.getElementById('h-mail').value,
betreff:document.getElementById('h-betreff').value,text:text})});
const d=await antwortLesen(a);if(!a.ok)throw new Error(d.fehler||'Fehler');
['h-name','h-mail','h-betreff','h-text'].forEach(i=>document.getElementById(i).value='');
document.getElementById('hand').classList.add('aus')}
catch(e){fehler(e.message)}
finally{fortschritt(false);await ladeListe()}}

function kurzInfo(v){
const teile=[];
VERF.forEach(f=>{const w=(v[f[0]]||'').trim();
if(w&&w.toLowerCase()!=='unbekannt'&&f[0]!=='sprachen')teile.push(w)});
return teile.length?teile.join(' · '):'keine Angaben zur Verfügbarkeit'}

async function ladeListe(){
let d;try{const a=await fetch('/api/liste');d=await antwortLesen(a);
if(!a.ok||!Array.isArray(d.eintraege))return}catch(e){return}
const ziel=document.getElementById('t-inhalt');
document.getElementById('t-zahl').textContent=
d.anzahl===0?'noch keine':(d.anzahl===1?'1 Bewerbung':d.anzahl+' Bewerbungen');
if(!d.anzahl){ziel.innerHTML=
'<p class="leer">Noch nichts verarbeitet. Lad oben eine Bewerbung hoch.</p>';return}
ziel.innerHTML='';
d.eintraege.forEach(e=>{
const p=PILL[e.status]||PILL.grau;
const reihe=document.createElement('div');
reihe.className='reihe'+(offeneZeilen.has(e.id)?' offen':'');

const kopf=document.createElement('div');kopf.className='kopfzeile';
kopf.onclick=()=>{if(offeneZeilen.has(e.id))offeneZeilen.delete(e.id);
else offeneZeilen.add(e.id);
reihe.classList.toggle('offen',offeneZeilen.has(e.id))};
kopf.innerHTML='<div class="ampel" style="background:'+p[3]+'"></div>'+
'<div class="wer"><b></b><small></small></div>'+
'<div class="wert"><b></b><small>'+p[2]+'</small></div>'+
'<div class="pfeilchen">▶</div>';
kopf.querySelector('.wer b').textContent=e.name||'Unbekannt';
kopf.querySelector('.wer small').textContent=e.status==='fehler'
?'Bewertung fehlgeschlagen, zum Lesen aufklappen':kurzInfo(e.verfuegbarkeit||{});
kopf.querySelector('.wert b').textContent=e.status==='fehler'?'–':e.gesamt;
reihe.appendChild(kopf);

const det=document.createElement('div');det.className='detail';
if(e.status!=='fehler')KRIT.forEach(k=>{const v=e.scores[k[0]]||0;
const z=document.createElement('div');z.className='zeile';
z.innerHTML='<div class="lab"><span><s></s><em>'+Number(k[2])+'%</em></span><b>'+Number(v)+'</b></div>'+
'<div class="spur"><i style="width:'+(Number(v)*10)+'%;background:'+farbe(v)+'"></i></div>';
const nameKrit=z.querySelector('.lab s');nameKrit.replaceWith(document.createTextNode(k[1]));
det.appendChild(z)});

if(e.status!=='fehler'){const vt=document.createElement('p');vt.className='blocktitel';
vt.textContent='Eckdaten';det.appendChild(vt);
const ul=document.createElement('ul');ul.className='verf';
VERF.forEach(f=>{const w=((e.verfuegbarkeit||{})[f[0]]||'unbekannt').trim();
const leer=w.toLowerCase()==='unbekannt';
const li=document.createElement('li');if(leer)li.className='offen';
const b=document.createElement('b');b.textContent=f[1];
const s=document.createElement('span');s.textContent=leer?'keine Angabe':w;
li.appendChild(b);li.appendChild(s);ul.appendChild(li)});
det.appendChild(ul)}

const fz=document.createElement('p');fz.className='fazit';
fz.textContent=e.zusammenfassung||'';det.appendChild(fz);
const fs=document.createElement('p');fs.className='fuss';
fs.textContent='Führerschein: '+(e.fuehrerschein||'unbekannt')+
(e.absender?'  ·  '+e.absender:'');det.appendChild(fs);

const kn=document.createElement('div');kn.className='knoepfe';
const link=document.createElement('a');link.href='/mail/'+e.id;
link.target='_blank';link.textContent='📄 Bewerbung öffnen';kn.appendChild(link);
const del=document.createElement('button');del.textContent='🗑 Entfernen';
let sicher=false;
del.onclick=async ev=>{ev.stopPropagation();
if(!sicher){sicher=true;del.textContent='Wirklich entfernen?';
setTimeout(()=>{sicher=false;del.textContent='🗑 Entfernen'},4000);return}
try{const a=await fetch('/api/loeschen/'+e.id,{method:'POST'});
if(!a.ok){const d=await antwortLesen(a);throw new Error(d.fehler||'Fehler')}}
catch(x){fehler('Entfernen fehlgeschlagen: '+x.message);return}
offeneZeilen.delete(e.id);ladeListe()};
kn.appendChild(del);det.appendChild(kn);

reihe.appendChild(det);ziel.appendChild(reihe)})}

ladeListe();
</script></body></html>"""

def fuer_skript(wert):
    """Wandelt Python-Daten in sicheres JavaScript, auch bei Sonderzeichen."""
    return json.dumps(wert, ensure_ascii=False)[1:-1].replace("</", "<\\/")


SEITE = SEITE.replace("{{KRITLISTE}}", fuer_skript([[k, t, g] for k, t, g, _ in KRITERIEN]))
SEITE = SEITE.replace("{{VERFLISTE}}", fuer_skript([[k, t] for k, t in VERFUEG_FELDER]))
