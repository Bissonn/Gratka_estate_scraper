"""
===============================================================
  Scraper Gratka.pl → Excel Dashboard
  Autor: [Twoje imię] | Fiverr Portfolio Project

  UŻYCIE:      python main.py
  WYMAGANIA:   pip install requests beautifulsoup4 openpyxl
===============================================================
"""

import requests
from bs4 import BeautifulSoup
import re
import time
import random
import logging
from dataclasses import dataclass
from typing import Optional
from datetime import datetime

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────
#  MODEL DANYCH
# ─────────────────────────────────────────────

@dataclass
class Nieruchomosc:
    id: str
    tytul: str
    miasto: str
    dzielnica: str
    typ: str
    transakcja: str
    cena: float
    powierzchnia: float
    liczba_pokoi: int
    pietro: Optional[int]
    rok_budowy: Optional[int]
    cena_za_m2: float
    url: str
    data_dodania: str
    opis: str


# ─────────────────────────────────────────────
#  KONFIGURACJA
# ─────────────────────────────────────────────

BASE_URL = "https://gratka.pl"

BASE_SEARCH   = "https://gratka.pl/nieruchomosci/{kategoria}"
SEARCH_URL_P1 = BASE_SEARCH + "?sort=newest"
SEARCH_URL    = BASE_SEARCH + "?page={page}&sort=newest"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pl-PL,pl;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Referer": "https://gratka.pl/",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)


# ─────────────────────────────────────────────
#  PARSOWANIE LISTY OFERT
# ─────────────────────────────────────────────

def parsuj_liste(html: str) -> list[dict]:
    """
    Wyciąga linki i podstawowe dane z listy wyników Gratka.pl.
    Gratka renderuje karty jako .property-card z linkiem .property-card__link.
    """
    soup = BeautifulSoup(html, "html.parser")

    # Każda karta to kontener z linkiem .property-card__link
    linki = soup.select("a.property-card__link[href*='/nieruchomosci/']")

    oferty = []
    seen = set()

    for link in linki:
        url = link.get("href", "")
        if not url:
            continue
        if not url.startswith("http"):
            url = BASE_URL + url

        # pomiń duplikaty (Gratka czasem powtarza featured oferty)
        if url in seen:
            continue
        seen.add(url)

        # Tytuł — z tytułu karty lub tekstu linku
        tytul_el = link.select_one(
            ".property-card__title, h2, h3, [class*='title']"
        )
        tytul = tytul_el.get_text(strip=True) if tytul_el else link.get_text(strip=True)[:80]

        # Cena — szukamy wewnątrz karty lub w sąsiednim elemencie
        # .property-card__link może być wewnątrz .card__outer
        karta = link.find_parent(class_=re.compile(r"card|property"))
        cena_txt = ""
        if karta:
            cena_el = karta.select_one(
                "[class*='price'], [data-cy*='price'], [class*='Price']"
            )
            cena_txt = cena_el.get_text(strip=True) if cena_el else ""

        # Lokalizacja
        lok_txt = ""
        if karta:
            lok_el = karta.select_one(
                "[class*='location'], [class*='address'], [class*='Location']"
            )
            lok_txt = lok_el.get_text(strip=True) if lok_el else ""

        # Parametry inline (piętro, powierzchnia) widoczne na karcie
        params_txt = ""
        if karta:
            for span in karta.select("[class*='cardPropertyInfo'], [class*='param']"):
                params_txt += " " + span.get_text(strip=True)

        oferty.append({
            "url": url,
            "tytul": tytul,
            "cena_txt": cena_txt,
            "lokalizacja": lok_txt,
            "params_txt": params_txt.strip(),
        })

    return oferty


# ─────────────────────────────────────────────
#  PARSOWANIE STRONY SZCZEGÓŁÓW
# ─────────────────────────────────────────────

def parsuj_oferte(url: str, podstawowe: dict) -> Optional[Nieruchomosc]:
    """
    Pobiera podstronę oferty i wyciąga pełne dane.
    """
    try:
        r = SESSION.get(url, timeout=15)
        r.raise_for_status()
    except Exception as e:
        log.warning(f"Błąd pobierania {url}: {e}")
        return None

    soup = BeautifulSoup(r.text, "html.parser")

    # ── ID ──────────────────────────────────────────────────────
    id_match = re.search(r"/ob/(\d+)", url)
    offer_id = f"GR{id_match.group(1)}" if id_match else \
               f"GR{abs(hash(url)) % 100000:05d}"

    # ── TYTUŁ ───────────────────────────────────────────────────
    tytul_el = soup.select_one("h1")
    tytul = tytul_el.get_text(strip=True) if tytul_el else podstawowe["tytul"]

    # ── CENA ────────────────────────────────────────────────────
    # Na stronie szczegółów cena jest w .priceInfo lub [data-cy*='price']
    cena_el = soup.select_one(
        "[class*='details-price__item']"
    )
    cena_txt = cena_el.get_text(strip=True) if cena_el else podstawowe.get("cena_txt", "")
    cena = float(re.sub(r"[^\d]", "", cena_txt) or 0)

    # ── TABELA PARAMETRÓW ────────────────────────────────────────
    # Gratka używa list .parameters lub tabel z [class*='detail']
    params: dict[str, str] = {}

    # Próbuj kilka możliwych struktur
    for item in soup.select(
        "[class*='parameters__item'], [class*='parametersList'] li, "
        "[class*='details__item'], [class*='offerFeature'], "
        "li[class*='parameter']"
    ):
        label_el = item.select_one(
            "[class*='label'], [class*='name'], strong, dt, b"
        )
        value_el = item.select_one(
            "[class*='value'], [class*='content'], span:last-child, dd"
        )
        if label_el and value_el:
            k = label_el.get_text(strip=True).lower()
            v = value_el.get_text(strip=True)
            if k and v and k != v.lower():
                params[k] = v

    # fallback: szukaj par klucz-wartość w dowolnych dt/dd
    if not params:
        for dt in soup.select("dt"):
            dd = dt.find_next_sibling("dd")
            if dd:
                params[dt.get_text(strip=True).lower()] = dd.get_text(strip=True)

    # ── WYCIĄGANIE WARTOŚCI Z PARAMS ────────────────────────────

    def znajdz(slowa_kluczowe: list[str]) -> str:
        for k, v in params.items():
            if any(s in k for s in slowa_kluczowe):
                return v
        return ""

    # Powierzchnia
    powierzchnia = 0.0
    pow_txt = znajdz(["powierzchni", "metra", "m²", "mkw", "area"])
    if pow_txt:
        m = re.search(r"([\d\s,\.]+)", pow_txt)
        if m:
            powierzchnia = float(m.group(1).strip().replace(" ", "").replace(",", "."))
    if not powierzchnia:
        # fallback z tytułu lub params_txt z karty
        for tekst in [tytul, podstawowe.get("params_txt", "")]:
            m = re.search(r"([\d,\.]+)\s*m[²2²]", tekst, re.IGNORECASE)
            if m:
                powierzchnia = float(m.group(1).replace(",", "."))
                break

    # Liczba pokoi
    liczba_pokoi = 0
    pok_txt = znajdz(["poko", "pokoi", "pokoje", "rooms"])
    if pok_txt:
        m = re.search(r"(\d+)", pok_txt)
        if m:
            liczba_pokoi = int(m.group(1))
    if not liczba_pokoi:
        for tekst in [tytul, podstawowe.get("params_txt", "")]:
            m = re.search(r"(\d+)\s*pok", tekst, re.IGNORECASE)
            if m:
                liczba_pokoi = int(m.group(1))
                break

    # Piętro
    pietro = None
    piet_txt = znajdz(["piętro", "pietro", "floor"])
    if piet_txt:
        # format: "3/10" lub "3" lub "parter" 
        if "parter" in piet_txt.lower():
            pietro = 0
        else:
            m = re.search(r"(\d+)", piet_txt)
            if m:
                pietro = int(m.group(1))
    # fallback z params_txt karty: "piętro 3/3"
    if pietro is None:
        m = re.search(r"piętro\s+(\d+)", podstawowe.get("params_txt", ""), re.IGNORECASE)
        if m:
            pietro = int(m.group(1))

    # Rok budowy
    rok_budowy = None
    rok_txt = znajdz(["rok bud", "rok pow", "year", "built"])
    if rok_txt:
        m = re.search(r"(19|20)\d{2}", rok_txt)
        if m:
            rok_budowy = int(m.group(0))

    # ── LOKALIZACJA ─────────────────────────────────────────────
    # Breadcrumbs: Polska → woj. → miasto → dzielnica
    breadcrumbs = [
        b.get_text(strip=True)
        for b in soup.select(
            "nav[aria-label*='bread'] a, .breadcrumb a, "
            "[class*='breadcrumb'] a, [class*='Breadcrumb'] a"
        )
        if b.get_text(strip=True) not in ("", "Nieruchomości", "Polska", "Gratka.pl", "Ogłoszenia")
    ]

    if len(breadcrumbs) >= 2:
        miasto = breadcrumbs[-2]
        dzielnica = breadcrumbs[-1]
    elif len(breadcrumbs) == 1:
        miasto = breadcrumbs[0]
        dzielnica = breadcrumbs[0]
    else:
        lok = podstawowe.get("lokalizacja", "")
        parts = [p.strip() for p in re.split(r"[,\-/]", lok) if p.strip()]
        miasto = parts[0] if parts else "Nieznane"
        dzielnica = parts[1] if len(parts) > 1 else miasto

    # ── CENA / M² ────────────────────────────────────────────────
    cena_za_m2 = 0.0
    cm2_txt = znajdz(["cena za m", "cena/m", "price/m"])
    if cm2_txt:
        m = re.search(r"([\d\s]+)", cm2_txt)
        if m:
            cena_za_m2 = float(re.sub(r"\s", "", m.group(1)) or 0)
    if not cena_za_m2 and cena > 0 and powierzchnia > 0:
        cena_za_m2 = round(cena / powierzchnia)

    # ── OPIS ────────────────────────────────────────────────────
    opis_el = soup.select_one(
        ".description__content, [class*='description__text'], "
        "[class*='offerDescription'], [data-cy='adPageDescription']"
    )
    opis = opis_el.get_text(strip=True)[:300] if opis_el else ""

    # ── DATA DODANIA ─────────────────────────────────────────────
    # Na karcie widzieliśmy: "Dodane: 2026.03.27"
    data_dodania = datetime.now().strftime("%Y-%m-%d")
    data_el = soup.select_one("time[datetime], [data-cy*='date'], [class*='date']")
    if data_el and data_el.get("datetime"):
        data_dodania = data_el["datetime"][:10]
    elif data_el:
        m = re.search(r"(\d{4}[\.\-]\d{2}[\.\-]\d{2})", data_el.get_text())
        if m:
            data_dodania = m.group(1).replace(".", "-")
    else:
        # szukaj tekstu "Dodane: YYYY.MM.DD" w całej stronie
        m = re.search(r"Dodane[:\s]+(\d{4})\.(\d{2})\.(\d{2})", soup.get_text())
        if m:
            data_dodania = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

    # ── TYP TRANSAKCJI I NIERUCHOMOŚCI ───────────────────────────
    transakcja = "Wynajem" if any(
        x in url.lower() for x in ("wynajem", "najem", "rent")
    ) else "Sprzedaż"

    if "dom" in url.lower() or "dom" in tytul.lower():
        typ = "Dom"
    elif liczba_pokoi == 1 or "kawalerka" in url.lower():
        typ = "Kawalerka"
    else:
        typ = "Mieszkanie"

    return Nieruchomosc(
        id=offer_id,
        tytul=tytul[:80],
        miasto=miasto,
        dzielnica=dzielnica,
        typ=typ,
        transakcja=transakcja,
        cena=cena,
        powierzchnia=powierzchnia,
        liczba_pokoi=max(liczba_pokoi, 1),
        pietro=pietro,
        rok_budowy=rok_budowy,
        cena_za_m2=cena_za_m2,
        url=url,
        data_dodania=data_dodania,
        opis=opis,
    )


# ─────────────────────────────────────────────
#  GŁÓWNA FUNKCJA
# ─────────────────────────────────────────────

def zbuduj_url(miasto: str, typ_nieruchomosci: str, transakcja: str) -> tuple[str, str]:
    """
    Buduje bazowy URL wyszukiwania Gratka.pl na podstawie filtrów.

    Gratka używa struktury:
      /nieruchomosci/{typ}                       — sprzedaż (brak słowa)
      /nieruchomosci/{typ}/wynajem               — wynajem
    Miasto jest parametrem ?lokalizacja= lub fragmentem ścieżki.

    Zwraca (url_strona1, url_kolejne_strony_z_{page}).
    """
    # Mapowanie typ nieruchomości → segment URL Gratka
    TYPY = {
        "Mieszkanie":  "mieszkania",
        "Dom":         "domy",
        "Kawalerka":   "kawalerki",
        "Działka":     "dzialki",
        "Lokal":       "lokale-uzytkowe",
    }
    segment = TYPY.get(typ_nieruchomosci, "mieszkania")

    if transakcja == "Wynajem":
        sciezka = f"nieruchomosci/{segment}/wynajem"
    else:
        sciezka = f"nieruchomosci/{segment}"

    # Miasto jako parametr zapytania
    params = ""
    if miasto.strip():
        # Gratka przyjmuje lokalizację jako ?lokalizacja=warszawa
        slug = miasto.strip().lower().replace(" ", "-")
        params = f"/{slug}"

    base = f"https://gratka.pl/{sciezka}"
    sep  = "?"

    url_p1 = f"{base}{params}"
    url_pn = f"{base}{params}{sep}page={{page}}"
    return url_p1, url_pn


def pobierz_oferty(
    liczba: int = 40,
    max_stron: int = 10,
    miasto: str = "",
    typ_nieruchomosci: str = "Mieszkanie",
    transakcja: str = "Sprzedaż",
    callback_postep=None,       # opcjonalna funkcja(tekst) do logowania w GUI
) -> tuple[list[Nieruchomosc], bool]:
    """
    Scrapuje Gratka.pl i zwraca (lista_ofert, czy_osiagnieto_limit).

    Parametry:
        liczba             — docelowa liczba ofert
        max_stron          — ile stron wyników przeglądnąć
        miasto             — np. "Warszawa", "Kraków" (puste = cała Polska)
        typ_nieruchomosci  — "Mieszkanie" / "Dom" / "Kawalerka" / "Działka" / "Lokal"
        transakcja         — "Sprzedaż" / "Wynajem"
        callback_postep    — funkcja(str) wywoływana przy każdym kroku (dla GUI)

    Zwraca:
        (wyniki, osiagnieto_limit)
        osiagnieto_limit=False oznacza że Gratka miała mniej ofert niż żądano
    """
    def log_gui(msg: str):
        log.info(msg)
        if callback_postep:
            callback_postep(msg)

    url_p1, url_pn = zbuduj_url(miasto, typ_nieruchomosci, transakcja)
    log_gui(f"Szukam: {typ_nieruchomosci} | {transakcja} | {miasto or 'cała Polska'}")
    log_gui(f"URL: {url_p1}")

    linki: list[dict] = []
    wyczerpano_strony = False

    for page in range(1, max_stron + 1):
        if len(linki) >= liczba:
            break

        url = url_p1 if page == 1 else url_pn.format(page=page)
        log_gui(f"Strona {page}...")

        try:
            r = SESSION.get(url, timeout=15)
            r.raise_for_status()
        except Exception as e:
            log_gui(f"Błąd strony {page}: {e}")
            break

        nowe = parsuj_liste(r.text)
        if not nowe:
            log_gui("Nie ma więcej ofert - koniec wyników Gratka.")
            wyczerpano_strony = True
            break

        linki.extend(nowe)
        log_gui(f"  → {len(nowe)} linków (łącznie: {len(linki)})")
        time.sleep(random.uniform(1.5, 3.0))

    osiagnieto_limit = len(linki) >= liczba
    linki = linki[:liczba]

    wyniki: list[Nieruchomosc] = []
    for i, pods in enumerate(linki, 1):
        log_gui(f"Pobieram ofertę {i}/{len(linki)}...")
        oferta = parsuj_oferte(pods["url"], pods)
        if oferta and oferta.cena > 0:
            wyniki.append(oferta)
        else:
            log_gui(f"Nie podano ceny oferty {i}/{len(linki)}. Nie będzie wliczona do średniej.")
        time.sleep(random.uniform(1.0, 2.5))

    log_gui(f"Gotowe - pobrano {len(wyniki)} ofert.")
    return wyniki, osiagnieto_limit
