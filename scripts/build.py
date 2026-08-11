#!/usr/bin/env python3
"""
Liest die Lego-Excel-Liste, reichert jedes Set einmalig über die Rebrickable-API an
(mit persistentem JSON-Cache, damit bereits bekannte Sets nicht erneut abgefragt werden),
und generiert daraus eine statische Homepage (docs/index.html + docs/stats.html)
für GitHub Pages.

Aufruf:
    python scripts/build.py

Benötigt die Umgebungsvariable REBRICKABLE_API_KEY.
"""

import html
import json
import os
import sys
import time
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parent.parent
DATA_XLSX = ROOT / "data" / "Lego_Liste_bereinigt.xlsx"
CACHE_FILE = ROOT / "data" / "api_cache.json"
OUTPUT_DIR = ROOT / "docs"
OUTPUT_HTML = OUTPUT_DIR / "index.html"
STATS_HTML = OUTPUT_DIR / "stats.html"

API_KEY = os.environ.get("REBRICKABLE_API_KEY")
API_BASE = "https://rebrickable.com/api/v3/lego"

REQUIRED_COLUMNS = ["Set Number", "Zustand"]
OPTIONAL_COLUMNS = [
    "Identifier", "Name", "ID", "THEME", "Anzahl",
    "Bauanleitung", "Original-Karton", "Verkaufsstatus", "Preis",
]
THEME_CACHE_PREFIX = "theme:"

STATUS_LABELS = {
    "unverkäuflich": "Unverkäuflich",
    "verfügbar": "Verfügbar",
    "reserviert": "Reserviert",
    "verkauft": "Verkauft",
}
STATUS_CLASS = {
    "unverkäuflich": "status-unsellable",
    "verfügbar": "status-available",
    "reserviert": "status-reserved",
    "verkauft": "status-sold",
}

SHARED_CSS = """
  :root {
    --bg: #FAFAF7; --surface: #FFFFFF; --ink: #1C1C1E; --ink-muted: #6B6B68;
    --line: #E7E4DC; --brick: #D6001C; --brick-dark: #A1000F;
    --green: #2F7D46; --amber: #B9781A; --sold: #8C8C86; --unsellable: #B6B3AC;
    --font-display: 'Space Grotesk', sans-serif;
    --font-body: 'Inter', -apple-system, sans-serif;
    --font-mono: 'JetBrains Mono', monospace;
  }
  * { box-sizing: border-box; }
  html { scroll-behavior: smooth; }
  body {
    font-family: var(--font-body); background: var(--bg); color: var(--ink);
    margin: 0; padding: 0; -webkit-font-smoothing: antialiased;
  }
  a { color: inherit; text-decoration: none; }

  .site-header { position: sticky; top: 0; z-index: 50; background: var(--surface); border-bottom: 1px solid var(--line); }
  .header-inner { max-width: 1200px; margin: 0 auto; padding: 0 24px; display: flex; align-items: center; justify-content: space-between; height: 64px; }
  .brand { display: flex; align-items: center; gap: 10px; font-family: var(--font-display); font-weight: 700; font-size: 19px; }
  .brand .studs { display: inline-flex; gap: 3px; }
  .brand .studs span { width: 7px; height: 7px; border-radius: 50%; background: var(--brick); display: block; }
  .brand .studs span:nth-child(2) { background: var(--ink); opacity: .8; }
  .brand .studs span:nth-child(3) { background: var(--ink); opacity: .5; }

  nav.desktop-nav { display: flex; gap: 28px; font-size: 14px; font-weight: 500; }
  nav.desktop-nav a { color: var(--ink-muted); transition: color .15s; }
  nav.desktop-nav a:hover, nav.desktop-nav a.active { color: var(--brick); }

  .hamburger { display: none; flex-direction: column; justify-content: center; gap: 4px; width: 36px; height: 36px; border: none; background: none; cursor: pointer; padding: 0; }
  .hamburger span { display: block; width: 22px; height: 2px; background: var(--ink); transition: all .2s; margin: 0 auto; }
  .hamburger.open span:nth-child(1) { transform: translateY(6px) rotate(45deg); }
  .hamburger.open span:nth-child(2) { opacity: 0; }
  .hamburger.open span:nth-child(3) { transform: translateY(-6px) rotate(-45deg); }

  .mobile-nav { display: none; flex-direction: column; background: var(--surface); border-bottom: 1px solid var(--line); overflow: hidden; max-height: 0; transition: max-height .25s ease; }
  .mobile-nav.open { max-height: 280px; }
  .mobile-nav a { padding: 14px 24px; font-size: 15px; font-weight: 500; border-top: 1px solid var(--line); }

  footer { border-top: 1px solid var(--line); padding: 28px 24px; text-align: center; font-size: 12px; color: var(--ink-muted); }

  @media (max-width: 720px) {
    nav.desktop-nav { display: none; }
    .hamburger { display: flex; }
    .mobile-nav { display: flex; }
  }
"""

NAV_LINKS = [
    ("index.html#top", "Sammlung", "sammlung"),
    ("index.html#top", "Zum Verkauf", "verkauf"),
    ("stats.html", "Statistiken", "stats"),
    ("index.html#footer", "Kontakt", "kontakt"),
]


def render_header(active: str) -> str:
    def link(href, label, key, extra_attrs=""):
        cls = "active" if key == active else ""
        return f'<a href="{href}" class="{cls}" {extra_attrs}>{label}</a>'

    desktop_links = []
    mobile_links = []
    for href, label, key in NAV_LINKS:
        attrs = ""
        mobile_attrs = ""
        if key == "verkauf":
            attrs = 'onclick="if(window.__setStatusFilter){window.__setStatusFilter(\'forsale\'); return false;}"'
            mobile_attrs = ('id="mobileForSale" onclick="if(window.__setStatusFilter){'
                             "window.__setStatusFilter('forsale'); "
                             "document.getElementById('mobileNav').classList.remove('open'); "
                             "document.getElementById('hamburger').classList.remove('open'); "
                             "return false;}\"")
        desktop_links.append(link(href, label, key, attrs))
        mobile_links.append(link(href, label, key, mobile_attrs))

    return f"""
<header class="site-header">
  <div class="header-inner">
    <a class="brand" href="index.html#top">
      <span class="studs"><span></span><span></span><span></span></span>
      Brick Archive
    </a>
    <nav class="desktop-nav">
      {''.join(desktop_links)}
    </nav>
    <button class="hamburger" id="hamburger" aria-label="Menü öffnen" aria-expanded="false">
      <span></span><span></span><span></span>
    </button>
  </div>
  <nav class="mobile-nav" id="mobileNav">
    {''.join(mobile_links)}
  </nav>
</header>
<script>
  (function() {{
    const hamburger = document.getElementById('hamburger');
    const mobileNav = document.getElementById('mobileNav');
    hamburger.addEventListener('click', () => {{
      const isOpen = mobileNav.classList.toggle('open');
      hamburger.classList.toggle('open', isOpen);
      hamburger.setAttribute('aria-expanded', isOpen);
    }});
  }})();
</script>
"""


def render_footer() -> str:
    return """
<footer id="footer">
  Brick Archive — privat gepflegte Sammlung. Set-Daten via Rebrickable-API.
</footer>
"""


def load_cache() -> dict:
    if CACHE_FILE.exists():
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache: dict) -> None:
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2, ensure_ascii=False, sort_keys=True)


def fetch_set_data(set_num: str) -> dict | None:
    """Ruft die Rebrickable-API für eine echte Set-Nummer ab (z.B. '31120-1')."""
    if not API_KEY:
        print(f"  WARNUNG: kein API-Key gesetzt, überspringe {set_num}", file=sys.stderr)
        return None
    url = f"{API_BASE}/sets/{set_num}/"
    headers = {"Authorization": f"key {API_KEY}"}

    for attempt in range(4):
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            if resp.status_code == 200:
                data = resp.json()
                return {
                    "name": data.get("name"),
                    "year": data.get("year"),
                    "num_parts": data.get("num_parts"),
                    "set_img_url": data.get("set_img_url"),
                    "set_url": data.get("set_url"),
                    "theme_id": data.get("theme_id"),
                    "fetched_at": int(time.time()),
                }
            elif resp.status_code == 404:
                print(f"  Set {set_num} nicht bei Rebrickable gefunden (404)", file=sys.stderr)
                return {"error": "not_found", "fetched_at": int(time.time())}
            elif resp.status_code == 429:
                wait = 5 * (attempt + 1)
                print(f"  Rate-Limit erreicht bei {set_num}, warte {wait}s (Versuch {attempt + 1}/4)...", file=sys.stderr)
                time.sleep(wait)
                continue
            else:
                print(f"  Fehler {resp.status_code} bei Set {set_num}", file=sys.stderr)
                return None
        except requests.RequestException as e:
            print(f"  Netzwerkfehler bei Set {set_num}: {e}", file=sys.stderr)
            return None

    print(f"  Set {set_num}: nach mehreren Versuchen weiterhin Rate-Limit, überspringe für diesen Lauf.", file=sys.stderr)
    return None


def fetch_theme_name(theme_id, cache: dict) -> str | None:
    """Löst eine Theme-ID über die Rebrickable-API in einen Klartextnamen auf (gecacht, inkl. parent_id)."""
    if theme_id is None or pd.isna(theme_id):
        return None
    key = f"{THEME_CACHE_PREFIX}{int(theme_id)}"
    if key in cache and "error" not in cache[key]:
        return cache[key].get("name")
    if not API_KEY:
        return None
    url = f"{API_BASE}/themes/{int(theme_id)}/"
    headers = {"Authorization": f"key {API_KEY}"}

    for attempt in range(4):
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            if resp.status_code == 200:
                data = resp.json()
                cache[key] = {
                    "name": data.get("name"),
                    "parent_id": data.get("parent_id"),
                    "fetched_at": int(time.time()),
                }
                time.sleep(1.1)
                return cache[key]["name"]
            elif resp.status_code == 429:
                wait = 5 * (attempt + 1)
                print(f"  Rate-Limit erreicht bei Theme {theme_id}, warte {wait}s (Versuch {attempt + 1}/4)...", file=sys.stderr)
                time.sleep(wait)
                continue
            else:
                cache[key] = {"error": f"http_{resp.status_code}", "fetched_at": int(time.time())}
                time.sleep(1.1)
                return None
        except requests.RequestException:
            return None

    print(f"  Theme {theme_id}: nach mehreren Versuchen weiterhin Rate-Limit, überspringe für diesen Lauf.", file=sys.stderr)
    return None


def theme_breadcrumb(theme_id, cache: dict) -> tuple[str, str]:
    """
    Baut aus der Theme-Hierarchie (parent_id-Kette) einen Breadcrumb-Pfad, z.B.
    "Star Wars > Ultimate Collector Series". Gibt (leaf_name, full_path) zurück.
    Sets, die Rebrickable direkt unter einem Ober-Theme einordnet (z.B. sehr neue
    Sets ohne eigenes Subtheme), liefern hier logischerweise nur die eine Ebene.
    """
    if theme_id is None or pd.isna(theme_id):
        return ("Unbekannt", "Unbekannt")

    chain = []
    current_id = int(theme_id)
    seen = set()
    for _ in range(6):  # Sicherheitslimit gegen Zirkelbezüge
        if current_id in seen:
            break
        seen.add(current_id)
        key = f"{THEME_CACHE_PREFIX}{current_id}"
        entry = cache.get(key)
        if not entry or "error" in entry:
            break
        chain.append(entry.get("name") or "?")
        parent_id = entry.get("parent_id")
        if parent_id is None or pd.isna(parent_id):
            break
        current_id = int(parent_id)

    if not chain:
        return ("Unbekannt", "Unbekannt")

    chain.reverse()  # von oberster Ebene zur spezifischsten
    leaf = chain[-1]
    full_path = " > ".join(chain)
    return (leaf, full_path)


def assign_missing_identifiers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Vergibt automatisch einen Identifier ({Set Number}-N) für Zeilen, bei denen
    die Identifier-Spalte leer gelassen wurde. So reicht beim Hinzufügen eines
    neuen Sets die reine Set-Nummer.
    """
    df = df.copy()
    if "Identifier" not in df.columns:
        df["Identifier"] = pd.NA

    used_per_set = {}
    for idx, row in df.iterrows():
        ident = row.get("Identifier")
        set_num = str(row["Set Number"])
        if pd.notna(ident) and str(ident).strip():
            used_per_set.setdefault(set_num, set()).add(str(ident).strip())

    for idx, row in df.iterrows():
        ident = row.get("Identifier")
        if pd.notna(ident) and str(ident).strip():
            continue
        set_num = str(row["Set Number"])
        taken = used_per_set.setdefault(set_num, set())
        n = 1
        while f"{set_num}-{n}" in taken:
            n += 1
        new_id = f"{set_num}-{n}"
        taken.add(new_id)
        df.at[idx, "Identifier"] = new_id

    return df


def real_set_number(row) -> str:
    """
    Ermittelt die tatsächliche Rebrickable-Set-Nummer.
    Mehrfach-Exemplare (z.B. Identifier '31120-2', '31120-3') teilen sich alle
    dieselben Katalog-Daten der Basisnummer '{Set Number}-1'.
    """
    return f"{row['Set Number']}-1"


def enrich_data(df: pd.DataFrame, cache: dict) -> dict:
    """Holt für jede eindeutige Set-Nummer die API-Daten (aus Cache oder live), inkl. Theme-Namen."""
    unique_set_nums = sorted(set(real_set_number(row) for _, row in df.iterrows()))
    new_calls = 0

    for set_num in unique_set_nums:
        if set_num in cache and "error" not in cache[set_num]:
            continue
        print(f"Rufe API ab für {set_num} ...")
        result = fetch_set_data(set_num)
        if result is not None:
            cache[set_num] = result
            new_calls += 1
        time.sleep(1.1)  # Rebrickable erlaubt im Schnitt ~1 Anfrage/Sekunde

    # Theme-Namen für alle vorkommenden theme_ids UND deren komplette Parent-Kette auflösen
    theme_ids = set()
    for set_num in unique_set_nums:
        tid = cache.get(set_num, {}).get("theme_id")
        if tid is not None:
            theme_ids.add(int(tid))

    resolved = set()
    to_resolve = list(theme_ids)
    while to_resolve:
        tid = to_resolve.pop()
        if tid in resolved:
            continue
        resolved.add(tid)
        fetch_theme_name(tid, cache)
        entry = cache.get(f"{THEME_CACHE_PREFIX}{tid}", {})
        parent_id = entry.get("parent_id")
        if parent_id is not None and not pd.isna(parent_id) and int(parent_id) not in resolved:
            to_resolve.append(int(parent_id))

    print(f"{new_calls} neue API-Aufrufe, {len(unique_set_nums) - new_calls} aus Cache.")
    return cache


def normalize_status(raw) -> str:
    s = str(raw).strip() if raw is not None and pd.notna(raw) else ""
    return s if s in STATUS_LABELS else "unverkäuflich"


def build_card_record(row, cache: dict) -> dict:
    set_num = real_set_number(row)
    api = cache.get(set_num, {})
    status = normalize_status(row.get("Verkaufsstatus"))
    preis = row.get("Preis", "")
    has_price = preis not in ("", None) and pd.notna(preis)

    # Name/Theme: primär von der API, Excel nur als Fallback (z.B. falls API-Aufruf fehlschlug)
    api_name = api.get("name")
    excel_name = row.get("Name")
    name = api_name or (str(excel_name) if pd.notna(excel_name) else f"Set {row['Set Number']}")

    theme_id = api.get("theme_id")
    if theme_id is not None:
        theme_leaf, theme_path = theme_breadcrumb(theme_id, cache)
    else:
        theme_leaf, theme_path = (None, None)
    excel_theme = row.get("THEME")
    excel_theme_str = str(excel_theme) if pd.notna(excel_theme) else None
    theme = theme_leaf or excel_theme_str or "Unbekannt"
    theme_display = theme_path or excel_theme_str or "Unbekannt"

    return {
        "identifier": str(row["Identifier"]),
        "set_number": str(row["Set Number"]),
        "name": name,
        "theme": theme,
        "theme_display": theme_display,
        "year": api.get("year", ""),
        "num_parts": api.get("num_parts", ""),
        "img": api.get("set_img_url", ""),
        "set_url": api.get("set_url", ""),
        "zustand_neu": str(row["Zustand"]).strip() == "J",
        "bauanleitung": str(row.get("Bauanleitung", "")).strip() == "J",
        "originalkarton": str(row.get("Original-Karton", "")).strip() == "J",
        "status": status,
        "status_label": STATUS_LABELS[status],
        "price": str(preis) if has_price else "",
    }


def render_card(rec: dict) -> str:
    status_class = STATUS_CLASS[rec["status"]]
    zustand_label = "Neu / ungeöffnet" if rec["zustand_neu"] else "Gebraucht / geöffnet"

    if rec["status"] == "unverkäuflich":
        sale_html = '<div class="unsellable-label">Unverkäuflich</div>'
    else:
        price_html = (f'<div class="price">CHF {rec["price"]}</div>'
                       if rec["price"] else '<div class="price price-muted">Preis auf Anfrage</div>')
        buy_html = ""
        if rec["status"] == "verfügbar":
            buy_html = ('<button class="buy-btn" onclick="event.stopPropagation(); '
                        "alert('Kauf-Workflow folgt in einem späteren Schritt.')\">"
                        "Zum Verkauf anfragen</button>")
        sale_html = price_html + buy_html

    badges = []
    if rec["bauanleitung"]:
        badges.append('<span class="tag"><i class="dot"></i>Bauanleitung</span>')
    if rec["originalkarton"]:
        badges.append('<span class="tag"><i class="dot"></i>Originalkarton</span>')
    badges_html = f'<div class="tags">{"".join(badges)}</div>' if badges else ""

    data_json = html.escape(json.dumps(rec, ensure_ascii=False), quote=True)
    forsale_flag = "0" if rec["status"] == "unverkäuflich" else "1"
    zustand_flag = "neu" if rec["zustand_neu"] else "gebraucht"

    return f"""
        <article class="card" tabindex="0" role="button" aria-label="Details zu {html.escape(rec['name'])}"
                 data-theme="{html.escape(rec['theme'])}" data-status="{rec['status']}" data-forsale="{forsale_flag}"
                 data-zustand="{zustand_flag}"
                 data-name="{html.escape(rec['name'].lower())}" data-set='{data_json}'>
          <div class="img-wrap">
            <img src="{rec['img']}" alt="{html.escape(rec['name'])}" loading="lazy" onerror="this.parentElement.classList.add('img-fallback')">
            <span class="badge {status_class}">{rec['status_label']}</span>
          </div>
          <div class="card-body">
            <div class="card-top">
              <span class="set-num">#{rec['set_number']}</span>
              {f'<span class="set-num">{rec["year"]}</span>' if rec['year'] else ''}
            </div>
            <h3>{html.escape(rec['name'])}</h3>
            <div class="theme-tag">{html.escape(rec['theme_display'])}</div>
            <div class="specs">
              <span>{zustand_label}</span>
              {f'<span>{rec["num_parts"]} Teile</span>' if rec['num_parts'] else ''}
            </div>
            {badges_html}
            {sale_html}
          </div>
        </article>"""


CARD_CSS = """
  .hero { max-width: 1200px; margin: 0 auto; padding: 40px 24px 8px; }
  .hero h1 { font-family: var(--font-display); font-size: clamp(26px, 4vw, 36px); font-weight: 700; margin: 0 0 8px; letter-spacing: -0.01em; }
  .hero p { color: var(--ink-muted); font-size: 15px; margin: 0; }
  .hero .stats-link { display: inline-block; margin-top: 14px; font-size: 13px; color: var(--brick); font-weight: 600; }

  .filter-bar { position: sticky; top: 64px; z-index: 40; background: var(--bg); border-bottom: 1px solid var(--line); }
  .filter-inner { max-width: 1200px; margin: 0 auto; padding: 14px 24px; display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
  .filter-inner input, .filter-inner select { font-family: var(--font-body); padding: 9px 12px; border: 1px solid var(--line); border-radius: 8px; font-size: 14px; background: var(--surface); color: var(--ink); }
  .filter-inner input { flex: 1; min-width: 180px; }
  .filter-inner input:focus, .filter-inner select:focus { outline: 2px solid var(--brick); outline-offset: 1px; }
  .stats-line { font-size: 12px; color: var(--ink-muted); font-family: var(--font-mono); padding: 10px 24px 0; max-width: 1200px; margin: 0 auto; }

  .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(230px, 1fr)); gap: 18px; max-width: 1200px; margin: 0 auto; padding: 16px 24px 64px; }
  .card { background: var(--surface); border: 1px solid var(--line); border-radius: 12px; overflow: hidden; display: flex; flex-direction: column; cursor: pointer; transition: border-color .15s, box-shadow .15s, transform .15s; }
  .card:hover { border-color: var(--ink); box-shadow: 0 6px 20px rgba(0,0,0,0.06); transform: translateY(-2px); }
  .card:focus-visible { outline: 2px solid var(--brick); outline-offset: 2px; }
  .img-wrap { position: relative; background: #F1F0EA; aspect-ratio: 4/3; display: flex; align-items: center; justify-content: center; }
  .img-wrap img { max-width: 82%; max-height: 82%; object-fit: contain; }
  .img-wrap.img-fallback::before { content: ''; position: absolute; inset: 0; background-image: radial-gradient(var(--line) 1.5px, transparent 1.5px); background-size: 14px 14px; }
  .badge { position: absolute; top: 10px; right: 10px; font-size: 10px; font-weight: 600; padding: 4px 9px; border-radius: 20px; color: white; text-transform: uppercase; letter-spacing: .03em; }
  .status-available { background: var(--green); }
  .status-reserved { background: var(--amber); }
  .status-sold { background: var(--sold); }
  .status-unsellable { background: var(--unsellable); color: #4a473f; }

  .card-body { padding: 14px 16px 16px; display: flex; flex-direction: column; gap: 6px; }
  .card-top { display: flex; justify-content: space-between; font-family: var(--font-mono); font-size: 11px; color: var(--ink-muted); }
  .card-body h3 { font-family: var(--font-display); margin: 2px 0 0; font-size: 15px; font-weight: 500; line-height: 1.3; }
  .theme-tag { font-size: 11px; color: var(--ink-muted); }
  .specs { display: flex; gap: 10px; font-size: 12px; color: var(--ink-muted); margin-top: 2px; }
  .tags { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 2px; }
  .tag { display: flex; align-items: center; gap: 5px; font-size: 11px; color: var(--ink-muted); }
  .tag .dot { width: 5px; height: 5px; border-radius: 50%; background: var(--green); display: inline-block; }
  .price { font-family: var(--font-mono); font-size: 17px; font-weight: 500; margin-top: 6px; }
  .price-muted { font-size: 13px; color: var(--ink-muted); font-weight: 400; }
  .unsellable-label { font-size: 12px; color: var(--ink-muted); margin-top: 6px; font-style: italic; }
  .buy-btn { margin-top: 6px; width: 100%; padding: 10px; border: none; border-radius: 8px; background: var(--brick); color: white; font-family: var(--font-body); font-weight: 600; font-size: 13px; cursor: pointer; transition: background .15s; }
  .buy-btn:hover { background: var(--brick-dark); }

  .empty-state { text-align: center; padding: 64px 24px; color: var(--ink-muted); display: none; }

  .modal-overlay { position: fixed; inset: 0; background: rgba(20,18,15,0.55); display: none; align-items: center; justify-content: center; z-index: 100; padding: 20px; }
  .modal-overlay.open { display: flex; }
  .modal-box { background: var(--surface); border-radius: 16px; max-width: 640px; width: 100%; max-height: 88vh; overflow-y: auto; position: relative; }
  .modal-close { position: absolute; top: 14px; right: 14px; width: 34px; height: 34px; border-radius: 50%; border: none; background: var(--bg); font-size: 18px; cursor: pointer; z-index: 2; }
  .modal-img { background: #F1F0EA; aspect-ratio: 16/10; display: flex; align-items: center; justify-content: center; border-radius: 16px 16px 0 0; }
  .modal-img img { max-width: 70%; max-height: 70%; object-fit: contain; }
  .modal-body { padding: 24px 28px 28px; }
  .modal-body .badge { position: static; display: inline-block; margin-bottom: 10px; }
  .modal-body h2 { font-family: var(--font-display); font-size: 22px; margin: 0 0 4px; }
  .modal-body .modal-theme { color: var(--ink-muted); font-size: 13px; margin-bottom: 16px; }
  .modal-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px 20px; margin-bottom: 16px; }
  .modal-grid div { font-size: 13px; }
  .modal-grid span { display: block; color: var(--ink-muted); font-size: 11px; text-transform: uppercase; letter-spacing: .03em; margin-bottom: 2px; }
  .modal-price { font-family: var(--font-mono); font-size: 24px; font-weight: 500; margin: 12px 0; }
  .modal-body .buy-btn { width: 100%; padding: 13px; font-size: 14px; }
  .modal-body .external-link { display: inline-block; margin-top: 14px; font-size: 13px; color: var(--ink-muted); }
  .modal-body .external-link:hover { color: var(--brick); }

  @media (max-width: 720px) {
    .filter-bar { top: 64px; }
    .grid { grid-template-columns: repeat(auto-fill, minmax(150px, 1fr)); gap: 12px; padding: 12px 16px 48px; }
    .hero { padding: 28px 16px 4px; }
    .filter-inner, .stats-line { padding-left: 16px; padding-right: 16px; }
  }
"""

CARD_SCRIPT = """
  const search = document.getElementById('search');
  const themeFilter = document.getElementById('themeFilter');
  const statusFilter = document.getElementById('statusFilter');
  const zustandFilter = document.getElementById('zustandFilter');
  const cards = Array.from(document.querySelectorAll('.card'));
  const stats = document.getElementById('stats');
  const emptyState = document.getElementById('emptyState');

  function applyFilters() {
    const q = search.value.toLowerCase();
    const theme = themeFilter.value;
    const status = statusFilter.value;
    const zustand = zustandFilter.value;
    let visible = 0;
    cards.forEach(c => {
      let statusMatch = true;
      if (status === 'forsale') statusMatch = c.dataset.forsale === '1';
      else if (status) statusMatch = c.dataset.status === status;
      const show = c.dataset.name.includes(q)
        && (!theme || c.dataset.theme === theme)
        && (!zustand || c.dataset.zustand === zustand)
        && statusMatch;
      c.style.display = show ? '' : 'none';
      if (show) visible++;
    });
    stats.textContent = visible + ' von ' + cards.length + ' Sets angezeigt';
    emptyState.style.display = visible === 0 ? 'block' : 'none';
  }

  window.__setStatusFilter = function(value) {
    statusFilter.value = value;
    applyFilters();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  search.addEventListener('input', applyFilters);
  themeFilter.addEventListener('change', applyFilters);
  statusFilter.addEventListener('change', applyFilters);
  zustandFilter.addEventListener('change', applyFilters);
  applyFilters();

  const overlay = document.getElementById('modalOverlay');
  const modalBox = document.getElementById('modalBox');

  function statusClass(status) {
    return { 'verfügbar': 'status-available', 'reserviert': 'status-reserved', 'verkauft': 'status-sold', 'unverkäuflich': 'status-unsellable' }[status] || 'status-unsellable';
  }

  function openModalWithData(rec) {
    let saleHtml;
    if (rec.status === 'unverkäuflich') {
      saleHtml = '<div class="unsellable-label" style="font-size:14px;">Dieses Set steht nicht zum Verkauf.</div>';
    } else {
      const priceHtml = rec.price
        ? '<div class="modal-price">CHF ' + rec.price + '</div>'
        : '<div class="modal-price price-muted" style="font-size:15px;">Preis auf Anfrage</div>';
      const buyHtml = rec.status === 'verfügbar'
        ? '<button class="buy-btn" onclick="alert(\\'Kauf-Workflow folgt in einem späteren Schritt.\\')">Zum Verkauf anfragen</button>'
        : '';
      saleHtml = priceHtml + buyHtml;
    }

    modalBox.innerHTML = `
      <button class="modal-close" onclick="closeModal()" aria-label="Schließen">✕</button>
      <div class="modal-img">${rec.img ? '<img src="' + rec.img + '" alt="' + rec.name + '">' : ''}</div>
      <div class="modal-body">
        <span class="badge ${statusClass(rec.status)}">${rec.status_label}</span>
        <h2>${rec.name}</h2>
        <div class="modal-theme">#${rec.set_number} — ${rec.theme_display}</div>
        <div class="modal-grid">
          <div><span>Zustand</span>${rec.zustand_neu ? 'Neu / ungeöffnet' : 'Gebraucht / geöffnet'}</div>
          <div><span>Erscheinungsjahr</span>${rec.year || '—'}</div>
          <div><span>Teile</span>${rec.num_parts || '—'}</div>
          <div><span>Identifier</span>${rec.identifier}</div>
          <div><span>Bauanleitung</span>${rec.bauanleitung ? 'Vorhanden' : 'Nicht vorhanden'}</div>
          <div><span>Originalkarton</span>${rec.originalkarton ? 'Vorhanden' : 'Nicht vorhanden'}</div>
        </div>
        ${saleHtml}
        ${rec.set_url ? '<a class="external-link" href="' + rec.set_url + '" target="_blank" rel="noopener">Details bei Rebrickable →</a>' : ''}
      </div>
    `;
    overlay.classList.add('open');
    document.body.style.overflow = 'hidden';
  }

  function closeModal() {
    overlay.classList.remove('open');
    document.body.style.overflow = '';
  }
  window.closeModal = closeModal;

  cards.forEach(c => {
    const open = () => openModalWithData(JSON.parse(c.dataset.set));
    c.addEventListener('click', open);
    c.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); open(); } });
  });
  overlay.addEventListener('click', (e) => { if (e.target === overlay) closeModal(); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeModal(); });
"""


def build_html(df: pd.DataFrame, cache: dict) -> str:
    themes = sorted(df["THEME"].dropna().unique())
    theme_options = "".join(f'<option value="{html.escape(t)}">{html.escape(t)}</option>' for t in themes)

    records = [build_card_record(row, cache) for _, row in df.iterrows()]
    cards_html = "\n".join(render_card(r) for r in records)

    return f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Brick Archive — Meine Lego-Sammlung</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
{SHARED_CSS}
{CARD_CSS}
</style>
</head>
<body>
{render_header('sammlung')}

<div id="top"></div>
<section class="hero">
  <h1>Meine Lego-Sammlung</h1>
  <p>Eine kuratierte Übersicht aller Sets — durchsuchbar, filterbar, teils käuflich.</p>
  <a class="stats-link" href="stats.html">Alle Zahlen zur Sammlung ansehen →</a>
</section>

<div class="filter-bar">
  <div class="filter-inner">
    <input type="text" id="search" placeholder="Set suchen…">
    <select id="themeFilter">
      <option value="">Alle Themen</option>
      {theme_options}
    </select>
    <select id="statusFilter">
      <option value="">Alle Status</option>
      <option value="unverkäuflich">Unverkäuflich</option>
      <option value="forsale">Zum Verkauf</option>
      <option value="verfügbar">Verfügbar</option>
      <option value="reserviert">Reserviert</option>
      <option value="verkauft">Verkauft</option>
    </select>
    <select id="zustandFilter">
      <option value="">Neu &amp; Gebraucht</option>
      <option value="neu">Neu / OVP</option>
      <option value="gebraucht">Gebraucht</option>
    </select>
  </div>
</div>
<div class="stats-line" id="stats"></div>

<main class="grid" id="grid">
{cards_html}
</main>
<p class="empty-state" id="emptyState">Keine Sets gefunden. Versuch andere Filter.</p>

{render_footer()}

<div class="modal-overlay" id="modalOverlay">
  <div class="modal-box" id="modalBox"></div>
</div>

<script>
{CARD_SCRIPT}
</script>
</body>
</html>"""


def build_stats_html(df: pd.DataFrame, cache: dict) -> str:
    records = [build_card_record(row, cache) for _, row in df.iterrows()]

    themes_all = sorted(set(r["theme"] for r in records))
    theme_filter_options = "".join(
        f'<option value="{html.escape(t)}">{html.escape(t)}</option>' for t in themes_all
    )

    # Kompakte Records fürs Frontend (nur was die Stats-Seite braucht)
    compact = []
    for r in records:
        compact.append({
            "identifier": r["identifier"],
            "name": r["name"],
            "theme": r["theme"],
            "zustand": "neu" if r["zustand_neu"] else "gebraucht",
            "year": r["year"] or None,
            "num_parts": int(r["num_parts"]) if r["num_parts"] else 0,
            "status": r["status"],
            "status_label": r["status_label"],
            "price": r["price"],
        })
    records_json = json.dumps(compact, ensure_ascii=False).replace("</", "<\\/")

    return f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Statistiken — Brick Archive</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
{SHARED_CSS}
  .page {{ max-width: 900px; margin: 0 auto; padding: 40px 24px 64px; }}
  .page h1 {{ font-family: var(--font-display); font-size: clamp(24px, 4vw, 32px); margin: 0 0 8px; }}
  .page > p.intro {{ color: var(--ink-muted); font-size: 14px; margin: 0 0 24px; }}
  .stat-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 16px; margin-bottom: 28px; }}
  .stat-card {{ background: var(--surface); border: 1px solid var(--line); border-radius: 12px; padding: 20px; }}
  .stat-card .num {{ font-family: var(--font-mono); font-size: 28px; font-weight: 500; display: block; }}
  .stat-card .label {{ font-size: 12px; color: var(--ink-muted); margin-top: 4px; }}
  .filter-bar {{ display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 32px; background: var(--surface); border: 1px solid var(--line); border-radius: 10px; padding: 12px 14px; }}
  .filter-bar select, .filter-bar input {{ font-family: var(--font-body); padding: 8px 11px; border: 1px solid var(--line); border-radius: 8px; font-size: 13px; background: var(--bg); color: var(--ink); }}
  .filter-bar input {{ flex: 1; min-width: 160px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
  th, td {{ text-align: left; padding: 10px 12px; border-bottom: 1px solid var(--line); }}
  th {{ font-family: var(--font-mono); font-size: 11px; text-transform: uppercase; color: var(--ink-muted); font-weight: 500; }}
  h2 {{ font-family: var(--font-display); font-size: 18px; margin: 32px 0 12px; }}
  h2:first-of-type {{ margin-top: 0; }}
  .table-stats-line {{ font-size: 12px; color: var(--ink-muted); font-family: var(--font-mono); margin-bottom: 10px; }}
  #fullTable {{ min-width: 640px; }}
  .empty-note {{ color: var(--ink-muted); font-size: 13px; padding: 16px 0; }}
</style>
</head>
<body>
{render_header('stats')}

<main class="page">
  <h1>Statistiken zur Sammlung</h1>
  <p class="intro">Filtere nach Zustand oder Verkaufsstatus — die Zahlen unten passen sich live an.</p>

  <div class="filter-bar">
    <input type="text" id="search" placeholder="Set suchen…">
    <select id="zustandFilter">
      <option value="">Neu &amp; Gebraucht</option>
      <option value="neu">Neu / OVP</option>
      <option value="gebraucht">Gebraucht</option>
    </select>
    <select id="statusFilter">
      <option value="">Alle Status</option>
      <option value="unverkäuflich">Unverkäuflich</option>
      <option value="verfügbar">Verfügbar</option>
      <option value="reserviert">Reserviert</option>
      <option value="verkauft">Verkauft</option>
    </select>
    <select id="themeFilter">
      <option value="">Alle Themen</option>
      {theme_filter_options}
    </select>
  </div>

  <div class="stat-grid" id="statGrid">
    <div class="stat-card"><span class="num" id="statSets">0</span><span class="label">Sets (gefiltert)</span></div>
    <div class="stat-card"><span class="num" id="statParts">0</span><span class="label">Teile (gefiltert)</span></div>
    <div class="stat-card"><span class="num" id="statThemes">0</span><span class="label">Themenwelten (gefiltert)</span></div>
  </div>

  <h2>Nach Themenwelt</h2>
  <table>
    <thead><tr><th>Theme</th><th>Sets</th><th>Teile</th></tr></thead>
    <tbody id="themeTableBody"></tbody>
  </table>
  <p class="empty-note" id="themeEmptyNote" style="display:none;">Keine Sets für diese Filterkombination.</p>

  <h2>Alle Sets (gefiltert)</h2>
  <div class="table-stats-line" id="tableStats"></div>
  <div style="overflow-x:auto;">
  <table id="fullTable">
    <thead>
      <tr>
        <th>Identifier</th><th>Name</th><th>Theme</th><th>Zustand</th>
        <th>Jahr</th><th>Teile</th><th>Status</th><th>Preis</th>
      </tr>
    </thead>
    <tbody id="fullTableBody"></tbody>
  </table>
  </div>
</main>

<script>
  const RECORDS = {records_json};

  const search = document.getElementById('search');
  const zustandFilter = document.getElementById('zustandFilter');
  const statusFilter = document.getElementById('statusFilter');
  const themeFilter = document.getElementById('themeFilter');

  const statSets = document.getElementById('statSets');
  const statParts = document.getElementById('statParts');
  const statThemes = document.getElementById('statThemes');
  const themeTableBody = document.getElementById('themeTableBody');
  const themeEmptyNote = document.getElementById('themeEmptyNote');
  const fullTableBody = document.getElementById('fullTableBody');
  const tableStats = document.getElementById('tableStats');

  function applyFilters() {{
    const q = search.value.toLowerCase();
    const zustand = zustandFilter.value;
    const status = statusFilter.value;
    const theme = themeFilter.value;

    const filtered = RECORDS.filter(r =>
      r.name.toLowerCase().includes(q)
      && (!zustand || r.zustand === zustand)
      && (!status || r.status === status)
      && (!theme || r.theme === theme)
    );

    // Kennzahlen
    statSets.textContent = filtered.length;
    statParts.textContent = filtered.reduce((sum, r) => sum + r.num_parts, 0).toLocaleString('de-CH');
    statThemes.textContent = new Set(filtered.map(r => r.theme)).size;

    // Themenwelt-Aggregation
    const byTheme = {{}};
    filtered.forEach(r => {{
      if (!byTheme[r.theme]) byTheme[r.theme] = {{ count: 0, parts: 0 }};
      byTheme[r.theme].count += 1;
      byTheme[r.theme].parts += r.num_parts;
    }});
    const themeRows = Object.entries(byTheme).sort((a, b) => b[1].count - a[1].count);
    themeTableBody.innerHTML = themeRows.map(([t, v]) =>
      '<tr><td>' + t + '</td><td>' + v.count + '</td><td>' + v.parts.toLocaleString('de-CH') + '</td></tr>'
    ).join('');
    themeEmptyNote.style.display = themeRows.length === 0 ? 'block' : 'none';

    // Volle Tabelle
    fullTableBody.innerHTML = filtered.map(r =>
      '<tr><td>' + r.identifier + '</td><td>' + r.name + '</td><td>' + r.theme + '</td>' +
      '<td>' + (r.zustand === 'neu' ? 'Neu / OVP' : 'Gebraucht') + '</td>' +
      '<td>' + (r.year || '—') + '</td><td>' + (r.num_parts || '—') + '</td>' +
      '<td>' + r.status_label + '</td><td>' + (r.price ? 'CHF ' + r.price : '—') + '</td></tr>'
    ).join('');
    tableStats.textContent = filtered.length + ' von ' + RECORDS.length + ' Sets angezeigt';
  }}

  search.addEventListener('input', applyFilters);
  zustandFilter.addEventListener('change', applyFilters);
  statusFilter.addEventListener('change', applyFilters);
  themeFilter.addEventListener('change', applyFilters);
  applyFilters();
</script>

{render_footer()}
</body>
</html>"""


def main():
    if not DATA_XLSX.exists():
        print(f"Excel-Datei nicht gefunden: {DATA_XLSX}", file=sys.stderr)
        sys.exit(1)

    df = pd.read_excel(DATA_XLSX)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        print(f"Fehlende Pflichtspalten in Excel: {missing}", file=sys.stderr)
        sys.exit(1)

    # Optionale Spalten ergänzen, falls sie (noch) nicht existieren
    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA
    if "Anzahl" in df.columns:
        df["Anzahl"] = df["Anzahl"].fillna(1)

    # Fehlende Identifier automatisch vergeben ({Set Number}-N)
    df = assign_missing_identifiers(df)

    cache = load_cache()
    cache = enrich_data(df, cache)
    save_cache(cache)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    html_index = build_html(df, cache)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html_index)

    html_stats = build_stats_html(df, cache)
    with open(STATS_HTML, "w", encoding="utf-8") as f:
        f.write(html_stats)

    print(f"Fertig. Seiten geschrieben nach {OUTPUT_HTML} und {STATS_HTML}")


if __name__ == "__main__":
    main()
