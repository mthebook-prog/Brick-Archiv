#!/usr/bin/env python3
"""
Liest die Lego-Excel-Liste, reichert jedes Set einmalig über die Rebrickable-API an
(mit persistentem JSON-Cache, damit bereits bekannte Sets nicht erneut abgefragt werden),
und generiert daraus eine statische HTML-Seite (docs/index.html) für GitHub Pages.

Aufruf:
    python scripts/build.py

Benötigt die Umgebungsvariable REBRICKABLE_API_KEY.
"""

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

API_KEY = os.environ.get("REBRICKABLE_API_KEY")
API_BASE = "https://rebrickable.com/api/v3/lego"

REQUIRED_COLUMNS = [
    "Identifier", "Set Number", "Name", "ID", "THEME", "Anzahl",
    "Zustand", "Bauanleitung", "Original-Karton",
]


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
        else:
            print(f"  Fehler {resp.status_code} bei Set {set_num}", file=sys.stderr)
            return None
    except requests.RequestException as e:
        print(f"  Netzwerkfehler bei Set {set_num}: {e}", file=sys.stderr)
        return None


def real_set_number(row) -> str:
    """
    Ermittelt die tatsächliche Rebrickable-Set-Nummer.
    Mehrfach-Exemplare (z.B. Identifier '31120-2', '31120-3') teilen sich alle
    dieselben Katalog-Daten der Basisnummer '{Set Number}-1'.
    """
    return f"{row['Set Number']}-1"


def enrich_data(df: pd.DataFrame, cache: dict) -> dict:
    """Holt für jede eindeutige Set-Nummer die API-Daten (aus Cache oder live)."""
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
            time.sleep(0.2)  # kleine Pause, freundlich zur API

    print(f"{new_calls} neue API-Aufrufe, {len(unique_set_nums) - new_calls} aus Cache.")
    return cache


def build_html(df: pd.DataFrame, cache: dict) -> str:
    themes = sorted(df["THEME"].dropna().unique())
    theme_options = "".join(f'<option value="{t}">{t}</option>' for t in themes)
    theme_chip_counts = df["THEME"].value_counts()
    total_parts = 0

    cards = []
    for _, row in df.iterrows():
        set_num = real_set_number(row)
        api = cache.get(set_num, {})
        img = api.get("set_img_url", "")
        year = api.get("year", "")
        parts = api.get("num_parts", "")
        set_url = api.get("set_url", "")
        if isinstance(parts, (int, float)) and not pd.isna(parts):
            total_parts += int(parts)

        zustand_label = "Neu / ungeöffnet" if str(row["Zustand"]).strip() == "J" else "Gebraucht / geöffnet"
        status = str(row.get("Verkaufsstatus", "verfügbar")).strip() or "verfügbar"
        status_class = {
            "verfügbar": "status-available",
            "reserviert": "status-reserved",
            "verkauft": "status-sold",
        }.get(status, "status-available")

        preis = row.get("Preis", "")
        has_price = preis not in ("", None) and pd.notna(preis)
        preis_html = f'<div class="price">CHF {preis}</div>' if has_price else ""

        for_sale_html = ""
        if status == "verfügbar" and has_price:
            for_sale_html = '''
                <button class="buy-btn" onclick="alert('Kauf-Workflow folgt in einem späteren Schritt.')">
                  Zum Verkauf anfragen
                </button>'''

        badges = []
        if str(row.get('Bauanleitung', '')).strip() == 'J':
            badges.append('<span class="tag"><i class="dot"></i>Bauanleitung</span>')
        if str(row.get('Original-Karton', '')).strip() == 'J':
            badges.append('<span class="tag"><i class="dot"></i>Originalkarton</span>')
        badges_html = "".join(badges)

        cards.append(f"""
        <article class="card" data-theme="{row['THEME']}" data-status="{status}" data-name="{row['Name'].lower()}">
          <div class="img-wrap">
            <img src="{img}" alt="{row['Name']}" loading="lazy" onerror="this.parentElement.classList.add('img-fallback')">
            <span class="badge {status_class}">{status}</span>
          </div>
          <div class="card-body">
            <div class="card-top">
              <span class="set-num">#{row['Set Number']}</span>
              {f'<span class="set-num">{year}</span>' if year else ''}
            </div>
            <h3>{row['Name']}</h3>
            <div class="theme-tag">{row['THEME']}</div>
            <div class="specs">
              <span>{zustand_label}</span>
              {f'<span>{parts} Teile</span>' if parts else ''}
            </div>
            {f'<div class="tags">{badges_html}</div>' if badges_html else ''}
            {preis_html}
            {for_sale_html}
            {f'<a class="external-link" href="{set_url}" target="_blank" rel="noopener">Details bei Rebrickable →</a>' if set_url else ''}
          </div>
        </article>""")

    cards_html = "\n".join(cards)
    n_available = int((df.get("Verkaufsstatus", pd.Series(dtype=str)).fillna("verfügbar").str.strip() == "verfügbar").sum()) if len(df) else 0

    return f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Meine Lego-Sammlung</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root {{
    --bg: #FAFAF7; --surface: #FFFFFF; --ink: #1C1C1E; --ink-muted: #6B6B68;
    --line: #E7E4DC; --brick: #D6001C; --brick-dark: #A1000F;
    --green: #2F7D46; --amber: #B9781A; --sold: #8C8C86;
    --font-display: 'Space Grotesk', sans-serif;
    --font-body: 'Inter', -apple-system, sans-serif;
    --font-mono: 'JetBrains Mono', monospace;
  }}
  * {{ box-sizing: border-box; }}
  html {{ scroll-behavior: smooth; }}
  body {{
    font-family: var(--font-body); background: var(--bg); color: var(--ink);
    margin: 0; padding: 0; -webkit-font-smoothing: antialiased;
  }}
  a {{ color: inherit; text-decoration: none; }}

  /* ---- Header ---- */
  .site-header {{
    position: sticky; top: 0; z-index: 50;
    background: var(--surface); border-bottom: 1px solid var(--line);
  }}
  .header-inner {{
    max-width: 1200px; margin: 0 auto; padding: 0 24px;
    display: flex; align-items: center; justify-content: space-between; height: 64px;
  }}
  .brand {{ display: flex; align-items: center; gap: 10px; font-family: var(--font-display); font-weight: 700; font-size: 19px; }}
  .brand .studs {{ display: inline-flex; gap: 3px; }}
  .brand .studs span {{ width: 7px; height: 7px; border-radius: 50%; background: var(--brick); display: block; }}
  .brand .studs span:nth-child(2) {{ background: var(--ink); opacity: .8; }}
  .brand .studs span:nth-child(3) {{ background: var(--ink); opacity: .5; }}

  nav.desktop-nav {{ display: flex; gap: 28px; font-size: 14px; font-weight: 500; }}
  nav.desktop-nav a {{ color: var(--ink-muted); transition: color .15s; }}
  nav.desktop-nav a:hover {{ color: var(--brick); }}

  .hamburger {{
    display: none; flex-direction: column; justify-content: center; gap: 4px;
    width: 36px; height: 36px; border: none; background: none; cursor: pointer; padding: 0;
  }}
  .hamburger span {{ display: block; width: 22px; height: 2px; background: var(--ink); transition: all .2s; margin: 0 auto; }}
  .hamburger.open span:nth-child(1) {{ transform: translateY(6px) rotate(45deg); }}
  .hamburger.open span:nth-child(2) {{ opacity: 0; }}
  .hamburger.open span:nth-child(3) {{ transform: translateY(-6px) rotate(-45deg); }}

  .mobile-nav {{
    display: none; flex-direction: column; background: var(--surface);
    border-bottom: 1px solid var(--line); overflow: hidden;
    max-height: 0; transition: max-height .25s ease;
  }}
  .mobile-nav.open {{ max-height: 240px; }}
  .mobile-nav a {{ padding: 14px 24px; font-size: 15px; font-weight: 500; border-top: 1px solid var(--line); }}

  /* ---- Hero / intro ---- */
  .hero {{ max-width: 1200px; margin: 0 auto; padding: 40px 24px 8px; }}
  .hero h1 {{ font-family: var(--font-display); font-size: clamp(26px, 4vw, 36px); font-weight: 700; margin: 0 0 8px; letter-spacing: -0.01em; }}
  .hero p {{ color: var(--ink-muted); font-size: 15px; margin: 0; }}
  .stat-row {{ display: flex; gap: 24px; margin-top: 20px; flex-wrap: wrap; }}
  .stat {{ font-family: var(--font-mono); }}
  .stat .num {{ font-size: 22px; font-weight: 500; display: block; }}
  .stat .label {{ font-size: 12px; color: var(--ink-muted); }}

  /* ---- Filter bar ---- */
  .filter-bar {{
    position: sticky; top: 64px; z-index: 40; background: var(--bg);
    border-bottom: 1px solid var(--line);
  }}
  .filter-inner {{
    max-width: 1200px; margin: 0 auto; padding: 14px 24px;
    display: flex; gap: 10px; flex-wrap: wrap; align-items: center;
  }}
  .filter-inner input, .filter-inner select {{
    font-family: var(--font-body); padding: 9px 12px; border: 1px solid var(--line);
    border-radius: 8px; font-size: 14px; background: var(--surface); color: var(--ink);
  }}
  .filter-inner input {{ flex: 1; min-width: 180px; }}
  .filter-inner input:focus, .filter-inner select:focus {{ outline: 2px solid var(--brick); outline-offset: 1px; }}
  .stats-line {{ font-size: 12px; color: var(--ink-muted); font-family: var(--font-mono); padding: 10px 24px 0; max-width: 1200px; margin: 0 auto; }}

  /* ---- Grid ---- */
  .grid {{
    display: grid; grid-template-columns: repeat(auto-fill, minmax(230px, 1fr));
    gap: 18px; max-width: 1200px; margin: 0 auto; padding: 16px 24px 64px;
  }}
  .card {{
    background: var(--surface); border: 1px solid var(--line); border-radius: 12px;
    overflow: hidden; display: flex; flex-direction: column; transition: border-color .15s, transform .15s;
  }}
  .card:hover {{ border-color: var(--ink); }}
  .img-wrap {{ position: relative; background: #F1F0EA; aspect-ratio: 4/3; display: flex; align-items: center; justify-content: center; }}
  .img-wrap img {{ max-width: 82%; max-height: 82%; object-fit: contain; }}
  .img-wrap.img-fallback::before {{
    content: ''; position: absolute; inset: 0;
    background-image: radial-gradient(var(--line) 1.5px, transparent 1.5px);
    background-size: 14px 14px;
  }}
  .badge {{
    position: absolute; top: 10px; right: 10px; font-size: 10px; font-weight: 600;
    padding: 4px 9px; border-radius: 20px; color: white; text-transform: uppercase; letter-spacing: .03em;
  }}
  .status-available {{ background: var(--green); }}
  .status-reserved {{ background: var(--amber); }}
  .status-sold {{ background: var(--sold); }}

  .card-body {{ padding: 14px 16px 16px; display: flex; flex-direction: column; gap: 6px; }}
  .card-top {{ display: flex; justify-content: space-between; font-family: var(--font-mono); font-size: 11px; color: var(--ink-muted); }}
  .card-body h3 {{ font-family: var(--font-display); margin: 2px 0 0; font-size: 15px; font-weight: 500; line-height: 1.3; }}
  .theme-tag {{ font-size: 11px; color: var(--ink-muted); }}
  .specs {{ display: flex; gap: 10px; font-size: 12px; color: var(--ink-muted); margin-top: 2px; }}
  .tags {{ display: flex; gap: 8px; flex-wrap: wrap; margin-top: 2px; }}
  .tag {{ display: flex; align-items: center; gap: 5px; font-size: 11px; color: var(--ink-muted); }}
  .tag .dot {{ width: 5px; height: 5px; border-radius: 50%; background: var(--green); display: inline-block; }}
  .price {{ font-family: var(--font-mono); font-size: 17px; font-weight: 500; margin-top: 6px; }}
  .buy-btn {{
    margin-top: 2px; padding: 10px; border: none; border-radius: 8px; background: var(--brick);
    color: white; font-family: var(--font-body); font-weight: 600; font-size: 13px; cursor: pointer; transition: background .15s;
  }}
  .buy-btn:hover {{ background: var(--brick-dark); }}
  .external-link {{ font-size: 12px; color: var(--ink-muted); margin-top: 2px; }}
  .external-link:hover {{ color: var(--brick); }}

  .empty-state {{ text-align: center; padding: 64px 24px; color: var(--ink-muted); display: none; }}

  footer {{ border-top: 1px solid var(--line); padding: 28px 24px; text-align: center; font-size: 12px; color: var(--ink-muted); }}

  @media (max-width: 720px) {{
    nav.desktop-nav {{ display: none; }}
    .hamburger {{ display: flex; }}
    .mobile-nav {{ display: flex; }}
    .filter-bar {{ top: 64px; }}
    .grid {{ grid-template-columns: repeat(auto-fill, minmax(150px, 1fr)); gap: 12px; padding: 12px 16px 48px; }}
    .hero {{ padding: 28px 16px 4px; }}
    .filter-inner, .stats-line {{ padding-left: 16px; padding-right: 16px; }}
  }}
</style>
</head>
<body>

<header class="site-header">
  <div class="header-inner">
    <a class="brand" href="#top">
      <span class="studs"><span></span><span></span><span></span></span>
      Brick Archive
    </a>
    <nav class="desktop-nav">
      <a href="#top">Sammlung</a>
      <a href="#top" onclick="document.getElementById('statusFilter').value='verfügbar'; document.getElementById('statusFilter').dispatchEvent(new Event('change')); return true;">Zum Verkauf</a>
      <a href="#footer">Kontakt</a>
    </nav>
    <button class="hamburger" id="hamburger" aria-label="Menü öffnen" aria-expanded="false">
      <span></span><span></span><span></span>
    </button>
  </div>
  <nav class="mobile-nav" id="mobileNav">
    <a href="#top">Sammlung</a>
    <a href="#top" id="mobileForSale">Zum Verkauf</a>
    <a href="#footer">Kontakt</a>
  </nav>
</header>

<div id="top"></div>
<section class="hero">
  <h1>Meine Lego-Sammlung</h1>
  <p>Eine kuratierte Übersicht aller Sets — durchsuchbar, filterbar, teils käuflich.</p>
  <div class="stat-row">
    <div class="stat"><span class="num">{len(df)}</span><span class="label">Sets gesamt</span></div>
    <div class="stat"><span class="num">{n_available}</span><span class="label">Verfügbar</span></div>
    <div class="stat"><span class="num">{total_parts:,}</span><span class="label">Teile gesamt</span></div>
    <div class="stat"><span class="num">{len(themes)}</span><span class="label">Themenwelten</span></div>
  </div>
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
      <option value="verfügbar">Verfügbar</option>
      <option value="reserviert">Reserviert</option>
      <option value="verkauft">Verkauft</option>
    </select>
  </div>
</div>
<div class="stats-line" id="stats"></div>

<main class="grid" id="grid">
{cards_html}
</main>
<p class="empty-state" id="emptyState">Keine Sets gefunden. Versuch andere Filter.</p>

<footer id="footer">
  Brick Archive — privat gepflegte Sammlung. Rebrickable-Daten via API.
</footer>

<script>
  const hamburger = document.getElementById('hamburger');
  const mobileNav = document.getElementById('mobileNav');
  hamburger.addEventListener('click', () => {{
    const isOpen = mobileNav.classList.toggle('open');
    hamburger.classList.toggle('open', isOpen);
    hamburger.setAttribute('aria-expanded', isOpen);
  }});
  document.getElementById('mobileForSale').addEventListener('click', () => {{
    document.getElementById('statusFilter').value = 'verfügbar';
    document.getElementById('statusFilter').dispatchEvent(new Event('change'));
    mobileNav.classList.remove('open');
    hamburger.classList.remove('open');
  }});

  const search = document.getElementById('search');
  const themeFilter = document.getElementById('themeFilter');
  const statusFilter = document.getElementById('statusFilter');
  const cards = Array.from(document.querySelectorAll('.card'));
  const stats = document.getElementById('stats');
  const emptyState = document.getElementById('emptyState');

  function applyFilters() {{
    const q = search.value.toLowerCase();
    const theme = themeFilter.value;
    const status = statusFilter.value;
    let visible = 0;
    cards.forEach(c => {{
      const show = c.dataset.name.includes(q)
        && (!theme || c.dataset.theme === theme)
        && (!status || c.dataset.status === status);
      c.style.display = show ? '' : 'none';
      if (show) visible++;
    }});
    stats.textContent = visible + ' von ' + cards.length + ' Sets angezeigt';
    emptyState.style.display = visible === 0 ? 'block' : 'none';
  }}

  search.addEventListener('input', applyFilters);
  themeFilter.addEventListener('change', applyFilters);
  statusFilter.addEventListener('change', applyFilters);
  applyFilters();
</script>
</body>
</html>"""


def main():
    if not DATA_XLSX.exists():
        print(f"Excel-Datei nicht gefunden: {DATA_XLSX}", file=sys.stderr)
        sys.exit(1)

    df = pd.read_excel(DATA_XLSX)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        print(f"Fehlende Spalten in Excel: {missing}", file=sys.stderr)
        sys.exit(1)

    cache = load_cache()
    cache = enrich_data(df, cache)
    save_cache(cache)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    html = build_html(df, cache)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Fertig. Seite geschrieben nach {OUTPUT_HTML}")


if __name__ == "__main__":
    main()
