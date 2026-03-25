#!/usr/bin/env python3
"""
Dynasty League Standings Tracker
Fetches Yahoo NBA Fantasy standings and generates an HTML report.

Usage:
    uv run python main.py           # fetch standings, write index.html
    uv run python main.py --debug   # also save raw HTML to debug.html
"""

import json
import re
import argparse
import sys
import getpass
import sqlite3
import shutil
import tempfile
import base64
from pathlib import Path
from datetime import date, timedelta
from textwrap import dedent

import requests
from bs4 import BeautifulSoup

# ── Config ────────────────────────────────────────────────────────────────────
LEAGUE_ID     = "2197"
NUM_TEAMS     = 12
TEAMLOG_BASE  = f"https://basketball.fantasysports.yahoo.com/nba/{LEAGUE_ID}"
SNAPSHOTS_DIR = Path(__file__).parent / "snapshots"
OUTPUT_FILE   = Path(__file__).parent / "index.html"
TODAY         = date.today()
SEVEN_AGO     = TODAY - timedelta(days=7)

STAT_COEFFS = {
    "FGM":  2.25,
    "FGA": -0.25,
    "FTM":  1.1,
    "FTA": -0.1,
    "3PTM": 1.0,
    "OREB": 1.5,
    "DREB": 1.0,
    "AST":  1.5,
    "ST":   2.5,
    "BLK":  2.0,
    "TO":  -0.5,
    "TECH":-1.0,
    "FF":  -2.0,
}


# ── Cookie extraction (Chrome SQLite + Windows DPAPI) ─────────────────────────

MANUAL_COOKIES_FILE = Path(__file__).parent / "cookies.txt"


def _chrome_aes_key() -> bytes | None:
    """Decrypt Chrome's AES master key from Local State using Windows DPAPI."""
    try:
        import win32crypt
    except ImportError:
        return None
    user = getpass.getuser()
    ls = Path(f"C:/Users/{user}/AppData/Local/Google/Chrome/User Data/Local State")
    if not ls.exists():
        return None
    state = json.loads(ls.read_text(encoding="utf-8"))
    b64 = state.get("os_crypt", {}).get("encrypted_key", "")
    if not b64:
        return None
    enc = base64.b64decode(b64)
    if enc[:5] != b"DPAPI":
        return None
    try:
        _, key = win32crypt.CryptUnprotectData(enc[5:], None, None, None, 0)
        return key
    except Exception:
        return None


def _decrypt_value(enc: bytes, key: bytes) -> str | None:
    if not enc:
        return ""
    if enc[:3] == b"v10":
        try:
            from Crypto.Cipher import AES
            cipher = AES.new(key, AES.MODE_GCM, nonce=enc[3:15])
            return cipher.decrypt_and_verify(enc[15:-16], enc[-16:]).decode("utf-8", errors="ignore")
        except Exception:
            return None
    if enc[:3] == b"v20":
        # Chrome 127+ app-bound encryption — cannot decrypt without Chrome binary
        return None
    try:
        import win32crypt
        _, val = win32crypt.CryptUnprotectData(enc, None, None, None, 0)
        return val.decode("utf-8", errors="ignore")
    except Exception:
        return None


def extract_chrome_cookies() -> list[dict] | None:
    """Read Yahoo cookies from Chrome's SQLite database."""
    key = _chrome_aes_key()
    if key is None:
        return None

    user = getpass.getuser()
    base = Path(f"C:/Users/{user}/AppData/Local/Google/Chrome/User Data/Default")
    db_path = base / "Network" / "Cookies"
    if not db_path.exists():
        db_path = base / "Cookies"
    if not db_path.exists():
        return None

    tmp = Path(tempfile.mktemp(suffix=".db"))
    try:
        shutil.copy2(db_path, tmp)
    except PermissionError:
        # Chrome is running and has the DB locked — fall back to manual cookies
        return None
    try:
        con = sqlite3.connect(str(tmp))
        rows = con.execute(
            "SELECT host_key, name, encrypted_value, path "
            "FROM cookies WHERE host_key LIKE '%.yahoo.com'"
        ).fetchall()
        con.close()
    finally:
        tmp.unlink(missing_ok=True)

    cookies, skipped = [], 0
    for host, name, enc, path in rows:
        val = _decrypt_value(enc, key)
        if val is None:
            skipped += 1
            continue
        cookies.append({"name": name, "value": val, "domain": host, "path": path})

    if skipped:
        print(f"  {skipped} cookies skipped (Chrome 127+ app-bound encryption).")
    print(f"  Extracted {len(cookies)} Yahoo cookies from Chrome.")
    return cookies or None


def load_manual_cookies() -> list[dict] | None:
    """
    Load cookies from cookies.txt (raw Cookie header string).
    Format: name=value; name2=value2; ...
    """
    if not MANUAL_COOKIES_FILE.exists():
        return None
    raw = MANUAL_COOKIES_FILE.read_text(encoding="utf-8").strip()
    cookies = []
    for part in raw.split(";"):
        part = part.strip()
        if "=" in part:
            name, _, value = part.partition("=")
            cookies.append({"name": name.strip(), "value": value.strip(), "domain": ".yahoo.com", "path": "/"})
    print(f"  Loaded {len(cookies)} cookies from cookies.txt.")
    return cookies or None


def print_cookie_instructions() -> None:
    print("""
ERROR: Could not extract Yahoo cookies automatically.

One-time manual setup:
  1. Open Chrome and go to: https://basketball.fantasysports.yahoo.com/nba/2197/standings
  2. Press F12 → Network tab → refresh the page
  3. Click the first request in the list
  4. Under "Request Headers", find the "Cookie:" line
  5. Copy everything after "Cookie: "
  6. Paste it into a file called cookies.txt in this folder

The script will use cookies.txt on every run (re-export if Yahoo logs you out).
""")


UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")




def _load_cookie_header() -> str:
    cookies = extract_chrome_cookies() or load_manual_cookies()
    if not cookies:
        print_cookie_instructions()
        sys.exit(1)
    return "; ".join(f"{c['name']}={c['value']}" for c in cookies)


def fetch_all_from_teamlog(cookie_header: str) -> list[dict]:
    """
    Scrape each team's teamlog page (SSR) to get team name, GP, and computed pts.
    Team IDs 1..NUM_TEAMS map directly to the URL path.
    Stats come from the <tfoot> totals row; pts computed via STAT_COEFFS.
    """
    hdrs = {
        "User-Agent": UA,
        "Cookie": cookie_header,
        "Accept": "text/html,application/xhtml+xml",
    }
    teams = []
    print(f"Fetching teamlog pages for {NUM_TEAMS} teams ...")

    for team_id in range(1, NUM_TEAMS + 1):
        url = f"{TEAMLOG_BASE}/{team_id}/teamlog"
        try:
            r = requests.get(url, headers=hdrs, timeout=20, allow_redirects=True)
            if not r.ok:
                print(f"  Warning: HTTP {r.status_code} for team {team_id}")
                continue
            if "login.yahoo.com" in r.url:
                print("ERROR: Not authenticated — re-export cookies.txt")
                sys.exit(1)

            soup = BeautifulSoup(r.text, "lxml")

            # Team name: first <span class="F-reset Nowrap"> in the page
            span = soup.find("span", class_=lambda c: c and "F-reset" in c and "Nowrap" in c)
            raw = span.get_text(strip=True) if span else f"Team {team_id}"
            # Strip any private-use / non-printable Unicode characters (e.g. icon glyphs)
            name = "".join(c for c in raw if c.isprintable() and not (0xE000 <= ord(c) <= 0xF8FF))

            # Find GP* column index from second header row
            header_row = soup.select_one("thead tr:nth-child(2)") or soup.select_one("thead tr")
            if not header_row:
                print(f"  Warning: no header row for team {team_id}")
                continue
            hdrcells = [th.get_text(strip=True) for th in header_row.find_all(["th", "td"])]
            # Build stat name -> column index map (strip * from GP*)
            col_map = {}
            for ci, h in enumerate(hdrcells):
                key = h.strip().upper().rstrip("*")
                if key:
                    col_map[key] = ci

            tfoot = soup.find("tfoot")
            if not tfoot:
                print(f"  Warning: no tfoot for {name}")
                continue
            cells = [td.get_text(strip=True) for td in tfoot.find_all(["td", "th"])]

            # GP
            gp_col = col_map.get("GP")
            gp = _to_int(cells[gp_col]) if gp_col is not None and gp_col < len(cells) else None

            # Pts via coefficients
            pts = 0.0
            for stat, coeff in STAT_COEFFS.items():
                ci = col_map.get(stat.upper())
                if ci is not None and ci < len(cells):
                    pts += ((_to_float(cells[ci]) or 0.0) * coeff)

            teams.append({"name": name, "gp": gp, "pts": round(pts, 2)})

        except SystemExit:
            raise
        except Exception as e:
            print(f"  Warning: error for team {team_id}: {e}")

    print(f"  Got data for {len(teams)}/{NUM_TEAMS} teams.")
    return teams


def _to_int(v):
    try: return int(v)
    except: return None

def _to_float(v):
    try: return float(str(v).replace(",", ""))
    except: return None



# ── Snapshots ─────────────────────────────────────────────────────────────────

def save_snapshot(teams: list[dict]) -> None:
    SNAPSHOTS_DIR.mkdir(exist_ok=True)
    path = SNAPSHOTS_DIR / f"{TODAY}.json"
    data = {t["name"]: {"gp": t["gp"], "pts": t["pts"]} for t in teams}
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"  Snapshot saved: {path.name}")


def load_all_snapshots() -> dict:
    """Return {date_str: {team_name: {gp, pts}}} for all snapshots except today."""
    result = {}
    if SNAPSHOTS_DIR.exists():
        for f in sorted(SNAPSHOTS_DIR.glob("*.json"), reverse=True):
            if f.stem != str(TODAY):
                try:
                    result[f.stem] = json.loads(f.read_text(encoding="utf-8"))
                except Exception:
                    pass
    return result


# ── HTML ──────────────────────────────────────────────────────────────────────

def build_html(teams: list[dict], snapshots: dict) -> str:
    today_str = TODAY.strftime("%A, %B %d %Y")

    # Build rows — delta cells are empty placeholders, filled by JS
    rows_html = ""
    for t in teams:
        name = t["name"]
        gp   = t["gp"]
        pts  = t["pts"]
        avg  = round(pts / gp, 2) if (gp and pts is not None) else None

        def fmt(v, dec=0):
            if v is None:
                return '<span class="na">\u2014</span>'
            if dec:
                return f"{v:,.{dec}f}"
            return f"{int(v):,}"

        # Escape name for use as HTML attribute value
        name_attr = name.replace('"', '&quot;')
        rows_html += f"""
        <tr data-name="{name_attr}">
          <td class="rank"></td>
          <td class="team">{name}</td>
          <td>{fmt(gp)}</td>
          <td>{fmt(pts, 1)}</td>
          <td>{fmt(avg, 2)}</td>
          <td class="sep delta-gp"><span class="na">\u2014</span></td>
          <td class="delta-pts"><span class="na">\u2014</span></td>
          <td class="delta-avg"><span class="na">\u2014</span></td>
        </tr>"""

    # Build dropdown options: sorted newest-first, labeled "X days ago (Mon DD)"
    options_html = '<option value="">— No comparison —</option>\n'
    for date_str in sorted(snapshots.keys(), reverse=True):
        try:
            snap_date = date.fromisoformat(date_str)
            days = (TODAY - snap_date).days
            label_date = snap_date.strftime("%b %d")
            options_html += f'        <option value="{date_str}">{days} day{"s" if days != 1 else ""} ago ({label_date})</option>\n'
        except ValueError:
            pass

    # Embed all snapshot data + current data as JS
    current_js  = json.dumps({t["name"]: {"gp": t["gp"], "pts": t["pts"]} for t in teams})
    snapshots_js = json.dumps(snapshots)

    return dedent(f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1.0" />
      <title>Dynasty League Standings</title>
      <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
          background: #0f1117;
          color: #e2e8f0;
          padding: 2rem;
        }}
        h1 {{
          font-size: 1.6rem;
          font-weight: 700;
          color: #fff;
          margin-bottom: .3rem;
        }}
        .subtitle {{
          font-size: .85rem;
          color: #64748b;
          margin-bottom: 1.2rem;
        }}
        .controls {{
          display: flex;
          align-items: center;
          gap: .75rem;
          margin-bottom: 1.5rem;
        }}
        .controls label {{
          font-size: .82rem;
          color: #64748b;
        }}
        .controls select {{
          background: #1e2535;
          color: #e2e8f0;
          border: 1px solid #2d3a52;
          border-radius: 6px;
          padding: .4rem .75rem;
          font-size: .82rem;
          cursor: pointer;
          outline: none;
        }}
        .controls select:hover {{ border-color: #4b5568; }}
        table {{
          width: 100%;
          border-collapse: collapse;
          font-size: .9rem;
        }}
        .group-header th {{
          background: #151b2b;
          color: #4b5568;
          font-size: .7rem;
          padding: .35rem 1rem;
          text-align: center;
        }}
        .group-header th.blank {{ background: #0f1117; }}
        thead th {{
          background: #1e2535;
          color: #94a3b8;
          font-size: .75rem;
          text-transform: uppercase;
          letter-spacing: .06em;
          padding: .65rem 1rem;
          text-align: right;
          white-space: nowrap;
          cursor: pointer;
          user-select: none;
          border-bottom: 2px solid #2d3a52;
        }}
        thead th:hover {{ color: #e2e8f0; }}
        thead th.left {{ text-align: left; }}
        .sort-asc::after  {{ content: " \u25b2"; font-size: .65rem; }}
        .sort-desc::after {{ content: " \u25bc"; font-size: .65rem; }}
        tbody tr {{ border-bottom: 1px solid #1a2236; transition: background .1s; }}
        tbody tr:hover {{ background: #1a2236; }}
        td {{
          padding: .7rem 1rem;
          text-align: right;
          white-space: nowrap;
        }}
        td.rank {{ text-align: left; color: #4b5568; font-size: .8rem; width: 2.2rem; padding-right: 0; }}
        td.team {{ text-align: left; font-weight: 600; color: #f1f5f9; min-width: 160px; }}
        td.sep  {{ border-left: 2px solid #2d3a52; }}
        .na {{ color: #374151; }}
        .footer {{ margin-top: 1.25rem; font-size: .72rem; color: #374151; }}
      </style>
    </head>
    <body>
      <h1>Dynasty League</h1>
      <div class="subtitle">{today_str}</div>

      <div class="controls">
        <label for="snapselect">Compare against:</label>
        <select id="snapselect" onchange="updateDeltas()">
        {options_html}
        </select>
        <span id="period-label" style="font-size:.82rem;color:#64748b;"></span>
      </div>

      <table id="tbl">
        <thead>
          <tr class="group-header">
            <th class="blank" colspan="2"></th>
            <th colspan="3">Overall</th>
            <th class="sep" id="period-header" colspan="3">Comparison Period</th>
          </tr>
          <tr>
            <th class="left" style="width:2.2rem">#</th>
            <th class="left" onclick="sort(1)">Team</th>
            <th onclick="sort(2)">GP</th>
            <th onclick="sort(3)">Pts</th>
            <th onclick="sort(4)">Avg</th>
            <th class="sep" onclick="sort(5)">GP</th>
            <th onclick="sort(6)">Pts</th>
            <th onclick="sort(7)">Avg</th>
          </tr>
        </thead>
        <tbody>
{rows_html}
        </tbody>
      </table>

      <div class="footer">
        Auto-generated · Yahoo Fantasy Sports · League {LEAGUE_ID}
      </div>

      <script>
        const CURRENT   = {current_js};
        const SNAPSHOTS = {snapshots_js};

        let lastCol = 3, lastAsc = false;

        function fmtNum(v, dec) {{
          if (v === null || v === undefined) return '<span class="na">\u2014</span>';
          if (dec) return v.toLocaleString('en-US', {{minimumFractionDigits: dec, maximumFractionDigits: dec}});
          return Math.round(v).toLocaleString('en-US');
        }}

        function updateDeltas() {{
          const dateStr = document.getElementById('snapselect').value;
          const snap    = dateStr ? SNAPSHOTS[dateStr] : null;

          // Update column group header
          if (dateStr) {{
            const d     = new Date(dateStr + 'T00:00:00');
            const today = new Date('{TODAY}T00:00:00');
            const days  = Math.round((today - d) / 86400000);
            document.getElementById('period-header').textContent = 'Last ' + days + ' Day' + (days !== 1 ? 's' : '');
            document.getElementById('period-label').textContent  = 'vs snapshot from ' + d.toLocaleDateString('en-US', {{month:'short', day:'numeric'}});
          }} else {{
            document.getElementById('period-header').textContent = 'Comparison Period';
            document.getElementById('period-label').textContent  = '';
          }}

          document.querySelectorAll('#tbl tbody tr').forEach(row => {{
            const name = row.dataset.name;
            const cur  = CURRENT[name];
            const old  = snap && snap[name];

            let dgp, dpts, davg;
            if (cur && old) {{
              dgp  = (cur.gp  != null && old.gp  != null) ? cur.gp  - old.gp  : null;
              dpts = (cur.pts != null && old.pts != null) ? cur.pts - old.pts : null;
              davg = (dgp && dpts != null)                ? dpts / dgp        : null;
            }}

            row.querySelector('.delta-gp' ).innerHTML = fmtNum(dgp,  0);
            row.querySelector('.delta-pts').innerHTML = fmtNum(dpts, 1);
            row.querySelector('.delta-avg').innerHTML = fmtNum(davg, 2);
          }});

          sort(lastCol, lastAsc);
        }}

        function sort(col, forceAsc) {{
          const tbody = document.querySelector('#tbl tbody');
          const rows  = Array.from(tbody.rows);
          const asc   = (forceAsc !== undefined) ? forceAsc : (col === lastCol ? !lastAsc : false);
          lastCol = col; lastAsc = asc;

          rows.sort((a, b) => {{
            const av = a.cells[col].innerText.replace(/,/g, '').trim();
            const bv = b.cells[col].innerText.replace(/,/g, '').trim();
            const an = parseFloat(av), bn = parseFloat(bv);
            if (isNaN(an) && isNaN(bn)) return asc ? av.localeCompare(bv) : bv.localeCompare(av);
            if (isNaN(an)) return 1;
            if (isNaN(bn)) return -1;
            return asc ? an - bn : bn - an;
          }});

          rows.forEach((r, i) => {{ tbody.appendChild(r); r.cells[0].textContent = i + 1; }});

          document.querySelectorAll('thead th').forEach(th => th.className = th.className.replace(/ sort-(asc|desc)/g, ''));
          document.querySelectorAll('thead tr:last-child th')[col].classList.add(asc ? 'sort-asc' : 'sort-desc');
        }}

        // Auto-select the snapshot closest to 7 days ago on load
        (function() {{
          const dates = Object.keys(SNAPSHOTS).sort();
          if (!dates.length) return;
          const today = new Date('{TODAY}T00:00:00');
          let best = dates[0], bestDiff = Infinity;
          dates.forEach(d => {{
            const diff = Math.abs(Math.round((today - new Date(d + 'T00:00:00')) / 86400000) - 7);
            if (diff < bestDiff) {{ bestDiff = diff; best = d; }}
          }});
          document.getElementById('snapselect').value = best;
          updateDeltas();
        }})();
      </script>
    </body>
    </html>
    """).strip()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    argparse.ArgumentParser(description="Dynasty League standings scraper").parse_args()

    print("=== Dynasty League Standings ===")

    print("Loading cookies ...")
    cookie_header = _load_cookie_header()

    teams = fetch_all_from_teamlog(cookie_header)

    if not teams:
        print("No team data found. Run with --debug and open debug.html to inspect the page.")
        sys.exit(1)

    print(f"  Found {len(teams)} teams.")
    save_snapshot(teams)

    snapshots = load_all_snapshots()
    if snapshots:
        print(f"  Loaded {len(snapshots)} historical snapshot(s): {', '.join(sorted(snapshots)[-3:])}")
    else:
        print("  No historical snapshots yet (comparison dropdown will be empty)")

    OUTPUT_FILE.write_text(build_html(teams, snapshots), encoding="utf-8")
    print(f"  Report written: {OUTPUT_FILE}")
    print("Done. Open index.html in your browser.")


if __name__ == "__main__":
    main()
