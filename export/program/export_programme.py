"""Export the current website programme to an A4 PDF (standard, booklet or three-fold leaflet)."""
import argparse
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from html import escape
from urllib.parse import urljoin

from bs4 import BeautifulSoup
import yaml

HERE = Path(__file__).resolve().parent


def find_browser():
    for root in (os.getenv("ProgramFiles(x86)"), os.getenv("ProgramFiles"),
                 os.getenv("LocalAppData")):
        if root:
            for relative in ("Microsoft/Edge/Application/msedge.exe",
                             "Google/Chrome/Application/chrome.exe"):
                path = Path(root) / relative
                if path.is_file():
                    return path
    for name in ("msedge", "chrome", "google-chrome", "chromium"):
        if path := shutil.which(name):
            return Path(path)
    raise ValueError("Install Microsoft Edge or Google Chrome to print the PDF.")


def render(root, settings, booklet=False, trifold=False):
    site = yaml.safe_load((root / "_config.yml").read_text(encoding="utf-8-sig"))
    source = (root / "index.md").read_text(encoding="utf-8-sig")
    _, metadata, body = source.split("---", 2)
    page = yaml.safe_load(metadata)
    soup = BeautifulSoup(body, "html.parser")

    if not settings["show_day_buttons"]:
        for nav in soup.select(".day-nav"):
            nav.decompose()
    if not settings["show_intro"]:
        for intro in soup.select(".intro"):
            intro.decompose()

    overrides = settings["chairs"]
    matched = set()
    for day in soup.select('h2[id^="day-"]'):
        table = day.find_next_sibling("table")
        if settings["start_each_day_on_new_page"] and day["id"] != "day-1":
            table["class"] = ["new-page"]
        rows = table.select("tr")
        # Keep each section heading and chair together with the next item.
        pending = []
        for row in rows:
            classes = row.get("class", [])
            if "chair-row" in classes:
                next_row = row.find_next_sibling("tr")
                time = next_row.td.get_text(strip=True) if next_row and next_row.td else ""
                key = f'{day["id"]} {time}'
                if key in overrides and key in matched:
                    raise ValueError(f"More than one chair row matches {key}")
                matched.add(key)
                if name := overrides.get(key):
                    row.td.string = f"Πρόεδρος: {name}"
                elif settings["hide_empty_chairs"] and not row.td.get_text(strip=True).partition(":")[2].strip():
                    row.decompose()
                    continue
            if row.find("th"):
                continue
            pending.append(row)
            if "chair-row" not in classes and "session-start" not in classes:
                group = soup.new_tag("tbody", attrs={"class": "print-group"})
                table.append(group)
                for item in pending:
                    group.append(item)
                pending.clear()

        # Put the heading in the first print group so it travels with the
        # opening item when the day starts near the foot of a page.
        if first_group := table.select_one(".print-group"):
            heading_row = soup.new_tag("tr", attrs={"class": "day-heading"})
            cell = soup.new_tag("td", colspan=4)
            heading_row.append(cell)
            cell.append(day)
            first_group.insert(0, heading_row)

    if unused := {key for key, name in overrides.items() if name} - matched:
        raise ValueError(f"Chair slots missing from index.md: {', '.join(sorted(unused))}")

    for link in soup.select("a[href]"):
        href = link["href"]
        if re.match(r"\{\{\s*site\.baseurl\s*\}\}", href):
            if base := settings["abstract_base_url"]:
                path = href.split("}}", 1)[1].lstrip("/")
                link["href"] = urljoin(base.rstrip("/") + "/", path)
            else:
                link.unwrap()

    def site_text(key):
        return escape(str(site[key]), quote=True)

    logos = ""
    if settings["show_logos"]:
        images = root / "assets" / "images"
        logos = (
            f'<div class="society-banner"><img class="society-logo" src="{(HERE / "society-logo-original.png").as_uri()}" alt="{site_text("organiser")}"></div>'
            f'<img class="conference-logo" src="{(images / "arch_logo_dark-gray.svg").as_uri()}" alt="">'
        )

    venue = ""
    if settings["show_map"]:
        image = ""
        if filename := settings["map_image"]:
            path = (HERE / filename).resolve()
            if not path.is_file():
                raise ValueError(f"Map image not found: {path}")
            image = f'<img class="map-image" src="{path.as_uri()}" alt="Χάρτης">'
        venue = f"""
<section class="venue-map">
<h2>Τόπος διεξαγωγής</h2>
<p>{site_text("venue_name")}<br>
<a href="{site_text("venue_map_url")}">{site_text("venue_address")} · Προβολή χάρτη</a></p>
{image}
</section>"""

    if trifold:
        return render_trifold(soup, venue, page, site, site_text, settings)

    return f"""<!doctype html>
<html lang="{site_text('lang')}">
<head>
<meta charset="utf-8">
<title>{escape(page['title'])} · {site_text('title')}</title>
<link rel="stylesheet" href="{(HERE / 'print.css').as_uri()}">
</head>
<body class="{'booklet' if booklet else 'standard'}">
<header>
{logos}
<p class="conference-title">{site_text("title")}</p>
<p>{site_text("subtitle")}</p>
<p class="place-date">{site_text("place_date")} · {site_text("venue_address")}</p>
<h1>{escape(page["title"])}</h1>
</header>
<main>{soup}{venue}</main>
</body>
</html>
"""


TRIFOLD_SCRIPT = r"""
(async () => {
  try { await document.fonts.ready; } catch (e) {}
  const opts = JSON.parse(document.getElementById("trifold-options").textContent);
  const flow = document.getElementById("flow");
  const units = [...flow.querySelectorAll(":scope > table > tbody.print-group, :scope > .venue-map")];
  const panels = opts.order.map(id => document.querySelector(`.panel[data-panel="${id}"] .panel-body`));
  const colgroup = '<colgroup><col class="c-time"><col class="c-item"></colgroup>';

  function overflows(panel) { return panel.scrollHeight > panel.clientHeight + 0.5; }

  function layout(pt) {
    document.documentElement.style.setProperty("--trifold-size", pt + "pt");
    panels.forEach(p => p.replaceChildren());
    flow.append(...units);
    let i = 0, table = null, day = "";
    for (const unit of units) {
      const heading = unit.querySelector("h2");
      if (heading) day = heading.textContent.trim();
      for (;;) {
        if (i >= panels.length) return false;
        const panel = panels[i];
        let added = [];
        if (unit.tagName === "TBODY") {
          if (!table || table.parentNode !== panel) {
            table = document.createElement("table");
            table.innerHTML = colgroup;
            panel.append(table);
            added.push(table);
            if (!heading && day) {
              const cont = document.createElement("tbody");
              cont.innerHTML = `<tr class="day-cont"><td colspan="2">${day} (συνέχεια)</td></tr>`;
              table.append(cont);
              added.push(cont);
            }
          }
          table.append(unit);
        } else {
          panel.append(unit);
          table = null;
        }
        if (!overflows(panel)) break;
        // does not fit: move to the next panel
        unit.remove();
        added.forEach(el => el.remove());
        if (table && !table.querySelector("tbody")) table = null;
        if (!panel.children.length) return false;   // larger than an empty panel
        table = null;
        i += 1;
      }
    }
    return true;
  }

  let fitted = false;
  for (let pt = opts.max; pt >= opts.min - 1e-9; pt -= 0.25) {
    if (layout(+pt.toFixed(2))) { fitted = true; break; }
  }
  if (!fitted) {
    // Still too long at the smallest size: add extra panels so that nothing is lost.
    document.body.classList.add("trifold-overflow");
    const extra = document.createElement("section");
    extra.className = "sheet extra";
    for (let n = 0; n < 6; n++) {
      const panel = document.createElement("div");
      panel.className = "panel";
      panel.innerHTML = '<div class="panel-body"></div>';
      extra.append(panel);
      panels.push(panel.firstChild);
    }
    document.body.append(extra);
    layout(opts.min);
  }
  flow.remove();
  document.documentElement.dataset.trifold = fitted ? "ok" : "overflow";
})();
"""


def render_trifold(soup, venue, page, site, site_text, settings):
    """A4 landscape, folded in three (roll fold).

    Page 1 (outside): A inside flap | B back | C front cover (laid out like the poster).
    Page 2 (inside):  D | E | F.
    The programme flows through the panels in settings["trifold"]["order"] and the type size
    shrinks until everything fits.
    """
    options = {"order": ["D", "E", "F", "A"], "max_font_pt": 9, "min_font_pt": 6}
    options.update(settings.get("trifold", {}))

    # Two columns fit a 99 mm panel better than four: time | speakers + title.
    for table in soup.select("table"):
        for header in table.select("tr"):
            if header.find("th"):
                header.decompose()
        for row in table.select("tr"):
            cells = row.find_all("td", recursive=False)
            if len(cells) != 4:
                if cells:
                    cells[0]["colspan"] = 2
                continue
            start, end, who, what = cells
            times = "–".join(t for t in (start.get_text(strip=True), end.get_text(strip=True)) if t)
            start.string = times
            if who.get_text(strip=True):
                who.name = "span"
                who["class"] = ["speakers"]
                what.insert(0, who)
                if what.contents[1:]:
                    who.insert_after(soup.new_tag("br"))
            else:
                who.decompose()
            end.decompose()

    def image(key, css_class, alt=""):
        if not (filename := options.get(key)):
            return ""
        path = (HERE / filename).resolve()
        if not path.is_file():
            raise ValueError(f"Image for trifold.{key} not found: {path}")
        return f'<img class="{css_class}" src="{path.as_uri()}" alt="{escape(alt, quote=True)}">'

    title_lines = options.get("cover_title") or [str(site["title"])]
    title = "<br>".join(escape(line) for line in title_lines)
    arch = image("cover_image", "cover-image")
    cover = f"""
<div class="cover">
<div class="cover-top">{image("cover_logo", "cover-logo", str(site["organiser"]))}</div>
<p class="cover-label">{escape(options.get("cover_label", ""))}</p>
<p class="conference-title">{title}</p>
<p class="subtitle">{site_text("subtitle")}</p>
{f'<div class="cover-image-box">{arch}</div>' if arch else ""}
<p class="cover-programme">{escape(page["title"])}</p>
<p class="place-date">{site_text("place_date")}</p>
<p class="venue">{site_text("venue_name")}<br>{site_text("venue_address")}</p>
{image("cover_qr", "cover-qr", "QR")}
</div>"""

    order = json.dumps({"order": options["order"], "max": options["max_font_pt"],
                        "min": options["min_font_pt"]})
    panel = lambda key, inner="": f'<div class="panel" data-panel="{key}"><div class="panel-body">{inner}</div></div>'
    return f"""<!doctype html>
<html lang="{site_text('lang')}">
<head>
<meta charset="utf-8">
<title>{escape(page['title'])} · {site_text('title')}</title>
<link rel="stylesheet" href="{(HERE / 'print.css').as_uri()}">
</head>
<body class="trifold">
<section class="sheet outside" style="--cover-bg: {escape(options.get('cover_background', '#f4f1e8'), quote=True)}; --cover-image-bg: {escape(options.get('cover_image_background', '#4e1818'), quote=True)}">{panel("A")}{panel("B")}<div class="panel cover-panel" data-panel="C">{cover}</div></section>
<section class="sheet inside">{panel("D")}{panel("E")}{panel("F")}</section>
<div id="flow">{soup}{venue}</div>
<script type="application/json" id="trifold-options">{order}</script>
<script>{TRIFOLD_SCRIPT}</script>
</body>
</html>
"""


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=HERE.parents[1], help="Folder containing index.md")
    parser.add_argument("--booklet", action="store_true", help="Use larger type and mirrored booklet margins")
    parser.add_argument("--trifold", action="store_true",
                        help="A4 landscape leaflet folded in three (outside and inside pages)")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        settings = json.loads((HERE / "settings.json").read_text(encoding="utf-8-sig"))
        if args.booklet and args.trifold:
            raise ValueError("Choose either --booklet or --trifold.")
        html = render(args.source.resolve(), settings, booklet=args.booklet, trifold=args.trifold)
        filename = ("programme-trifold.pdf" if args.trifold else
                    "programme-booklet.pdf" if args.booklet else "programme.pdf")
        output = (args.output or HERE / filename).resolve()
        # Publish the PDF only after the browser has produced a complete file.
        with tempfile.TemporaryDirectory(prefix="programme-", dir=output.parent) as temp:
            temp = Path(temp)
            document = temp / "programme.html"
            document.write_text(html, encoding="utf-8")
            pdf = temp / "programme.pdf"
            command = [
                str(find_browser()), "--headless", "--disable-gpu", "--no-first-run",
                "--no-pdf-header-footer", f"--user-data-dir={temp / 'profile'}",
                # time for the leaflet script to wait for the fonts and fit the panels
                "--virtual-time-budget=15000",
                f"--print-to-pdf={pdf}", document.as_uri(),
            ]
            result = subprocess.run(command, capture_output=True, text=True,
                                    errors="replace", timeout=90)
            if result.returncode or not pdf.is_file() or pdf.stat().st_size == 0:
                raise ValueError(f"Browser could not print the PDF.\n{result.stderr.strip()}")
            pdf.replace(output)
    except (OSError, ValueError, TypeError, KeyError, yaml.YAMLError, subprocess.TimeoutExpired) as error:
        parser.exit(1, f"Export failed: {error}\n")
    print(output)


if __name__ == "__main__":
    sys.stdout.reconfigure(errors="backslashreplace")
    main()
