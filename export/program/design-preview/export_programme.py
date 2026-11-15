"""Export the current website programme to an A4 PDF."""
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
    for name in ("msedge", "chrome", "chromium"):
        if path := shutil.which(name):
            return Path(path)
    raise ValueError("Install Microsoft Edge or Google Chrome to print the PDF.")


def render(root, settings, booklet=False):
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

        widths = table.select("th")
        widths[1].decompose()
        for cell, width in zip((widths[0], *widths[2:]), (19, 24, 57)):
            cell["style"] = f"width: {width}%"
        for row in table.select("tr"):
            cells = row.find_all("td", recursive=False)
            if len(cells) == 4:
                start, end = (cell.get_text(strip=True) for cell in cells[:2])
                cells[0].string = "–".join(filter(None, (start, end)))
                cells[0]["class"] = ["time-range"]
                cells[1].decompose()
            elif len(cells) == 1:
                cells[0]["colspan"] = "3"

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


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=HERE.parents[2], help="Folder containing index.md")
    parser.add_argument("--booklet", action="store_true", help="Use larger type and mirrored booklet margins")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        settings = json.loads((HERE / "settings.json").read_text(encoding="utf-8-sig"))
        html = render(args.source.resolve(), settings, booklet=args.booklet)
        filename = "programme-design-preview.pdf" if args.booklet else "programme-design-preview-standard.pdf"
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
