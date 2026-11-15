"""Build the conference's A4 book of abstracts from the website sources."""
import argparse
from html import escape
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import unicodedata

from bs4 import BeautifulSoup
import markdown
import pymupdf
import yaml

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
PROGRAM = HERE.parent / "program"
sys.path.insert(0, str(PROGRAM))
from export_programme import find_browser


MARKER = re.compile(r"ΠΕΡΙΛΗΨΗ\s+(\d+)")
BACK_LINK = re.compile(r"(?m)^---[ \t]*\n\s*\[Επιστροφή στο πρόγραμμα[^\n]*\n?\s*\Z")
ABSTRACT_HEADINGS = {"περίληψη": "GREEK", "abstract": "LATIN"}


def read_markdown(path):
    text = path.read_text(encoding="utf-8-sig")
    opening, metadata, body = re.split(r"(?m)^---[ \t]*$", text, maxsplit=2)
    if opening.strip():
        raise ValueError(f"Expected YAML front matter in {path}")
    return yaml.safe_load(metadata), body.strip()


def render_abstract(content, title):
    html = markdown.markdown(content)
    soup = BeautifulSoup(html, "html.parser")
    sections = [
        (heading, ABSTRACT_HEADINGS.get(heading.get_text(strip=True).casefold()))
        for heading in soup.find_all("h2", recursive=False)
    ]
    if {language for _, language in sections if language} != {"GREEK", "LATIN"}:
        return html

    scripts = [unicodedata.name(char, "").split(" ", 1)[0]
               for char in title if char.isalpha()]
    greek, latin = scripts.count("GREEK"), scripts.count("LATIN")
    if greek == latin:
        raise ValueError(f"Could not determine the language of the title: {title}")
    keep = "GREEK" if greek > latin else "LATIN"
    for heading, language in sections:
        if language and language != keep:
            for sibling in heading.find_next_siblings():
                if sibling.name == "h2":
                    break
                sibling.decompose()
            heading.decompose()
    return str(soup)


def read_programme(root):
    _, body = read_markdown(root / "index.md")
    soup = BeautifulSoup(body, "html.parser")
    days = []
    missing = []
    seen = set()
    for heading in soup.select('h2[id^="day-"]'):
        label = heading.get_text(strip=True)
        papers = []
        session = ""
        for row in heading.find_next_sibling("table").select("tr"):
            if "session-start" in row.get("class", []):
                session = row.get_text(" ", strip=True)
                continue
            link = row.select_one('a[href*="/abstracts/"]')
            if link:
                slug = Path(link["href"]).stem
                if slug in seen:
                    raise ValueError(f"Abstract listed twice: {slug}")
                seen.add(slug)
                meta, source = read_markdown(root / "abstracts" / f"{slug}.md")
                content = BACK_LINK.sub("", source).strip()
                if not content:
                    raise ValueError(f"Abstract file is empty: {slug}")
                papers.append({
                    "number": len(seen),
                    "title": meta["title"],
                    "authors": meta["authors"],
                    "affiliation": meta.get("affiliation") or "",
                    "body": render_abstract(content, meta["title"]),
                    "pending": content.splitlines()[-1].strip("* ") == "Η περίληψη θα αναρτηθεί σύντομα.",
                    "session": session,
                })
                continue
            cells = row.find_all("td", recursive=False)
            if len(cells) == 4:
                author, title = (cell.get_text(" ", strip=True) for cell in cells[2:])
                if author and title:
                    missing.append(f"{author} — {title}")
        if papers:
            days.append((label, papers))
    if not days:
        raise ValueError("No linked abstracts found in index.md.")
    return days, missing


def render(root, site, days, missing, page_numbers):
    title = escape(site["title"])
    cover_title = title.replace(" από τον ", "<br>από τον ", 1)
    sections = []
    articles = []
    for day, papers in days:
        entries = []
        last_session = None
        for paper in papers:
            number = paper["number"]
            session = paper["session"]
            authors = escape(paper["authors"])
            paper_title = escape(paper["title"])
            if session != last_session and session:
                entries.append(f'<li class="toc-session">{escape(session)}</li>')
            last_session = session
            page = page_numbers.get(number, "–")
            status = ' <small class="pending">· Εκκρεμεί</small>' if paper["pending"] else ""
            entries.append(f"""<li><a class="entry" href="#abstract-{number}">
<span><strong>{authors}</strong>{status}<br>{paper_title}</span>
<span class="entry-page">{page}</span></a></li>""")
            affiliation = paper["affiliation"]
            affiliation_html = (
                f'<p class="affiliation">{escape(affiliation)}</p>' if affiliation else ""
            )
            articles.append(f"""<article class="abstract" id="abstract-{number}">
<header class="article-header">
<p class="kicker">ΠΕΡΙΛΗΨΗ {number:02}</p>
<p class="day">{escape(day)}</p>
<p class="session">{escape(session)}</p>
<h2>{paper_title}</h2>
<p class="authors">{authors}</p>
{affiliation_html}
</header>
<div class="abstract-body">{paper["body"]}</div>
</article>""")
        sections.append(f"""<section class="contents">
<h2>Περιεχόμενα</h2>
<p class="contents-day">{escape(day)}</p>
<ul>{"".join(entries)}</ul>
</section>""")
    if missing:
        note = "<br>".join(escape(item) for item in missing)
        sections[-1] = sections[-1].replace(
            "</section>",
            f'<p class="availability">Χωρίς διαθέσιμη περίληψη<br>{note}</p></section>',
        )
    seal = (PROGRAM / "society-logo-original.png").as_uri()
    arch = (root / "assets/images/arch_logo_dark-gray.svg").as_uri()
    css = (HERE / "abstracts-book.css").as_uri()
    return f"""<!doctype html>
<html lang="el">
<head>
<meta charset="utf-8">
<title>Βιβλίο περιλήψεων · {title}</title>
<link rel="stylesheet" href="{css}">
</head>
<body>
<section class="cover">
<header class="society">
<div class="society-seal"><img src="{seal}" alt=""></div>
<div class="society-wordmark"><img src="{seal}" alt="{escape(site['organiser'])} · έτος ίδρυσης 1837"></div>
</header>
<div class="cover-arch">
<svg class="construction" viewBox="0 0 277 108" preserveAspectRatio="none" aria-hidden="true">
      <g fill="none" stroke="#dcae94" stroke-width=".25" opacity=".48">
        <path d="M8 9H269V99H8Z M8 54H269 M138.5 9V99" stroke-dasharray="2 2"/>
        <path d="M66 99V83a72.5 72.5 0 0 1 145 0v16 M77 99V83a61.5 61.5 0 0 1 123 0v16 M88 99V83a50.5 50.5 0 0 1 101 0v16"/>
        <path d="M8 83H269 M24 9V99 M253 9V99"/>
      </g>
      <g fill="#dcae94" opacity=".8">
        <circle cx="24" cy="54" r=".9"/><circle cx="253" cy="54" r=".9"/>
      </g>
    </svg>
<img src="{arch}" alt="">
</div>
<div class="cover-title">
<p class="kicker">ΣΥΝΕΔΡΙΟ</p>
<h1>{cover_title}</h1>
<p class="subtitle">{escape(site["subtitle"])}</p>
<p class="book-label">Βιβλίο περιλήψεων</p>
</div>
<div class="cover-venue">
<p><strong>{escape(site["place_date"])}</strong></p>
<p>{escape(site["venue_name"])}</p>
<p>{escape(site["venue_address"])}</p>
</div>
</section>
{"".join(sections)}
{"".join(articles)}
</body>
</html>"""


def print_pdf(browser, work, html, filename):
    document = work / "abstracts.html"
    document.write_text(html, encoding="utf-8")
    pdf = work / filename
    command = [
        str(browser), "--headless", "--disable-gpu", "--no-first-run",
        "--no-pdf-header-footer", f"--user-data-dir={work / 'profile'}",
        f"--print-to-pdf={pdf}", document.as_uri(),
    ]
    result = subprocess.run(command, capture_output=True, text=True,
                            errors="replace", timeout=120)
    if result.returncode or not pdf.is_file() or not pdf.stat().st_size:
        raise ValueError(f"Browser could not print the PDF.\n{result.stderr.strip()}")
    return pdf


def abstract_pages(doc):
    return {
        int(number): page.number + 1
        for page in doc
        for number in MARKER.findall(page.get_text())
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=HERE / "abstracts-book.pdf")
    args = parser.parse_args()
    try:
        site = yaml.safe_load((ROOT / "_config.yml").read_text(encoding="utf-8-sig"))
        days, missing = read_programme(ROOT)
        count = sum(len(papers) for _, papers in days)
        output = args.output.resolve()
        browser = find_browser()
        with tempfile.TemporaryDirectory(prefix="abstracts-", dir=output.parent) as temp:
            work = Path(temp)
            draft = print_pdf(browser, work, render(ROOT, site, days, missing, {}), "draft.pdf")
            with pymupdf.open(draft) as doc:
                pages = abstract_pages(doc)
            if set(pages) != set(range(1, count + 1)):
                raise ValueError("Could not locate every abstract in the rendered PDF.")
            numbered = print_pdf(
                browser, work, render(ROOT, site, days, missing, pages), "numbered.pdf"
            )
            with pymupdf.open(numbered) as doc:
                if abstract_pages(doc) != pages:
                    raise ValueError("Contents pagination changed when page numbers were added.")
                outline = [[1, "Περιεχόμενα", 2]]
                for day, papers in days:
                    outline.append([1, day, pages[papers[0]["number"]]])
                    for paper in papers:
                        label = f"{paper['authors']} · {paper['title']}"
                        outline.append([2, label, pages[paper["number"]]])
                doc.set_toc(outline)
                doc.save(work / "final.pdf", deflate=True)
            (work / "final.pdf").replace(output)
        pending = sum(paper["pending"] for _, papers in days for paper in papers)
        print(f"{count - pending} abstracts and {pending} pending entries exported to {output}")
        for item in missing:
            print(f"No linked abstract: {item}")
    except (OSError, ValueError, TypeError, KeyError, yaml.YAMLError, subprocess.TimeoutExpired) as error:
        parser.exit(1, f"Export failed: {error}\n")


if __name__ == "__main__":
    sys.stdout.reconfigure(errors="backslashreplace")
    main()
