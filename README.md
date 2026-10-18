# Οι πόλεις της Ελλάδας από τον Αύγουστο έως την Τετραρχία

Website for the programme and abstracts of the conference
«Οι πόλεις της Ελλάδας από τον Αύγουστο έως την Τετραρχία: Πολιτισμός – κοινωνία – θεσμοί»
(Εν Αθήναις Αρχαιολογική Εταιρεία, Αθήνα, 3–5 Δεκεμβρίου 2026).

Built with Jekyll and published with GitHub Pages.

## Structure

```
_config.yml          site title, organiser, dates
_layouts/            page templates (default, page, abstract)
assets/css/style.css styles
index.md             programme (GENERATED from the Excel – do not edit by hand)
abstracts/*.md       one page per paper
_build/build.py      regenerates index.md from the Excel (not published)
```

## Updating the programme

1. Edit the Excel (`Programma_synedriou_v2.xlsx`).
2. In `_build`, run `python build.py` (see `_build/README.md`).
3. Upload/commit the changed `index.md` and `abstracts/*.md`.

## Editing an abstract

Edit the text in `abstracts/<Name>.md` below the front matter. The `title` and
`authors` lines come from the Excel and are overwritten by `build.py`; change
affiliations in the `affiliation:` line.


