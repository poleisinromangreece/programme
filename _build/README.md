# Programme build

The programme page (`index.md`) is generated from the Excel programme
(`Programma_synedriou_v2.xlsx`). The Excel is the only source for order, times,
speakers and titles.

Rebuild after any change to the Excel (from any folder):

```
python build.py
```

Without arguments it uses the newest `Programma_synedriou*.xlsx` found in the
`markdown` folder or the folders above it (Abstracts, the conference folder).
To use a specific file:

```
python build.py "D:\path\to\Programma_synedriou_v3.xlsx"
```

(requires Python 3 with `openpyxl`, which Anaconda already includes)

What it does:

- rewrites `index.md` (days, sessions, times, breaks, linked titles);
- updates `title`, `authors` and the "back to programme" link of every file in
  `abstracts/` (the abstract text itself is never changed);
- warns about talks with no abstract file, duplicated speakers, and abstract
  files no longer in the programme.

Chairs: a row with `Πρόεδρος:` in column D (Speakers) is shown as a
"Πρόεδρος: …" line, with the name taken from column E. Leave E empty
to show the line with no name.

Settings at the top of `build.py`: dates of the three days, page intro, URL
pattern of the abstracts, and the speaker → abstract-file table (`SLUGS`).
A new speaker needs a new line in `SLUGS` and a new file in `abstracts/`.
Affiliations are kept in the `affiliation:` field of each abstract and are
added to `authors` in parentheses.
