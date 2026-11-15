# Conference book of abstracts

Run in PowerShell from the `export/abstracts-book` folder:

```powershell
python .\export_abstracts.py
```

The script reads the programme order from `../../index.md`, conference details
from `../../_config.yml`, and the linked Markdown abstracts from `../../abstracts`.
When both «Περίληψη» and «Abstract» are present, it prints the version that
matches the title's language. Affiliations and bibliographies are included
when present. Programme items with a speaker and title but no abstract link
are listed in the contents. The script reads the source files without editing them.

The browser finder and Society artwork are shared with `../program`.

- Edit abstract wording in the corresponding Markdown file.
- Use the existing Excel programme build to change the programme order or the
  titles and names that it synchronises.
- Edit `abstracts-book.css` for type sizes, colours, margins and cover styling.
- Run `python .\export_abstracts.py --output .\abstracts-proof.pdf` to save
  under another filename.

The exporter needs Python, Microsoft Edge or Chromium, Brill installed locally,
and the Python packages below. The PDF embeds the fonts used by the browser.

```powershell
python -m pip install beautifulsoup4 markdown PyMuPDF PyYAML
```
