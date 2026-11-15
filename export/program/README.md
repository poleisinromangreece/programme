# Programme PDF export

Open PowerShell in this folder and run:

```powershell
python .\export_programme.py
```

This creates or replaces `programme.pdf` here. You need Python 3.10 or later,
Microsoft Edge or Google Chrome, and the Python packages below. If the script
reports a missing Python package, install the dependencies once:

```powershell
python -m pip install beautifulsoup4 PyYAML
```

## Edit the PDF

Open `settings.json` in a text editor, save your changes, and run the script again.

- **Πρόεδροι** — type each name inside the empty quotes beside the day and
  start time of that chair's first talk. For example, replace
  `"day-1 09:15": ""` with `"day-1 09:15": "Όνομα Επώνυμο"`.
  Blank settings use the programme's existing names when available.
  Set `hide_empty_chairs` to `false` to print empty «Πρόεδρος:» lines.
  Update the time in the key if the first talk in a chaired block is rescheduled.
- **Map** — set `show_map` to `true` to add a venue section with a clickable
  map link. For a map that appears on paper, save an image in this folder and
  put its filename in `map_image`. The conference address appears in the header.
- **Day buttons** — set `show_day_buttons` to `true` to print the buttons.
  The dates remain as headings when the buttons are hidden.
- **Logos, introduction, and page breaks** — use `show_logos`, `show_intro`,
  and `start_each_day_on_new_page`. Use `print.css` for fonts, colours, margins,
  and spacing.
- **Abstract links** — put the published programme's base URL in
  `abstract_base_url` to make talk titles clickable.

The script reads the current `index.md` and `_config.yml`. For changes to talks,
times, or speakers, rebuild the website programme from Excel as usual before
exporting again. The export reads the website source files without editing them.

To save under another filename, run
`python .\export_programme.py --output .\programme-review.pdf`.

## Booklet edition

Run `python .\export_programme.py --booklet` to save `programme-booklet.pdf`.
This edition uses 15 pt Brill for the main programme with mirrored 15 mm
inside and 10 mm outside margins. The PDF is in reading order. Use the PDF
viewer's booklet print option with A3 paper to keep the finished pages at A4.

## Three-column design preview

The `design-preview` folder has its own copy of `settings.json`. From that
folder, use `python .\export_programme.py` for the standard preview and
`python .\export_programme.py --booklet` for the booklet preview.
