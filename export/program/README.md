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

## Three-fold leaflet

Run `python .\export_programme.py --trifold` to save `programme-trifold.pdf`:
two A4 landscape pages, each divided into three panels of 99 mm.

- Page 1 is the outside: inside flap | back | front cover.
- Page 2 is the inside: three panels.

Print both pages on one sheet, double-sided, **flip on short edge**, at
100 % (no "fit to page"). With the outside facing you, fold the left third
(the inside flap) in first, then fold the cover over it.

The programme runs through the inside panels and then the inside flap, and the
type size is reduced (from 9 pt down to 6 pt) until it fits. Adjust this in the
`trifold` section of `settings.json`:

- `order` — the panels the programme flows through: `D`, `E`, `F` are the
  inside panels from left to right, `A` is the inside flap and `B` the back.
  Add `"B"` if the programme does not fit at a readable size.
- `max_font_pt`, `min_font_pt` — largest and smallest type size to try. If it
  still does not fit at the smallest size, an extra page is added so that
  nothing is lost.
- The cover is laid out like the poster: `cover_logo`, `cover_label`,
  `cover_title` (one entry per line), `cover_image` (the arch drawing),
  `cover_qr` (links to https://www.archetai.gr/poleisinromangreece/),
  `cover_background` and `cover_image_background`. The images are the
  `cover-*.png` files in this folder; leave a setting empty to drop that item.

Chair names, the map and the other settings above apply to the leaflet too.

## Three-column design preview

The `design-preview` folder has its own copy of `settings.json`. From that
folder, use `python .\export_programme.py` for the standard preview and
`python .\export_programme.py --booklet` for the booklet preview.
