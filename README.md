# MCNP_outp_viewer

Interactive viewer for MCNP `outp` text files built with PySide6.

## Features

- Open one or more MCNP `outp` text files
- Browse detected tallies in a sortable/filterable list
- Inspect simple tallies, FS-segment tallies, and energy-bin tallies
- Apply multiplier and additive offset to extracted tally values
- Export the current table to CSV
- Export multiple checked tallies to a single Excel workbook

## Requirements

- Python 3.10+
- PySide6
- pandas
- openpyxl

Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python mcnp_outp_viewer.py
```

## Notes

- This tool uses text-pattern parsing because MCNP `outp` files are plain text.
- It works best when the output file contains echoed input cards.
