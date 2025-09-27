# Yu-Gi-Oh! Excuses Bingo Generator

This Python script generates printable 5x5 bingo boards for Yu-Gi-Oh! tournaments, tracking common excuses players give after losing. It supports both PDF (print-ready) and Excel (editable) output formats.

## Features

- Generates multiple bingo boards at once.
- PDF export includes a uniform header and one board per page.
- Excel export preserves wrapped text and formatting.
- Each board contains a "FREE SPACE" in the center.
- Configurable number of boards via command line.

## Requirements

Install the required Python modules:

```pip install -r requirements.txt```

- `openpyxl` – for Excel export  
- `reportlab` – for PDF export

## Usage

Default: generate 5 PDF boards:

```python ygo_excuse_bingo.py```

Generate 10 PDF boards:

```python ygo_excuse_bingo.py --num 10```

Generate Excel boards only:

```python ygo_excuse_bingo.py --excel```

Generate both PDF and Excel boards:

```python ygo_excuse_bingo.py --num 20 --both```

## Disclaimer

I produced this almost entirely with ChatGPT as I was feeling lazy. So if it doesn't work, blame ChatGPT.

ChatGPT itself produced the following disclaimer when I asked it to add a disclaimer:

This script and its contents were produced by ChatGPT (OpenAI) and are intended for personal or entertainment use only.
