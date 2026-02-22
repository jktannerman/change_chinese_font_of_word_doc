# Chinese Font Changer

Applies an East Asian font to every Chinese character in a `.docx` file, without touching the fonts on any Latin or other text. Covers body paragraphs, tables (including nested), headers, footers, and text boxes.

## Requirements

- Python 3.13
- `python-docx`

```
py -3.13 -m pip install -r requirements.txt
```

## Usage

### GUI (recommended for non-technical users)

Double-click **`launch.bat`** to open the window.

1. Click **Browse…** and select your `.docx` file.
2. Choose a font from the dropdown, or select **Other…** to type a custom font name.
3. Click **Convert**.

The output is saved as `<original name>_modified.docx` in the same folder. You'll be offered the option to open that folder in Explorer when done.

**Tip:** You can also drag a `.docx` file directly onto `launch.bat` to pre-fill the file path.

### Command line

```
py -3.13 change_chinese_font.py <input.docx> [--font <font_name>] [--output <out.docx>]
```

| Argument | Default | Description |
|---|---|---|
| `input` | *(required)* | Path to the source `.docx` file |
| `--font` | `FangSong` | East Asian font name to apply |
| `--output` | `<input>_modified.docx` | Path for the output file |

**Examples:**

```
py -3.13 change_chinese_font.py report.docx
py -3.13 change_chinese_font.py report.docx --font SimSun
py -3.13 change_chinese_font.py report.docx --font KaiTi --output report_final.docx
```

## How it works

Word's OOXML format stores font information per character category. Setting only the `w:eastAsia` font attribute on a run applies the chosen font exclusively to CJK characters in that run, leaving the Latin font (and all other formatting) untouched.

Chinese characters are detected using a regex covering all relevant Unicode blocks: CJK Unified Ideographs, Extensions A–F, Compatibility Ideographs, CJK Symbols & Punctuation, and Fullwidth Forms.
