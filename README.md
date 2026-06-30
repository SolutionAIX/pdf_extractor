# PDF → Excel extractor (MOVE logistics-center brochures)

Reads MOVE logistics-center property PDFs and produces one Excel row per
**property page**, auto-filling these columns:

`사용가능여부 · 중개인/임대인명 · 창고명 · 주소 · 행정구역_도 · 행정구역_시 ·
대지면적 · 연면적 · 건축면적 · 준공연도 · 공실현황 · 시설특이사항 · 기타 ·
정보확인일자 · 사진`

The tool handles **two kinds of PDF automatically**, per page:

1. **Text-layer PDFs** (e.g. the full `dsv_main.pdf` catalog) — values are read
   straight from the embedded text, so the data is exact (no OCR error).
2. **Image-only / scanned PDFs** (e.g. `dsvtest2.pdf`) — values are recovered
   with OCR (Tesseract, Korean + English).

**Non-property pages are skipped automatically.** A page only becomes a row if
it looks like a property sheet (has the `General Information` + `Space
Availability` tables, an address, and an `N.` title). Cover pages, the table of
contents and region dividers are ignored — e.g. `dsv_main.pdf` yields 296 rows
from 328 pages.

The two image columns — **공실현황** (Space Availability table) and **사진**
(building photo) — are cropped from the page and embedded directly into the
cells.

## Install

```bash
# system OCR engine + Korean language data (macOS)
brew install tesseract tesseract-lang

# python deps (a venv already exists in ./venv)
./venv/bin/pip install -r requirements.txt
```

## Usage

```bash
./venv/bin/python main.py dsv_main.pdf -o dsv_main.xlsx --broker S1   # 296 rows
./venv/bin/python main.py dsvtest2.pdf                                # OCR path
./venv/bin/python main.py input.pdf --date 01/07/2026
```

Options: `-o/--output`, `--broker` (fills 중개인/임대인명), `--date dd/mm/yyyy`
(default today), `--dpi` (OCR render DPI, default 300), `--img-dpi` (cropped
image DPI, default 150), `--img-width` (embedded image width px, default 360).

### Web UI (drag & drop)

A Streamlit app wraps the same extractor: drag-drop a PDF, preview the result
table and cropped images, then download the `.xlsx`.

```bash
./venv/bin/streamlit run app.py
```

Then open the shown URL (default http://localhost:8501). Set 중개인/임대인명,
정보 확인일자 and image sizes in the sidebar before/while uploading.

| Column | Source / default |
| --- | --- |
| 창고명 | title headline (OCR) |
| 주소 / 행정구역_도 / 행정구역_시 | `소재지` row, address parsed into province + city |
| 대지면적 / 연면적 | General-Information rows |
| 준공연도 | `준공년도` row, default **미정** |
| 공실현황 | cropped **Space Availability** table image |
| 시설특이사항 | `비고`/SPEC middle column (OCR) |
| 정보확인일자 | `--date`, default **today (dd/mm/yyyy)** |
| 사진 | cropped top-left building photo |
| 사용가능여부 / 중개인/임대인명 / 건축면적 / 기타 | left blank (set `--broker` to fill 중개인/임대인명) |

Cropped images are also saved alongside the workbook in
`<output>_images/`.

## Notes / limitations

- For text-layer PDFs the extracted values are exact. OCR is only used for
  scanned/image pages, where bold headlines are ~90% accurate — spot-check 창고명.
- The layout constants in `main.py` are tuned to the MOVE template (the OCR
  `*_BOX` fractions and the text-mode crop anchors). A different brochure layout
  may need those adjusted.
- A row is created per property page; the workbook embeds 2 images per row, so a
  large catalog produces a large `.xlsx` (≈90 MB for 296 rows). Lower `--img-dpi`
  or `--img-width` to shrink it.
