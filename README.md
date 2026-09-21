# DDX'26 San Diego — Conference Report

Static single page, DDX chrome (black / white / #FFF204, Helvetica Neue stack, Tinos italic for quotes).
Built 2026-09-20 from four sources:

- **Photos** — Dropbox folder "DDX 2026 - San Diego" (link in the 2026-09-19 "What a Day!" attendee email), 264 files by Desmond Chua & Nicole Gosé. 17 exported to `assets/img/photos/` at ≤1400px (hero ≤2200px). Full set kept only in the session scratchpad.
- **Speaker quotes** — Notion "Speakers, VIPs & Jury Profiles (Step 1)", property "Quote - The Future of Innovation", `Which Event? = San Diego`. 15 quotes used verbatim (Brian LeDuc's row says "just updating headshot" and was skipped). Faces reused from `../ddx-agenda-sandiego/assets/img/speakers/`; `harriet-swan.png` is the left half of `harriet-katie.webp`.
- **Programme** — `../ddx-agenda-sandiego/assets/js/agenda.js` (mirrors ddxconference.com/sd-agenda + Notion Sessions Step 3). Website copy (theme "The Inflection Point", 25 speakers / 6 workshops / 3 panels, 3,000+ alumni) from ddxconference.com/sandiego.
- **2026 statistics** — see the Sources list at the foot of the page; every figure has its publication month and URL.

Open `index.html` directly; no build step. Deploy the folder as-is (GitHub Pages: add `.nojekyll`).

**To confirm before sending out:** the "300+ leaders registered" tile (Luma showed 319 approved on 14 Sep; replace with the final checked-in count), and the London/Dubai dates, which are still "late 2026 / early 2027" on the public site.

## Recap deck (PDF)

`deck.html` is a 16-slide, 1920×1080 version of the report; `DDX26-San-Diego-Recap.pdf` is its render.
Rebuild after editing `deck.html`:

```bash
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --disable-gpu --no-pdf-header-footer --virtual-time-budget=15000 --print-to-pdf="$PWD/DDX26-San-Diego-Recap.pdf" "file://$PWD/deck.html"
```

`@page{size:1920px 1080px;margin:0}` gives 1440×810 pt pages (exact 16:9). No poppler on this Mac: preview pages with the PDFKit script in the session scratchpad (`pdfrender.swift`).

**Partner logos** (`assets/img/partners/`): sliced from the two white-on-transparent strips on ddxconference.com/sandiego (`Group+2147203277+(1).png`, `Group+2147203290.png`, fetched at `?format=2500w`, delivered as WebP) with the scratchpad `slice.swift` (segments by transparent columns, gap ≥ 40px). The strip's last mark is a cropped Mondo Robot; the real logo is `mondo-robot.svg` from framerusercontent. No logo artwork exists for Don Norman Design Award, J.S. Held, Edenspiekermann, Appsfactory, Friends of Figma, SDXD or UX Speakeasy — they stay as text on slide 15.

## PowerPoint version

`build-pptx.js` (pptxgenjs, LAYOUT_WIDE 13.33×7.5 in, Arial + Times New Roman so it renders the same on every machine) rebuilds `DDX26-San-Diego-Recap.pptx` from the same content as `deck.html`. It expects an `img/` folder next to it holding PNG/JPG copies of every asset (speaker `.webp` files converted with `sips -s format png`; `mondo-robot.svg` rasterised the same way). Run `npm install pptxgenjs` in that folder, then `node build-pptx.js out.pptx`. Validate with the pptx skill's `validate.py` under Python 3.13 (the system python3 is too old for its `match` statements). Keynote automation was declined on this Mac, so the PPTX was checked structurally (validator + every shape inside the slide bounds), not visually; the PDF is the visually verified version.

## Recap carousel (social)

`carousel.html` → `carousel/` holds the 10-slide 1080×1350 (4:5) recap: `DDX26-San-Diego-Recap-Carousel.pdf` for a LinkedIn document post, `DDX26-SD-Recap-01…10.jpg` for Instagram (zipped too), `caption.md` with LinkedIn and Instagram copy + tag list. Same Chrome `--print-to-pdf` route with `@page{size:1080px 1350px}`; JPGs rendered from the PDF at 1080px with the PDFKit script and cropped to exactly 1350px tall (PDFKit rounds the page to 1351).
