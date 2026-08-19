"""Render part of a drawing sheet to PNG so its text can be read by eye.

The load diagrams this pipeline starts from are AutoCAD PDF exports whose text
has been converted to outlines: ``pdfplumber`` reports **zero** characters on
every page against seven to nine thousand curves, and ``pdftotext`` returns
nothing at all.  So no text extraction is possible -- the only way to read a
board name or a circuit description is to look at it.

An A1 sheet is 2384 x 1684 pt.  Rendered whole it is far too coarse to read, so
this tool renders a *fraction* of the page at a chosen scale.  Measured on
``01 BLD_Y1 Only Load Diagram.pdf``:

  * scale 2.0 over the whole sheet -- layout is legible, no text is
  * scale 2.6 over a quarter of the width (~190 DPI) -- every label reads clearly
  * scale 5.0 over a narrow strip -- reads breaker ratings and ditto marks

pypdfium2 does the rendering.  poppler is not used: this machine has
``pdftotext`` but no ``pdftoppm``, so anything built on poppler page rendering
would not run here.

Cable specs and load descriptions on these sheets are set rotated 90 degrees,
reading bottom to top, so ``--rotate cw`` is usually wanted for those strips.

    python pdfcrop.py sheet.pdf 2 --box 0.05 0.03 0.26 0.26 --scale 2.6
    python pdfcrop.py sheet.pdf 1 --box 0.05 0.02 0.11 0.20 --scale 9 --rotate cw
    python pdfcrop.py sheet.pdf 1 --grid 3 4 --scale 2.6 -o tiles
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pypdfium2 as pdfium

# Above this, a render is slower to produce than it is useful and the file gets
# unwieldy; the reader downsamples it again anyway.
MAX_PIXELS = 40_000_000


def render_box(
    pdf_path: str, page_no: int, box: tuple[float, float, float, float], scale: float
) -> "pdfium.PdfBitmap":
    """Render the fractional region ``box`` = (left, top, right, bottom).

    Fractions are of the page's width and height with the origin at the top
    left, which is how a region is described when reading off a screen.  pdfium
    wants margins to trim from each edge with the origin at the bottom left, so
    the two are converted here rather than at every call site.
    """
    left, top, right, bottom = box
    if not 0.0 <= left < right <= 1.0 or not 0.0 <= top < bottom <= 1.0:
        raise ValueError(f"box out of order or out of range: {box}")

    document = pdfium.PdfDocument(pdf_path)
    if not 1 <= page_no <= len(document):
        raise ValueError(f"page {page_no} out of range (1..{len(document)})")
    page = document[page_no - 1]
    width, height = page.get_width(), page.get_height()

    pixels = (right - left) * width * scale * (bottom - top) * height * scale
    if pixels > MAX_PIXELS:
        raise ValueError(
            f"that box at scale {scale} is {pixels / 1e6:.0f} megapixels; "
            f"narrow the box or lower the scale"
        )

    crop = (left * width, (1.0 - bottom) * height, (1.0 - right) * width, top * height)
    return page.render(scale=scale, crop=crop)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf")
    parser.add_argument("page", type=int, help="1-based page number")
    parser.add_argument(
        "--box", type=float, nargs=4, metavar=("L", "T", "R", "B"),
        default=[0.0, 0.0, 1.0, 1.0],
        help="fractional region of the page, origin top-left (default: whole page)",
    )
    parser.add_argument("--scale", type=float, default=2.6,
                        help="render scale; 2.6 is ~190 DPI on an A1 sheet")
    parser.add_argument("--grid", type=int, nargs=2, metavar=("ROWS", "COLS"),
                        help="split --box into tiles and render each one")
    parser.add_argument("--overlap", type=float, default=0.01,
                        help="fraction of a tile to extend on each side, so a "
                             "label sitting on a seam is whole in one of them")
    parser.add_argument("--rotate", choices=("cw", "ccw"), default=None,
                        help="turn the render so bottom-to-top text reads "
                             "left to right; 'cw' is the one these sheets need")
    parser.add_argument("-o", "--out", default="crop",
                        help="output PNG path, or directory prefix with --grid")
    args = parser.parse_args(argv)

    left, top, right, bottom = args.box
    turn = {None: 0, "cw": -90, "ccw": 90}[args.rotate]

    if not args.grid:
        out = Path(args.out).with_suffix(".png")
        out.parent.mkdir(parents=True, exist_ok=True)
        image = render_box(args.pdf, args.page, (left, top, right, bottom), args.scale).to_pil()
        if turn:
            image = image.rotate(turn, expand=True)
        image.save(out)
        print(f"{out}  {image.width}x{image.height}")
        return 0

    rows, cols = args.grid
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    tile_w = (right - left) / cols
    tile_h = (bottom - top) / rows
    for row in range(rows):
        for col in range(cols):
            box = (
                max(0.0, left + col * tile_w - args.overlap * tile_w),
                max(0.0, top + row * tile_h - args.overlap * tile_h),
                min(1.0, left + (col + 1) * tile_w + args.overlap * tile_w),
                min(1.0, top + (row + 1) * tile_h + args.overlap * tile_h),
            )
            path = out_dir / f"p{args.page}_r{row + 1}c{col + 1}.png"
            image = render_box(args.pdf, args.page, box, args.scale).to_pil()
            if turn:
                image = image.rotate(turn, expand=True)
            image.save(path)
            print(f"{path}  {image.width}x{image.height}  box={[round(v, 4) for v in box]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
