"""Rasterize existing native Excel/PPT PDFs; never opens Office or rewrites Word."""
import argparse, hashlib, json, struct, sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '.tmp/office-libs'))
try:
    import pymupdf
except ImportError as exc:
    raise SystemExit('Install local dependency: python -m pip install --target .tmp/office-libs PyMuPDF') from exc
def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--width', type=int, default=1600)
    parser.add_argument('--preset', nargs='+', choices=['excel-base-styles', 'ppt-base-styles'], default=['excel-base-styles', 'ppt-base-styles'])
    args = parser.parse_args()
    if not 800 <= args.width <= 2400:
        parser.error('--width must be between 800 and 2400')
    atlas = ROOT / 'docs/style-atlas'
    expected = {'excel-base-styles': 2, 'ppt-base-styles': 10}
    files = []
    for preset in dict.fromkeys(args.preset):
        source = atlas / (preset + '.pdf')
        destination = atlas / preset
        destination.mkdir(exist_ok=True)
        images = []
        with pymupdf.open(source) as pdf:
            if pdf.page_count != expected[preset]:
                raise ValueError(f'{preset}: expected {expected[preset]} pages, found {pdf.page_count}')
            for index, page in enumerate(pdf):
                rect = page.rect
                scale = args.width / rect.width
                pixmap = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), colorspace=pymupdf.csRGB, alpha=False)
                height = round(args.width * rect.height / rect.width)
                if pixmap.width != args.width:
                    pixmap = pymupdf.Pixmap(pixmap, args.width, height)
                path = destination / f'page-{index + 1:02d}.png'
                data = pixmap.tobytes('png')
                if len(data) <= 24 or data[:8] != b'\x89PNG\r\n\x1a\n':
                    raise ValueError(f'Invalid or empty PNG: {path}')
                width, actual_height = struct.unpack('>II', data[16:24])
                if width != args.width or abs(actual_height - height) > 1:
                    raise ValueError(f'Unexpected raster size: {width}x{actual_height}')
                path.write_bytes(data)
                if path.stat().st_size != len(data):
                    raise ValueError(f'PNG write incomplete: {path}')
                images.append({'page': index + 1, 'path': str(path), 'width': width, 'height': actual_height, 'requestedWidth': args.width, 'bytes': len(data), 'pdfWidthPoints': rect.width, 'pdfHeightPoints': rect.height, 'orientation': 'landscape' if rect.width > rect.height else 'portrait'})
            files.append({'presetId': preset, 'pdf': str(source), 'pageCount': pdf.page_count, 'renderer': 'PyMuPDF', 'pdfSHA256': hashlib.sha256(source.read_bytes()).hexdigest(), 'images': images})
        print('Rendered', preset, 'pages', len(images), 'sizes', [(i['width'], i['height'], i['bytes']) for i in images])
    report = {'scope': 'native-pdf-rendering', 'renderer': 'PyMuPDF', 'rendererVersion': pymupdf.VersionBind, 'files': files}
    (atlas / 'pdf-render-evidence.json').write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
if __name__ == '__main__':
    main()