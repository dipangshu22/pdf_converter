"""
font_manager.py
Resolves the right TTF/OTF font for a given target language.
Bundled DejaVuSans covers Latin/Cyrillic/Greek.
All other scripts are downloaded from Google Fonts (Noto) on first use
and cached in the project's fonts/ folder.
"""

import logging
import urllib.request
from pathlib import Path

logger = logging.getLogger(__name__)

FONTS_DIR = Path(__file__).parent / 'fonts'

# ── Remote font catalogue ────────────────────────────────────────────────────
# Maps local filename → stable download URL (Noto fonts via GitHub raw)
_BASE = 'https://raw.githubusercontent.com/googlefonts/noto-fonts/main/hinted/ttf'
_CJK  = 'https://raw.githubusercontent.com/googlefonts/noto-cjk/main/Sans/SubsetOTF'

FONT_DOWNLOADS: dict[str, str] = {
    # Indic
    'NotoSansDevanagari-Regular.ttf': f'{_BASE}/NotoSansDevanagari/NotoSansDevanagari-Regular.ttf',
    'NotoSansBengali-Regular.ttf':    f'{_BASE}/NotoSansBengali/NotoSansBengali-Regular.ttf',
    'NotoSansGujarati-Regular.ttf':   f'{_BASE}/NotoSansGujarati/NotoSansGujarati-Regular.ttf',
    'NotoSansGurmukhi-Regular.ttf':   f'{_BASE}/NotoSansGurmukhi/NotoSansGurmukhi-Regular.ttf',
    'NotoSansTamil-Regular.ttf':      f'{_BASE}/NotoSansTamil/NotoSansTamil-Regular.ttf',
    'NotoSansTelugu-Regular.ttf':     f'{_BASE}/NotoSansTelugu/NotoSansTelugu-Regular.ttf',
    'NotoSansKannada-Regular.ttf':    f'{_BASE}/NotoSansKannada/NotoSansKannada-Regular.ttf',
    'NotoSansMalayalam-Regular.ttf':  f'{_BASE}/NotoSansMalayalam/NotoSansMalayalam-Regular.ttf',
    'NotoSansSinhala-Regular.ttf':    f'{_BASE}/NotoSansSinhala/NotoSansSinhala-Regular.ttf',
    # Arabic-script
    'NotoSansArabic-Regular.ttf':     f'{_BASE}/NotoSansArabic/NotoSansArabic-Regular.ttf',
    # South-East Asian
    'NotoSansThai-Regular.ttf':       f'{_BASE}/NotoSansThai/NotoSansThai-Regular.ttf',
    'NotoSansKhmer-Regular.ttf':      f'{_BASE}/NotoSansKhmer/NotoSansKhmer-Regular.ttf',
    'NotoSansMyanmarRegular.ttf':     f'{_BASE}/NotoSansMyanmar/NotoSansMyanmar-Regular.ttf',
    'NotoSansLao-Regular.ttf':        f'{_BASE}/NotoSansLao/NotoSansLao-Regular.ttf',
    # CJK (subset OTF — ~5-8 MB each)
    'NotoSansSC-Regular.otf':         f'{_CJK}/SC/NotoSansSC-Regular.otf',
    'NotoSansTC-Regular.otf':         f'{_CJK}/TC/NotoSansTC-Regular.otf',
    'NotoSansJP-Regular.otf':         f'{_CJK}/JP/NotoSansJP-Regular.otf',
    'NotoSansKR-Regular.otf':         f'{_CJK}/KR/NotoSansKR-Regular.otf',
    # Mongolian / Georgian / Armenian use DejaVu — already bundled
}

# ── Language → font file mapping ─────────────────────────────────────────────
LANGUAGE_FONT: dict[str, str] = {
    # Indic — Devanagari
    'hindi':      'NotoSansDevanagari-Regular.ttf',
    'nepali':     'NotoSansDevanagari-Regular.ttf',
    'marathi':    'NotoSansDevanagari-Regular.ttf',
    'sanskrit':   'NotoSansDevanagari-Regular.ttf',
    # Indic — others
    'bengali':    'NotoSansBengali-Regular.ttf',
    'gujarati':   'NotoSansGujarati-Regular.ttf',
    'punjabi':    'NotoSansGurmukhi-Regular.ttf',
    'tamil':      'NotoSansTamil-Regular.ttf',
    'telugu':     'NotoSansTelugu-Regular.ttf',
    'kannada':    'NotoSansKannada-Regular.ttf',
    'malayalam':  'NotoSansMalayalam-Regular.ttf',
    'sinhala':    'NotoSansSinhala-Regular.ttf',
    # Arabic-script
    'arabic':     'NotoSansArabic-Regular.ttf',
    'persian':    'NotoSansArabic-Regular.ttf',
    'urdu':       'NotoSansArabic-Regular.ttf',
    'pashto':     'NotoSansArabic-Regular.ttf',
    'sindhi':     'NotoSansArabic-Regular.ttf',
    'uyghur':     'NotoSansArabic-Regular.ttf',
    # South-East Asian
    'thai':       'NotoSansThai-Regular.ttf',
    'khmer':      'NotoSansKhmer-Regular.ttf',
    'myanmar':    'NotoSansMyanmarRegular.ttf',
    'lao':        'NotoSansLao-Regular.ttf',
    # CJK
    'chinese':              'NotoSansSC-Regular.otf',
    'chinese simplified':   'NotoSansSC-Regular.otf',
    'chinese traditional':  'NotoSansTC-Regular.otf',
    'mandarin':             'NotoSansSC-Regular.otf',
    'cantonese':            'NotoSansSC-Regular.otf',
    'japanese':             'NotoSansJP-Regular.otf',
    'korean':               'NotoSansKR-Regular.otf',
}

# Bundled fallback (covers Latin, Cyrillic, Greek, Hebrew, Armenian, Georgian…)
_BUNDLED = 'DejaVuSans.ttf'


def get_font_path(target_language: str) -> str:
    """
    Return the absolute path to the best font for *target_language*.
    Downloads and caches the font on first call if not already present.
    Raises RuntimeError if no font can be found or downloaded.
    """
    FONTS_DIR.mkdir(exist_ok=True)
    lang = target_language.lower().strip()

    # Exact match first, then partial
    font_file = LANGUAGE_FONT.get(lang)
    if not font_file:
        for key, fname in LANGUAGE_FONT.items():
            if key in lang or lang in key:
                font_file = fname
                break

    # Default to bundled DejaVu for everything else
    if not font_file:
        font_file = _BUNDLED

    font_path = FONTS_DIR / font_file

    # Already on disk — done
    if font_path.exists():
        logger.info(f'Using font: {font_file}')
        return str(font_path)

    # Try to download
    url = FONT_DOWNLOADS.get(font_file)
    if url:
        logger.info(f'Downloading font {font_file} …')
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'pdf-converter/1.0'})
            with urllib.request.urlopen(req, timeout=30) as resp, \
                 open(font_path, 'wb') as out:
                out.write(resp.read())
            logger.info(f'Font saved: {font_path}')
            return str(font_path)
        except Exception as exc:
            logger.warning(f'Font download failed ({font_file}): {exc}')
            font_path.unlink(missing_ok=True)  # remove partial file

    # Last resort: bundled DejaVu
    fallback = FONTS_DIR / _BUNDLED
    if fallback.exists():
        logger.warning(f'Falling back to {_BUNDLED} for language "{target_language}"')
        return str(fallback)

    raise RuntimeError(
        f'No font available for "{target_language}" and bundled DejaVuSans is missing. '
        f'Please ensure fonts/DejaVuSans.ttf exists in the project directory.'
    )
