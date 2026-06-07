"""OCR extraction using pytesseract."""

from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

import pytesseract
from PIL import Image

from fm26_screenshot_exporter.config import OcrToken
from fm26_screenshot_exporter.image_preprocess import to_pil


def check_tesseract() -> dict[str, str | bool]:
    """Check whether tesseract is installed and return version info."""
    path = shutil.which("tesseract")
    if not path:
        return {"installed": False, "path": "", "version": ""}

    try:
        result = subprocess.run(
            ["tesseract", "--version"],
            capture_output=True,
            text=True,
            check=True,
        )
        version_line = result.stdout.splitlines()[0] if result.stdout else ""
    except (subprocess.CalledProcessError, FileNotFoundError):
        version_line = ""

    return {"installed": True, "path": path, "version": version_line}


def run_ocr(image, *, lang: str = "eng") -> list[OcrToken]:
    """Run OCR with word-level bounding boxes via pytesseract.image_to_data."""
    pil_image = to_pil(image) if not isinstance(image, Image.Image) else image
    data = pytesseract.image_to_data(pil_image, lang=lang, output_type=pytesseract.Output.DICT)

    tokens: list[OcrToken] = []
    n = len(data["text"])
    for i in range(n):
        text = (data["text"][i] or "").strip()
        if not text:
            continue
        try:
            conf = float(data["conf"][i])
        except (ValueError, TypeError):
            conf = -1.0
        if conf < 0:
            continue

        tokens.append(
            OcrToken(
                text=text,
                confidence=conf,
                x=int(data["left"][i]),
                y=int(data["top"][i]),
                width=int(data["width"][i]),
                height=int(data["height"][i]),
                line_num=int(data["line_num"][i]),
                block_num=int(data["block_num"][i]),
                par_num=int(data["par_num"][i]),
            )
        )
    return tokens


def filter_tokens_by_confidence(
    tokens: list[OcrToken],
    min_confidence: float,
) -> list[OcrToken]:
    return [t for t in tokens if t.confidence >= min_confidence]


def save_ocr_json(tokens: list[OcrToken], path: str | Path) -> Path:
    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    payload = [t.model_dump() for t in tokens]
    out.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return out


def load_ocr_json(path: str | Path) -> list[OcrToken]:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    return [OcrToken.model_validate(item) for item in data]
