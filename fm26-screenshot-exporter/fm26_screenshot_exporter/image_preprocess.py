"""Screenshot preprocessing for improved OCR accuracy."""

from __future__ import annotations

from pathlib import Path

import cv2
import numpy as np
from PIL import Image

from fm26_screenshot_exporter.config import CropRegion


def load_image(path: str | Path) -> np.ndarray:
    """Load an image as a BGR numpy array."""
    image = cv2.imread(str(path))
    if image is None:
        raise FileNotFoundError(f"Could not read image: {path}")
    return image


def save_image(image: np.ndarray, path: str | Path) -> Path:
    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    cv2.imwrite(str(out), image)
    return out


def crop_region(image: np.ndarray, crop: CropRegion | None) -> np.ndarray:
    if crop is None:
        return image
    h, w = image.shape[:2]
    x1 = max(0, crop.x)
    y1 = max(0, crop.y)
    x2 = min(w, crop.x + crop.width)
    y2 = min(h, crop.y + crop.height)
    if x2 <= x1 or y2 <= y1:
        raise ValueError(f"Invalid crop region: {crop}")
    return image[y1:y2, x1:x2]


def preprocess_screenshot(
    image: np.ndarray,
    *,
    crop: CropRegion | None = None,
    upscale_factor: float = 2.0,
) -> np.ndarray:
    """
    Apply FM screenshot preprocessing pipeline:
    crop → upscale → grayscale → denoise → sharpen → adaptive threshold.
    """
    working = crop_region(image, crop)

    if upscale_factor != 1.0:
        h, w = working.shape[:2]
        working = cv2.resize(
            working,
            (int(w * upscale_factor), int(h * upscale_factor)),
            interpolation=cv2.INTER_CUBIC,
        )

    gray = cv2.cvtColor(working, cv2.COLOR_BGR2GRAY)
    denoised = cv2.fastNlMeansDenoising(gray, h=10, templateWindowSize=7, searchWindowSize=21)

    sharpen_kernel = np.array(
        [[0, -1, 0], [-1, 5, -1], [0, -1, 0]],
        dtype=np.float32,
    )
    sharpened = cv2.filter2D(denoised, -1, sharpen_kernel)

    thresholded = cv2.adaptiveThreshold(
        sharpened,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31,
        8,
    )
    return thresholded


def preprocess_for_ocr(
    image: np.ndarray,
    *,
    crop: CropRegion | None = None,
    upscale_factor: float = 2.0,
) -> tuple[np.ndarray, np.ndarray]:
    """
    Return (preprocessed_for_tesseract, color_cropped_for_debug).
    Tesseract works better on upscaled grayscale without harsh thresholding.
    """
    working = crop_region(image, crop)
    color = working.copy()

    if upscale_factor != 1.0:
        h, w = working.shape[:2]
        working = cv2.resize(
            working,
            (int(w * upscale_factor), int(h * upscale_factor)),
            interpolation=cv2.INTER_CUBIC,
        )
        color = cv2.resize(
            color,
            (int(w * upscale_factor), int(h * upscale_factor)),
            interpolation=cv2.INTER_CUBIC,
        )

    gray = cv2.cvtColor(working, cv2.COLOR_BGR2GRAY)
    denoised = cv2.fastNlMeansDenoising(gray, h=10, templateWindowSize=7, searchWindowSize=21)
    sharpen_kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]], dtype=np.float32)
    sharpened = cv2.filter2D(denoised, -1, sharpen_kernel)
    return sharpened, color


def to_pil(image: np.ndarray) -> Image.Image:
    if len(image.shape) == 2:
        return Image.fromarray(image)
    rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    return Image.fromarray(rgb)
