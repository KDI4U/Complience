import asyncio
import logging
import os
import re
from pathlib import Path
from typing import Dict, List, Union

import ollama

# Configure logger for captcha reader module
logger = logging.getLogger("read_captcha")
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter("[%(asctime)s] [%(levelname)s] [read_captcha]: %(message)s", datefmt="%H:%M:%S")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)


class CaptchaReader:
    """
    Standalone CAPTCHA Reader powered by Ollama Vision Models (Default: gemma3:4b).
    Extracts text from CAPTCHA images synchronously or asynchronously.
    """

    DEFAULT_MODEL = "gemma4:12b"
    DEFAULT_PROMPT = (
        "Read the CAPTCHA text from this image. "
        "It contains a short 6-7 character alphanumeric code. "
        "Reply ONLY with these exact characters. "
        "Do NOT add any extra words, spaces, or punctuation."
    )

    def __init__(self, model: str = DEFAULT_MODEL, prompt: str = DEFAULT_PROMPT):
        self.model = model
        self.prompt = prompt

    def clean_text(self, raw_text: str) -> str:
        """Removes non-alphanumeric characters from extracted string."""
        return re.sub(r"[^a-zA-Z0-9]", "", raw_text)

    def read_captcha(self, image_path: Union[str, Path]) -> str:
        """
        Synchronously reads a CAPTCHA image using Ollama.

        :param image_path: Path to the CAPTCHA image file.
        :return: Cleaned CAPTCHA text string.
        """
        path = Path(image_path).resolve()
        if not path.exists():
            logger.error(f"File not found: {path}")
            return ""

        try:
            logger.info(f"Reading CAPTCHA from image: {path.name} (Model: {self.model})")
            kwargs = {
                "model": self.model,
                "messages": [{
                    "role": "user",
                    "content": self.prompt,
                    "images": [str(path)]
                }],
                "options": {
                    "num_predict": 25,
                    "temperature": 0.0
                }
            }
            try:
                response = ollama.chat(**kwargs, think=False)
            except TypeError:
                response = ollama.chat(**kwargs)

            raw_text = ""
            if hasattr(response, "message") and response.message:
                raw_text = (getattr(response.message, "content", "") or "").strip()
                if not raw_text and getattr(response.message, "thinking", None):
                    candidates = re.findall(r"['\"]([a-zA-Z0-9]{4,7})['\"]", response.message.thinking)
                    if candidates:
                        raw_text = candidates[-1]
            elif isinstance(response, dict):
                msg = response.get("message", {})
                raw_text = (msg.get("content") or "").strip()
                if not raw_text and msg.get("thinking"):
                    candidates = re.findall(r"['\"]([a-zA-Z0-9]{4,7})['\"]", msg.get("thinking"))
                    if candidates:
                        raw_text = candidates[-1]

            cleaned_text = self.clean_text(raw_text)
            logger.info(f"Successfully extracted CAPTCHA: '{cleaned_text}' (Raw: '{raw_text}')")
            return cleaned_text
        except Exception as e:
            logger.error(f"Failed to read CAPTCHA from {path.name}: {e}")
            return ""

    async def read_captcha_async(self, image_path: Union[str, Path]) -> str:
        """
        Asynchronously reads a CAPTCHA image.

        :param image_path: Path to the CAPTCHA image file.
        :return: Cleaned CAPTCHA text string.
        """
        return await asyncio.to_thread(self.read_captcha, image_path)

    def read_captcha_folder(self, folder_path: Union[str, Path], extensions: List[str] = [".png", ".jpg", ".jpeg"]) -> Dict[str, str]:
        """
        Reads all CAPTCHA images in a specified folder.

        :param folder_path: Folder containing CAPTCHA images.
        :param extensions: Allowed image file extensions.
        :return: Dictionary mapping image filename to extracted CAPTCHA text.
        """
        folder = Path(folder_path).resolve()
        if not folder.exists() or not folder.is_dir():
            logger.error(f"Folder not found or invalid: {folder}")
            return {}

        results = {}
        image_files = [f for f in folder.iterdir() if f.suffix.lower() in extensions]
        logger.info(f"Found {len(image_files)} images in folder: {folder}")

        for idx, img in enumerate(image_files, 1):
            logger.info(f"Processing [{idx}/{len(image_files)}]: {img.name}")
            captcha_text = self.read_captcha(img)
            results[img.name] = captcha_text

        return results


# Functional helpers for quick one-line calls
def read_single_captcha(image_path: Union[str, Path], model: str = CaptchaReader.DEFAULT_MODEL) -> str:
    reader = CaptchaReader(model=model)
    return reader.read_captcha(image_path)

async def read_single_captcha_async(image_path: Union[str, Path], model: str = CaptchaReader.DEFAULT_MODEL) -> str:
    reader = CaptchaReader(model=model)
    return await reader.read_captcha_async(image_path)

