import argparse
import sys
from pathlib import Path

# Configure Windows console encoding safely
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Add current folder to path so imports work smoothly
CURRENT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(CURRENT_DIR))

from captcha_reader import CaptchaReader, read_single_captcha


def find_sample_images() -> list[Path]:
    """Scans project for existing sample CAPTCHA images to test."""
    possible_dirs = [
        CURRENT_DIR / "sample_images",
        CURRENT_DIR.parent / "data" / "captcha",
        CURRENT_DIR.parent.parent / "read_captcha" / "sample_images",
    ]
    samples = []
    for d in possible_dirs:
        if d.exists() and d.is_dir():
            found = list(d.glob("*.png")) + list(d.glob("*.jpg"))
            if found:
                samples.extend(found)
    return samples


def main():
    parser = argparse.ArgumentParser(
        description="Standalone CAPTCHA Reader Tool (Ollama AI Vision)",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--image", "-i", type=str, help="Path to a single CAPTCHA image file")
    parser.add_argument("--folder", "-f", type=str, help="Path to a folder containing CAPTCHA images")
    parser.add_argument("--model", "-m", type=str, default="gemma3:4b", help="Ollama Vision model name")
    parser.add_argument("--test", "-t", action="store_true", help="Run automated test using available sample CAPTCHA images")

    args = parser.parse_args()
    reader = CaptchaReader(model=args.model)

    print("=" * 60)
    print("           STANDALONE CAPTCHA READER (OLLAMA AI)")
    print("=" * 60)
    print(f"Ollama Model : {args.model}")

    if args.image:
        image_path = Path(args.image).resolve()
        print(f"\nTarget Image : {image_path}")
        result = reader.read_captcha(image_path)
        print("-" * 60)
        print(f"RESULT: {result}")
        print("=" * 60)
        return

    if args.folder:
        folder_path = Path(args.folder).resolve()
        print(f"\nTarget Folder: {folder_path}")
        results = reader.read_captcha_folder(folder_path)
        print("\n" + "=" * 60)
        print("SUMMARY RESULTS:")
        print("=" * 60)
        for img_name, code in results.items():
            print(f"  {img_name:<40} -> {code}")
        print("=" * 60)
        return

    if args.test or not (args.image or args.folder):
        print("\nRunning Test Mode on available sample images...")
        samples = find_sample_images()
        if not samples:
            print("[X] No sample images found in project.")
            print("Usage Examples:")
            print("  python main.py --image path/to/captcha.png")
            print("  python main.py --folder path/to/captcha_folder")
            return

        print(f"Found {len(samples)} sample image(s). Testing up to 3 samples:\n")
        for sample in samples[:3]:
            result = reader.read_captcha(sample)
            print(f"[IMAGE]  : {sample.name}")
            print(f"[RESULT] : {result}")
            print("-" * 40)
        print("=" * 60)


if __name__ == "__main__":
    main()

