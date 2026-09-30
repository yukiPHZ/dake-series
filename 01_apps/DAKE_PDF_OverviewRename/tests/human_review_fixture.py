"""Create synthetic, numbered documents for the fixed Human Review folder."""
import argparse
import json
from pathlib import Path
from PIL import Image, ImageDraw
from pypdf import PdfReader


def generate(folder: Path, count: int):
    folder.mkdir(parents=True, exist_ok=True)
    manifest = []
    for i in range(1, count+1):
        path = folder / f"synthetic_{i:04d}.pdf"
        if path.exists():
            raise FileExistsError(path)
        pages = 1 + i % 5
        image = Image.new("RGB", (700, 910), "white")
        draw = ImageDraw.Draw(image)
        draw.rectangle((20, 20, 680, 890), outline=(i%200, 80, 150), width=8)
        draw.text((60, 70), f"DAKE SYNTHETIC DOCUMENT {i:04d}", fill="black", font_size=30)
        draw.text((60, 130), f"Expected pages: {pages}", fill="black", font_size=24)
        draw.text((60, 190), "No confidential data", fill="black", font_size=24)
        image.save(path, "PDF", resolution=100, save_all=True, append_images=[image]*(pages-1))
        assert len(PdfReader(path).pages) == pages
        manifest.append(dict(name=path.name, pages=pages, bytes=path.stat().st_size))
    (folder / "expected.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return manifest


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("folder", type=Path)
    parser.add_argument("--count", type=int, default=400)
    args = parser.parse_args()
    print(json.dumps({"count":len(generate(args.folder, args.count)), "folder":str(args.folder)}))
