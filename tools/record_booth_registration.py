from __future__ import annotations

import argparse
import re
from pathlib import Path
from urllib.parse import urlparse

from store.release_pipeline_core import ReleasePipeline, validate_product_id, write_text_atomic


REQUIRED_CONFIRMATION = "RECORD BOOTH REGISTRATION"
EXPECTED_HOST = "peakheadz.booth.pm"


class SafetyStop(RuntimeError):
    pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Record a BOOTH product URL back to a DAKE ORIGINAL.md.")
    parser.add_argument("product_id", help="DAKE product id")
    parser.add_argument("--booth-url", required=True, help="BOOTH product URL, for example https://peakheadz.booth.pm/items/1234567")
    parser.add_argument("--apply", action="store_true", help="Write the BOOTH URL and status fields to ORIGINAL.md.")
    parser.add_argument("--confirm-product-id", default="", help="Must match product_id when --apply is used.")
    parser.add_argument("--confirmation-text", default="", help="Must be RECORD BOOTH REGISTRATION <product_id> when --apply is used.")
    return parser.parse_args()


def validate_booth_url(value: str) -> str:
    if value != value.strip() or any(char in value for char in "\r\n\t "):
        raise SafetyStop("booth_url must not contain whitespace")
    parsed = urlparse(value)
    if parsed.scheme != "https":
        raise SafetyStop("booth_url must use HTTPS")
    if parsed.netloc != EXPECTED_HOST:
        raise SafetyStop(f"booth_url host must be {EXPECTED_HOST}")
    if not re.fullmatch(r"/items/[0-9]+", parsed.path):
        raise SafetyStop("booth_url path must be /items/<digits>")
    if parsed.params or parsed.query or parsed.fragment:
        raise SafetyStop("booth_url must not include params, query, or fragment")
    return value


def metadata_value(text: str, key: str) -> str:
    match = re.search(rf"^\s*[-*]\s*{re.escape(key)}\s*:\s*(.+?)\s*$", text, re.MULTILINE)
    if not match:
        return ""
    value = match.group(1).strip()
    if len(value) >= 2 and value.startswith("`") and value.endswith("`"):
        value = value[1:-1].strip()
    return value


def unset(value: str) -> bool:
    return value.strip().lower() in {"", "not set", "none", "null", "-", "unset", "未設定"}


def replace_metadata_line(text: str, key: str, value: str, *, required: bool = True) -> tuple[str, bool]:
    pattern = re.compile(rf"^(\s*[-*]\s*{re.escape(key)}\s*:\s*)(.*?)\s*$", re.MULTILINE)
    matches = list(pattern.finditer(text))
    if not matches:
        if required:
            raise SafetyStop(f"missing metadata line: {key}")
        return text, False
    if len(matches) != 1:
        raise SafetyStop(f"expected exactly one {key} line, found {len(matches)}")
    return pattern.sub(rf"\g<1>{value}", text, count=1), True


def find_duplicate_url(pipeline: ReleasePipeline, product_id: str, booth_url: str) -> str:
    sources, duplicates = pipeline.discover_sources()
    if duplicates:
        raise SafetyStop("duplicate product id exists in source tree")
    for other_id, source in sources.items():
        if other_id == product_id:
            continue
        if source.booth_url == booth_url:
            return other_id
    return ""


def build_updated_original(text: str, booth_url: str) -> str:
    updated, _ = replace_metadata_line(text, "booth_url", booth_url)
    updated, _ = replace_metadata_line(updated, "status", "available")
    updated, _ = replace_metadata_line(updated, "payment_status", "booth_only")
    updated, _ = replace_metadata_line(updated, "distribution", "BOOTH / manual private download", required=False)
    return updated


def main() -> int:
    args = parse_args()
    id_errors = validate_product_id(args.product_id)
    if id_errors:
        for error in id_errors:
            print(error)
        return 1
    try:
        booth_url = validate_booth_url(args.booth_url)
        pipeline = ReleasePipeline()
        sources, duplicates = pipeline.discover_sources()
        if args.product_id in duplicates:
            raise SafetyStop("duplicate product id in ORIGINAL.md files")
        source = sources.get(args.product_id)
        if source is None:
            raise SafetyStop(f"Product id not found: {args.product_id}")
        if source.product_type != "pack":
            raise SafetyStop("BOOTH registration recording is currently limited to pack products")
        duplicate_owner = find_duplicate_url(pipeline, args.product_id, booth_url)
        if duplicate_owner:
            raise SafetyStop(f"booth_url is already used by {duplicate_owner}")
        current_booth_url = metadata_value(source.text, "booth_url")
        if current_booth_url and not unset(current_booth_url):
            raise SafetyStop("ORIGINAL.md already has a BOOTH URL; refusing to overwrite")
        if metadata_value(source.text, "payment_status") != "preparing":
            raise SafetyStop("payment_status must be preparing before BOOTH registration is recorded")
        if not unset(metadata_value(source.text, "stripe_payment_link")):
            raise SafetyStop("stripe_payment_link must remain unset before this BOOTH registration step")

        updated = build_updated_original(source.text, booth_url)
        print(f"product_id={args.product_id}")
        print(f"booth_url={booth_url}")
        print("status_after=available")
        print("payment_status_after=booth_only")
        print(f"source_original={source.source_original}")
        print(f"apply={args.apply}")

        if not args.apply:
            return 0
        expected_text = f"{REQUIRED_CONFIRMATION} {args.product_id}"
        if args.confirm_product_id != args.product_id:
            raise SafetyStop(f"--confirm-product-id must be {args.product_id}")
        if args.confirmation_text != expected_text:
            raise SafetyStop(f'--confirmation-text must be "{expected_text}"')
        write_text_atomic(source.path, updated)
        verified = source.path.read_text(encoding="utf-8")
        if metadata_value(verified, "booth_url") != booth_url:
            raise SafetyStop("post-apply booth_url verification failed")
        if metadata_value(verified, "status") != "available":
            raise SafetyStop("post-apply status verification failed")
        if metadata_value(verified, "payment_status") != "booth_only":
            raise SafetyStop("post-apply payment_status verification failed")
        print("applied=True")
        return 0
    except SafetyStop as exc:
        print(str(exc))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
