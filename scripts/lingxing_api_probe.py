from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path
from typing import Any

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from lingxing_rpa_runner import (
    DEFAULT_CONFIG_FILE_NAME,
    LingxingPlaywrightAutomation,
    load_login_credentials,
    normalize_text,
)


def safe_json_loads(text: str) -> Any:
    try:
        return json.loads(text)
    except Exception:
        return None


def summarize_json(value: Any) -> Any:
    if isinstance(value, dict):
        return {
            "type": "dict",
            "keys": list(value.keys())[:40],
            "size": len(value),
        }
    if isinstance(value, list):
        first = value[0] if value else None
        return {
            "type": "list",
            "size": len(value),
            "first": summarize_json(first),
        }
    return {"type": type(value).__name__}


def should_store_full_body(url: str) -> bool:
    return any(
        token in url
        for token in [
            "/api/fba_shipment/showShipment_v2",
            "/amz-sta-server/inbound-plan/detail",
            "/amz-sta-server/inbound-shipment/shipmentLabelPage",
            "/amz-sta-server/inbound-shipment/upsAddress",
            "exportPackingListV2",
        ]
    )


def sanitize_header_snapshot(headers: dict[str, str]) -> dict[str, str]:
    sensitive_tokens = ("cookie", "token", "authorization", "secret", "password", "csrf")
    output: dict[str, str] = {}
    for key, value in headers.items():
        lowered = key.lower()
        if any(token in lowered for token in sensitive_tokens):
            output[key] = "<redacted>"
        else:
            output[key] = value[:300]
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description="Probe Lingxing internal APIs without saving cookies or headers.")
    parser.add_argument("--fba", required=True)
    parser.add_argument("--resource-dir", default=".")
    parser.add_argument("--profile-dir", default="data/browser/profile_playwright")
    parser.add_argument("--config-file")
    parser.add_argument("--output-dir", default="data/probes")
    parser.add_argument("--download-first", action="store_true", help="Click the first card download button to capture export API payload.")
    args = parser.parse_args()

    resource_dir = Path(args.resource_dir).resolve()
    profile_dir = Path(args.profile_dir).resolve()
    config_path = Path(args.config_file).resolve() if args.config_file else resource_dir / DEFAULT_CONFIG_FILE_NAME
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    credentials = load_login_credentials(config_path)
    automation = LingxingPlaywrightAutomation(profile_dir=profile_dir, credentials=credentials)
    events: list[dict[str, Any]] = []

    def record_response(response: Any) -> None:
        try:
            request = response.request
            resource_type = request.resource_type
            url = response.url
            if resource_type not in {"xhr", "fetch"}:
                return
            if "lingxing" not in url and "lingxingerp" not in url:
                return
            content_type = response.headers.get("content-type", "")
            item: dict[str, Any] = {
                "time": time.strftime("%Y-%m-%d %H:%M:%S"),
                "url": url,
                "method": request.method,
                "status": response.status,
                "resource_type": resource_type,
                "content_type": content_type,
                "post_data": normalize_text(request.post_data or "")[:2000],
            }
            if should_store_full_body(url):
                item["request_headers"] = sanitize_header_snapshot(dict(request.headers))
            if any(token in content_type.lower() for token in ["json", "text", "javascript"]):
                try:
                    text = response.text()
                except Exception as exc:
                    item["body_error"] = str(exc)
                else:
                    normalized = normalize_text(text)
                    parsed = safe_json_loads(text)
                    item["body_contains_fba"] = args.fba.upper() in normalized.upper()
                    item["body_preview"] = normalized[:2000]
                    item["json_summary"] = summarize_json(parsed) if parsed is not None else None
                    if should_store_full_body(url):
                        item["body_text"] = text
            events.append(item)
        except Exception as exc:
            events.append({"probe_error": str(exc)})

    try:
        automation.start()
        automation.page.on("response", record_response)
        automation.search_shipment(args.fba)
        automation.open_shipment_detail(args.fba)
        automation.ensure_box_labels_ready(args.fba)
        cards = automation._find_shipment_cards()
        download_result = None
        if args.download_first and cards:
            probe_download_dir = output_dir / "downloads" / args.fba
            probe_download_dir.mkdir(parents=True, exist_ok=True)
            card_text = normalize_text(cards[0].inner_text(timeout=2000))
            warehouse_code = automation._extract_warehouse_code(card_text, 1)
            download_result = automation._download_card_packing_list(
                card=cards[0],
                fba_code=args.fba,
                index=1,
                warehouse_code=warehouse_code,
                download_dir=probe_download_dir,
                defer_raw_download=True,
            )
        time.sleep(2)
        report = {
            "fba": args.fba,
            "current_url": automation.page.url,
            "page_title": automation.page.title(),
            "shipment_card_count": len(cards),
            "download_result": {
                key: value
                for key, value in (download_result or {}).items()
                if not key.startswith("_")
            } if download_result else None,
            "events": events,
        }
    finally:
        automation.close()

    output_path = output_dir / f"lingxing_api_probe_{args.fba}_{int(time.time())}.json"
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(output_path)
    print(json.dumps({
        "fba": args.fba,
        "shipment_card_count": report.get("shipment_card_count"),
        "event_count": len(events),
        "output_path": str(output_path),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
