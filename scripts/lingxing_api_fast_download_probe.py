from __future__ import annotations

import argparse
import json
import sys
import time
import uuid
from datetime import timedelta
from pathlib import Path
from typing import Any

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from lingxing_excel_processor import process_workbooks
from lingxing_rpa_runner import (
    DEFAULT_CONFIG_FILE_NAME,
    LingxingPlaywrightAutomation,
    beijing_now,
    build_download_filename,
    elapsed_seconds,
    filename_from_content_disposition,
    is_valid_xlsx_payload,
    load_login_credentials,
    sanitize_filename_part,
)


ERP_BASE = "https://erp.lingxing.com"
GW_BASE = "https://gw.lingxingerp.com"


class ApiProbeError(RuntimeError):
    pass


def clone_api_headers(headers: dict[str, str]) -> dict[str, str]:
    output = {
        key: value
        for key, value in headers.items()
        if key.lower() not in {"content-length", "accept-encoding"}
    }
    output["content-type"] = "application/json;charset=UTF-8"
    output["accept"] = "application/json, text/plain, */*"
    output["x-ak-request-id"] = str(uuid.uuid4())
    return output


def post_json(context: Any, headers: dict[str, str], url: str, payload: dict[str, Any]) -> tuple[dict[str, Any], float]:
    start = time.perf_counter()
    response = context.request.post(
        url,
        data=json.dumps(payload, ensure_ascii=False),
        headers=clone_api_headers(headers),
        timeout=120000,
    )
    body = response.body()
    if not response.ok:
        raise ApiProbeError(f"API HTTP {response.status}: {url}")
    try:
        data = json.loads(body.decode("utf-8"))
    except Exception as exc:
        raise ApiProbeError(f"API did not return JSON: {url}") from exc
    if data.get("code") not in (1, 200, "1", "200", None) and data.get("success") is not True:
        raise ApiProbeError(f"API business error: {url}: {data.get('msg') or data.get('message')}")
    return data, elapsed_seconds(start)


def download_excel(context: Any, headers: dict[str, str], url: str, payload: dict[str, Any]) -> tuple[str, bytes, float]:
    start = time.perf_counter()
    response = context.request.post(
        url,
        data=json.dumps(payload, ensure_ascii=False),
        headers=clone_api_headers(headers),
        timeout=120000,
    )
    body = response.body()
    if not response.ok:
        raise ApiProbeError(f"Export HTTP {response.status}: {url}")
    if not is_valid_xlsx_payload(body):
        raise ApiProbeError("Export response is not a valid xlsx")
    suggested = filename_from_content_disposition(response.headers.get("content-disposition")) or "packing_list.xlsx"
    return suggested, body, elapsed_seconds(start)


def search_shipment(context: Any, headers: dict[str, str], fba_code: str) -> tuple[dict[str, Any], float]:
    end = beijing_now().date()
    start = end - timedelta(days=90)
    payload = {
        "search_field_time": "create_date",
        "is_sta": "",
        "is_awd": "",
        "ship_mode": "",
        "step": [],
        "is_closed": "",
        "application_diff": "",
        "received_diff": "",
        "application_received_diff": "",
        "is_relate_packing_task_sn": "",
        "is_add_tracking": "",
        "delivery_order_status": [],
        "box_type": "",
        "is_uploaded_box": "",
        "sta_transportation_mode": "",
        "delivery_mode": "",
        "carrier_type": "",
        "create_uids": [],
        "principal_uids": [],
        "is_store_diff": "",
        "search_field": "shipment_id",
        "search_value": fba_code,
        "shipment_status": [],
        "is_relate_shipment": "",
        "start_date": start.isoformat(),
        "end_date": end.isoformat(),
        "seniorSearchList": [],
        "shipment_type": [],
        "offset": 0,
        "length": 20,
        "req_time_sequence": "/api/fba_shipment/showShipment_v2$$1",
    }
    data, seconds = post_json(context, headers, f"{ERP_BASE}/api/fba_shipment/showShipment_v2", payload)
    rows = ((data.get("data") or {}).get("list") or [])
    exact = [row for row in rows if str(row.get("shipment_id", "")).upper() == fba_code.upper()]
    if len(exact) != 1:
        raise ApiProbeError(f"Expected exactly one shipment for {fba_code}, got {len(exact)}")
    return exact[0], seconds


def get_plan_detail(context: Any, headers: dict[str, str], local_task_id: str) -> tuple[dict[str, Any], float]:
    payload = {
        "localTaskId": local_task_id,
        "req_time_sequence": "/amz-sta-server/inbound-plan/detail$$1",
    }
    data, seconds = post_json(context, headers, f"{GW_BASE}/amz-sta-server/inbound-plan/detail", payload)
    detail = data.get("data") or {}
    if not detail.get("inboundPlanId"):
        raise ApiProbeError("Missing inboundPlanId")
    return detail, seconds


def get_label_page(context: Any, headers: dict[str, str], sid: int, inbound_plan_id: str) -> tuple[list[dict[str, Any]], float]:
    payload = {
        "inboundPlanId": inbound_plan_id,
        "sid": sid,
        "req_time_sequence": "/amz-sta-server/inbound-shipment/shipmentLabelPage$$1",
    }
    data, seconds = post_json(context, headers, f"{GW_BASE}/amz-sta-server/inbound-shipment/shipmentLabelPage", payload)
    shipments = data.get("data") or []
    if not shipments:
        raise ApiProbeError("No shipment label records")
    return shipments, seconds


def capture_api_headers(automation: LingxingPlaywrightAutomation) -> tuple[dict[str, str], float]:
    captured: dict[str, str] = {}

    def on_request(request: Any) -> None:
        if captured:
            return
        if "/api/fba_shipment/showShipment_v2" not in request.url:
            return
        captured.update(request.headers)

    start = time.perf_counter()
    automation.page.on("request", on_request)
    automation.open_fba_shipments_page()
    deadline = time.time() + 12
    while time.time() < deadline and not captured:
        automation.page.wait_for_timeout(200)
    if not captured:
        raise ApiProbeError("Could not capture Lingxing API headers")
    return captured, elapsed_seconds(start)


def run_probe(fba_code: str, resource_dir: Path, profile_dir: Path, config_path: Path | None, job_dir: Path) -> dict[str, Any]:
    timings: dict[str, float] = {}
    job_dir.mkdir(parents=True, exist_ok=True)
    download_dir = job_dir / "downloads" / sanitize_filename_part(fba_code)
    output_dir = job_dir / "output" / sanitize_filename_part(fba_code)
    download_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    credentials = load_login_credentials(config_path) if config_path else None
    automation = LingxingPlaywrightAutomation(profile_dir=profile_dir, credentials=credentials)
    overall_start = time.perf_counter()
    downloaded_files: list[dict[str, Any]] = []
    try:
        start = time.perf_counter()
        automation.start()
        automation.ensure_logged_in()
        timings["login_ready_seconds"] = elapsed_seconds(start)
        api_headers, timings["header_capture_seconds"] = capture_api_headers(automation)

        shipment, timings["search_api_seconds"] = search_shipment(automation.context, api_headers, fba_code)
        local_task_id = str(shipment.get("local_sta_id") or "")
        sid = int(shipment["sid"])
        if not local_task_id:
            raise ApiProbeError("Missing local_sta_id")

        detail, timings["plan_detail_api_seconds"] = get_plan_detail(automation.context, api_headers, local_task_id)
        shipments, timings["label_page_api_seconds"] = get_label_page(
            automation.context,
            api_headers,
            sid=sid,
            inbound_plan_id=detail["inboundPlanId"],
        )

        export_start = time.perf_counter()
        for index, label in enumerate(shipments, start=1):
            warehouse_code = sanitize_filename_part(str(label.get("warehouseId") or f"WAREHOUSE{index:02d}"))
            payload = {
                "isBatch": 0,
                "isPic": 0,
                "packingListBOList": [
                    {
                        "localTaskId": local_task_id,
                        "packingGroupId": None,
                        "shipmentId": label["shipmentId"],
                    }
                ],
                "sid": sid,
            }
            suggested_name, body, seconds = download_excel(
                automation.context,
                api_headers,
                f"{GW_BASE}/amz-sta-server/inbound-packing/exportPackingListV2",
                payload,
            )
            suffix = Path(suggested_name).suffix or ".xlsx"
            target_path = download_dir / build_download_filename(fba_code, index, warehouse_code, suggested_name, suffix)
            target_path.write_bytes(body)
            downloaded_files.append(
                {
                    "sequence": index,
                    "warehouse_code": warehouse_code,
                    "shipment_confirmation_id": label.get("shipmentConfirmationId"),
                    "source_name": suggested_name,
                    "saved_name": target_path.name,
                    "saved_path": str(target_path),
                    "api_export_seconds": seconds,
                }
            )
        timings["export_api_seconds"] = elapsed_seconds(export_start)
    finally:
        automation.close()

    process_start = time.perf_counter()
    process_report = process_workbooks(resource_dir, download_dir, output_dir)
    timings["excel_process_seconds"] = elapsed_seconds(process_start)
    timings["total_seconds"] = elapsed_seconds(overall_start)

    return {
        "fba": fba_code,
        "status": "success",
        "shipment_count": len(downloaded_files),
        "downloaded_files": downloaded_files,
        "process_report": process_report,
        "timings": timings,
        "download_dir": str(download_dir),
        "output_dir": str(output_dir),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Experimental direct Lingxing API fast download probe.")
    parser.add_argument("--fba", required=True)
    parser.add_argument("--resource-dir", default=".")
    parser.add_argument("--profile-dir", default="data/browser/profile_playwright")
    parser.add_argument("--config-file")
    parser.add_argument("--job-dir", default="data/probes/api_fast_job")
    args = parser.parse_args()

    resource_dir = Path(args.resource_dir).resolve()
    profile_dir = Path(args.profile_dir).resolve()
    config_path = Path(args.config_file).resolve() if args.config_file else resource_dir / DEFAULT_CONFIG_FILE_NAME
    job_dir = Path(args.job_dir).resolve()

    report = run_probe(args.fba, resource_dir, profile_dir, config_path, job_dir)
    report_path = job_dir / f"api_fast_probe_{sanitize_filename_part(args.fba)}.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(report_path)
    print(json.dumps({
        "status": report["status"],
        "fba": report["fba"],
        "shipment_count": report["shipment_count"],
        "timings": report["timings"],
        "output_dir": report["output_dir"],
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
