import asyncio
import json
import logging
import os
import threading
from typing import Dict, Optional

import azure.functions as func

from main import main as run_sync


app = func.FunctionApp()
_sync_lock = threading.Lock()


def _parse_bool(value: Optional[str]) -> Optional[bool]:
    if value is None:
        return None
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "y", "on"}:
        return True
    if normalized in {"0", "false", "no", "n", "off"}:
        return False
    raise ValueError(f"Invalid boolean value '{value}'")


def _run_sync(overrides: Optional[Dict[str, str]] = None) -> int:
    overrides = overrides or {}
    with _sync_lock:
        previous_values: Dict[str, Optional[str]] = {}
        for key, value in overrides.items():
            previous_values[key] = os.environ.get(key)
            os.environ[key] = value

        try:
            return asyncio.run(run_sync())
        finally:
            for key, previous in previous_values.items():
                if previous is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = previous


def _get_request_value(req: func.HttpRequest, name: str) -> Optional[str]:
    if name in req.params:
        return req.params.get(name)

    try:
        payload = req.get_json()
    except ValueError:
        payload = {}

    value = payload.get(name)
    if value is None:
        return None
    return str(value)


def _build_overrides(req: func.HttpRequest) -> dict[str, str]:
    overrides: dict[str, str] = {}

    force_full_sync = _parse_bool(_get_request_value(req, "force_full_sync"))
    if force_full_sync is not None:
        overrides["FORCE_FULL_SYNC"] = str(force_full_sync).lower()

    dry_run = _parse_bool(_get_request_value(req, "dry_run"))
    if dry_run is not None:
        overrides["DRY_RUN"] = str(dry_run).lower()

    delete_orphaned_blobs = _parse_bool(_get_request_value(req, "delete_orphaned_blobs"))
    if delete_orphaned_blobs is not None:
        overrides["DELETE_ORPHANED_BLOBS"] = str(delete_orphaned_blobs).lower()

    sync_permissions = _parse_bool(_get_request_value(req, "sync_permissions"))
    if sync_permissions is not None:
        overrides["SYNC_PERMISSIONS"] = str(sync_permissions).lower()

    return overrides


@app.function_name(name="sharepoint_sync_http")
@app.http_trigger(
    route="sharepoint-sync",
    arg_name="req",
    methods=["GET", "POST"],
    auth_level=func.AuthLevel.FUNCTION,
)
def sharepoint_sync_http(req: func.HttpRequest) -> func.HttpResponse:
    try:
        overrides = _build_overrides(req)
    except ValueError as exc:
        return func.HttpResponse(
            body=json.dumps({
                "status": "error",
                "message": str(exc),
                "allowed_values": ["true", "false", "1", "0", "yes", "no"],
            }),
            status_code=400,
            mimetype="application/json",
        )

    logging.info("HTTP trigger received for SharePoint sync", extra={"overrides": overrides})

    exit_code = _run_sync(overrides)
    status_code = 200 if exit_code == 0 else (400 if exit_code == 2 else 500)

    return func.HttpResponse(
        body=json.dumps({
            "status": "ok" if exit_code == 0 else "error",
            "exit_code": exit_code,
            "applied_overrides": overrides,
        }),
        status_code=status_code,
        mimetype="application/json",
    )
