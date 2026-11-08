"""JSON-only bridge for the app; credentials arrive through stdin, never arguments."""
import argparse
import contextlib
from datetime import datetime, timezone
import json
import logging
import re
import os
from pathlib import Path
import shutil
import sys
import tempfile
import time
import uuid

import buildinglink_backup as backup


LIVE_PATH = None
STARTED = datetime.now(timezone.utc).isoformat()


class Cancelled(Exception):
    pass


def emit(kind, **values):
    event = dict(type=kind, **values)
    if LIVE_PATH is not None:
        status = dict(event, running=kind != "complete", updated=datetime.now(timezone.utc).isoformat(), started=STARTED)
        temporary = LIVE_PATH.with_name(LIVE_PATH.name + "." + uuid.uuid4().hex + ".tmp")
        try:
            LIVE_PATH.parent.mkdir(parents=True, exist_ok=True)
            temporary.write_text(json.dumps(status), encoding="utf-8")
            os.replace(temporary, LIVE_PATH)
        except OSError:
            pass
        finally:
            temporary.unlink(missing_ok=True)
    print(json.dumps(event), flush=True)


def validate_settings(settings):
    buildings = [b for b in settings.get("Buildings", []) if b.get("Enabled", True)]
    if not buildings:
        raise ValueError("Select at least one corporation.")
    paths = []
    identities = {"DEMO 1001": "901001", "DEMO 2002": "902002"}
    seen = set()
    for b in buildings:
        for key in ("Corporation", "PropertyName", "FacilityId", "OccupantFolder", "UnitFolder"):
            if not str(b.get(key, "")).strip():
                raise ValueError("A corporation is missing its identity or destination folders.")
        if not re.fullmatch(r"[A-Za-z0-9 -]{1,60}", b["Corporation"]) or not str(b["FacilityId"]).isdigit():
            raise ValueError("Invalid corporation identity.")
        if b["Corporation"] not in identities or str(b["FacilityId"]) != identities[b["Corporation"]]:
            raise ValueError("This downloader supports DEMO 1001 and DEMO 2002 with their assigned BuildingLink properties only.")
        if b["Corporation"] in seen:
            raise ValueError("Choose each supported corporation only once.")
        seen.add(b["Corporation"])
        for key in ("OccupantFolder", "UnitFolder"):
            path = Path(b[key])
            if not path.is_absolute():
                raise ValueError("Backup destinations must be full paths.")
            paths.append(os.path.normcase(str(path.resolve())))
    if len(set(paths)) != len(paths):
        raise ValueError("Use separate destination folders for each corporation and report.")
    return buildings


def validate_download(path, report_type, corporation):
    backup.validate_workbook(path, report_type, corporation)
    if path.suffix.lower() == ".xls":
        content = path.read_bytes()
        if content.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
            return  # Legacy binary Excel; report identity is checked before export.
        text = content.decode("utf-8", errors="replace").lower()
        if ("<table" not in text and "urn:schemas-microsoft-com:office:spreadsheet" not in text) or "unit" not in text:
            raise RuntimeError("Unit Data download is not a recognized Excel report.")


def publish(local, folder, report_type, corporation):
    """Validate and publish through a sibling temporary file; retain older snapshots."""
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    destination = folder / local.name
    if destination.exists():
        destination = folder / f"{local.stem}_{time.time_ns()}{local.suffix}"
    temp = folder / f".{uuid.uuid4().hex}{local.suffix}"
    try:
        shutil.copyfile(local, temp)
        validate_download(temp, report_type, corporation)
        # No overwrite even if another process chose the same filename.
        if destination.exists():
            raise FileExistsError("A backup with this name already exists.")
        temp.rename(destination)
        return str(destination)
    finally:
        temp.unlink(missing_ok=True)


def execute(settings, credentials, cancelled=lambda: False):
    buildings = validate_settings(settings)
    username, password = credentials.get("username", ""), credentials.get("password", "")
    if not username or not password:
        raise ValueError("Save a BuildingLink username and password in Manage credentials first.")
    total = len(buildings) * 2
    files, failures = [], []
    current_stage = "Starting download"
    milestones = [0] * len(buildings)
    building_index = 0

    def percentage():
        # A failed corporation retains only its completed milestones. No timer
        # estimates, skipped work, or failure can turn the bar into 100%.
        return min(99, sum(milestones) / len(buildings))

    def stage(message, milestone, check_cancel=True, kind="progress", **extra):
        nonlocal current_stage
        if check_cancel and cancelled():
            raise Cancelled()
        current_stage = message
        milestones[building_index] = max(milestones[building_index], milestone)
        emit(kind, message=message, completed=len(files), total=total,
             progress_percent=percentage(), corporation=buildings[building_index]["Corporation"], **extra)

    with backup.single_instance(), backup.sync_playwright() as playwright:
        for building_index, b in enumerate(buildings):
            context, profile, manager = None, None, None
            corp = b["Corporation"]
            try:
                stage(f"Signing in to {corp}: starting hidden browser", 0)
                context, page, profile = backup.launch_clean_edge(playwright, backup.PORTAL_URL, True, 20)
                stage(f"Signing in to {corp}: opening sign-in page", 4)

                def login_progress(step):
                    messages = {
                        "login_form": ("entering saved login", 6),
                        "login_submitted": ("login submitted; waiting for BuildingLink", 10),
                        "selecting_building": ("selecting corporation", 14),
                    }
                    detail, milestone = messages[step]
                    stage(f"Signing in to {corp}: {detail}", milestone)

                manager = backup.login_and_open_building(context, page, username, password, b["PropertyName"], 60, 20, progress=login_progress)
                stage(f"Signing in to {corp}: verifying active corporation", 18)
                backup.verify_or_switch_building(manager, b["PropertyName"], b["FacilityId"], 20)
                stage(f"{corp}: corporation verified; opening Occupant Data", 20)
                with tempfile.TemporaryDirectory(prefix="PropertyManager-BuildingLink-") as directory:
                    staging = Path(directory)
                    occupant_page = backup.open_occupant_report(context, manager, b["PropertyName"], headless=True)
                    stage(f"Downloading {corp} Occupant Data: requesting Excel export", 25)
                    download = backup.download_occupant(occupant_page, 60)
                    stage(f"Downloading {corp} Occupant Data: receiving export", 35)
                    occupant = backup.save_staged_download(download, staging, corp, "Occupant-Data")
                    stage(f"{corp}: validating Occupant Data", 40)
                    validate_download(occupant, "Occupant-Data", corp)
                    stage(f"Downloading {corp} Unit Data: opening report", 45)
                    unit_page, surface = backup.open_unit_report(context, occupant_page, b["PropertyName"], 60, headless=True)
                    stage(f"Downloading {corp} Unit Data: requesting Excel export", 50)
                    download = backup.download_unit(unit_page, surface, 60, context=context, headless=True)
                    stage(f"Downloading {corp} Unit Data: receiving export", 65)
                    unit = backup.save_staged_download(download, staging, corp, "Unit-Data")
                    stage(f"{corp}: validating Unit Data", 70)
                    validate_download(unit, "Unit-Data", corp)
                    stage(f"{corp}: both reports verified; saving Occupant Data", 80)
                    # Once saving begins, complete this corporation's pair before checking cancellation.
                    for local, folder, kind, milestone in ((occupant,b["OccupantFolder"],"Occupant-Data",90),(unit,b["UnitFolder"],"Unit-Data",98)):
                        path = publish(local, folder, kind, corp)
                        files.append(dict(corporation=corp, report=kind, path=path))
                        stage(f"{corp} {kind} saved", milestone, check_cancel=False, kind="file", path=path)
                        if kind == "Occupant-Data":
                            stage(f"{corp}: saving Unit Data", milestone, check_cancel=False)
                    stage(f"{corp}: reports saved; closing hidden browser", 98, check_cancel=False)
            except Cancelled:
                return dict(success=False, total=total, progress_percent=percentage(), cancelled=True, message="Download cancelled. Previously saved reports were kept.", files=files, failures=failures)
            except Exception:
                # Locator errors can contain filled values, tokens, or redirected URLs.
                # Retain the stage and actionable guidance, never the raw exception.
                guidance = "Check saved login, Microsoft Edge, and any required sign-in verification." if current_stage.startswith("Signing in") else "Check BuildingLink report access and destination folder availability."
                failures.append(dict(corporation=corp, reason=current_stage+" failed. "+guidance))
                emit("error", message=f"{corp}: download failed. Check the saved login, BuildingLink access, and destination folders.",
                     completed=len(files), total=total, progress_percent=percentage())
                # Avoid repeating a rejected login against another building.
                if manager is None:
                    break
            else:
                milestones[building_index] = 100
            finally:
                backup.close_clean_edge(context, profile)
    return dict(success=not failures, total=total, progress_percent=100 if not failures else percentage(), cancelled=False, files=files, failures=failures,
                message=f"{len(files)} of {total} reports saved." + (" Some reports need attention." if failures else ""))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--settings", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--cancel", type=Path, required=True)
    args = parser.parse_args()
    global LIVE_PATH
    LIVE_PATH = args.result.with_name("buildinglink-progress.json")
    # Original helper logging stays off: stdout is reserved for structured events.
    logging.disable(logging.CRITICAL)
    result = None
    try:
        credentials = json.loads(sys.stdin.readline())
        settings = json.loads(args.settings.read_text(encoding="utf-8-sig"))
        result = execute(settings, credentials, args.cancel.exists)
    except Exception:
        result = dict(success=False, message="BuildingLink download could not start. Check credentials, settings, Microsoft Edge, and download components.", files=[], failures=[])
    result["finished"] = datetime.now(timezone.utc).isoformat()
    args.result.parent.mkdir(parents=True, exist_ok=True)
    temp = args.result.with_name(args.result.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        temp.write_text(json.dumps(result, indent=2), encoding="utf-8")
        os.replace(temp, args.result)
    finally:
        temp.unlink(missing_ok=True)
    emit("complete", **result)
    return 0 if result.get("success") else 2 if result.get("cancelled") else 1


if __name__ == "__main__":
    raise SystemExit("Public showcase: live worker execution is disabled.")
    raise SystemExit(main())
