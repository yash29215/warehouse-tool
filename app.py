"""
app.py — Flask web server for Warehouse Tools.
Serves the single-page app and all API endpoints.
"""

import json
import os
import platform
import posixpath
import shutil
import subprocess
import tempfile
import threading
import time
import uuid
from collections import defaultdict

import pandas as pd
import paramiko
from flask import (Flask, Response, jsonify, request,
                   send_file, send_from_directory)

import warehouse_logic as wl

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "downloads")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# True when this Flask process itself is running on a Mac. Map Sync uses this
# to decide whether it can copy files on its own local disk (Windows/Linux —
# presumably co-located with the robot folders) or must SSH into the remote
# Linux VM instead (Mac — no local access to those paths, and RoboShop
# doesn't run on Mac at all).
IS_MAC_SERVER = platform.system() == "Darwin"

# In-memory store for SSE progress streams
_sse_streams: dict[str, list] = {}
_sse_lock = threading.Lock()


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def _save_upload(file_obj, prefix="upload") -> str:
    ext = os.path.splitext(file_obj.filename)[1].lower() or ".bin"
    name = f"{prefix}_{uuid.uuid4().hex}{ext}"
    path = os.path.join(UPLOAD_DIR, name)
    file_obj.save(path)
    return path


def _dl_path(name: str) -> str:
    return os.path.join(DOWNLOAD_DIR, name)


def _load_json_file(path: str):
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def _push_event(stream_id: str, data: dict):
    with _sse_lock:
        if stream_id in _sse_streams:
            _sse_streams[stream_id].append(data)


def _close_stream(stream_id: str):
    _push_event(stream_id, {"type": "done"})


def _detect_binary_smap(path: str):
    """
    Detect the newer binary/ZIP-packaged .smap export format.

    Plain-JSON .smap files begin with whitespace + `{`; ZIP-packaged ones
    begin with the local-file-header signature `PK\\x03\\x04` (or, very rarely,
    the central-directory `PK\\x05\\x06` / spanned-archive `PK\\x07\\x08`).
    Returns a user-facing message when the file is binary, else None.
    """
    try:
        with open(path, "rb") as fh:
            magic = fh.read(4)
    except OSError:
        return None
    if len(magic) >= 2 and magic[:2] == b"PK":
        return ("This .smap is a binary ZIP-packaged file (newer export format), "
                "not plain JSON. Re-export from your map software in 'plain JSON' "
                "format, or unzip the .smap and use the .json inside.")
    return None


def _ssh_connect(host: str, port: int, username: str,
                  key_path: str = "", key_passphrase: str = "",
                  password: str = "") -> paramiko.SSHClient:
    """Opens an SSH connection to the RoboShop VM. Uses AutoAddPolicy — these
    are internal automation VMs the user names explicitly, not arbitrary
    hosts, so we skip strict host-key verification for convenience.

    Prefers a private key (key_path) when given — many VMs (root login in
    particular) only accept publickey auth and never even offer password
    auth. Falls back to password only when no key path is supplied.
    """
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    if key_path:
        client.connect(host, port=port, username=username,
                        key_filename=os.path.expanduser(key_path),
                        passphrase=key_passphrase or None, timeout=15)
    else:
        client.connect(host, port=port, username=username, password=password, timeout=15)
    return client


def _ssh_exec(client: paramiko.SSHClient, command: str) -> tuple[int, str, str]:
    """Runs one command over an open SSH connection, blocking until it exits.
    Returns (exit_code, stdout, stderr)."""
    _, stdout, stderr = client.exec_command(command)
    exit_code = stdout.channel.recv_exit_status()
    return exit_code, stdout.read().decode("utf-8", "replace"), stderr.read().decode("utf-8", "replace")


# Standalone discovery script — mirrors discover_projects()/discover_projects_impl()
# from bind_parking_and_push_map.py / roboshop_agent.py, run on-demand over SSH
# (via `python3 -`) instead of needing a persistent agent process on the VM.
_DISCOVER_PROJECTS_SCRIPT = r"""
import os, json, re

def find_ip(data):
    ip_pattern = re.compile(r"(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(?:\.(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}")
    if isinstance(data, dict):
        for v in data.values():
            r = find_ip(v)
            if r: return r
    elif isinstance(data, list):
        for v in data:
            r = find_ip(v)
            if r: return r
    elif isinstance(data, str):
        m = ip_pattern.search(data)
        if m: return m.group()
    return None

def get_robots_root():
    override = os.environ.get("ROBOSHOP_ROBOTS_ROOT")
    if override and os.path.isdir(override):
        return override
    home = os.path.expanduser("~")
    for c in [
        os.path.join(home, ".local", "share", "RoboshopPro", "appInfo", "robots"),
        os.path.join(home, ".config", "RoboshopPro", "appInfo", "robots"),
        os.path.join(home, "RoboshopPro", "appInfo", "robots"),
        "/opt/RoboshopPro/appInfo/robots",
    ]:
        if os.path.isdir(c):
            return c
    return None

robots_root = get_robots_root()
projects = []
if robots_root and os.path.exists(robots_root):
    for category in os.listdir(robots_root):
        category_path = os.path.join(robots_root, category)
        if not os.path.isdir(category_path):
            continue
        for project_id in os.listdir(category_path):
            project_path = os.path.join(category_path, project_id)
            if not os.path.isdir(project_path):
                continue
            ip = "Unknown IP"
            robot_info = os.path.join(project_path, "robot_info.json")
            if os.path.exists(robot_info):
                try:
                    with open(robot_info, encoding="utf-8") as f:
                        ip_found = find_ip(json.load(f))
                        if ip_found:
                            ip = ip_found
                except Exception:
                    pass
            projects.append({"display": f"{category} ({ip})", "path": project_path, "ip": ip})

projects.sort(key=lambda p: (p["ip"] == "Unknown IP", p["display"].lower()))
print(json.dumps({"count": len(projects), "projects": projects}))
"""


# ─────────────────────────────────────────────
# Static / index
# ─────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("templates", "index.html")


@app.route("/static/<path:path>")
def static_files(path):
    return send_from_directory("static", path)


@app.route("/api/platform")
def platform_info():
    """Lets the frontend know whether this server is running on a Mac, so
    Map Sync can show/require the remote-agent fields (see IS_MAC_SERVER)."""
    return jsonify({"is_mac": IS_MAC_SERVER})


# ─────────────────────────────────────────────
# Downloads
# ─────────────────────────────────────────────

@app.route("/api/downloads/<filename>")
def download_file(filename):
    safe = os.path.basename(filename)
    path = os.path.join(DOWNLOAD_DIR, safe)
    if not os.path.exists(path):
        return jsonify({"error": "File not found"}), 404
    return send_file(path, as_attachment=True, download_name=safe)


@app.route("/api/sample-excel")
def sample_excel():
    """Generate and return a sample Point-to-Point Excel file."""
    out = _dl_path("sample_point_to_point.xlsx")
    wl.build_sample_excel(out)
    return send_file(out, as_attachment=True,
                     download_name="sample_point_to_point.xlsx")


# ─────────────────────────────────────────────
# Server-side directory listing for the native-style file picker.
# Browsers cannot expose the local disk path of a chosen file, so this lets
# the user navigate the server's filesystem (= their own machine, locally)
# and pick a real on-disk path that the backend can then read directly.
# ─────────────────────────────────────────────

_PICKER_EXTS = {".json", ".smap", ".xlsx"}


@app.route("/api/fs/list")
def fs_list():
    path = request.args.get("path", "").strip()
    ext_filter = request.args.get("ext", "").strip().lower()  # e.g. ".smap,.json"
    allowed = (set(e if e.startswith(".") else "." + e
                   for e in ext_filter.split(",") if e)
               if ext_filter else _PICKER_EXTS)

    if not path:
        # Defeat Pylance's static narrowing of os.name so the Unix branch is
        # analyzed too (we ship the same code on all platforms).
        platform_name = getattr(os, "name")
        if platform_name == "nt":
            import string
            drives = [f"{d}:\\" for d in string.ascii_uppercase
                      if os.path.exists(f"{d}:\\")]
            return jsonify({
                "path": "", "parent": None,
                "folders": drives, "files": [], "is_root": True
            })
        path = os.path.expanduser("~")

    if not os.path.isabs(path):
        return jsonify({"error": "Path must be absolute"}), 400
    if not os.path.isdir(path):
        return jsonify({"error": f"Not a directory: {path}"}), 400

    try:
        entries = os.listdir(path)
    except PermissionError:
        return jsonify({"error": "Permission denied"}), 403
    except OSError as e:
        return jsonify({"error": str(e)}), 400

    folders, files = [], []
    for name in entries:
        if name.startswith("."):
            continue
        full = os.path.join(path, name)
        try:
            if os.path.isdir(full):
                folders.append(name)
            elif os.path.isfile(full):
                ext = os.path.splitext(name)[1].lower()
                if ext in allowed:
                    files.append(name)
        except OSError:
            pass

    parent = os.path.dirname(path)
    if parent == path:  # already at root
        parent = "" if os.name == "nt" else None

    return jsonify({
        "path": path,
        "parent": parent,
        "folders": sorted(folders, key=str.lower),
        "files": sorted(files, key=str.lower),
        "is_root": False,
    })


# ─────────────────────────────────────────────
# Tab 1 — Each Pick
# ─────────────────────────────────────────────

@app.route("/api/each-pick/process", methods=["POST"])
def each_pick_process():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    include_work = request.form.get("include_work", "true").lower() == "true"
    include_resource = request.form.get("include_resource", "true").lower() == "true"

    path = _save_upload(f, "each_pick")
    try:
        bin_msg = _detect_binary_smap(path)
        if bin_msg:
            return jsonify({"error": bin_msg}), 422
        try:
            data = _load_json_file(path)
        except Exception as e:
            return jsonify({"error": f"Cannot parse JSON: {e}"}), 400
    finally:
        try:
            os.remove(path)
        except OSError:
            pass

    result = wl.run_marker_extraction(data, include_work, include_resource)
    if result["error"]:
        return jsonify({"error": result["error"]}), 422

    out_name = f"each_pick_{uuid.uuid4().hex}.xlsx"
    out_path = _dl_path(out_name)
    try:
        wl.save_marker_extraction_to_excel(
            result["sheets"], result.get("validation", {}), out_path)
    except Exception as e:
        return jsonify({"error": f"Failed to write Excel: {e}"}), 500

    return jsonify({
        "log": result["log"],
        "download": f"/api/downloads/{out_name}",
        "filename": "marker_extraction.xlsx"
    })


# ─────────────────────────────────────────────
# Tab 2 — Case Pick
# ─────────────────────────────────────────────

@app.route("/api/case-pick/analyze", methods=["POST"])
def case_pick_analyze():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    orientation = request.form.get("orientation", "horizontal")
    cross_aisle = request.form.get("cross_aisle", "false").lower() == "true"
    cross_aisle_sensitivity = float(request.form.get("sensitivity", "1.5"))
    # cross_aisle_markers: JSON array of point names
    ca_markers_raw = request.form.get("cross_aisle_markers", "[]")
    try:
        ca_markers = set(json.loads(ca_markers_raw))
    except Exception:
        ca_markers = set()

    path = _save_upload(f, "case_pick")
    try:
        bin_msg = _detect_binary_smap(path)
        if bin_msg:
            return jsonify({"error": bin_msg}), 422
        try:
            data = _load_json_file(path)
        except Exception as e:
            return jsonify({"error": f"Cannot parse JSON: {e}"}), 400
    finally:
        try:
            os.remove(path)
        except OSError:
            pass

    ok, msg = wl.validate_json_for_pick_sequence(data)
    if not ok:
        return jsonify({"error": msg}), 422

    graph, node_positions = wl.build_graph_from_json(data)
    aps = wl.get_action_points(data)
    if not aps:
        return jsonify({"error": "No ActionPoints found in JSON"}), 422

    aisles, sort_axis = wl.auto_group_aisles(aps, orientation)
    aisles = [a for a in aisles if a]

    if cross_aisle:
        aisles = wl.detect_all_cross_aisles(aisles, sort_axis, cross_aisle_sensitivity)

    if ca_markers:
        all_pts = [{"name": p.get("instanceName"), "x": float(p["pos"]["x"]),
                    "y": float(p["pos"]["y"])}
                   for p in data.get("advancedPointList", [])
                   if p.get("instanceName") and "pos" in p
                   and "x" in p["pos"] and "y" in p["pos"]]
        aisles = wl.split_aisles_by_markers(aisles, sort_axis, all_pts, ca_markers)

    aisles = [a for a in aisles if a]

    all_points_list = [
        {"name": p.get("instanceName"),
         "type": "AP" if p.get("className") == "ActionPoint" else "LM",
         "x": float(p["pos"]["x"]), "y": float(p["pos"]["y"])}
        for p in data.get("advancedPointList", [])
        if p.get("instanceName") and "pos" in p
        and "x" in p["pos"] and "y" in p["pos"]
    ]

    aisle_summary = [
        {
            "index": i + 1,
            "first": a[0]["name"],
            "last": a[-1]["name"],
            "count": len({ap["name"] for ap in a}),
            "aps": [ap["name"] for ap in a],
        }
        for i, a in enumerate(aisles)
    ]

    ap_names = [ap["name"] for ap in aps if ap["name"] in node_positions]

    return jsonify({
        "aisle_count": len(aisles),
        "ap_count": len(ap_names),
        "aisles": aisle_summary,
        "sort_axis": sort_axis,
        "all_points": all_points_list,
        "ap_names": ap_names,
        "log": [
            f"Graph built: {len(graph)} nodes",
            f"Grouped into {len(aisles)} aisles ({orientation})",
            f"Total APs: {len(ap_names)}",
        ]
    })


@app.route("/api/case-pick/generate", methods=["POST"])
def case_pick_generate():
    """
    Expects multipart/form-data with:
      file        — map JSON
      template    — (optional) template Excel
      mode        — "automatic" | "manual"
      orientation — "horizontal" | "vertical"
      loading_point — AP name (automatic mode)
      aisle_directions — JSON [{index, direction}] (manual mode)
      excluded_aps — JSON [name, ...]
      loading_aps  — JSON [name, ...]
      unloading_aps — JSON [name, ...]
      cross_aisle_markers — JSON [name, ...]
      cross_aisle — bool
      sensitivity — float
      duplication — int
    """
    f = request.files.get("file")
    template_file = request.files.get("template")
    mode = request.form.get("mode", "automatic")
    orientation = request.form.get("orientation", "horizontal")
    loading_point_name = request.form.get("loading_point", "")
    cross_aisle = request.form.get("cross_aisle", "false").lower() == "true"
    sensitivity = float(request.form.get("sensitivity", "1.5"))
    duplication = int(request.form.get("duplication", "0"))

    # Optional path inputs (user runs the app locally, so the server can read /
    # write any path the user supplies)
    map_path_in = request.form.get("map_path", "").strip()
    output_path_in = request.form.get("output_path", "").strip() or "case_pick_sequence.xlsx"

    def _parse_json_list(key):
        try:
            return json.loads(request.form.get(key, "[]"))
        except Exception:
            return []

    excluded_aps = set(_parse_json_list("excluded_aps"))
    loading_aps = set(_parse_json_list("loading_aps"))
    unloading_aps = set(_parse_json_list("unloading_aps"))
    ca_markers = set(_parse_json_list("cross_aisle_markers"))
    aisle_directions = _parse_json_list("aisle_directions")
    # Manual mode: frontend sends the exact aisle structure it showed in the table
    # so per-aisle directions stay aligned. Each entry: {aps:[name,...], direction:"f2l"|"l2f"}.
    manual_aisles = _parse_json_list("manual_aisles")

    template_path = ""
    if template_file and template_file.filename:
        template_path = _save_upload(template_file, "template")

    # ── Resolve input source ─────────────────────────────────────────
    # Prefer a real on-disk path; fall back to the uploaded file.
    data = None
    map_path_resolved = ""    # set when we used a path on disk (used for output-dir derivation)
    if map_path_in and os.path.exists(map_path_in):
        bin_msg = _detect_binary_smap(map_path_in)
        if bin_msg:
            return jsonify({"error": bin_msg}), 422
        try:
            data = _load_json_file(map_path_in)
            map_path_resolved = os.path.abspath(map_path_in)
        except Exception as e:
            return jsonify({"error": f"Cannot read map at path: {e}"}), 400
    elif f is not None:
        upload_tmp = _save_upload(f, "gen_map")
        try:
            bin_msg = _detect_binary_smap(upload_tmp)
            if bin_msg:
                return jsonify({"error": bin_msg}), 422
            try:
                data = _load_json_file(upload_tmp)
            except Exception as e:
                return jsonify({"error": f"Cannot parse JSON: {e}"}), 400
        finally:
            try:
                os.remove(upload_tmp)
            except OSError:
                pass
    else:
        return jsonify({"error": "No map file uploaded and no Map Path provided"}), 400

    ok, msg = wl.validate_json_for_pick_sequence(data)
    if not ok:
        return jsonify({"error": msg}), 422

    graph, node_positions = wl.build_graph_from_json(data)
    aps = wl.get_action_points(data)
    all_excluded = excluded_aps | loading_aps | unloading_aps

    aps_filtered = [ap for ap in aps if ap["name"] not in all_excluded]
    aisles, sort_axis = wl.auto_group_aisles(aps_filtered, orientation)
    aisles = [a for a in aisles if a]

    if cross_aisle:
        aisles = wl.detect_all_cross_aisles(aisles, sort_axis, sensitivity)

    if ca_markers:
        all_pts = [{"name": p.get("instanceName"), "x": float(p["pos"]["x"]),
                    "y": float(p["pos"]["y"])}
                   for p in data.get("advancedPointList", [])
                   if p.get("instanceName") and "pos" in p
                   and "x" in p["pos"] and "y" in p["pos"]]
        aisles = wl.split_aisles_by_markers(aisles, sort_axis, all_pts, ca_markers)

    aisles = [[ap for ap in a if ap["name"] in node_positions] for a in aisles]
    aisles = [a for a in aisles if a]

    log = []

    aisle_groups = None  # populated in manual branch; stays None for automatic
    if mode == "automatic":
        if not loading_point_name:
            return jsonify({"error": "loading_point is required for automatic mode"}), 400
        sequence = wl.generate_automatic_sequence(
            aisles, loading_point_name, graph, node_positions,
            log_fn=log.append)
        flat = [ap["name"] for ap in sequence]
        log.append(f"Automatic sequence: {len(flat)} APs")
    else:
        # MANUAL MODE — honour the exact aisle structure + direction the user
        # configured in the table. Use the frontend-supplied `manual_aisles` if
        # available (the only way to guarantee index alignment); otherwise fall
        # back to recomputing + applying `aisle_directions` by index.
        # Also build `aisle_groups` (one (name, [ap…]) tuple per aisle) so the
        # writer can emit the "aisle_AP" sheet.
        flat = []
        aisle_groups = []
        if manual_aisles:
            for i, entry in enumerate(manual_aisles, 1):
                aps_list = entry.get("aps", []) or []
                direction = entry.get("direction", "f2l")
                raw_name = (entry.get("name") or "").strip()
                clean = [n for n in aps_list
                         if n not in all_excluded and n in node_positions]
                if not clean:
                    continue
                if direction == "l2f":
                    clean = list(reversed(clean))
                flat.extend(clean)
                aisle_groups.append((raw_name or f"Aisle {i}", clean))
        else:
            dir_map = {d["index"] - 1: d["direction"] for d in aisle_directions}
            for i, aisle in enumerate(aisles):
                direction = dir_map.get(i, "f2l")
                clean = [ap for ap in aisle if ap["name"] not in all_excluded]
                if not clean:
                    continue
                if direction == "l2f":
                    clean = list(reversed(clean))
                names = [ap["name"] for ap in clean]
                flat.extend(names)
                aisle_groups.append((f"Aisle {i + 1}", names))
        log.append(f"Manual sequence: {len(flat)} APs")

    # ── Resolve output destination ──────────────────────────────────
    # - User entered a path with a directory (absolute OR with separator) → write there.
    # - User entered only a filename AND we read the map from disk → save next to the map.
    # - User entered only a filename AND map came from upload → save to server downloads dir
    #   (browser will download it).
    has_dir = bool(os.path.dirname(output_path_in))
    save_to_disk = False
    if has_dir:
        final_path = output_path_in
        save_to_disk = True
    elif map_path_resolved:
        final_path = os.path.join(os.path.dirname(map_path_resolved), output_path_in)
        save_to_disk = True
    else:
        # bare filename + no map path → server downloads dir, browser downloads
        out_name = f"case_pick_{uuid.uuid4().hex}_{output_path_in}"
        final_path = _dl_path(out_name)

    if save_to_disk:
        try:
            parent = os.path.dirname(final_path)
            if parent:
                os.makedirs(parent, exist_ok=True)
        except OSError as e:
            return jsonify({"error": f"Cannot create output directory: {e}"}), 400

    try:
        wl.save_sequence_to_excel(
            flat, final_path,
            loading_aps, unloading_aps,
            data, graph, node_positions,
            duplication=duplication,
            template_path=template_path,
            aisle_groups=aisle_groups if mode != "automatic" else None,
        )
    except Exception as e:
        return jsonify({"error": f"Failed to write Excel: {e}"}), 500
    finally:
        if template_path:
            try:
                os.remove(template_path)
            except OSError:
                pass

    log.append(f"Saved {len(flat)} AP(s) to Excel")
    log.append(f"Output: {os.path.abspath(final_path)}")

    if save_to_disk:
        return jsonify({
            "ap_count": len(flat),
            "log": log,
            "saved_to": os.path.abspath(final_path),
        })
    return jsonify({
        "ap_count": len(flat),
        "log": log,
        "download": f"/api/downloads/{os.path.basename(final_path)}",
        "filename": output_path_in,
    })


# ─────────────────────────────────────────────
# Tab 3 — Point to Point
# ─────────────────────────────────────────────

@app.route("/api/point-to-point/generate", methods=["POST"])
def point_to_point_generate():
    """
    Expects multipart/form-data:
      json_file      — map JSON / SMAP
      excel_file     — zone config Excel
      entry_exit_type — "LM" | "AP"
      ndeep          — "true" | "false"
      drop_seq       — "f2l" | "l2f"  (N-Deep only)
      zone_scope     — "BOTH"|"ENTRY"|"EXIT"
      sub_scope      — "BOTH"|"ENTRY"|"EXIT"
      loc_scope      — "BOTH"|"ENTRY"|"EXIT"
      create_sequence — "true"|"false"
      subzone_directions — JSON {subzone_name: direction}
    """
    if "json_file" not in request.files:
        return jsonify({"error": "json_file is required"}), 400

    json_f = request.files["json_file"]
    excel_f = request.files.get("excel_file")
    drawn_zones_raw = request.form.get("drawn_zones", "").strip()

    if not excel_f and not drawn_zones_raw:
        return jsonify({"error": "Provide either excel_file or drawn_zones"}), 400

    entry_exit_type = request.form.get("entry_exit_type", "LM")
    ndeep = request.form.get("ndeep", "false").lower() == "true"
    drop_seq = request.form.get("drop_seq", "f2l")
    zone_scope = request.form.get("zone_scope", "BOTH")
    sub_scope = request.form.get("sub_scope", "BOTH")
    loc_scope = request.form.get("loc_scope", "BOTH")
    create_sequence = request.form.get("create_sequence", "false").lower() == "true"
    try:
        subzone_dirs = json.loads(request.form.get("subzone_directions", "{}"))
    except Exception:
        subzone_dirs = {}

    json_path = _save_upload(json_f, "p2p_json")

    try:
        bin_msg = _detect_binary_smap(json_path)
        if bin_msg:
            return jsonify({"error": bin_msg}), 422
        try:
            mapping = _load_json_file(json_path)
        except Exception as e:
            return jsonify({"error": f"Cannot parse JSON: {e}"}), 400
    finally:
        try:
            os.remove(json_path)
        except OSError:
            pass

    # ── Build the zone-definition DataFrame ─────────────────────────
    if drawn_zones_raw:
        # User drew zones on the map → synthesise the DataFrame in-memory.
        try:
            drawn = json.loads(drawn_zones_raw)
        except Exception as e:
            return jsonify({"error": f"Bad drawn_zones JSON: {e}"}), 400
        if not drawn:
            return jsonify({"error": "drawn_zones is empty"}), 400
        rows = []
        for z in drawn:
            row = {
                "Zone name":  z.get("name",  ""),
                "zone entry": z.get("entry", ""),
                "zone exit":  z.get("exit",  ""),
                "min x": float(z.get("minX", 0)),
                "max x": float(z.get("maxX", 0)),
                "min y": float(z.get("minY", 0)),
                "max y": float(z.get("maxY", 0)),
            }
            if not ndeep:
                subzones = z.get("subzones", [])
                if subzones:
                    # User drew explicit subzones — create one row per subzone
                    for sz in subzones:
                        sz_row = dict(row)
                        sz_row["subzone name"]   = sz.get("name", "")
                        sz_row["subzone entery"] = sz.get("entry", "")
                        sz_row["subzone exit"]   = sz.get("exit",  "")
                        sz_row["min x"] = float(sz.get("minX", row["min x"]))
                        sz_row["max x"] = float(sz.get("maxX", row["max x"]))
                        sz_row["min y"] = float(sz.get("minY", row["min y"]))
                        sz_row["max y"] = float(sz.get("maxY", row["max y"]))
                        rows.append(sz_row)
                    continue  # skip the outer rows.append — zone rows already added
                else:
                    # No subzones drawn → auto-create one named <zone>_1
                    row["subzone name"]   = f"{z.get('name', '')}_1"
                    row["subzone entery"] = z.get("entry", "")
                    row["subzone exit"]   = z.get("exit",  "")
            rows.append(row)
        df = pd.DataFrame(rows)
    else:
        excel_path = _save_upload(excel_f, "p2p_excel")
        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            return jsonify({"error": f"Cannot read Excel: {e}"}), 400
        finally:
            try:
                os.remove(excel_path)
            except OSError:
                pass

    log = []
    graph, node_positions = wl.build_graph_from_json(mapping)
    log.append(f"Graph: {len(graph)} nodes")

    try:
        if ndeep:
            ok, msg = wl.validate_excel_for_ndeep(df)
            if not ok:
                return jsonify({"error": msg}), 422
            log.append("N-Deep mode")
            result = wl.generate_ndeep_zone_configuration(
                mapping, df, drop_seq,
                zone_scope, sub_scope, loc_scope,
                entry_exit_type, graph, node_positions,
                log_fn=log.append)
        elif create_sequence:
            ok, msg = wl.validate_excel_for_zone(df)
            if not ok:
                return jsonify({"error": msg}), 422
            log.append("Standard mode with sequences")
            seq_configs = {sz: {"direction": d} for sz, d in subzone_dirs.items()}
            result = wl.generate_zone_configuration_with_sequences(
                mapping, df, seq_configs, entry_exit_type, graph, node_positions)
        else:
            ok, msg = wl.validate_excel_for_zone(df)
            if not ok:
                return jsonify({"error": msg}), 422
            log.append("Standard mode")
            result = wl.generate_zone_configuration(
                mapping, df, entry_exit_type, graph, node_positions)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    out_name = f"zone_config_{uuid.uuid4().hex}.json"
    out_path = _dl_path(out_name)
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(result, fh, indent=2)

    n_zones = len(result.get("LocationDataMapping", {}).get("zones", []))
    log.append(f"Generated {n_zones} zone(s)")

    return jsonify({
        "zones": n_zones,
        "log": log,
        "download": f"/api/downloads/{out_name}",
        "filename": "zone_configuration.json"
    })


# ─────────────────────────────────────────────
# Tab 4 — Map Validator (SSE streaming)
# ─────────────────────────────────────────────

@app.route("/api/map-validator/validate", methods=["POST"])
def map_validator_validate():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    detailed_review = request.form.get("detailed_review") in ("1", "true", "True", "on")
    stream_id = uuid.uuid4().hex
    path = _save_upload(f, "mapval")

    with _sse_lock:
        _sse_streams[stream_id] = []

    def _worker():
        try:
            _push_event(stream_id, {"type": "log", "level": "info",
                                     "msg": "Loading file…"})
            # Detect binary (ZIP-packaged) .smap files before attempting JSON parse
            bin_msg = _detect_binary_smap(path)
            if bin_msg:
                _push_event(stream_id, {"type": "log", "level": "err",
                                         "msg": bin_msg})
                _close_stream(stream_id)
                return
            try:
                data = _load_json_file(path)
            except Exception as e:
                _push_event(stream_id, {"type": "log", "level": "err",
                                         "msg": f"Cannot parse file: {e}"})
                _close_stream(stream_id)
                return

            _push_event(stream_id, {"type": "log", "level": "info",
                                     "msg": "Running validation…"})
            result = wl.run_map_validation(data, detailed_review=detailed_review)

            if detailed_review:
                n_removed = result.get("n_removed", 0)
                if n_removed:
                    names = [p["name"] for p in result["removed_points"]]
                    preview = ", ".join(names[:10])
                    if n_removed > 10:
                        preview += f", +{n_removed - 10} more"
                    _push_event(stream_id, {
                        "type": "log", "level": "info",
                        "msg": f"Detailed Map Review: removed {n_removed} bidirectional-only "
                               f"point(s): {preview}"
                    })
                else:
                    _push_event(stream_id, {"type": "log", "level": "info",
                                             "msg": "Detailed Map Review: no bidirectional-only "
                                                    "points to remove."})

            total = result["total"]
            n_issues = result["n_issues"]
            n_conn = result["n_conn"]
            n_sccs = result["n_sccs"]

            if n_issues == 0:
                _push_event(stream_id, {
                    "type": "log", "level": "ok",
                    "msg": f"All {total} points have valid entry and exit paths."
                })
            else:
                _push_event(stream_id, {
                    "type": "log", "level": "warn",
                    "msg": f"Found {n_issues} entry/exit issue(s)."
                })

            _push_event(stream_id, {"type": "log", "level": "info",
                                     "msg": "Checking global reachability…"})

            if n_conn == 0:
                _push_event(stream_id, {"type": "log", "level": "ok",
                                         "msg": "All points are mutually reachable."})
            else:
                _push_event(stream_id, {
                    "type": "log", "level": "warn",
                    "msg": f"Graph has {n_sccs} disconnected group(s), "
                           f"{n_conn} unreachable direction(s)."
                })

            _push_event(stream_id, {"type": "result", "data": result})

        except Exception as e:
            _push_event(stream_id, {"type": "log", "level": "err",
                                     "msg": f"Validation error: {e}"})
        finally:
            try:
                os.remove(path)
            except OSError:
                pass
            _close_stream(stream_id)

    threading.Thread(target=_worker, daemon=True).start()
    return jsonify({"stream_id": stream_id})


@app.route("/api/map-validator/stream/<stream_id>")
def map_validator_stream(stream_id):
    """SSE endpoint — client subscribes after getting stream_id."""
    def _generate():
        sent = 0
        timeout = 300  # 5 min max
        deadline = time.time() + timeout
        while time.time() < deadline:
            with _sse_lock:
                events = _sse_streams.get(stream_id, [])
                new_events = events[sent:]
            for ev in new_events:
                sent += 1
                yield f"data: {json.dumps(ev)}\n\n"
                if ev.get("type") == "done":
                    with _sse_lock:
                        _sse_streams.pop(stream_id, None)
                    return
            time.sleep(0.15)
        yield f"data: {json.dumps({'type': 'done', 'timeout': True})}\n\n"
        with _sse_lock:
            _sse_streams.pop(stream_id, None)

    return Response(_generate(), mimetype="text/event-stream",
                    headers={
                        "Cache-Control": "no-cache",
                        "X-Accel-Buffering": "no",
                    })


# ─────────────────────────────────────────────
# Tab 5 — Orderfile Generator
# ─────────────────────────────────────────────

@app.route("/api/orderfile/analyze", methods=["POST"])
def orderfile_analyze():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    path = _save_upload(f, "orderfile")
    try:
        df = pd.read_excel(path)
    except Exception as e:
        return jsonify({"error": f"Cannot read Excel: {e}"}), 400
    finally:
        try:
            os.remove(path)
        except OSError:
            pass

    df.columns = [c.strip().lower() for c in df.columns]
    missing = [c for c in ['order', 'line', 'unit'] if c not in df.columns]
    if missing:
        return jsonify({
            "error": f"Missing required columns: {', '.join(missing)}. "
                     f"Found: {list(df.columns)}"
        }), 422

    df = df[['order', 'line', 'unit']].dropna()
    df['line'] = pd.to_numeric(df['line'], errors='coerce').fillna(0)
    df['unit'] = pd.to_numeric(df['unit'], errors='coerce').fillna(0)

    return jsonify({
        "raw_orders": int(df['order'].nunique()),
        "raw_lines":  int(df['line'].sum()),
        "raw_units":  float(df['unit'].sum()),
    })


@app.route("/api/orderfile/generate", methods=["POST"])
def orderfile_generate():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    path = _save_upload(f, "orderfile")
    try:
        df_raw = pd.read_excel(path)
    except Exception as e:
        return jsonify({"error": f"Cannot read Excel: {e}"}), 400
    finally:
        try:
            os.remove(path)
        except OSError:
            pass

    df_raw.columns = [c.strip().lower() for c in df_raw.columns]
    missing = [c for c in ['order', 'line', 'unit'] if c not in df_raw.columns]
    if missing:
        return jsonify({"error": f"Missing required columns: {', '.join(missing)}"}), 422

    df_raw = df_raw[['order', 'line', 'unit']].dropna()
    df_raw['line'] = pd.to_numeric(df_raw['line'], errors='coerce').fillna(0)
    df_raw['unit'] = pd.to_numeric(df_raw['unit'], errors='coerce').fillna(0)

    try:
        target_orders = int(request.form.get("target_orders", 100))
        target_lines  = int(request.form.get("target_lines", 500))
        target_units  = float(request.form.get("target_units", 5000))
    except (ValueError, TypeError) as e:
        return jsonify({"error": f"Invalid target parameters: {e}"}), 400

    use_limits = request.form.get("use_limits", "false").lower() == "true"

    def _opt_float(key):
        v = request.form.get(key, "").strip()
        return float(v) if use_limits and v else None

    max_qty_per_order = _opt_float("max_qty_per_order")
    min_qty_per_order = _opt_float("min_qty_per_order")
    max_qty_per_line  = _opt_float("max_qty_per_line")

    stream_id = uuid.uuid4().hex
    with _sse_lock:
        _sse_streams[stream_id] = []

    def _worker():
        try:
            _push_event(stream_id, {"type": "progress", "pct": 1,
                                     "msg": "Initializing optimizer…"})
            optimizer = wl.OrderfileOptimizer(
                df_raw, target_orders, target_lines, target_units,
                max_qty_per_order=max_qty_per_order,
                min_qty_per_order=min_qty_per_order,
                max_qty_per_line=max_qty_per_line,
            )

            def _progress_cb(pct, msg):
                _push_event(stream_id, {"type": "progress", "pct": pct, "msg": msg})

            result_df, actual_orders, actual_lines, actual_qty = optimizer.optimize(
                progress_cb=_progress_cb, total_time=120
            )

            if result_df.empty:
                _push_event(stream_id, {"type": "log", "level": "warn",
                                         "msg": "No suitable subset found. Try adjusting targets."})
                _close_stream(stream_id)
                return

            out_name = f"orderfile_{uuid.uuid4().hex}.xlsx"
            out_path = _dl_path(out_name)
            wl.orderfile_to_excel(result_df, out_path)

            _push_event(stream_id, {
                "type": "result",
                "stats": {
                    "actual_orders": actual_orders,
                    "actual_lines":  actual_lines,
                    "actual_qty":    actual_qty,
                    "target_orders": target_orders,
                    "target_lines":  target_lines,
                    "target_units":  target_units,
                },
                "download": f"/api/downloads/{out_name}",
                "filename": "orderfile_output.xlsx",
            })
        except Exception as e:
            _push_event(stream_id, {"type": "log", "level": "err",
                                     "msg": f"Optimization error: {e}"})
        finally:
            _close_stream(stream_id)

    threading.Thread(target=_worker, daemon=True).start()
    return jsonify({"stream_id": stream_id})


@app.route("/api/orderfile/stream/<stream_id>")
def orderfile_stream(stream_id):
    def _generate():
        sent = 0
        deadline = time.time() + 300
        while time.time() < deadline:
            with _sse_lock:
                new_events = _sse_streams.get(stream_id, [])[sent:]
            for ev in new_events:
                sent += 1
                yield f"data: {json.dumps(ev)}\n\n"
                if ev.get("type") == "done":
                    with _sse_lock:
                        _sse_streams.pop(stream_id, None)
                    return
            time.sleep(0.15)
        yield f"data: {json.dumps({'type': 'done', 'timeout': True})}\n\n"
        with _sse_lock:
            _sse_streams.pop(stream_id, None)

    return Response(_generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ─────────────────────────────────────────────
# Tab 6 — Robot Map Sync
# ─────────────────────────────────────────────

@app.route("/api/map-sync/vm-projects", methods=["POST"])
def map_sync_vm_projects():
    """SSHes into the RoboShop VM and runs a one-off discovery script (no
    persistent agent process needed) so the browser gets a pickable list of
    RoboShop systems (name + IP + path) instead of the user typing a raw
    path blind."""
    data = request.get_json(silent=True) or {}
    host           = (data.get("ssh_host") or "").strip()
    username       = (data.get("ssh_username") or "").strip()
    key_path       = (data.get("ssh_key_path") or "").strip()
    key_passphrase = data.get("ssh_key_passphrase") or ""
    password       = data.get("ssh_password") or ""
    try:
        port = int(data.get("ssh_port") or 22)
    except (TypeError, ValueError):
        return jsonify({"error": "Invalid SSH port"}), 400

    if not host:
        return jsonify({"error": "VM host is required"}), 400
    if not username:
        return jsonify({"error": "SSH username is required"}), 400

    try:
        client = _ssh_connect(host, port, username, key_path, key_passphrase, password)
    except Exception as e:
        return jsonify({"error": f"Could not connect over SSH: {e}"}), 502

    try:
        stdin, stdout, stderr = client.exec_command("python3 -")
        stdin.write(_DISCOVER_PROJECTS_SCRIPT)
        stdin.channel.shutdown_write()
        exit_code = stdout.channel.recv_exit_status()
        out = stdout.read().decode("utf-8", "replace")
        err = stderr.read().decode("utf-8", "replace")
        if exit_code != 0:
            return jsonify({"error": f"Discovery script failed: {err or out}"}), 502
        payload = json.loads(out)
    except Exception as e:
        return jsonify({"error": f"Discovery failed: {e}"}), 502
    finally:
        client.close()

    return jsonify(payload)


@app.route("/api/map-sync/execute", methods=["POST"])
def map_sync_execute():
    import shlex
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    base_dir     = request.form.get("base_dir",   "").strip()
    prefix       = request.form.get("prefix",     "").strip()
    sub_folder   = request.form.get("sub_folder", "").strip()
    ssh_host       = request.form.get("ssh_host",     "").strip()
    ssh_username   = request.form.get("ssh_username", "").strip()
    ssh_key_path   = request.form.get("ssh_key_path", "").strip()
    ssh_key_pass   = request.form.get("ssh_key_passphrase", "")
    ssh_password   = request.form.get("ssh_password", "")
    # Explicit client-side choice ("I'm using a Mac" checkbox) rather than
    # IS_MAC_SERVER (the Flask process's own OS) — the two can disagree, e.g.
    # when warehouse-tool runs inside a Linux Docker container on a Mac host.
    use_remote_agent = request.form.get("use_remote_agent") in ("1", "true", "True", "on")

    try:
        start = int(request.form.get("start", 1))
        count = int(request.form.get("count", 0))
        ssh_port = int(request.form.get("ssh_port") or 22)
    except ValueError as e:
        return jsonify({"error": f"Invalid start/count/port: {e}"}), 400

    if not base_dir:
        return jsonify({"error": "Parent directory is required"}), 400
    if not prefix:
        return jsonify({"error": "Folder prefix is required"}), 400
    if count <= 0:
        return jsonify({"error": "Number of robots must be > 0"}), 400

    if use_remote_agent:
        # No local filesystem access to the robot folders from here — the
        # path only makes sense on the VM, so skip the local os.path.isdir
        # check and require SSH credentials for that VM instead.
        if not ssh_host:
            return jsonify({"error": "VM host is required"}), 400
        if not ssh_username:
            return jsonify({"error": "SSH username is required"}), 400
    elif not os.path.isdir(base_dir):
        return jsonify({"error": f"Directory not found: {base_dir}"}), 400

    src_path = _save_upload(f, "mapsync")
    filename = f.filename or os.path.basename(src_path)
    stream_id = uuid.uuid4().hex

    with _sse_lock:
        _sse_streams[stream_id] = []

    def _worker():
        success = failed = 0
        try:
            if use_remote_agent:
                _push_event(stream_id, {"type": "log", "level": "info",
                                         "msg": f"Connecting to {ssh_host} over SSH…"})
                try:
                    client = _ssh_connect(ssh_host, ssh_port, ssh_username,
                                           ssh_key_path, ssh_key_pass, ssh_password)
                    sftp = client.open_sftp()
                except Exception as e:
                    _push_event(stream_id, {"type": "log", "level": "err",
                                             "msg": f"SSH connection failed: {e}"})
                    _push_event(stream_id, {"type": "result", "success": 0,
                                             "failed": count, "total": count})
                    return

                try:
                    for i in range(count):
                        robot_id    = f"{prefix}{start + i}"
                        folder_path = posixpath.join(base_dir, robot_id, sub_folder) if sub_folder \
                                      else posixpath.join(base_dir, robot_id)
                        dest = posixpath.join(folder_path, filename)
                        try:
                            exit_code, _, err = _ssh_exec(client, f"mkdir -p {shlex.quote(folder_path)}")
                            if exit_code != 0:
                                raise RuntimeError(err.strip() or f"mkdir exited {exit_code}")
                            sftp.put(src_path, dest)
                            success += 1
                            _push_event(stream_id, {
                                "type": "progress",
                                "pct": int((i + 1) / count * 100),
                                "msg": f"[OK]   {folder_path}",
                                "ok": True,
                            })
                        except Exception as e:
                            failed += 1
                            _push_event(stream_id, {
                                "type": "progress",
                                "pct": int((i + 1) / count * 100),
                                "msg": f"[FAIL] {folder_path}  —  {e}",
                                "ok": False,
                            })
                finally:
                    sftp.close()
                    client.close()

                _push_event(stream_id, {
                    "type": "result", "success": success, "failed": failed, "total": count,
                })
                return

            for i in range(count):
                robot_id    = f"{prefix}{start + i}"
                folder_path = os.path.join(base_dir, robot_id, sub_folder) if sub_folder \
                              else os.path.join(base_dir, robot_id)
                dest = os.path.join(folder_path, filename)
                try:
                    os.makedirs(folder_path, exist_ok=True)
                    shutil.copy2(src_path, dest)
                    success += 1
                    _push_event(stream_id, {
                        "type": "progress",
                        "pct": int((i + 1) / count * 100),
                        "msg": f"[OK]   {folder_path}",
                        "ok": True,
                    })
                except Exception as e:
                    failed += 1
                    _push_event(stream_id, {
                        "type": "progress",
                        "pct": int((i + 1) / count * 100),
                        "msg": f"[FAIL] {folder_path}  —  {e}",
                        "ok": False,
                    })
            _push_event(stream_id, {
                "type": "result",
                "success": success,
                "failed": failed,
                "total": count,
            })
        except Exception as e:
            _push_event(stream_id, {"type": "log", "level": "err", "msg": f"Error: {e}"})
        finally:
            try:
                os.remove(src_path)
            except OSError:
                pass
            _close_stream(stream_id)

    threading.Thread(target=_worker, daemon=True).start()
    return jsonify({"stream_id": stream_id})


@app.route("/api/map-sync/stream/<stream_id>")
def map_sync_stream(stream_id):
    def _generate():
        sent = 0
        deadline = time.time() + 600  # 10 min for large fleets
        while time.time() < deadline:
            with _sse_lock:
                new_events = _sse_streams.get(stream_id, [])[sent:]
            for ev in new_events:
                sent += 1
                yield f"data: {json.dumps(ev)}\n\n"
                if ev.get("type") == "done":
                    with _sse_lock:
                        _sse_streams.pop(stream_id, None)
                    return
            time.sleep(0.1)
        yield f"data: {json.dumps({'type': 'done', 'timeout': True})}\n\n"
        with _sse_lock:
            _sse_streams.pop(stream_id, None)

    return Response(_generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ─────────────────────────────────────────────
# Tab 7 — DWG → Map
# ─────────────────────────────────────────────

# No macOS build of ODA File Converter exists, so this only works when
# warehouse-tool runs inside the project's Docker image built with
# --build-arg INSTALL_ODA=true (see the Dockerfile for setup steps).
ODA_FILE_CONVERTER = os.environ.get("ODA_FILE_CONVERTER_PATH", "/usr/bin/ODAFileConverter")


def _convert_dwg_to_dxf(dwg_path: str) -> str:
    """Converts a .dwg file to .dxf via the ODA File Converter CLI, run under
    a virtual display (xvfb) since it's a Qt GUI app even in batch mode.
    Returns the path to the resulting .dxf file, saved under UPLOAD_DIR.

    On success the tool exits on its own once it's done (confirmed against
    a real file — an empty input folder is the one case it never exits, but
    that can't happen here since we always seed exactly one file).
    On a bad/corrupt .dwg it still exits cleanly, but writes a
    "<name>.dxf.err" file instead of "<name>.dxf" — read that back so the
    caller gets ODA's actual reason instead of a generic failure message.
    """
    if not os.path.exists(ODA_FILE_CONVERTER):
        raise RuntimeError(
            f"ODA File Converter not found at {ODA_FILE_CONVERTER}. This step "
            "only runs inside the Docker image built with --build-arg "
            "INSTALL_ODA=true — see the Dockerfile for setup steps."
        )

    in_dir = tempfile.mkdtemp(prefix="dwg_in_")
    out_dir = tempfile.mkdtemp(prefix="dwg_out_")
    try:
        dwg_name = os.path.basename(dwg_path)
        shutil.copy2(dwg_path, os.path.join(in_dir, dwg_name))
        dxf_name = os.path.splitext(dwg_name)[0] + ".dxf"

        # ODAFileConverter <in_dir> <out_dir> <out_ver> <out_type> <recurse> <audit>
        cmd = ["xvfb-run", "-a", ODA_FILE_CONVERTER,
               in_dir, out_dir, "ACAD2018", "DXF", "0", "1"]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode != 0:
            raise RuntimeError(f"ODA File Converter failed: {result.stderr.strip() or result.stdout.strip()}")

        converted_path = os.path.join(out_dir, dxf_name)
        if not os.path.exists(converted_path):
            err_path = os.path.join(out_dir, dxf_name + ".err")
            if os.path.exists(err_path):
                with open(err_path, encoding="utf-8", errors="replace") as fh:
                    raise RuntimeError(f"ODA File Converter could not read this file: {fh.read().strip()}")
            raise RuntimeError("ODA File Converter did not produce a .dxf output file")

        final_path = os.path.join(UPLOAD_DIR, f"dwgconv_{uuid.uuid4().hex}.dxf")
        shutil.move(converted_path, final_path)
        return final_path
    finally:
        shutil.rmtree(in_dir, ignore_errors=True)
        shutil.rmtree(out_dir, ignore_errors=True)


@app.route("/api/dwg-map/list-blocks", methods=["POST"])
def dwg_map_list_blocks():
    """Step 1 of DWG -> SMAP: upload a .dwg (or .dxf directly) and get back
    the distinct block definition names found in it. Runs in the background
    and streams progress over SSE (see /api/dwg-map/stream) since DWG->DXF
    conversion can take a few seconds."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    ext = os.path.splitext(f.filename or "")[1].lower()
    if ext not in (".dwg", ".dxf"):
        return jsonify({"error": "Please upload a .dwg or .dxf file"}), 400

    src_path = _save_upload(f, "dwgmap")
    stream_id = uuid.uuid4().hex

    with _sse_lock:
        _sse_streams[stream_id] = []

    def _worker():
        dxf_path = src_path
        converted = False
        try:
            _push_event(stream_id, {"type": "log", "level": "info",
                                     "msg": "File uploaded, starting…"})
            if ext == ".dwg":
                _push_event(stream_id, {"type": "log", "level": "info",
                                         "msg": "Converting DWG → DXF via ODA File Converter "
                                                "(this can take a few seconds)…"})
                dxf_path = _convert_dwg_to_dxf(src_path)
                converted = True
                _push_event(stream_id, {"type": "log", "level": "ok",
                                         "msg": "Conversion complete."})
            else:
                _push_event(stream_id, {"type": "log", "level": "info",
                                         "msg": "Already DXF — skipping conversion."})

            _push_event(stream_id, {"type": "log", "level": "info",
                                     "msg": "Scanning for block definitions…"})
            blocks = wl.list_dwg_block_names(dxf_path)
            _push_event(stream_id, {"type": "log", "level": "ok",
                                     "msg": f"Found {len(blocks)} block definition(s)."})

            _push_event(stream_id, {"type": "log", "level": "info",
                                     "msg": "Rendering block previews…"})
            previews = wl.render_dwg_block_previews(dxf_path, blocks)
            _push_event(stream_id, {"type": "log", "level": "ok", "msg": "Done."})

            _push_event(stream_id, {"type": "result", "blocks": blocks, "count": len(blocks),
                                     "previews": previews})
        except Exception as e:
            _push_event(stream_id, {"type": "log", "level": "err", "msg": str(e)})
        finally:
            try:
                os.remove(src_path)
            except OSError:
                pass
            if converted:
                try:
                    os.remove(dxf_path)
                except OSError:
                    pass
            _close_stream(stream_id)

    threading.Thread(target=_worker, daemon=True).start()
    return jsonify({"stream_id": stream_id})


@app.route("/api/dwg-map/stream/<stream_id>")
def dwg_map_stream(stream_id):
    """SSE endpoint — client subscribes after getting stream_id."""
    def _generate():
        sent = 0
        deadline = time.time() + 300  # 5 min max
        while time.time() < deadline:
            with _sse_lock:
                new_events = _sse_streams.get(stream_id, [])[sent:]
            for ev in new_events:
                sent += 1
                yield f"data: {json.dumps(ev)}\n\n"
                if ev.get("type") == "done":
                    with _sse_lock:
                        _sse_streams.pop(stream_id, None)
                    return
            time.sleep(0.15)
        yield f"data: {json.dumps({'type': 'done', 'timeout': True})}\n\n"
        with _sse_lock:
            _sse_streams.pop(stream_id, None)

    return Response(_generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ─────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────

if __name__ == "__main__":
    print("Starting Warehouse Tools Web Server…")
    print("Open http://127.0.0.1:5000 in your browser")
    app.run(debug=True, host="0.0.0.0", port=5000, threaded=True)
