# Warehouse Tools — Web Application

Flask-based web conversion of the desktop Warehouse Tools app. All four tools are available as a single-page application with the Nexus Modern glassmorphic design.

---

## Quick Start

```bash
# 1. Create and activate a virtual environment (recommended)
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # Linux / macOS

# 2. Install dependencies
pip install -r requirements.txt

# 3. Start the server
python app.py

# 4. Open in browser
#    http://127.0.0.1:5000
```

---

## File Structure

```
web/
├── app.py                  Flask server — routes and API endpoints
├── warehouse_logic.py      Pure Python business logic (no UI imports)
├── requirements.txt        Python dependencies
├── README.md               This file
├── templates/
│   └── index.html          Single-page application (HTML shell)
├── static/
│   ├── css/
│   │   └── style.css       Nexus Modern glassmorphic theme
│   └── js/
│       └── app.js          Frontend logic (file upload, SSE consumer)
├── uploads/                Temp dir — uploaded files (auto-cleaned)
└── downloads/              Generated output files served for download
```

---

## Tabs & Features

### 1. Each Pick
Extracts WorkMarker / ResourceMarker coordinates from a warehouse JSON/SMAP file and writes them to a grouped Excel spreadsheet.

**Input:** `.json` or `.smap` map file  
**Options:** Include WorkMarkers, Include ResourceMarkers  
**Output:** `marker_extraction.xlsx` (sheets: `WorkMarkers_Grouped`, `ResourceMarkers_Grouped`, missing-code sheets)

---

### 2. Case Pick
Generates pick sequences from a warehouse map by grouping ActionPoints into aisles and optimising traversal order.

**Step 1 — Analyze:**
- Upload map JSON/SMAP
- Choose orientation (Horizontal / Vertical)
- Optionally enable Auto Cross-Aisle detection with configurable sensitivity

**Step 2 — Generate:**
- Mode: Automatic (Dijkstra-ordered by loading point) or Manual (per-aisle direction toggle)
- Enable duplication count for repeated picks per location
- Optionally upload a template Excel to populate existing sheets

**Output:** `case_pick_sequence.xlsx` with sheets:
- `location-marker-mapping` — AP names with duplication
- `marker-sequence` — pick order
- `aisle-sequence` — aisle groupings
- `marker-config` — marker metadata header
- `station-config` — loading/unloading station rows with nearest-LM waitMarker

---

### 3. Point to Point
Generates zone configuration JSON from a map file + zone definition Excel.

**Modes:**
- **Standard:** Excel must contain columns: `Zone name | zone entry | zone exit | subzone name | subzone entery | subzone exit | min x | max x | min y | max y`
- **N-Deep:** Excel must contain: `Zone name | zone entry | zone exit | min x | max x | min y | max y`. Subzones are auto-generated from spatial + connectivity analysis.

**Options:**
- Entry/Exit Type: `LM` (LocationMark, default) or `AP` (ActionPoint)
- N-Deep: Drop sequence `f2l` / `l2f`; Zone / Subzone / Location scope (`BOTH` / `ENTRY` / `EXIT`)
- Standard: Create sequences for subzones (direction per subzone)

**Sample Excel:** Click "⬇ Sample Excel" to download a two-sheet sample (`Standard` and `N-Deep`).

**Output:** `zone_configuration.json` — `LocationDataMapping.zones` structure.

---

### 4. Map Validator
Validates a warehouse map for path integrity and global reachability.

**Checks:**
1. Every point has at least one entry path and one exit path
2. All points are mutually reachable (Tarjan SCC + condensation BFS)

**Output:** Live-streamed log via Server-Sent Events; issue table showing points with missing paths or connectivity gaps; summary badges.

---

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `GET`  | `/` | Serve the SPA |
| `POST` | `/api/each-pick/process` | Extract markers → Excel |
| `POST` | `/api/case-pick/analyze` | Analyze aisles from map |
| `POST` | `/api/case-pick/generate` | Generate pick sequence Excel |
| `POST` | `/api/point-to-point/generate` | Generate zone config JSON |
| `POST` | `/api/map-validator/validate` | Start validation (returns stream_id) |
| `GET`  | `/api/map-validator/stream/<id>` | SSE stream for validation progress |
| `GET`  | `/api/sample-excel` | Download sample zone config Excel |
| `GET`  | `/api/downloads/<filename>` | Download a generated file |

---

## Architecture Notes

- **`warehouse_logic.py`** contains all computation: graph building, Dijkstra, Tarjan SCC, aisle grouping, sequence generation, Excel/JSON writers. It has **zero UI imports** — fully testable in isolation.
- **`app.py`** is a thin Flask layer: receives uploads, calls logic functions, returns JSON or streams Server-Sent Events.
- Map validation runs in a **background thread** so large maps (600+ points) never block the server. Progress is pushed to the SSE stream in real time.
- Uploaded files are saved to `uploads/` with a UUID name and deleted immediately after processing. Generated outputs in `downloads/` are retained until the server restarts.

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `flask` | Web server and routing |
| `openpyxl` | Excel file I/O |
| `pandas` | Excel reading, marker grouping |
| `numpy` | IQR-based cross-aisle threshold |

No external CSS frameworks — all styling is custom CSS using CSS custom properties.
