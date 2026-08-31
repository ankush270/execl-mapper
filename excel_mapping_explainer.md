# Excel Mapping Codebase — Complete Function Explainer

> Every function across all three files explained with purpose, inputs, outputs, and examples.

---

## Architecture Overview

```
HTTP Request
    ↓
excel_manage_services.py   (Layer 1 — Routes / API)
    ↓
excel_manage_component.py  (Layer 2 — Orchestration / Business Logic)
    ↓
excel_manage_helper.py     (Layer 3 — Core Algorithms / Excel + JSON logic)
```

Think of it as a restaurant:
- **Services** = Waiter (takes order, delivers food)
- **Component** = Chef (decides what to cook, coordinates)
- **Helper**    = Kitchen tools (knife, oven — does the actual work)

---

## FILE 1 — `excel_manage_services.py`
### Role: HTTP Route Definitions (the API layer)

This file only does ONE thing per route: call the component and return the result as JSON.
No business logic lives here. It's purely a thin HTTP wrapper.

---

### `excel_mapping_bp = Blueprint("excel_mapping", __name__)`

**What it is:** Creates a Flask Blueprint — a group of related routes that can be registered onto the main Flask app.

Think of it as a "folder" for all `/json/*` and `/excel/*` URLs.

---

### Route: `POST /json/upload`  →  `upload_json()`

```python
result = excel_component.upload_json(request.files.get("file"))
return {"success": True, "data": result}, 201
```

**What it does:**
- Takes a file from the `multipart/form-data` request (`request.files`)
- Passes it to the component
- Returns `201 Created` with the new `document_id`

**On error:** Returns `400 Bad Request` with the error message.

---

### Route: `GET /json`  →  `get_jsons()`

```python
result = excel_component.get_all_json()
return {"success": True, "data": result}, 200
```

**What it does:** Returns a list of all uploaded JSON documents (without the actual data — just metadata).

---

### Route: `GET /json/<document_id>`  →  `get_json_by_id(document_id)`

```python
result = excel_component.get_json(document_id)
```

**What it does:** Fetches one specific JSON document including its full data, by MongoDB `_id`.

---

### Route: `GET /excel/templates`  →  `get_templates_route()`

```python
result = excel_component.get_templates()
```

**What it does:** Lists all `.xlsx` files available in the `templates/` folder.

---

### Route: `GET /excel/templates/<file_name>/preview`  →  `preview_template_route()`

```python
result = excel_component.preview_template(file_name)
```

**What it does:** Opens the Excel file and returns a raw preview of the first 20 rows of every sheet. Used to inspect what the template looks like before mapping.

---

### Route: `GET /excel/templates/<file_name>/headers?sheet_name=X`  →  `get_template_headers_route()`

```python
sheet_name = request.args.get("sheet_name")
result = excel_component.get_template_headers(file_name, sheet_name)
```

**What it does:** Reads the headers from a specific sheet of a template, auto-detecting whether they're horizontal, vertical, or mixed.

**Query param:** `sheet_name` is required. Example: `?sheet_name=Sales`

---

### Route: `POST /excel/generate`  →  `generate_excel_route()`

```python
data = request.get_json()
result = excel_component.generate_excel(data)
return {"success": True, "data": result}, 201
```

**What it does:** The main action endpoint. Takes a JSON body with `document_id`, `template_file`, `sheet_name` and generates a filled Excel file.

**Returns:** `export_id` and `file_name` of the generated file.

---

### Route: `GET /excel/download/<export_id>`  →  `download_excel_route()`

```python
file_path = excel_component.get_export_path(export_id)
return send_file(file_path, as_attachment=True)
```

**What it does:** Looks up the file path for a previously generated Excel by its `export_id` and streams it as a downloadable file attachment.

---
---

## FILE 2 — `excel_manage_component.py`
### Role: Business Logic & Orchestration

This class (`ExcelManageComponent`) is the brain. It knows WHAT needs to happen and in what ORDER, but delegates the HOW to the helper.

---

### `__init__(self)`

```python
self.db = Database()
self.helper = ExcelManageHelper()
```

**What it does:** Initializes the component when it first loads. Creates:
- `self.db` — MongoDB connection (for reading/writing all collections)
- `self.helper` — Instance of ExcelManageHelper (for Excel/JSON logic)

If either fails (e.g., MongoDB down), the whole server startup fails with a logged error.

---

### `upload_json(self, file)`

**Input:** A file object from Flask's `request.files`

**Step by step:**
1. Checks that a file was actually provided
2. Validates the filename ends with `.json`
3. Calls `json.load(file)` — if invalid JSON, raises a clear `JSONDecodeError` BEFORE saving to DB
4. Saves the document to MongoDB `json_documents` collection with:
   - `file_name`, `json_data`, `created_at`
5. Returns `{"document_id": "...", "file_name": "..."}`

**Key design decision:** JSON is validated BEFORE going into MongoDB. Earlier, bad JSON would silently save and fail later.

---

### `get_all_json(self)`

**What it does:**
- Queries MongoDB `json_collection`
- Uses `{"json_data": 0}` projection — this means "return everything EXCEPT json_data"
- Sorts newest-first (`sort("created_at", -1)`)
- Returns a list of metadata objects (no actual JSON content — that would be too heavy)

**Why exclude json_data?** A list view should be lightweight. You don't need the full data just to show a list of uploaded files.

---

### `get_json(self, document_id)`

**What it does:**
- Converts the string `document_id` into a MongoDB `ObjectId`
- Finds ONE document in the collection
- Returns the full document including `json_data`

**Error:** If no document found → `"JSON document not found for id: <id>"`

---

### `get_templates(self)`

**What it does:**
- Looks at the `templates/` folder on disk (path from `Config.TEMPLATE_FOLDER`)
- Iterates all files in that folder
- Keeps only files that are `.xlsx`
- Returns a list of `{"template_name": "Invoice", "file_name": "Invoice.xlsx"}`

**Edge case handled:** If the folder doesn't exist at all → descriptive error telling you to create it.

---

### `preview_template(self, file_name)`

**What it does:**
1. Builds the full path: `templates_folder / file_name`
2. Checks the file exists and is `.xlsx`
3. Opens it with `openpyxl` in `data_only=True` mode (so formulas show values, not `=SUM(...)`)
4. Calls `helper.get_excel_preview(workbook)` — returns first 20 rows of each sheet
5. Closes the workbook in a `finally` block (so it ALWAYS closes, even if an error occurs)

---

### `get_template_headers(self, file_name, sheet_name)`

**What it does:**
1. **Validates** `sheet_name` is provided (not empty)
2. Loads the workbook
3. Calls `_resolve_sheet_name()` — case-insensitive sheet matching
4. Calls `helper.get_excel_headers(sheet)` — auto-detects orientation and returns headers

**Returns:** A list of header dicts, each with `header_name`, `excel_column`, `orientation`, and layout-specific extras.

---

### `save_mapping(self, document_id, template_file, sheet_name, mappings)`

**What it does:**
- Saves the completed mapping relationship to MongoDB `mapping_collection`
- A "mapping" is a record of which JSON field was mapped to which Excel column
- This creates an audit trail — you can look back and see how any export was generated

**Stored fields:**
- `document_id` — which JSON was used
- `template_file` — which Excel template
- `sheet_name` — which sheet
- `mappings` — the array of {json_path → excel_column} pairs
- `created_at` — timestamp

---

### `generate_excel(self, data)` — The Main Function

This is the most important function in the entire codebase. It is a 10-step pipeline:

**Input:** `{"document_id": "...", "template_file": "Invoice.xlsx", "sheet_name": "Sheet1"}`

**Step 1 — Validate input**
```python
if not document_id or not template_file or not sheet_name:
    raise ValueError("...")
```
All three fields are required. Fails fast with a clear error.

**Step 2 — Fetch JSON from DB**
```python
document = self.db.json_collection.find_one({"_id": ObjectId(document_id)})
```
Gets the previously uploaded JSON document.

**Step 3 — Load Excel template**
```python
workbook = load_workbook(template_path)
```
Opens the `.xlsx` template file from the `templates/` folder using `openpyxl`.

**Step 4 — Resolve sheet name (case-insensitive)**
```python
matched_sheet = self._resolve_sheet_name(workbook, sheet_name)
```
"sheet1" will match "Sheet1", "SHEET1" etc.

**Steps 5–6 — Auto-detect orientation and extract headers**
```python
excel_headers = self.helper.get_excel_headers(sheet)
orientation = excel_headers[0].get("orientation", "horizontal")
```
The helper automatically figures out if the template is horizontal, vertical, or mixed.

**Step 7A (MIXED branch) — Fill cross-tab matrix**
```python
mappings = self.helper.fill_mixed_layout(sheet, mixed_info, document["json_data"])
```
For cross-tab layouts, one dedicated handler does everything.

**Step 7B (HORIZONTAL/VERTICAL branch) — Create fuzzy mappings**
```python
mappings = self.helper.create_mapping(document["json_data"], excel_headers)
```
Pairs each Excel header to the best-matching JSON field using fuzzy name matching.

**Step 8 — Save the mapping record to DB**
```python
mapping_res = self.save_mapping(document_id, template_file, matched_sheet, mappings)
```
Creates an audit record.

**Step 9 — Extract values + write to sheet**
```python
extracted_data = self.helper.extract_mapping_data(document["json_data"], mappings)
self.helper.write_excel_data(sheet, mappings, extracted_data)
```
Reads values from JSON by path, writes them into the correct Excel cells.

**Step 10 — Save the file**
```python
file_name = f"{uuid.uuid4().hex}.xlsx"
workbook.save(output_path)
result = self.db.export_collection.insert_one(export)
```
Generates a unique filename, saves to `generated_excels/` folder, records the export in MongoDB.

**Finally block:** `workbook.close()` always runs — even if an exception occurs — to free the file handle.

---

### `get_export_path(self, export_id)`

**What it does:**
1. Finds the export record in `export_collection` by ID
2. Gets the `file_path` string
3. **New check:** Verifies the file actually STILL EXISTS on disk
4. Returns the path (for `send_file()` in the route)

**Edge case handled:** If the file was deleted manually from disk after export → clear error instead of crash.

---

### `_resolve_sheet_name(workbook, sheet_name)` — Static Method

```python
target = sheet_name.strip().lower()
for name in workbook.sheetnames:
    if name.strip().lower() == target:
        return name
return None
```

**What it does:** Compares the requested sheet name against all actual sheet names in a case-insensitive way.

**Solves:**
- `"Sheet1"` finds `"sheet1"`, `"SHEET1"`, `"Sheet 1"` → Wait, `"Sheet 1"` ≠ `"Sheet1"` (space matters). But `"SHEET1"` = `"sheet1"` ✅

---
---

## FILE 3 — `excel_manage_helper.py`
### Role: Core Algorithms — Excel Reading, JSON Traversal, Mapping Logic

This is the most complex file. All actual data processing happens here.

---

## SECTION A — Utilities (3 private helper methods)

---

### `_normalize_name(self, name: str) → str`

```python
return re.sub(r"[\s_\-]+", "", str(name).lower().strip())
```

**What it does:** The fuzzy matching engine. Converts any name into a canonical lowercase token by removing all spaces, underscores, and hyphens.

**Examples:**
| Input | Output |
|---|---|
| `"Customer Name"` | `"customername"` |
| `"customer_name"` | `"customername"` |
| `"CustomerName"` | `"customername"` |
| `"CUSTOMER-NAME"` | `"customername"` |

**Why it matters:** Without this, `"Customer Name"` (Excel header) and `"customer_name"` (JSON field) would never match, even though they represent the same concept.

---

### `_get_merged_cell_value(self, sheet, cell)`

```python
for merged_range in sheet.merged_cells.ranges:
    if cell.coordinate in merged_range:
        master = sheet.cell(merged_range.min_row, merged_range.min_col)
        return master.value
return cell.value
```

**The problem:** When you merge cells A1:C1 in Excel, openpyxl stores the value in A1 only. Reading B1 or C1 returns `None`.

**What it does:** Detects if a cell is part of a merged region. If yes, it finds the top-left "master" cell of that merged region and returns ITS value.

**Example:**
```
Excel: [   Company Name    ] ← A1:C1 merged
       Reading B1 normally → None
       _get_merged_cell_value(sheet, B1) → "Company Name" ✅
```

---

### `_get_type(self, value) → str`

```python
type_map = {bool: "boolean", dict: "object", list: "array",
            str: "string", int: "number", float: "number"}
return type_map.get(type(value), "unknown")
```

**What it does:** Maps Python types to JSON schema type names.
- `None` → `"null"`
- `True/False` → `"boolean"`
- `42` or `3.14` → `"number"`
- `"hello"` → `"string"`
- `{...}` → `"object"`
- `[...]` → `"array"`

Used when building the mapping metadata record (for Swagger docs / DB storage).

---

## SECTION B — Orientation Detection (3 methods)

---

### `detect_header_orientation(self, sheet) → str`

**The most intelligent method in the codebase.** Reads the Excel sheet structure and decides how headers are laid out.

**Algorithm:**

```
horizontal_score = count of non-empty cells in Row 1
col_a_score      = count of non-empty cells in Col A, rows 2 to 30
a1_empty         = is cell A1 blank?

IF  horizontal_score ≥ 2  AND  col_a_score ≥ 2  AND  A1 is empty
    → "mixed"   (cross-tab matrix layout)

ELIF  col_a_score ≥ 3  AND  col_a_score > horizontal_score
    → "vertical"  (form/transposed layout)

ELSE
    → "horizontal"  (traditional table layout)
```

**Why check A1?** In a cross-tab (mixed) layout, A1 is always empty because it's the intersection of the row-axis and column-axis — it has no data. This is the strongest signal for mixed detection.

**Merged cells:** Both scores use `_get_merged_cell_value()` so merged header cells are counted correctly.

---

### `_find_header_row(self, sheet, max_scan=15) → int`

```python
for row_idx in range(1, max_scan + 1):
    non_empty = [cell for cell in sheet[row_idx]
                 if ... and str(...).strip()]
    if len(non_empty) >= 2:
        return row_idx
return 1
```

**Problem it solves:** Many Excel templates have company logos, report titles, or blank rows above the actual header row. If you always read Row 1, you'd get the wrong data.

**What it does:** Scans rows downward until it finds the first row that has 2 or more non-empty cells. That's the real header row.

**Example:**
```
Row 1: [Company Logo]          ← 1 cell, skip
Row 2: (blank)                 ← 0 cells, skip
Row 3: Name | Email | Amount   ← 3 cells → This is the header row!
```
Returns `3`.

---

## SECTION C — Header Extraction (4 methods)

---

### `get_excel_headers(self, sheet) → list`

**The public entry point for header extraction.** Calls `detect_header_orientation()` and dispatches to the right private method.

| Orientation | Calls | Returns |
|---|---|---|
| `"horizontal"` | `_get_horizontal_headers()` | Flat list of header dicts |
| `"vertical"` | `_get_vertical_headers()` | Flat list of header dicts |
| `"mixed"` | `_get_mixed_info()` | Single-item list with a structured dict |

Every returned dict always has: `excel_column`, `header_name`, `orientation`.

---

### `_get_horizontal_headers(self, sheet) → list`

```python
header_row = self._find_header_row(sheet)
for cell in sheet[header_row]:
    value = self._get_merged_cell_value(sheet, cell)
    if value is not None and str(value).strip():
        headers.append({
            "column": cell.column_letter,
            "excel_column": cell.column_letter,   # e.g. "A", "B", "C"
            "header_name": str(value).strip(),
            "orientation": "horizontal",
            "header_row": header_row,
            "data_start_row": header_row + 1,     # where data begins
        })
```

**What it does:** Finds the real header row (using `_find_header_row`), then reads every non-empty cell in that row.

**`data_start_row`** tells the writer: "skip this many rows and start writing data here."

---

### `_get_vertical_headers(self, sheet) → list`

```python
for row_idx in range(1, sheet.max_row + 1):
    cell = sheet.cell(row=row_idx, column=1)   # Col A only
    value = self._get_merged_cell_value(sheet, cell)
    if value is not None and str(value).strip():
        headers.append({
            "excel_column": f"A{row_idx}",     # e.g. "A1", "A2", "A3"
            "header_name": str(value).strip(),
            "orientation": "vertical",
            "row": row_idx,                    # which row to write data to
            "data_start_col": 2,              # Column B = index 2
        })
```

**What it does:** Reads all non-empty cells in Column A as headers.

**Key design:** `excel_column` is set to `"A{row_idx}"` (like `"A3"`) instead of just `"A"`, because multiple headers all live in Column A. We need a UNIQUE key per header for the mapping pipeline.

---

### `_get_mixed_info(self, sheet) → dict`

**What it does:** Extracts BOTH dimensions for a cross-tab layout.

```python
# Col headers: Row 1, skip A1
col_headers = [{"name": "Q1", "column": "B", "col_idx": 2},
               {"name": "Q2", "column": "C", "col_idx": 3}]

# Row headers: Col A, from row 2 onwards
row_headers = [{"name": "North", "row": 2},
               {"name": "South", "row": 3}]

return {
    "orientation": "mixed",
    "excel_column": "__mixed__",      # sentinel — tells the system "don't use flat flow"
    "row_headers": row_headers,
    "col_headers": col_headers,
}
```

Returns a single dict (not a list of headers) because mixed layout needs a completely different processing pipeline.

---

## SECTION D — Preview

### `get_excel_preview(self, workbook, max_rows=20) → list`

```python
for sheet in workbook.worksheets:
    rows = []
    for row in sheet.iter_rows(values_only=True):
        rows.append(list(row))
        if len(rows) >= max_rows:
            break
    result.append({"sheet_name": sheet.title, "rows": rows})
```

**What it does:** Reads up to 20 rows from every sheet in the workbook and returns them as raw arrays. Used by the `/preview` API endpoint to let the user inspect the template before generating.

---

## SECTION E — JSON Traversal (2 methods)

---

### `get_json_value(self, data, path: str) → list`

**The JSONPath engine.** Given a path string, walks through a JSON document and returns matching values as a flat list.

**Supported path syntax:**
| Path | Meaning |
|---|---|
| `$.name` | Top-level key "name" |
| `$.address.city` | Nested: go into "address", get "city" |
| `$.orders[*].amount` | Expand array "orders", get "amount" from each item |

**How it works — step by step:**

```
path = "$.orders[*].amount"
parts = ["orders[*]", "amount"]   # after removing "$."

Start: values = [full_json_data]

Part "orders[*]":
  → Has "[*]", so key = "orders"
  → For each dict in values, get the list at "orders"
  → Extend next_values with ALL items in that list
  values = [order1, order2, order3]   # expanded!

Part "amount":
  → No "[*]", so simple key lookup
  → For each dict in values, get "amount"
  values = [500, 1000, 750]

Return: [500, 1000, 750]
```

---

### `_find_json_fields(self, value, path, fields) → None`

**Recursive JSON schema discovery.** Walks the entire JSON tree and collects every "leaf" field (scalar value) with its path.

**Algorithm:**

```
If value is a DICT:
    For each key, item:
        current_path = parent_path + "." + key
        If item is dict or list → recurse
        Else (scalar) → add to fields list

If value is a LIST:
    Look at FIRST item only (structure inference)
    For each key in that first item:
        path = parent_path + "[*]." + key
        If child is dict or list → recurse
        Else → add to fields list
    BREAK after first item
```

**Why only the first list item?** This is schema inference. We assume all items in a list have the same structure. Only the PATH matters here (not the value) — at write time, `get_json_value` will fetch values from ALL items using `[*]`.

**Example:**
```json
{"user": {"name": "Rahul", "age": 30},
 "orders": [{"id": 1, "amount": 500}, {"id": 2, "amount": 800}]}
```
Fields discovered:
- `{name: "name", path: "$.user.name"}`
- `{name: "age",  path: "$.user.age"}`
- `{name: "id",   path: "$.orders[*].id"}`
- `{name: "amount", path: "$.orders[*].amount"}`

---

## SECTION F — Mapping Creation

### `create_mapping(self, json_data, excel_headers) → list`

**The matching engine for horizontal and vertical layouts.** Pairs every Excel header to the best JSON field.

**Step by step:**

1. **Collect all JSON fields** using `_find_json_fields()`
2. **For each Excel header:**
   - Normalize its name: `"Customer Name"` → `"customername"`
   - Find ALL JSON fields whose normalized name matches
   - If **multiple matches** (duplicate key problem): pick the one with the deepest path (most `.` separators)
   - If **no match**: add to `unmatched_headers` list (warning, not error)
3. Build a mapping dict for each matched pair
4. Log all unmatched headers as a warning

**Mapping dict structure:**
```python
{
    "header_name": "Customer Name",     # display label
    "header_type": "string",            # JSON type
    "excel_column": "B",               # which Excel column
    "json_path": "$.customer.name",    # how to fetch the value
    "orientation": "horizontal",       # layout type
    "data_start_row": 2,               # where to start writing
}
```

---

## SECTION G — Data Extraction

### `extract_mapping_data(self, json_data, mappings) → dict`

```python
result = {}
for mapping in mappings:
    values = self.get_json_value(json_data, mapping["json_path"])
    result[mapping["excel_column"]] = values  # always a list
return result
```

**What it does:** Uses the mappings list to extract actual values from the JSON document.

**Returns:**
```python
{
    "B": ["Rahul", "Priya", "Amit"],    # Column B has 3 values
    "C": ["r@e", "p@e", "a@e"],         # Column C has 3 values
    "D": [500, 800, 1200],              # Column D has 3 values
}
```

This is a flat dict of `{excel_column: [list_of_values]}`.

---

## SECTION H — Data Writing (3 methods)

---

### `write_excel_data(self, sheet, mappings, extracted_data) → None`

**The dispatcher.** Reads the `orientation` from the first mapping and routes to the correct writer.

```python
orientation = mappings[0].get("orientation", "horizontal")
if orientation == "vertical":
    self._write_vertical_data(...)
else:
    self._write_horizontal_data(...)
```

> **Note:** For mixed layout, this method is NOT called. `fill_mixed_layout()` handles writing directly.

---

### `_write_horizontal_data(self, sheet, mappings, extracted_data) → None`

**Traditional row-by-row writer.**

**Step 1:** Calculate max rows across all columns:
```python
lengths = {"B": 3, "C": 3, "D": 2}   # D has fewer values!
max_rows = 3
```

**Step 2:** Warn about unequal lengths:
```
WARNING: Column 'Amount' (D) has 2 values but max across columns is 3.
         Missing rows written as blank (None).
```

**Step 3:** Write cell by cell:
```python
for index in range(3):           # 0, 1, 2
    for each mapping:
        sheet["B2"] = "Rahul"    # index=0, data_start_row=2
        sheet["B3"] = "Priya"    # index=1
        sheet["B4"] = "Amit"     # index=2
        sheet["D4"] = None       # D had no 3rd value → blank
```

---

### `_write_vertical_data(self, sheet, mappings, extracted_data) → None`

**Column-by-column writer for vertical layouts.**

For each mapping:
- Knows which `row` this header is in (e.g., "Name" is in Row 1)
- Gets values for that field
- Writes them to B, C, D... on that same row

```python
# "Name" is in Row 1, data_start_col = 2 (Col B)
sheet["B1"] = "Rahul"    # col_offset=0 → get_column_letter(2+0) = "B"
sheet["C1"] = "Priya"    # col_offset=1 → get_column_letter(2+1) = "C"
```

---

## SECTION I — Mixed / Cross-Tab Layout (4 methods)

---

### `fill_mixed_layout(self, sheet, mixed_info, json_data) → list`

**All-in-one handler for matrix layouts.** Does header reading + JSON matching + cell writing in one pass.

**Step by step:**
1. Get `row_headers` and `col_headers` from `mixed_info`
2. Call `_build_cross_tab_lookup()` → builds a lookup dict from JSON
3. For every `(row_header, col_header)` pair:
   - Find the value: `lookup["north"]["q1"]` → `100`
   - Write to the intersection cell: `sheet["B2"] = 100`
   - Record the mapping for DB audit

**Visual:**
```
lookup = {"north": {"q1": 100, "q2": 200},
          "south": {"q1": 400, "q2": 500}}

For North × Q1: sheet["B2"] = lookup["north"]["q1"] = 100
For North × Q2: sheet["C2"] = lookup["north"]["q2"] = 200
For South × Q1: sheet["B3"] = lookup["south"]["q1"] = 400
For South × Q2: sheet["C3"] = lookup["south"]["q2"] = 500
```

---

### `_build_cross_tab_lookup(self, json_data, row_headers, col_headers) → dict`

**The JSON parser for mixed layouts.** Converts any of the 3 supported JSON formats into a single flat lookup dict.

**Format A (nested dict) — checked first:**
```python
{"North": {"Q1": 100, ...}, "South": {...}}
# dict values are all dicts → Format A
# Keys become row labels
```

**Format C (wrapped array) — checked second:**
```python
{"data": [{"region": "North", ...}, ...]}
# dict has a list value → extract that list → treat as array
```

**Format B (plain array) — fallback:**
```python
[{"region": "North", "Q1": 100}, {"region": "South", "Q1": 400}]
# plain list → pass to _lookup_from_records
```

**Returns:**
```python
{"north": {"q1": 100, "q2": 200},
 "south": {"q1": 400, "q2": 500}}
# All keys are normalized (lowercase, no spaces/underscores)
```

---

### `_lookup_from_records(self, records, row_norms, col_norms) → dict`

**Auto-detects the "row key" field from an array of records.**

**Problem:** In an array like `[{"region": "North", "Q1": 100}]`, we need to know that `"region"` is the row identifier, not a data column.

**How it detects:**
```python
# Look at the first record
for field_name, field_val in records[0].items():
    if isinstance(field_val, str):
        if _normalize_name(field_val) in row_norms:
            row_key_field = field_name  # Found it! "region" → "north" matches row headers
            break
```

Once the row key is found, build the lookup by:
1. Iterating all records
2. Getting the row label from `row_key_field`
3. Collecting all other fields that match column headers

---

### `_get_type` (already covered above)

---

## Complete Call Chain — Single Request

```
POST /excel/generate
    ↓
generate_excel_route()
    ↓ request.get_json()
generate_excel(data)
    ├── db.json_collection.find_one()           [DB: fetch JSON]
    ├── load_workbook(template_path)            [openpyxl: open Excel]
    ├── _resolve_sheet_name(workbook, name)     [case-insensitive match]
    ├── helper.get_excel_headers(sheet)
    │       ├── detect_header_orientation()
    │       │       ├── _get_merged_cell_value()  [for Row 1 + Col A]
    │       │       └── returns "horizontal" / "vertical" / "mixed"
    │       └── _get_horizontal_headers() / _get_vertical_headers() / _get_mixed_info()
    │               └── _get_merged_cell_value()  [for each header cell]
    │
    ├── [MIXED branch]
    │       └── helper.fill_mixed_layout(sheet, mixed_info, json_data)
    │               ├── _build_cross_tab_lookup()
    │               │       └── _lookup_from_records()
    │               └── sheet["B2"] = value  [direct cell writes]
    │
    ├── [HORIZONTAL / VERTICAL branch]
    │       ├── helper.create_mapping(json_data, excel_headers)
    │       │       ├── _find_json_fields()         [recursive JSON scan]
    │       │       └── _normalize_name()           [fuzzy match]
    │       ├── helper.extract_mapping_data()
    │       │       └── get_json_value()            [JSONPath traversal]
    │       └── helper.write_excel_data()
    │               ├── _write_horizontal_data()    [row-by-row]
    │               └── _write_vertical_data()      [col-by-col]
    │
    ├── save_mapping()                          [DB: audit trail]
    ├── workbook.save(output_path)              [openpyxl: write file]
    ├── db.export_collection.insert_one()       [DB: record export]
    └── return {"export_id": ..., "file_name": ...}
```

---

## MongoDB Collections Used

| Collection | Stores | Written By | Read By |
|---|---|---|---|
| `json_documents` | Uploaded JSON files | `upload_json` | `generate_excel`, `get_json` |
| `mapping_collection` | JSON→Excel mappings | `save_mapping` | (audit only) |
| `export_collection` | Generated file records | `generate_excel` | `get_export_path` |
