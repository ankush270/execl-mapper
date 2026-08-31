"""
excel_manage_helper.py
======================
Handles ALL Excel ↔ JSON mapping logic.

Edge Cases Solved
─────────────────
1. Fuzzy name matching   → spaces / underscores / camelCase all equated
2. Duplicate JSON keys   → deepest (most-specific) path wins; warned in logs
3. Unequal array lengths → shorter columns padded with None; warned in logs
4. Merged cells          → master-cell value used transparently
5. Blank leading rows    → auto-scanned to find real header row (horizontal)
6. Null / bool values    → written as-is (None → blank, bool → Excel TRUE/FALSE)
7. Unmatched headers     → logged as warnings; Excel column stays blank (no crash)
8. Sheet name            → validated before use (done in component)

Orientation Support
───────────────────
HORIZONTAL (traditional table)
  Row N  │ Name  │ Email   │ Amount   ← headers in a row
  Row N+1│ Rahul │ r@e.com │ 500      ← data expands downward

VERTICAL (transposed / form layout)
  Col A   │ Col B    │ Col C
  Name    │ Rahul    │ Priya           ← headers in col A, data goes right
  Email   │ r@e.com  │ p@e.com
  Amount  │ 500      │ 800

MIXED / CROSS-TAB (matrix layout)
         │ Q1    │ Q2    │ Q3         ← col headers in Row 1 (A1 usually empty)
  North  │ 100   │ 200   │ 150        ← row headers in Col A, data at intersection
  South  │ 400   │ 500   │ 350
  East   │ 200   │ 300   │ 250

JSON structures supported for MIXED:
  A) Nested dict  : {"North": {"Q1": 100, "Q2": 200}, "South": {...}}
  B) Array + key  : [{"region": "North", "Q1": 100}, {"region": "South", ...}]
  C) Wrapped array: {"data": [{"region": "North", ...}, ...]}

Auto-detected via detect_header_orientation().
"""

import re
from openpyxl.utils import get_column_letter
from core_utils.logging import logger


class ExcelManageHelper:

    # ══════════════════════════════════════════════════════
    #  UTILITIES
    # ══════════════════════════════════════════════════════

    def _normalize_name(self, name: str) -> str:
        """
        Fuzzy name normalizer.
        Strips ALL whitespace, underscores and hyphens, then lowercases.

        Examples (all become the same token):
            "Customer Name"  →  "customername"
            "customer_name"  →  "customername"
            "CustomerName"   →  "customername"
            "CUSTOMER-NAME"  →  "customername"
        """
        return re.sub(r"[\s_\-]+", "", str(name).lower().strip())

    def _get_merged_cell_value(self, sheet, cell):
        """
        Return the correct value for merged-cell regions.
        openpyxl sets every non-master merged cell's value to None.
        This method finds the top-left master cell and returns its value.
        """
        for merged_range in sheet.merged_cells.ranges:
            if cell.coordinate in merged_range:
                master = sheet.cell(merged_range.min_row, merged_range.min_col)
                return master.value
        return cell.value

    def _get_type(self, value) -> str:
        """Map a Python value to a JSON-schema type string."""
        if value is None:
            return "null"
        type_map = {
            bool:  "boolean",
            dict:  "object",
            list:  "array",
            str:   "string",
            int:   "number",
            float: "number",
        }
        return type_map.get(type(value), "unknown")

    # ══════════════════════════════════════════════════════
    #  ORIENTATION DETECTION  (horizontal / vertical / mixed)
    # ══════════════════════════════════════════════════════

    def detect_header_orientation(self, sheet) -> str:
        """
        Decide whether headers are in a row (horizontal), a column (vertical),
        or BOTH (mixed / cross-tab matrix).

        Algorithm
        ─────────
        horizontal_score = non-empty cells in Row 1
        col_a_score      = non-empty cells in Col A from row 2 onwards
        a1_empty         = whether cell A1 is blank

        Decision tree:
          IF Row 1 has ≥2 headers  AND  Col A (row 2+) has ≥2 row labels
               AND  A1 is empty   → 'mixed'   (cross-tab matrix)
          ELIF Col A (row 2+) >= 3 AND col_a_score > horizontal_score
                                  → 'vertical'
          ELSE                    → 'horizontal'

        Merged-cell values are resolved before scoring.
        """
        # ── Row 1 score ──────────────────────────────────
        horizontal_score = 0
        for cell in sheet[1]:
            val = self._get_merged_cell_value(sheet, cell)
            if val is not None and str(val).strip():
                horizontal_score += 1

        # ── Col A score (row 2 onwards, up to 30 rows) ───
        max_scan = min(30, sheet.max_row)
        col_a_score = 0
        for row_idx in range(2, max_scan + 1):
            cell = sheet.cell(row=row_idx, column=1)
            val = self._get_merged_cell_value(sheet, cell)
            if val is not None and str(val).strip():
                col_a_score += 1

        # ── A1 empty? ────────────────────────────────────
        a1_val = self._get_merged_cell_value(sheet, sheet.cell(1, 1))
        a1_empty = (a1_val is None or not str(a1_val).strip())

        # ── Decision ─────────────────────────────────────
        if horizontal_score >= 2 and col_a_score >= 2 and a1_empty:
            orientation = "mixed"
        elif col_a_score >= 3 and col_a_score > horizontal_score:
            orientation = "vertical"
        else:
            orientation = "horizontal"

        logger.info(
            f"Orientation detection → {orientation} "
            f"(row1={horizontal_score}, colA_row2+={col_a_score}, A1_empty={a1_empty})"
        )
        return orientation

    def _find_header_row(self, sheet, max_scan: int = 15) -> int:
        """
        Scan downward (up to max_scan rows) to find the first row
        that contains ≥ 2 non-empty cells.  Skips blank title/logo rows.
        Returns 1-based row index (defaults to 1).
        """
        for row_idx in range(1, max_scan + 1):
            non_empty = [
                cell for cell in sheet[row_idx]
                if self._get_merged_cell_value(sheet, cell) is not None
                and str(self._get_merged_cell_value(sheet, cell)).strip()
            ]
            if len(non_empty) >= 2:
                return row_idx
        logger.warning(
            f"No header row with ≥2 non-empty cells found in first {max_scan} rows. "
            "Defaulting to row 1."
        )
        return 1

    # ══════════════════════════════════════════════════════
    #  HEADER EXTRACTION
    # ══════════════════════════════════════════════════════

    def get_excel_headers(self, sheet) -> list:
        """
        Auto-detect orientation and return a unified list of header dicts.

        For horizontal and vertical layouts, returns a flat list.
        For mixed layouts, returns a single-item list containing a
        structured dict (so callers can check ['orientation'] == 'mixed').

        Every dict always contains:
            orientation   – 'horizontal' | 'vertical' | 'mixed'
            header_name   – display name (raw, un-normalised)
            excel_column  – unique key used throughout the mapping pipeline

        Horizontal extras: header_row, data_start_row
        Vertical extras  : row, data_start_col
        Mixed extras     : row_headers list, col_headers list (see _get_mixed_info)
        """
        try:
            orientation = self.detect_header_orientation(sheet)
            if orientation == "horizontal":
                return self._get_horizontal_headers(sheet)
            if orientation == "vertical":
                return self._get_vertical_headers(sheet)
            # mixed
            return [self._get_mixed_info(sheet)]
        except Exception as exc:
            raise ValueError(f"Unable to read Excel headers: {exc}")

    def _get_horizontal_headers(self, sheet) -> list:
        """Headers across a row → data expands downward."""
        header_row = self._find_header_row(sheet)
        headers = []
        for cell in sheet[header_row]:
            value = self._get_merged_cell_value(sheet, cell)
            if value is not None and str(value).strip():
                headers.append({
                    "column":         cell.column_letter,
                    "excel_column":   cell.column_letter,
                    "header_name":    str(value).strip(),
                    "orientation":    "horizontal",
                    "header_row":     header_row,
                    "data_start_row": header_row + 1,
                })
        return headers

    def _get_vertical_headers(self, sheet) -> list:
        """
        Headers in Column A → data expands rightward (Col B, C, …).
        excel_column is set to 'A<row>' (e.g. 'A3') to be a unique key.
        """
        headers = []
        for row_idx in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row_idx, column=1)
            value = self._get_merged_cell_value(sheet, cell)
            if value is not None and str(value).strip():
                headers.append({
                    "column":        "A",
                    "excel_column":  f"A{row_idx}",  # unique per header
                    "header_name":   str(value).strip(),
                    "orientation":   "vertical",
                    "row":           row_idx,
                    "data_start_col": 2,              # B = 2
                })
        return headers

    def _get_mixed_info(self, sheet) -> dict:
        """
        Extract the full cross-tab structure:
            row_headers – Col A values from row 2 onwards
            col_headers – Row 1 values from col B onwards

        Returns a single structured dict (not a flat list) so that
        generate_excel can branch on orientation == 'mixed'.
        """
        # Column headers (Row 1, skip A1)
        col_headers = []
        for cell in sheet[1]:
            if cell.column == 1:
                continue  # skip A1 (usually empty label)
            val = self._get_merged_cell_value(sheet, cell)
            if val is not None and str(val).strip():
                col_headers.append({
                    "name":    str(val).strip(),
                    "column":  cell.column_letter,
                    "col_idx": cell.column,
                })

        # Row headers (Col A, from row 2 onwards)
        row_headers = []
        for row_idx in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row_idx, column=1)
            val = self._get_merged_cell_value(sheet, cell)
            if val is not None and str(val).strip():
                row_headers.append({
                    "name": str(val).strip(),
                    "row":  row_idx,
                })

        return {
            "orientation":  "mixed",
            "excel_column": "__mixed__",          # sentinel key
            "header_name":  "__cross_tab__",
            "row_headers":  row_headers,
            "col_headers":  col_headers,
        }

    # ══════════════════════════════════════════════════════
    #  PREVIEW
    # ══════════════════════════════════════════════════════

    def get_excel_preview(self, workbook, max_rows: int = 20) -> list:
        """Return up to max_rows rows from every sheet for preview."""
        try:
            result = []
            for sheet in workbook.worksheets:
                rows = []
                for row in sheet.iter_rows(values_only=True):
                    rows.append(list(row))
                    if len(rows) >= max_rows:
                        break
                result.append({"sheet_name": sheet.title, "rows": rows})
            return result
        except Exception as exc:
            raise ValueError(f"Unable to create Excel preview: {exc}")

    # ══════════════════════════════════════════════════════
    #  JSON PATH TRAVERSAL
    # ══════════════════════════════════════════════════════

    def get_json_value(self, data, path: str) -> list:
        """
        Walk a JSONPath-style string (e.g. '$.orders[*].amount') and
        return a flat list of matching leaf values.

        Supports:
            $.field            – simple key lookup
            $.parent.child     – nested key lookup
            $.list[*].field    – expand all items in an array
        """
        try:
            parts = path.replace("$.", "").split(".")
            values = [data]
            for part in parts:
                next_values = []
                if "[*]" in part:
                    key = part.replace("[*]", "")
                    for value in values:
                        if isinstance(value, dict):
                            items = value.get(key, [])
                            if isinstance(items, list):
                                next_values.extend(items)
                else:
                    for value in values:
                        if isinstance(value, dict) and part in value:
                            next_values.append(value[part])
                values = next_values
            return values
        except Exception as exc:
            raise ValueError(f"Unable to read JSON path '{path}': {exc}")

    def _find_json_fields(self, value, path: str, fields: list) -> None:
        """
        Recursively collect every scalar leaf field from the JSON tree.
        For arrays, structure is inferred from the FIRST element only
        (schema inference), but the path uses [*] so get_json_value will
        fetch values from ALL elements at write time.
        """
        try:
            if isinstance(value, dict):
                for key, item in value.items():
                    current_path = f"{path}.{key}"
                    if isinstance(item, (dict, list)):
                        self._find_json_fields(item, current_path, fields)
                    else:
                        fields.append({
                            "name":  key,
                            "type":  self._get_type(item),
                            "value": item,
                            "path":  current_path,
                        })
            elif isinstance(value, list):
                for item in value:
                    if isinstance(item, dict):
                        for key, child in item.items():
                            current_path = f"{path}[*].{key}"
                            if isinstance(child, (dict, list)):
                                self._find_json_fields(child, current_path, fields)
                            else:
                                fields.append({
                                    "name":  key,
                                    "type":  self._get_type(child),
                                    "value": child,
                                    "path":  current_path,
                                })
                        break  # structure inferred from first item only
        except Exception as exc:
            logger.error(f"Unable to analyze JSON structure: {exc}")
            raise ValueError(f"Unable to analyze JSON structure: {exc}")

    # ══════════════════════════════════════════════════════
    #  MAPPING CREATION  (horizontal / vertical — fuzzy + duplicate-safe)
    # ══════════════════════════════════════════════════════

    def create_mapping(self, json_data, excel_headers) -> list:
        """
        Build a mapping list pairing each Excel header to a JSON field.
        Used for horizontal and vertical layouts only.
        For mixed/cross-tab, use fill_mixed_layout() directly.

        Fuzzy matching
        ──────────────
        Names are normalised before comparison:
            Excel "Customer Name"  ↔  JSON "customer_name" / "CustomerName"

        Duplicate JSON key handling
        ───────────────────────────
        If multiple JSON fields share a normalised name, the deepest
        (most-nested) path wins and a warning is logged.

        Unmatched headers
        ─────────────────
        Logged as a warning; the Excel column stays blank — NOT a hard error.
        """
        try:
            json_fields: list = []
            self._find_json_fields(json_data, "$", json_fields)

            mappings = []
            unmatched_headers = []

            for header in excel_headers:
                header_norm = self._normalize_name(header["header_name"])

                matches = [
                    field for field in json_fields
                    if self._normalize_name(field["name"]) == header_norm
                ]

                if matches:
                    best = max(matches, key=lambda f: f["path"].count("."))
                    if len(matches) > 1:
                        logger.warning(
                            f"Multiple JSON fields match Excel header "
                            f"'{header['header_name']}'. "
                            f"Using deepest path: '{best['path']}'. "
                            f"All candidates: {[m['path'] for m in matches]}"
                        )
                    mappings.append({
                        "header_type":    best["type"],
                        "header_name":    header["header_name"],
                        "header_value":   best["value"],
                        "excel_column":   header["excel_column"],
                        "json_path":      best["path"],
                        "orientation":    header.get("orientation", "horizontal"),
                        "row":            header.get("row"),
                        "data_start_row": header.get("data_start_row", 2),
                        "data_start_col": header.get("data_start_col", 2),
                    })
                else:
                    unmatched_headers.append(header["header_name"])

            if unmatched_headers:
                logger.warning(
                    f"No JSON field matched for Excel header(s): {unmatched_headers}. "
                    "Those columns will remain blank in the output."
                )

            return mappings

        except Exception as exc:
            logger.error(f"Unable to create mapping: {exc}")
            raise ValueError(f"Unable to create mapping: {exc}")

    # ══════════════════════════════════════════════════════
    #  DATA EXTRACTION
    # ══════════════════════════════════════════════════════

    def extract_mapping_data(self, json_data, mappings) -> dict:
        """
        Return {excel_column: [values]} for every mapping.
        Scalars are wrapped in a list so downstream code is uniform.
        """
        try:
            result = {}
            for mapping in mappings:
                values = self.get_json_value(json_data, mapping["json_path"])
                if not isinstance(values, list):
                    values = [values]
                result[mapping["excel_column"]] = values
            return result
        except Exception as exc:
            logger.error(f"Unable to extract mapping data: {exc}")
            raise ValueError(f"Unable to extract mapping data: {exc}")

    # ══════════════════════════════════════════════════════
    #  DATA WRITING  (horizontal / vertical)
    # ══════════════════════════════════════════════════════

    def write_excel_data(self, sheet, mappings, extracted_data) -> None:
        """
        Dispatch to the correct writer based on orientation stored in mappings.
        NOTE: For mixed/cross-tab layouts use fill_mixed_layout() instead.
        """
        try:
            if not mappings:
                logger.warning("write_excel_data called with empty mappings. Nothing written.")
                return
            orientation = mappings[0].get("orientation", "horizontal")
            if orientation == "vertical":
                self._write_vertical_data(sheet, mappings, extracted_data)
            else:
                self._write_horizontal_data(sheet, mappings, extracted_data)
        except Exception as exc:
            logger.error(f"Unable to write Excel data: {exc}")
            raise ValueError(f"Unable to write Excel data: {exc}")

    def _write_horizontal_data(self, sheet, mappings, extracted_data) -> None:
        """
        Fill cells row-by-row (traditional table layout).
        Unequal array lengths → shorter columns padded with None + warning.
        """
        lengths = {
            m["excel_column"]: len(extracted_data.get(m["excel_column"], []))
            for m in mappings
        }
        max_rows = max(lengths.values(), default=0)

        if max_rows == 0:
            logger.warning("All mapped columns have 0 data values. Nothing written.")
            return

        for mapping in mappings:
            col = mapping["excel_column"]
            col_len = lengths.get(col, 0)
            if col_len != max_rows:
                logger.warning(
                    f"Column '{mapping['header_name']}' ({col}) has {col_len} values "
                    f"but max across columns is {max_rows}. "
                    "Missing rows written as blank (None)."
                )

        data_start_row = mappings[0].get("data_start_row", 2)

        for index in range(max_rows):
            for mapping in mappings:
                col    = mapping["excel_column"]
                values = extracted_data.get(col, [])
                value  = values[index] if index < len(values) else None
                sheet[f"{col}{data_start_row + index}"] = value

    def _write_vertical_data(self, sheet, mappings, extracted_data) -> None:
        """
        Fill cells column-by-column (transposed / form layout).
        Each label in Col A gets its value(s) written to B, C, D, … on that row.
        """
        for mapping in mappings:
            row = mapping.get("row")
            if row is None:
                logger.warning(
                    f"Vertical mapping for '{mapping['header_name']}' has no 'row'. Skipping."
                )
                continue

            col_key = mapping["excel_column"]
            values  = extracted_data.get(col_key, [])
            if not isinstance(values, list):
                values = [values]

            data_start_col = mapping.get("data_start_col", 2)
            for col_offset, value in enumerate(values):
                col_letter = get_column_letter(data_start_col + col_offset)
                sheet[f"{col_letter}{row}"] = value

    # ══════════════════════════════════════════════════════
    #  MIXED / CROSS-TAB LAYOUT
    # ══════════════════════════════════════════════════════

    def fill_mixed_layout(self, sheet, mixed_info: dict, json_data) -> list:
        """
        All-in-one handler for cross-tab (matrix) Excel layouts.

        Layout
        ──────
                 │ Q1  │ Q2  │ Q3       ← col headers (Row 1, col B onwards)
        North    │ 100 │ 200 │ 150      ← row headers (Col A, row 2 onwards)
        South    │ 400 │ 500 │ 350       + data at each intersection cell

        Supported JSON formats
        ──────────────────────
        A) Nested dict (keys = row labels):
               {"North": {"Q1": 100, "Q2": 200},
                "South": {"Q1": 400, "Q2": 500}}

        B) Array with one row-key field + column fields:
               [{"region": "North", "Q1": 100, "Q2": 200},
                {"region": "South", "Q1": 400, "Q2": 500}]
           (row-key field is auto-detected by matching its values to row headers)

        C) Wrapped array (one level of nesting):
               {"data": [{"region": "North", ...}, ...]}

        Returns
        ───────
        List of mapping dicts (one per cell written) for DB storage.
        """
        row_headers = mixed_info.get("row_headers", [])
        col_headers = mixed_info.get("col_headers", [])

        if not row_headers:
            raise ValueError(
                "Mixed layout: no row headers found in Column A. "
                "Expected labels starting from row 2."
            )
        if not col_headers:
            raise ValueError(
                "Mixed layout: no column headers found in Row 1. "
                "Expected labels starting from column B."
            )

        logger.info(
            f"Mixed layout: {len(row_headers)} row headers × "
            f"{len(col_headers)} col headers"
        )

        # Build lookup: {normalized_row_label: {normalized_col_label: value}}
        lookup = self._build_cross_tab_lookup(json_data, row_headers, col_headers)

        mappings  = []
        written   = 0
        skipped   = []

        for rh in row_headers:
            rh_norm  = self._normalize_name(rh["name"])
            row_data = lookup.get(rh_norm)

            if row_data is None:
                skipped.append(rh["name"])
                logger.warning(
                    f"Mixed layout: no JSON data matched row header '{rh['name']}'"
                )

            for ch in col_headers:
                ch_norm = self._normalize_name(ch["name"])
                value   = row_data.get(ch_norm) if row_data else None

                cell_addr = f"{ch['column']}{rh['row']}"
                sheet[cell_addr] = value

                if value is not None:
                    written += 1

                mappings.append({
                    "orientation":  "mixed",
                    "header_name":  f"{rh['name']} × {ch['name']}",
                    "excel_column": cell_addr,
                    "json_path":    f"$.{rh['name']}.{ch['name']}",
                    "row_header":   rh["name"],
                    "col_header":   ch["name"],
                    "value":        value,
                })

        if skipped:
            logger.warning(
                f"Mixed layout: {len(skipped)} row header(s) had no JSON match: {skipped}"
            )
        logger.info(f"Mixed layout: wrote {written} / {len(mappings)} cells")
        return mappings

    def _build_cross_tab_lookup(
        self, json_data, row_headers: list, col_headers: list
    ) -> dict:
        """
        Convert JSON into a flat lookup dict:
            {normalized_row_label: {normalized_col_label: value}}

        Handles formats A (nested dict), B (array+row-key), C (wrapped array).
        """
        row_norms = {self._normalize_name(rh["name"]) for rh in row_headers}
        col_norms = {self._normalize_name(ch["name"]) for ch in col_headers}
        lookup: dict = {}

        # ── Format A: {"North": {"Q1": 100}, "South": {...}} ─────────
        if isinstance(json_data, dict) and all(
            isinstance(v, dict) for v in json_data.values()
        ):
            for key, val in json_data.items():
                k_norm = self._normalize_name(str(key))
                if k_norm in row_norms:
                    lookup[k_norm] = {
                        self._normalize_name(fk): fv
                        for fk, fv in val.items()
                        if self._normalize_name(fk) in col_norms
                    }
            if lookup:
                logger.info("Mixed layout: JSON matched as nested-dict (Format A)")
                return lookup

        # ── Format C: {"data": [...]} or {"sales": [...]} etc. ────────
        records = None
        if isinstance(json_data, dict):
            for v in json_data.values():
                if isinstance(v, list) and v and isinstance(v[0], dict):
                    records = v
                    logger.info("Mixed layout: JSON matched as wrapped-array (Format C)")
                    break
        elif isinstance(json_data, list):
            records = json_data
            logger.info("Mixed layout: JSON matched as array (Format B/C)")

        # ── Format B: [{"region": "North", "Q1": 100, ...}, ...] ──────
        if records:
            lookup = self._lookup_from_records(records, row_norms, col_norms)

        if not lookup:
            logger.warning(
                "Mixed layout: Could not build cross-tab lookup from JSON. "
                f"Row headers expected: {[rh['name'] for rh in row_headers]}, "
                f"Col headers expected: {[ch['name'] for ch in col_headers]}"
            )
        return lookup

    def _lookup_from_records(
        self, records: list, row_norms: set, col_norms: set
    ) -> dict:
        """
        Build the cross-tab lookup from a list of dicts.
        Auto-detects which field is the 'row key' by finding a field
        whose values (normalised) intersect with row_norms.
        """
        if not records:
            return {}

        # Auto-detect the row-key field from the first record
        row_key_field = None
        for field_name, field_val in records[0].items():
            if isinstance(field_val, (str, int, float)):
                if self._normalize_name(str(field_val)) in row_norms:
                    row_key_field = field_name
                    break

        if row_key_field is None:
            logger.warning(
                "Mixed layout: Could not auto-detect row-key field in JSON records. "
                f"First record keys: {list(records[0].keys())}. "
                f"Expected a field whose value is one of: {row_norms}"
            )
            return {}

        logger.info(f"Mixed layout: row-key field auto-detected as '{row_key_field}'")

        lookup = {}
        for record in records:
            row_label = record.get(row_key_field)
            if row_label is None:
                continue
            rn = self._normalize_name(str(row_label))
            if rn not in row_norms:
                continue
            lookup[rn] = {
                self._normalize_name(fk): fv
                for fk, fv in record.items()
                if fk != row_key_field
                and self._normalize_name(fk) in col_norms
            }

        return lookup