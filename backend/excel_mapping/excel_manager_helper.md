# Detailed Function-by-Function Documentation: `excel_manage_helper.py`

Yeh document [`excel_manage_helper.py`](file:///e:/Projects/excelmapping/backend/excel_mapping/excel_manage_helper.py) file me maujood **sabhi Classes, Methods, aur Utility Functions** ki exhaustive, detailed guide hai.
Har function ke baare me 3 mukhya baatein samjhayi gayi hain:
1. **Kya karta hai?** (Functionality & Input/Output)
2. **Kyun karta hai?** (Internal Logic & Algorithm)
3. **Need kyun hai?** (Business Utility & Importance in System)

---

## 📌 Table of Contents
- [Class: ExcelManageHelper](#1-class-excelmanagehelper)
  - [1.1 to_object_id](#11-to_object_id)
  - [1.2 normalize_name](#12-normalize_name)
  - [1.3 name_tokens](#13-name_tokens)
  - [1.4 clean_value](#14-clean_value)
  - [1.5 join_json_path](#15-join_json_path)
  - [1.6 safe_template_path](#16-safe_template_path)
  - [1.7 analyze_template_workbook](#17-analyze_template_workbook)
  - [1.8 analyze_json](#18-analyze_json)
  - [1.9 _create_field_entry](#19-_create_field_entry)
  - [1.10 _walk_json](#110-_walk_json)
  - [1.11 _get_type](#111-_get_type)
  - [1.12 build_mappings](#112-build_mappings)
  - [1.13 _format_excel_field](#113-_format_excel_field)
  - [1.14 _flatten_excel_fields](#114-_flatten_excel_fields)
  - [1.15 _mapping_score](#115-_mapping_score)
  - [1.16 _types_compatible](#116-_types_compatible)
  - [1.17 get_json_value](#117-get_json_value)
  - [1.18 _parse_json_path](#118-_parse_json_path)
  - [1.19 populate_template](#119-populate_template)
- [Class: _WorkbookAnalyzerAdapter](#2-class-_workbookanalyzeradapter)
  - [2.1 analyze](#21-analyze)
  - [2.2 _analyze_rows](#22-_analyze_rows)
  - [2.3 _build_row_analysis](#23-_build_row_analysis)
  - [2.4 _find_label_cell](#24-_find_label_cell)
  - [2.5 _find_value_cell](#25-_find_value_cell)
  - [2.6 _discover_font_levels](#26-_discover_font_levels)
  - [2.7 _classify_rows](#27-_classify_rows)
  - [2.8 _detect_value_column](#28-_detect_value_column)
  - [2.9 _build_hierarchy](#29-_build_hierarchy)
  - [2.10 _close_section_ranges](#210-_close_section_ranges)
  - [2.11 Utility Methods (`_is_empty`, `_get_merge_info`, `_is_formula`, `_get_fill_color`, `_clean_label`)](#211-utility-methods)

---

## 1. Class: `ExcelManageHelper`

Is class me JSON processing, JSONPath evaluation, fuzzy mapping heuristics, security validation, aur Excel template filling ke core utility functions hain.

---

### 1.1 `to_object_id`
```python
def to_object_id(self, value: str | ObjectId) -> ObjectId:
```
- **Kya karta hai?** Given string ya `ObjectId` value ko MongoDB compatible `ObjectId` class instance me convert karta hai.
- **Kyun karta hai?** Agar input pehle se `ObjectId` hai to direct return karta hai. Agar valid string ID (`"651a2b3c..."`) hai to `ObjectId(value)` initialize karta hai, warna `ValueError("Invalid ObjectId.")` throw karta hai.
- **Need kyun hai?** MongoDB Queries me `_id` field raw string format support nahi karti. Frontend se string ID aati hai, use database format me parse karne ke liye iski zaroorat padti hai.

---

### 1.2 `normalize_name`
```python
@staticmethod
def normalize_name(value: Any) -> str:
```
- **Kya karta hai?** Kisi bhi Label/Key string me se special characters (`@`, `#`, `$`), hyphens (`-`), underscores (`_`), aur capitalization ko normalize karke clean lowercase string return karta hai.
- **Kyun karta hai?** Regex `re.sub` use karke extra whitespaces clean karta hai. E.g., `"First Name:"` ➔ `"first name"`, `"user_first_name"` ➔ `"user first name"`.
- **Need kyun hai?** Excel labels aur JSON keys alag-alag case aur format me hote hain. Exact comparison impossible hota hai, isliye fuzzy score calculation se pehle standard form chahiye.

---

### 1.3 `name_tokens`
```python
@staticmethod
def name_tokens(value: Any) -> set[str]:
```
- **Kya karta hai?** Normalized string ko individual words (tokens) ke Python `set` me split karta hai. E.g., `"policy effective date"` ➔ `{"policy", "effective", "date"}`.
- **Kyun karta hai?** Jaccard Similarity index (token intersection & union) compute karne ke liye input word list deta hai.
- **Need kyun hai?** Word order change hone par bhi matching perform karne ke liye (E.g., `"Date of Birth"` vs `"Birth Date"`).

---

### 1.4 `clean_value`
```python
def clean_value(self, value: Any) -> Any:
```
- **Kya karta hai?** Agar input Python `datetime` ya `date` object hai to use ISO format string (`"2026-09-04T00:00:00"`) me convert karta hai, baki types ko as-is return karta hai.
- **Kyun karta hai?** OpenPyXL cell writing aur JSON serialization me native date objects serialization errors throw kar sakte hain.
- **Need kyun hai?** Date-time values write karte waqt application crash na ho.

---

### 1.5 `join_json_path`
```python
def join_json_path(self, path: str, key: str) -> str:
```
- **Kya karta hai?** Parent JSONPath me naye child key ko safely append karta hai (E.g., `"$['user']"` + `"address"` ➔ `"$['user']['address']"`).
- **Kyun karta hai?** Key me quotes (`'`) aur backslashes (`\`) ko escape karta hai taaki JSONPath syntax corrupt na ho.
- **Need kyun hai?** Har JSON field ka global unique identifier path build karne ke liye.

---

### 1.6 `safe_template_path`
```python
def safe_template_path(self, file_name: str) -> Path:
```
- **Kya karta hai?** Client dwara bheje gaye template filename ko security perspective se check karta hai aur absolute disk path return karta hai.
- **Kyun karta hai?** Check karta hai ki file extension `.xlsx` ya `.xlsm` hai, filename me Path Traversal (`../`) attacks to nahi hain, aur file `Config.TEMPLATE_FOLDER` me exist karti hai ya nahi.
- **Need kyun hai?** Arbitrary file read & security path traversal vulnerabilities rokne ke liye.

---

### 1.7 `analyze_template_workbook`
```python
def analyze_template_workbook(self, workbook, *, file_name: str, only_sheet: str | None = None) -> dict[str, Any]:
```
- **Kya karta hai?** `_WorkbookAnalyzerAdapter` class instaniate karke workbook ke sabhi non-hidden sheets ya specific sheet ka structural schema extract karta hai.
- **Kyun karta hai?** Agar `only_sheet` param diya gaya ho to baki sheets filter out karke specific sheet schema return karta hai.
- **Need kyun hai?** API layer aur UI ko Excel Template ki complete hierarchy and fields represent karne ke liye.

---

### 1.8 `analyze_json`
```python
def analyze_json(self, data: Any) -> dict[str, Any]:
```
- **Kya karta hai?** Full JSON document ko traverse karke raw fields metadata list, array paths list, aur root type (`"object"` / `"array"`) ka schema dict banata hai.
- **Kyun karta hai?** Internal `_walk_json` recursive traversal process initiate karta hai.
- **Need kyun hai?** Mapping engine ko JSON side ka structural schema dene ke liye.

---

### 1.9 `_create_field_entry`
```python
def _create_field_entry(self, key: Any, child: Any, child_path: str, parent_path: str, parent_name: str | None) -> dict[str, Any]:
```
- **Kya karta hai?** JSON node element ke liye standardized dictionary item construct karta hai containing `name`, `normalized_name`, `type`, `path`, `parent_path`, `parent_name`, `value`, `array_depth`, aur `tokens`.
- **Kyun karta hai?** Code repetition ko avoid karta hai (DRY principle).
- **Need kyun hai?** Dictionary keys aur Array elements dono me identical structure metadata record construct karne ke liye.

---

### 1.10 `_walk_json`
```python
def _walk_json(self, value, path, parent_path, parent_name, fields, arrays, _depth=0):
```
- **Kya karta hai?** Recursively JSON tree ko Depth-First Traversal se scan karta hai.
- **Kyun karta hai?** Dictionary objects, nested lists, aur representative array items (`value[:20]`) ko iterate karke fields list me append karta hai. Recursion limit `_MAX_JSON_DEPTH = 50` set hai.
- **Need kyun hai?** Unstructured nested JSON ko flat list of field paths me unravel karne ke liye.

---

### 1.11 `_get_type`
```python
def _get_type(self, value: Any) -> str:
```
- **Kya karta hai?** Python object data type ko JSON Schema standard string (`"boolean"`, `"integer"`, `"number"`, `"string"`, `"array"`, `"object"`, `"null"`, `"unknown"`) me map karta hai.
- **Kyun karta hai?** Type map dict (`{type(None): "null", bool: "boolean", ...}`) lookup perform karta hai.
- **Need kyun hai?** Excel expected field type aur JSON actual data type ki compatibility cross-check karne ke liye.

---

### 1.12 `build_mappings`
```python
def build_mappings(self, *, json_schema: dict[str, Any], excel_sheet_schema: dict[str, Any]) -> list[dict[str, Any]]:
```
- **Kya karta hai?** Automated mapping algorithm running scoring engine between Excel fields and JSON fields. Returns sorted list of high-confidence mappings.
- **Kyun karta hai?** Har Excel field ke liye saare unused JSON fields ka `_mapping_score` calculate karta hai. Score $\ge 0.50$ (50% confidence) hone par best matching JSON path claim karta hai.
- **Need kyun hai?** Automated Zero-Touch JSON-to-Excel field mapping engine.

---

### 1.13 `_format_excel_field`
```python
def _format_excel_field(self, field: dict[str, Any], section_path: list[str]) -> dict[str, Any]:
```
- **Kya karta hai?** Excel cell metadata dictionary formatting karta hai (including column letter extraction e.g. `"B5"` ➔ `"B"`).
- **Kyun karta hai?** Section fields aur Orphan fields dono ko same dict schema me transform karta hai.
- **Need kyun hai?** `_flatten_excel_fields` method me duplicate dict creation code hatane ke liye.

---

### 1.14 `_flatten_excel_fields`
```python
def _flatten_excel_fields(self, sheet_schema: dict[str, Any]) -> list[dict[str, Any]]:
```
- **Kya karta hai?** Excel Tree Hierarchy (`sections` ➔ `children` ➔ `fields`) ko flat list me convert karta hai.
- **Kyun karta hai?** Inner recursive `visit_section` traversal se path tracking pass hoti hai.
- **Need kyun hai?** Mapping comparison loop me easy iteration ke liye flat Excel field list tayar karna.

---

### 1.15 `_mapping_score`
```python
def _mapping_score(self, excel_field: dict[str, Any], json_field: dict[str, Any]) -> float:
```
- **Kya karta hai?** Heuristic Scoring Engine compute karke 0.00 se 1.00 tak Match Confidence Score nikalta hai.
- **Kyun karta hai?** Weighted Scoring Formula:
  - Exact Name Match: **+0.70**
  - Jaccard Word Token Similarity: **+0.45 * similarity**
  - Section Context Path Overlap: **+0.20 max**
  - Data Type Compatibility: **+0.10 bonus / -0.05 penalty**
  - Array Repeating Field Bonus: **+0.02**
- **Need kyun hai?** Human-like intelligent auto-mapping suggestions deliver karne ke liye.

---

### 1.16 `_types_compatible`
```python
def _types_compatible(self, expected: str, actual: str) -> bool:
```
- **Kya karta hai?** Check karta hai ki Excel field expected type aur JSON actual type compatible hain ya nahi.
- **Kyun karta hai?** Type aliases map karta hai (`"int"` ➔ `"integer"`, `"double"` ➔ `"number"`). `"number"` matches `"integer"`, aur `"string"` matches any primitive type.
- **Need kyun hai?** Score me data type compatibility reward/penalty apply karne ke liye.

---

### 1.17 `get_json_value`
```python
def get_json_value(self, data: Any, path: str) -> list[Any]:
```
- **Kya karta hai?** Given JSON document me se specified JSONPath ke according exact value(s) query karke list return karta hai.
- **Kyun karta hai?** `_parse_json_path` ke tokens ke through JSON object/array traversing karta hai.
- **Need kyun hai?** Excel sheet fill karte waqt JSON source data se actual data extract karne ke liye.

---

### 1.18 `_parse_json_path`
```python
def _parse_json_path(self, path: str) -> list[str]:
```
- **Kya karta hai?** String JSONPath (e.g. `"$['policy']['number']"`) ko token list me parse karta hai (e.g. `['policy', 'number']`).
- **Kyun karta hai?** Regex `re.compile(r"\['((?:\\.|[^'])*)'\]|\[\*\]")` through brackets aur array wildcards `[*]` split karta hai.
- **Need kyun hai?** JSONPath parsing without external library dependencies.

---

### 1.19 `populate_template`
```python
def populate_template(self, *, workbook, sheet, json_data, mappings: list[dict[str, Any]]) -> dict[str, Any]:
```
- **Kya karta hai?** OpenPyXL sheet me actual values write karta hai (Scalar values as well as Array table expansion) aur write statistics summary dict return karta hai.
- **Kyun karta hai?**
  1. Mappings ko Scalar vs Repeating Groups me divide karta hai.
  2. Scalar values write karta hai.
  3. Array repeating groups ke liye section `safe_limit` calculate karta hai taaki niche ki fields overwrite na hon.
  4. Styles copy karke rows cell-by-cell populate karta hai.
- **Need kyun hai?** Excel Generation Core Feature execution.

---

## 2. Class: `_WorkbookAnalyzerAdapter`

Yeh internal class openpyxl Workbook object ko read karke uske visually formatted rows, labels, values, aur section hierarchy ko statistical heuristics se auto-detect karti hai.

---

### 2.1 `analyze`
```python
def analyze(self):
```
- **Kya karta hai?** Workbook ke sabhi non-hidden worksheets ko analyze karke JSON schema return karta hai.
- **Kyun karta hai?** Har sheet ke liye `_analyze_rows`, `_discover_font_levels`, `_classify_rows`, `_detect_value_column`, aur `_build_hierarchy` orchestrate karta hai.
- **Need kyun hai?** Excel Analyzer entry point.

---

### 2.2 `_analyze_rows`
```python
def _analyze_rows(self, ws):
```
- **Kya karta hai?** Worksheet ke 1 se `MAX_SCAN_ROWS` (10,000) tak ki non-empty rows aur columns extract karta hai.
- **Kyun karta hai?** Empty rows filter out karta hai aur populated column scores calculate karta hai.
- **Need kyun hai?** Scanning boundary setup.

---

### 2.3 `_build_row_analysis`
```python
def _build_row_analysis(self, row_number, cells):
```
- **Kya karta hai?** Row level features dictionary extract karta hai containing `label`, `value_cells`, `font_size`, `bold`, `fill_type`, `fill_color`, `indent`, `is_merged_header`.
- **Kyun karta hai?** Visual layout metadata capture karta hai.
- **Need kyun hai?** Row classification ML-like scoring rules ke liye features supply karna.

---

### 2.4 `_find_label_cell`
```python
def _find_label_cell(self, cells):
```
- **Kya karta hai?** Row ke andar pehli non-empty, non-merged cell find karta hai jise Header/Label mana ja sake.
- **Kyun karta hai?** `next((c for c in cells if not isinstance(c, MergedCell) and not self._is_empty(c)), None)` execute karta hai.
- **Need kyun hai?** Label cell location identification.

---

### 2.5 `_find_value_cell`
```python
def _find_value_cell(self, cells, label_cell):
```
- **Kya karta hai?** Label cell ke right side par pehli non-empty ya styled input cell locate karta hai.
- **Kyun karta hai?** Form layout detection (Key ➔ Value mapping cell).
- **Need kyun hai?** Data fill Target Cell determine karne ke liye.

---

### 2.6 `_discover_font_levels`
```python
def _discover_font_levels(self, rows):
```
- **Kya karta hai?** Sheet me present all font sizes ko sort karke header levels map karta hai. (Largest font = Level 0 Root Header).
- **Kyun karta hai?** Multi-level section hierarchy font size par depend karti hai.
- **Need kyun hai?** Section hierarchy levels assign karne ke liye.

---

### 2.7 `_classify_rows`
```python
def _classify_rows(self, rows, font_levels):
```
- **Kya karta hai?** Statistical Scoring Formula se har row ko classify karta hai as `"root_header"`, `"section"`, `"field"`, ya `"uncertain"`.
- **Kyun karta hai?** Feature Score Formulas:
  - `field_score` = 0.40 * has_data_value + 0.10 * indent + 0.10 * not_bold
  - `header_score` = 0.20 * bold + 0.20 * is_merged + 0.15 * solid_fill + 0.25 * max_font_size
- **Need kyun hai?** Visual formatting se structural semantics detect karna.

---

### 2.8 `_detect_value_column`
```python
def _detect_value_column(self, ws, column_scores):
```
- **Kya karta hai?** Primary Data Value Column letter detect karta hai (e.g. `"B"`).
- **Kyun karta hai?** `max(column_scores, key=column_scores.get)` calculate karta hai.
- **Need kyun hai?** Fallback target cell assignment jab specific value cell explicit na ho.

---

### 2.9 `_build_hierarchy`
```python
def _build_hierarchy(self, rows, value_column, max_row):
```
- **Kya karta hai?** Stack Data Structure use karke classified rows se Tree Hierarchy (`sections` ➔ `children` ➔ `fields`) construct karta hai.
- **Kyun karta hai?** Level comparison karke parent-child section nesting map karta hai.
- **Need kyun hai?** Tree Schema generation for Excel templates.

---

### 2.10 `_close_section_ranges`
```python
def _close_section_ranges(self, sections: list[dict[str, Any]], max_row: int):
```
- **Kya karta hai?** Recursively har section aur child section ka `end_row` boundary calculate karta hai.
- **Kyun karta hai?** Agla section start hone se 1 row pehle current section range end karta hai (`sections[i+1]["row"] - 1`).
- **Need kyun hai?** Array truncation limits aur table overflow boundary check ke liye.

---

### 2.11 Utility Methods

#### `_is_empty(self, cell)`
- **Kya/Kyun/Need:** Cell empty (None / Merged / Blank String) hai ya nahi check karta hai taaki invalid cells ignore ho sakein.

#### `_get_merge_info(self, cell)`
- **Kya/Kyun/Need:** Check karta hai cell merged range ka anchor cell hai ya nahi, taaki merged header formatting detection ho sake.

#### `_is_formula(self, cell)`
- **Kya/Kyun/Need:** Check karta hai value `=` se start hoti hai ya nahi, taaki Excel formulas label/value parsing me mix na hon.

#### `_get_fill_color(self, cell)`
- **Kya/Kyun/Need:** Solid background fill colors (RGB / Indexed / Theme) extract karta hai header detection weightage ke liye.

#### `_clean_label(self, value)`
- **Kya/Kyun/Need:** `" ".join(str(value).split())` se newlines aur extra spaces clean karta hai taaki uniform label string mil sake.

---
*Created automatically by Antigravity AI Code Assistant.*
