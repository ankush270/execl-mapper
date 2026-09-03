# Detailed Function-by-Function Documentation: `excel_manage_component.py`

Yeh document [`excel_manage_component.py`](file:///e:/Projects/excelmapping/backend/excel_mapping/excel_manage_component.py) file me maujood **`ExcelManageComponent` Class ke sabhi Methods** ki exhaustive, detailed guide hai.
Har method ke baare me 3 mukhya baatein samjhayi gayi hain:
1. **Kya karta hai?** (Functionality & Input/Output)
2. **Kyun karta hai?** (Internal Logic & Workflow)
3. **Need kyun hai?** (Business Utility & System Importance)

---

## 📌 Table of Contents
- [Class: ExcelManageComponent](#class-excelmanagecomponent)
  - [1. `__init__`](#1-__init__)
  - [2. `upload_json`](#2-upload_json)
  - [3. `get_all_json`](#3-get_all_json)
  - [4. `get_json`](#4-get_json)
  - [5. `get_templates`](#5-get_templates)
  - [6. `analyze_template`](#6-analyze_template)
  - [7. `_build_and_store_mapping`](#7-_build_and_store_mapping)
  - [8. `create_mapping`](#8-create_mapping)
  - [9. `generate_excel`](#9-generate_excel)
  - [10. `get_export_path`](#10-get_export_path)

---

## Class: `ExcelManageComponent`

`ExcelManageComponent` application ki main **Business Logic Layer** hai jo API routes (`excel_manage_services.py`), Database (`Database()`), File Storage, aur Algorithmic Helper (`ExcelManageHelper()`) ke beech bridge ka kaam karti hai.

---

### 1. `__init__`
```python
def __init__(self) -> None:
    self.db = Database()
    self.helper = ExcelManageHelper()
```
- **Kya karta hai?** Class instance initialize karte waqt MongoDB database connection (`self.db`) aur Excel helper utility instance (`self.helper`) ko attach karta hai.
- **Kyun karta hai?** Database collections (`json_collection`, `mapping_collection`, `export_collection`) aur helper functions ko pure component class me reuse karne ke liye initialization karta hai.
- **Need kyun hai?** Sabhi business logic methods ko DB access aur helper calculations ki zaroorat padti hai.

---

### 2. `upload_json`
```python
def upload_json(self, file) -> dict[str, str]:
```
- **Kya karta hai?** User upload ki gayi raw `.json` file ko validate, parse, aur MongoDB `json_collection` me save karta hai.
- **Kyun karta hai?**
  1. Check karta hai file Uploaded hai ya nahi (`file.filename`).
  2. Check karta hai extension `.json` hai ya nahi.
  3. `json.load(file)` se syntax validity check karta hai (`json.JSONDecodeError` handling).
  4. Document record construct karta hai: `{"file_name": ..., "json_data": ..., "created_at": datetime.now(timezone.utc)}`.
  5. MongoDB me insert karke logger me inserted ID print karta hai aur `{"document_id": ..., "file_name": ...}` return karta hai.
- **Need kyun hai?** Core JSON Upload Feature API service.

---

### 3. `get_all_json`
```python
def get_all_json(self) -> list[dict[str, Any]]:
```
- **Kya karta hai?** MongoDB me uploaded sabhi JSON files ki list (`document_id`, `file_name`, `created_at`) return karta hai, sorted by creation date descending (`created_at: -1`).
- **Kyun karta hai?** Network payload aur RAM optimize karne ke liye Query Projection use karta hai (`{"json_data": 0}`), jisse heavy raw JSON content database se fetch nahi hota.
- **Need kyun hai?** Frontend dashboard aur JSON selection dropdown list ke liye fast response deliver karna.

---

### 4. `get_json`
```python
def get_json(self, document_id: str) -> dict[str, Any]:
```
- **Kya karta hai?** Given `document_id` se specific JSON document ka full record including raw `"json_data"` content payload database se fetch karke return karta hai.
- **Kyun karta hai?** `self.helper.to_object_id(document_id)` se string ID ko MongoDB `ObjectId` format me convert karta hai. Agar document exist nahi karta to `ValueError("JSON not found.")` throw karta hai.
- **Need kyun hai?** Specific JSON preview/viewing feature on frontend UI.

---

### 5. `get_templates`
```python
def get_templates(self) -> list[dict[str, str]]:
```
- **Kya karta hai?** Server ke `Config.TEMPLATE_FOLDER` directory se sabhi available Excel templates (`.xlsx`, `.xlsm`) ki list return karta hai.
- **Kyun karta hai?** Directory existence check karta hai, files ko alphabetically sort karta hai, aur extension filter out karke `[{"template_name": ..., "file_name": ...}]` list banata hai.
- **Need kyun hai?** Frontend UI par Template selection list display karne ke liye.

---

### 6. `analyze_template`
```python
def analyze_template(self, file) -> dict[str, Any]:
```
- **Kya karta hai?** User uploaded external Excel template file ko analyze karke uska complete sheet, section, aur field structure schema JSON return karta hai.
- **Kyun karta hai?**
  1. Temporary disk file (`NamedTemporaryFile`) banakar uploaded stream save karta hai.
  2. OpenPyXL `load_workbook` se file load karta hai (`keep_vba` enabled for `.xlsm`).
  3. `self.helper.analyze_template_workbook` call karke structural schema parse karta hai.
  4. `finally` block me workbook instance close karta hai aur temporary file disk se `unlink()` (delete) karta hai.
- **Need kyun hai?** Dynamic Excel Template Structure Analysis Feature.

---

### 7. `_build_and_store_mapping`
```python
def _build_and_store_mapping(self, document, template_path, sheet_name, workbook) -> tuple[str, list[dict[str, Any]]]:
```
- **Kya karta hai?** Helper logic orchestrate karke JSON Schema aur Excel Sheet Schema analyze karta hai, automated field mappings generate karta hai, aur record ko MongoDB `mapping_collection` me save karta hai.
- **Kyun karta hai?**
  1. `analyze_template_workbook` se target Excel sheet schema extract karta hai.
  2. `analyze_json` se JSON schema extract karta hai.
  3. `build_mappings` call karke automated pairings compute karta hai.
  4. Mappings empty hone par `ValueError("No JSON fields could be mapped...")` throw karta hai.
  5. `mapping_doc` construct karke MongoDB me insert karta hai.
- **Need kyun hai?** Code reusability (DRY) between `create_mapping` and `generate_excel` methods.

---

### 8. `create_mapping`
```python
def create_mapping(self, document_id: str, template_file: str, sheet_name: str) -> dict[str, Any]:
```
- **Kya karta hai?** Specific JSON document ID, Template File, aur Sheet Name ke beech explicitly mapping record create karke database me save karta hai.
- **Kyun karta hai?** Validates JSON document existence in MongoDB, validates template file path via `safe_template_path`, verifies sheet existence in workbook, aur `_build_and_store_mapping` call karta hai.
- **Need kyun hai?** Standalone Mapping creation API endpoint.

---

### 9. `generate_excel`
```python
def generate_excel(self, data: dict[str, Any]) -> dict[str, Any]:
```
- **Kya karta hai?** End-to-End Excel Generation Feature: Provided JSON data and mapping rules ke dwara target Excel template me data write karke naye generated `.xlsx` / `.xlsm` file ko disk par save karta hai aur export metadata MongoDB `export_collection` me store karta hai.
- **Kyun karta hai?**
  1. `document_id`, `template_file`, `sheet_name` validate karta hai.
  2. Agar `mapping_id` diya gaya ho to MongoDB se mapping record fetch karta hai; warna `_build_and_store_mapping` se fresh mapping calculate karta hai.
  3. `self.helper.populate_template` execute karke cell values write karta hai.
  4. Unique output filename `uuid.uuid4().hex` se `Config.GENERATED_FOLDER` directory me file save karta hai.
  5. Export record Mongo me insert karta hai aur written statistics return karta hai.
- **Need kyun hai?** Application ka core main output feature — Final Population & Excel File Generation.

---

### 10. `get_export_path`
```python
def get_export_path(self, export_id: str) -> str:
```
- **Kya karta hai?** Given `export_id` se MongoDB `export_collection` lookup karke server disk par saved generated Excel file ka absolute path return karta hai.
- **Kyun karta hai?**
  1. `export_id` ko `ObjectId` me parse karke database record search karta hai.
  2. Verify karta hai file physical disk par exist karti hai ya nahi (`file_path.exists()`).
  3. Agar file delete ho chuki ho to `ValueError("Generated Excel file no longer exists.")` throw karta hai.
- **Need kyun hai?** User ko final generated Excel file browser me download karwane ke liye (`/excel_download/<export_id>`).

---
*Created automatically by Antigravity AI Code Assistant.*
