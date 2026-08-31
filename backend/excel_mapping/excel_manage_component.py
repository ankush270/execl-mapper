from datetime import datetime, timezone
import json
import uuid
from pathlib import Path
from bson import ObjectId
from openpyxl import load_workbook

from core_utils.db import Database
from core_utils.logging import logger
from core_utils.config import Config
from excel_mapping.excel_manage_helper import ExcelManageHelper


class ExcelManageComponent:
    def __init__(self):
        try:
            self.db = Database()
            self.helper = ExcelManageHelper()
        except Exception as exc:
            logger.error(f"Component initialization failed: {exc}")
            raise

    # ──────────────────── JSON Upload / Retrieval ───────────────────

    def upload_json(self, file):
        try:
            if not file:
                raise ValueError("JSON file is required")
            if not file.filename.lower().endswith(".json"):
                raise ValueError("Only JSON files are allowed")

            # Validate that the content is valid JSON before saving
            try:
                json_data = json.load(file)
            except json.JSONDecodeError as e:
                raise ValueError(f"Invalid JSON content: {e}")

            document = {
                "file_name":  file.filename,
                "json_data":  json_data,
                "created_at": datetime.now(timezone.utc),
            }
            result = self.db.json_collection.insert_one(document)
            logger.info(f"JSON uploaded: {result.inserted_id}")
            return {"document_id": str(result.inserted_id), "file_name": file.filename}
        except Exception as exc:
            logger.error(f"JSON upload failed: {exc}")
            raise

    def get_all_json(self):
        try:
            documents = (
                self.db.json_collection
                .find({}, {"json_data": 0})
                .sort("created_at", -1)
            )
            result = []
            for document in documents:
                result.append({
                    "document_id": str(document["_id"]),
                    "file_name":   document["file_name"],
                    "created_at":  document["created_at"],
                })
            return result
        except Exception as exc:
            logger.error(f"Failed to get JSON list: {exc}")
            raise

    def get_json(self, document_id):
        try:
            document = self.db.json_collection.find_one({"_id": ObjectId(document_id)})
            if not document:
                raise ValueError(f"JSON document not found for id: {document_id}")
            return {
                "document_id": str(document["_id"]),
                "file_name":   document["file_name"],
                "json_data":   document["json_data"],
            }
        except Exception as exc:
            logger.error(f"Failed to get JSON: {exc}")
            raise

    # ──────────────────── Template Handling ─────────────────────────

    def get_templates(self):
        try:
            template_folder = Path(Config.TEMPLATE_FOLDER)
            if not template_folder.exists():
                raise ValueError(
                    f"Template folder not found at: {template_folder}. "
                    "Please create the folder and add .xlsx templates."
                )
            templates = []
            for file in template_folder.iterdir():
                if file.is_file() and file.suffix.lower() == ".xlsx":
                    templates.append({
                        "template_name": file.stem,
                        "file_name":     file.name,
                    })
            if not templates:
                logger.warning(f"No .xlsx templates found in {template_folder}")
            return templates
        except Exception as exc:
            logger.error(f"Failed to get templates: {exc}")
            raise

    def preview_template(self, file_name):
        try:
            file_path = Path(Config.TEMPLATE_FOLDER) / file_name
            if not file_path.exists():
                raise ValueError(f"Excel template '{file_name}' not found in templates folder")
            if file_path.suffix.lower() != ".xlsx":
                raise ValueError(f"Invalid file type '{file_path.suffix}'. Only .xlsx is supported")

            workbook = load_workbook(file_path, data_only=True)
            try:
                return self.helper.get_excel_preview(workbook)
            finally:
                workbook.close()
        except Exception as exc:
            logger.error(f"Template preview failed: {exc}")
            raise

    def get_template_headers(self, file_name, sheet_name):
        try:
            # Guard: sheet_name is a required query param
            if not sheet_name or not sheet_name.strip():
                raise ValueError(
                    "sheet_name query parameter is required. "
                    "Use /excel/templates/<file_name>/preview to see available sheet names."
                )

            file_path = Path(Config.TEMPLATE_FOLDER) / file_name
            if not file_path.exists():
                raise ValueError(f"Excel template '{file_name}' not found")

            workbook = load_workbook(file_path, data_only=True)
            try:
                # Case-insensitive sheet name matching
                matched_sheet = self._resolve_sheet_name(workbook, sheet_name)
                if matched_sheet is None:
                    available = ", ".join(workbook.sheetnames)
                    raise ValueError(
                        f"Sheet '{sheet_name}' not found. "
                        f"Available sheets: [{available}]"
                    )
                sheet = workbook[matched_sheet]
                headers = self.helper.get_excel_headers(sheet)
                if not headers:
                    logger.warning(
                        f"No headers found in sheet '{matched_sheet}'. "
                        "The sheet may be empty or use an unsupported layout."
                    )
                return headers
            finally:
                workbook.close()
        except Exception as exc:
            logger.error(f"Failed to get template headers: {exc}")
            raise

    # ──────────────────── Mapping & Generation ──────────────────────

    def save_mapping(self, document_id, template_file, sheet_name, mappings):
        try:
            if not document_id or not template_file or not sheet_name or not mappings:
                raise ValueError(
                    "document_id, template_file, sheet_name and mappings are all required"
                )
            mapping = {
                "document_id":   ObjectId(document_id),
                "template_file": template_file,
                "sheet_name":    sheet_name,
                "mappings":      mappings,
                "created_at":    datetime.now(timezone.utc),
            }
            result = self.db.mapping_collection.insert_one(mapping)
            logger.info(f"Mapping created: {result.inserted_id}")
            return {"mapping_id": str(result.inserted_id)}
        except Exception as exc:
            logger.error(f"Mapping creation failed: {exc}")
            raise

    def generate_excel(self, data):
        """
        Main orchestration method.

        Steps
        ─────
        1. Validate request payload
        2. Fetch JSON document from DB
        3. Load Excel template
        4. Resolve sheet (case-insensitive)
        5. Auto-detect header orientation (horizontal / vertical)
        6. Extract Excel headers (merged-cell safe)
        7. Create fuzzy JSON ↔ header mappings
        8. Extract JSON values via JSONPath
        9. Write values into sheet (with unequal-length protection)
        10. Save output and record in DB
        """
        workbook = None
        try:
            document_id  = data.get("document_id")
            template_file = data.get("template_file")
            sheet_name   = data.get("sheet_name")

            if not document_id or not template_file or not sheet_name:
                raise ValueError(
                    "Request body must contain: document_id, template_file, sheet_name"
                )

            # Step 2: Fetch JSON
            document = self.db.json_collection.find_one({"_id": ObjectId(document_id)})
            if not document:
                raise ValueError(f"JSON document not found for id: {document_id}")

            # Step 3: Load template
            template_path = Path(Config.TEMPLATE_FOLDER) / Path(template_file).name
            if not template_path.exists():
                raise ValueError(
                    f"Excel template '{template_file}' not found in templates folder"
                )
            if template_path.suffix.lower() != ".xlsx":
                raise ValueError(
                    f"Invalid template extension '{template_path.suffix}'. Use .xlsx"
                )

            workbook = load_workbook(template_path)

            # Step 4: Resolve sheet name (case-insensitive)
            matched_sheet = self._resolve_sheet_name(workbook, sheet_name)
            if matched_sheet is None:
                available = ", ".join(workbook.sheetnames)
                raise ValueError(
                    f"Sheet '{sheet_name}' does not exist. "
                    f"Available sheets: [{available}]"
                )
            sheet = workbook[matched_sheet]

            # Steps 5–6: Detect orientation + extract headers
            excel_headers = self.helper.get_excel_headers(sheet)
            if not excel_headers:
                raise ValueError(
                    f"No headers found in sheet '{matched_sheet}'. "
                    "The sheet may be empty or have an unsupported layout."
                )

            # Determine orientation from first header entry
            orientation = excel_headers[0].get("orientation", "horizontal")

            # ── MIXED / CROSS-TAB branch ──────────────────────────────────────
            if orientation == "mixed":
                mixed_info = excel_headers[0]  # single structured dict
                logger.info(
                    f"Mixed (cross-tab) layout detected in sheet '{matched_sheet}'. "
                    f"{len(mixed_info['row_headers'])} row headers × "
                    f"{len(mixed_info['col_headers'])} col headers."
                )
                mappings = self.helper.fill_mixed_layout(
                    sheet, mixed_info, document["json_data"]
                )
                # Check at least one cell was filled
                if not any(m.get("value") is not None for m in mappings):
                    row_labels = [rh["name"] for rh in mixed_info["row_headers"]]
                    col_labels = [ch["name"] for ch in mixed_info["col_headers"]]
                    raise ValueError(
                        "Mixed layout detected but no data could be written. "
                        "JSON structure does not match the cross-tab layout.\n"
                        f"Expected row labels : {row_labels}\n"
                        f"Expected col labels : {col_labels}\n"
                        "Supported JSON formats:\n"
                        "  A) Nested dict : {{\"North\": {{\"Q1\": 100, ...}}, ...}}\n"
                        "  B) Array+key   : [{{\"region\": \"North\", \"Q1\": 100, ...}}, ...]\n"
                        "  C) Wrapped     : {{\"data\": [...]}}"
                    )

            # ── HORIZONTAL / VERTICAL branch ──────────────────────────────────
            else:
                # Step 7: Create mappings (fuzzy matching)
                mappings = self.helper.create_mapping(document["json_data"], excel_headers)
                if not mappings:
                    json_fields = []
                    self.helper._find_json_fields(document["json_data"], "$", json_fields)
                    json_names   = [f["name"] for f in json_fields]
                    excel_names  = [h["header_name"] for h in excel_headers]
                    raise ValueError(
                        f"No mapping could be created between JSON and Excel template.\n"
                        f"JSON fields   : {json_names}\n"
                        f"Excel headers : {excel_names}\n"
                        "Tip: Field names don't need to be identical — spaces, underscores "
                        "and casing are ignored — but at least one pair must share the same word."
                    )

                # Step 9: Extract values + write
                extracted_data = self.helper.extract_mapping_data(
                    document["json_data"], mappings
                )
                self.helper.write_excel_data(sheet, mappings, extracted_data)

            # Step 8: Save mapping record (common to all orientations)
            mapping_res = self.save_mapping(
                document_id, template_file, matched_sheet, mappings
            )
            mapping_id = mapping_res["mapping_id"]

            # Step 10: Save file
            generated_folder = Path(Config.GENERATED_FOLDER)
            generated_folder.mkdir(parents=True, exist_ok=True)
            file_name   = f"{uuid.uuid4().hex}.xlsx"
            output_path = generated_folder / file_name
            workbook.save(output_path)

            export = {
                "document_id":   ObjectId(document_id),
                "mapping_id":    mapping_id,
                "template_file": template_file,
                "sheet_name":    matched_sheet,
                "file_name":     file_name,
                "file_path":     str(output_path),
                "created_at":    datetime.now(timezone.utc),
            }
            result = self.db.export_collection.insert_one(export)
            logger.info(f"Excel generated: {result.inserted_id}")
            return {"export_id": str(result.inserted_id), "file_name": file_name}

        except Exception as exc:
            logger.error(f"Excel generation failed: {exc}")
            raise
        finally:
            if workbook:
                try:
                    workbook.close()
                except Exception as exc:
                    logger.error(f"Failed to close workbook: {exc}")

    def get_export_path(self, export_id):
        try:
            export = self.db.export_collection.find_one({"_id": ObjectId(export_id)})
            if not export:
                raise ValueError(f"Export record not found for id: {export_id}")
            file_path = Path(export["file_path"])
            if not file_path.exists():
                raise ValueError(
                    f"Generated Excel file no longer exists on disk: {file_path}. "
                    "It may have been deleted."
                )
            return str(file_path)
        except Exception as exc:
            logger.error(f"Failed to get export path: {exc}")
            raise

    # ──────────────────── Private Helpers ───────────────────────────

    @staticmethod
    def _resolve_sheet_name(workbook, sheet_name: str):
        """
        Case-insensitive sheet name resolution.
        Returns the exact sheet name from the workbook, or None if not found.

        Solves: "Sheet 1" vs "sheet1" vs "SHEET1" mismatches.
        """
        target = sheet_name.strip().lower()
        for name in workbook.sheetnames:
            if name.strip().lower() == target:
                return name
        return None