from flask import Blueprint, request, send_file, jsonify
from flasgger import swag_from
from excel_mapping.excel_manage_component import ExcelManageComponent
from excel_mapping.excel_manage_spec import ExcelMappingSpec

excel_mapping_bp = Blueprint("excel_mapping",__name__,)
excel_component = ExcelManageComponent()

@excel_mapping_bp.post("/json_upload")
@swag_from(ExcelMappingSpec.upload_json())
def upload_json():
    try:
        files = request.files.getlist("file")
        if not files:
            raise ValueError("At least one JSON file is required.")
        results = []
        for file in files:
            results.append(excel_component.upload_json(file))

        return {"success": True,"data": results,}, 201

    except Exception as exc:
        return {"success": False,"message": str(exc),}, 400


@excel_mapping_bp.get("/json")
@swag_from(ExcelMappingSpec.get_all_json())
def get_jsons():
    try:
        result = excel_component.get_all_json()
        return {"success": True,"data": result,}, 200
    except Exception as exc:
        return {"success": False,"message": str(exc),}, 500


@excel_mapping_bp.get("/json/<document_id>")
@swag_from(ExcelMappingSpec.get_json())
def get_json_by_id(document_id):
    try:
        result = excel_component.get_json(document_id)
        return {"success": True,"data": result,}, 200
    except Exception as exc:
        return {"success": False,"message": str(exc),}, 400


@excel_mapping_bp.get("/excel_templates")
@swag_from(ExcelMappingSpec.get_templates())
def get_templates_route():
    try:
        result = excel_component.get_templates()
        return {"success": True,"data": result,}, 200
    except Exception as exc:
        return {"success": False,"message": str(exc),}, 500


@excel_mapping_bp.get("/excel_templates/<path:filename>")
def download_template_route(filename):
    """Serve a template file for frontend preview."""
    try:
        file_path = excel_component.helper.safe_template_path(filename)
        return send_file(file_path, as_attachment=False)
    except ValueError as exc:
        return {"success": False, "message": str(exc)}, 404
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 500

@excel_mapping_bp.post("/excel_generate")
@swag_from(ExcelMappingSpec.generate_excel())
def generate_excel_route():
    try:
        data = request.get_json(silent=True) or {}
        required = ["document_id","template_file","sheet_name",]
        missing = [key for key in required if not data.get(key)]
        if missing:
            raise ValueError("Missing required fields: "+ ", ".join(missing))
        result = excel_component.generate_excel(data)
        return {"success": True,"data": result,}, 201
    except Exception as exc:
        return {"success": False,"message": str(exc),}, 400


@excel_mapping_bp.get("/excel_download/<export_id>")
@swag_from(ExcelMappingSpec.download_excel())
def download_excel_route(export_id):
    try:
        file_path = (excel_component.get_export_path(export_id))
        return send_file(file_path, as_attachment=True)
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 404


@excel_mapping_bp.post("/excel_analyze")
@swag_from(ExcelMappingSpec.analyze_template())
def analyze_template_route():
    try:
        file = request.files.get("file")
        if not file:
            raise ValueError("Excel template file is required.")
        result = excel_component.analyze_template(file)
        return {"success": True, "data": result}, 200
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 400


@excel_mapping_bp.post("/excel_mapping")
@swag_from(ExcelMappingSpec.create_mapping())
def create_mapping_route():
    try:
        data = request.get_json(silent=True) or {}
        required = ["document_id", "template_file", "sheet_name"]
        missing = [key for key in required if not data.get(key)]
        if missing:
            raise ValueError("Missing required fields: " + ", ".join(missing))
        result = excel_component.create_mapping(
            document_id=data["document_id"],
            template_file=data["template_file"],
            sheet_name=data["sheet_name"],
        )
        return {"success": True, "data": result}, 201
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 400




