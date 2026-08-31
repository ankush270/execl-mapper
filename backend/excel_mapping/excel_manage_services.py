from flask import Blueprint, request, send_file
from flasgger import swag_from

from excel_mapping.excel_manage_component import ExcelManageComponent
from excel_mapping.excel_manage_spec import ExcelMappingSpec

excel_mapping_bp = Blueprint("excel_mapping", __name__)
excel_component = ExcelManageComponent()

@excel_mapping_bp.post("/json/upload")
@swag_from(ExcelMappingSpec.upload_json())
def upload_json():
    try:
        result = excel_component.upload_json(request.files.get("file"))
        return {"success": True, "data": result}, 201
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 400

@excel_mapping_bp.get("/json")
@swag_from(ExcelMappingSpec.get_all_json())
def get_jsons():
    try:
        result = excel_component.get_all_json()
        return {"success": True, "data": result}, 200
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 500

@excel_mapping_bp.get("/json/<document_id>")
@swag_from(ExcelMappingSpec.get_json())
def get_json_by_id(document_id):
    try:
        result = excel_component.get_json(document_id)
        return {"success": True, "data": result}, 200
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 400

@excel_mapping_bp.get("/excel/templates")
@swag_from(ExcelMappingSpec.get_templates())
def get_templates_route():
    try:
        result = excel_component.get_templates()
        return {"success": True, "data": result}, 200
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 500

@excel_mapping_bp.get("/excel/templates/<file_name>/preview")
@swag_from(ExcelMappingSpec.preview_template())
def preview_template_route(file_name):
    try:
        result = excel_component.preview_template(file_name)
        return {"success": True, "data": result}, 200
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 400

@excel_mapping_bp.get("/excel/templates/<file_name>/headers")
@swag_from(ExcelMappingSpec.get_template_headers())
def get_template_headers_route(file_name):
    try:
        sheet_name = request.args.get("sheet_name")
        result = excel_component.get_template_headers(file_name, sheet_name)
        return {"success": True, "data": result}, 200
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 400

@excel_mapping_bp.post("/excel/generate")
@swag_from(ExcelMappingSpec.generate_excel())
def generate_excel_route():
    try:
        data = request.get_json()
        result = excel_component.generate_excel(data)
        return {"success": True, "data": result}, 201
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 400

@excel_mapping_bp.get("/excel/download/<export_id>")
@swag_from(ExcelMappingSpec.download_excel())
def download_excel_route(export_id):
    try:
        file_path = excel_component.get_export_path(export_id)
        return send_file(file_path, as_attachment=True)
    except Exception as exc:
        return {"success": False, "message": str(exc)}, 404