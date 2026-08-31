class ExcelMappingSpec:
    @staticmethod
    def upload_json():
        return {
            "tags": ["JSON Documents"],
            "summary": "Upload a JSON document",
            "description": "Uploads a raw JSON file to MongoDB to be used for dynamic Excel mappings.",
            "parameters": [
                {
                    "name": "file",
                    "in": "formData",
                    "type": "file",
                    "required": True,
                    "description": "The JSON file to upload"
                }
            ],
            "responses": {
                "201": {
                    "description": "JSON document uploaded successfully",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": True},
                            "data": {
                                "type": "object",
                                "properties": {
                                    "document_id": {"type": "string", "example": "651a2b3c4d5e6f7a8b9c0d1e"},
                                    "file_name": {"type": "string", "example": "data.json"}
                                }
                            }
                        }
                    }
                },
                "400": {
                    "description": "Missing file, incorrect extension or failed upload",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": False},
                            "message": {"type": "string", "example": "Only JSON files are allowed"}
                        }
                    }
                }
            }
        }

    @staticmethod
    def get_all_json():
        return {
            "tags": ["JSON Documents"],
            "summary": "Get all uploaded JSON documents metadata",
            "description": "Retrieves metadata of all uploaded JSON files, sorted by creation date descending.",
            "responses": {
                "200": {
                    "description": "A list of JSON documents metadata",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": True},
                            "data": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "document_id": {"type": "string", "example": "651a2b3c4d5e6f7a8b9c0d1e"},
                                        "file_name": {"type": "string", "example": "data.json"},
                                        "created_at": {"type": "string", "format": "date-time", "example": "2026-08-31T13:25:00Z"}
                                    }
                                }
                            }
                        }
                    }
                },
                "500": {
                    "description": "Internal Server Error",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": False},
                            "message": {"type": "string", "example": "Failed to get JSON list: database error"}
                        }
                    }
                }
            }
        }

    @staticmethod
    def get_json():
        return {
            "tags": ["JSON Documents"],
            "summary": "Get specific JSON document details",
            "description": "Retrieves the complete uploaded JSON document by ID, including its parsed content data.",
            "parameters": [
                {
                    "name": "document_id",
                    "in": "path",
                    "type": "string",
                    "required": True,
                    "description": "The unique document object ID in MongoDB"
                }
            ],
            "responses": {
                "200": {
                    "description": "JSON document details retrieved successfully",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": True},
                            "data": {
                                "type": "object",
                                "properties": {
                                    "document_id": {"type": "string", "example": "651a2b3c4d5e6f7a8b9c0d1e"},
                                    "file_name": {"type": "string", "example": "data.json"},
                                    "json_data": {"type": "object", "description": "Raw data content of the JSON file"}
                                }
                            }
                        }
                    }
                },
                "400": {
                    "description": "Invalid document ID or document not found",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": False},
                            "message": {"type": "string", "example": "JSON not found"}
                        }
                    }
                }
            }
        }

    @staticmethod
    def get_templates():
        return {
            "tags": ["Excel Templates"],
            "summary": "Get all Excel templates",
            "description": "Lists all available Excel (.xlsx) template files in the templates directory.",
            "responses": {
                "200": {
                    "description": "A list of Excel templates",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": True},
                            "data": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "template_name": {"type": "string", "example": "SalesTemplate"},
                                        "file_name": {"type": "string", "example": "SalesTemplate.xlsx"}
                                    }
                                }
                            }
                        }
                    }
                },
                "500": {
                    "description": "Failed to read templates directory",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": False},
                            "message": {"type": "string", "example": "Template folder not found"}
                        }
                    }
                }
            }
        }

    @staticmethod
    def preview_template():
        return {
            "tags": ["Excel Templates"],
            "summary": "Preview Excel template contents",
            "description": "Generates a structured preview showing sheets and limited rows/cols in the template Excel.",
            "parameters": [
                {
                    "name": "file_name",
                    "in": "path",
                    "type": "string",
                    "required": True,
                    "description": "The exact name of the Excel template file"
                }
            ],
            "responses": {
                "200": {
                    "description": "Excel template preview retrieved successfully",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": True},
                            "data": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "sheet_name": {"type": "string", "example": "Sheet1"},
                                        "rows": {
                                            "type": "array",
                                            "items": {
                                                "type": "array",
                                                "items": {"type": "string", "example": "ColumnHeader1"}
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                },
                "400": {
                    "description": "Excel template file not found or invalid format",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": False},
                            "message": {"type": "string", "example": "Excel template not found"}
                        }
                    }
                }
            }
        }

    @staticmethod
    def get_template_headers():
        return {
            "tags": ["Excel Templates"],
            "summary": "Get headers of a specific sheet in a template",
            "description": "Extracts the header columns (from row 1) of the designated template sheet.",
            "parameters": [
                {
                    "name": "file_name",
                    "in": "path",
                    "type": "string",
                    "required": True,
                    "description": "The template Excel file name"
                },
                {
                    "name": "sheet_name",
                    "in": "query",
                    "type": "string",
                    "required": True,
                    "description": "The specific sheet inside the template to inspect"
                }
            ],
            "responses": {
                "200": {
                    "description": "Template headers retrieved successfully",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": True},
                            "data": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string", "example": "A"},
                                        "excel_column": {"type": "string", "example": "A"},
                                        "header_name": {"type": "string", "example": "User Name"}
                                    }
                                }
                            }
                        }
                    }
                },
                "400": {
                    "description": "Template file or sheet not found",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": False},
                            "message": {"type": "string", "example": "Sheet not found"}
                        }
                    }
                }
            }
        }

    @staticmethod
    def generate_excel():
        return {
            "tags": ["Excel Generation"],
            "summary": "Generate Excel spreadsheet using mappings",
            "description": "Generates a mapped Excel spreadsheet by correlating fields from a JSON document to the headers of an Excel template sheet, saving it locally.",
            "parameters": [
                {
                    "name": "body",
                    "in": "body",
                    "required": True,
                    "schema": {
                        "type": "object",
                        "required": ["document_id", "template_file", "sheet_name"],
                        "properties": {
                            "document_id": {"type": "string", "example": "651a2b3c4d5e6f7a8b9c0d1e"},
                            "template_file": {"type": "string", "example": "SalesTemplate.xlsx"},
                            "sheet_name": {"type": "string", "example": "Sheet1"}
                        }
                    }
                }
            ],
            "responses": {
                "201": {
                    "description": "Excel spreadsheet generated successfully",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": True},
                            "data": {
                                "type": "object",
                                "properties": {
                                    "export_id": {"type": "string", "example": "651a2b3c4d5e6f7a8b9c0d2f"},
                                    "file_name": {"type": "string", "example": "1a2b3c4d5e6f7a8b9c0d.xlsx"}
                                }
                            }
                        }
                    }
                },
                "400": {
                    "description": "Invalid parameters, missing dependencies or template mismatch",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": False},
                            "message": {"type": "string", "example": "No mapping could be created"}
                        }
                    }
                }
            }
        }

    @staticmethod
    def download_excel():
        return {
            "tags": ["Excel Generation"],
            "summary": "Download generated Excel file",
            "description": "Streams the generated Excel spreadsheet file back to the client as an attachment.",
            "parameters": [
                {
                    "name": "export_id",
                    "in": "path",
                    "type": "string",
                    "required": True,
                    "description": "The unique ID of the generated Excel export"
                }
            ],
            "responses": {
                "200": {
                    "description": "Excel file downloaded successfully",
                    "headers": {
                        "Content-Disposition": {
                            "type": "string",
                            "description": "attachment; filename=..."
                        },
                        "Content-Type": {
                            "type": "string",
                            "description": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        }
                    }
                },
                "404": {
                    "description": "Generated export record not found or file missing",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": False},
                            "message": {"type": "string", "example": "Export not found"}
                        }
                    }
                }
            }
        }
