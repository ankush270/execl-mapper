class ExcelMappingSpec:
    @staticmethod
    def upload_json():
        return {
            "tags": ["JSON Documents"],
            "summary": "Upload JSON document(s)",
            "description": "Uploads one or multiple raw JSON files to MongoDB to be used for dynamic Excel mappings.",
            "parameters": [
                {
                    "name": "file",
                    "in": "formData",
                    "type": "file",
                    "required": True,
                    "description": "The JSON file(s) to upload"
                }
            ],
            "responses": {
                "201": {
                    "description": "JSON document(s) uploaded successfully",
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
                                        "file_name": {"type": "string", "example": "data.json"}
                                    }
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
                            "message": {"type": "string", "example": "Only JSON files are allowed."}
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
                            "message": {"type": "string", "example": "Failed to retrieve JSON list."}
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
                            "message": {"type": "string", "example": "JSON not found."}
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
            "description": "Lists all available Excel (.xlsx, .xlsm) template files in the templates directory.",
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
                            "message": {"type": "string", "example": "Template folder not found."}
                        }
                    }
                }
            }
        }

    @staticmethod
    def analyze_template():
        return {
            "tags": ["Excel Templates"],
            "summary": "Analyze Excel template structure",
            "description": "Parses an uploaded Excel template file (.xlsx, .xlsm) and extracts sheet schemas, section hierarchies, and fields.",
            "parameters": [
                {
                    "name": "file",
                    "in": "formData",
                    "type": "file",
                    "required": True,
                    "description": "The Excel template file to analyze"
                }
            ],
            "responses": {
                "200": {
                    "description": "Excel template schema extracted successfully",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "file_name": {"type": "string", "example": "SalesTemplate.xlsx"},
                            "sheets": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "name": {"type": "string", "example": "Sheet1"},
                                        "max_row": {"type": "integer", "example": 50},
                                        "max_column": {"type": "integer", "example": 10},
                                        "value_column": {"type": "string", "example": "B"},
                                        "sections": {"type": "array", "items": {"type": "object"}},
                                        "orphan_fields": {"type": "array", "items": {"type": "object"}},
                                        "needs_review": {"type": "array", "items": {"type": "integer"}}
                                    }
                                }
                            }
                        }
                    }
                },
                "400": {
                    "description": "Invalid template file or format",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": False},
                            "message": {"type": "string", "example": "Only .xlsx and .xlsm templates are supported."}
                        }
                    }
                }
            }
        }

    @staticmethod
    def create_mapping():
        return {
            "tags": ["Excel Mappings"],
            "summary": "Create mapping between JSON and Excel template",
            "description": "Generates and saves field mappings between a specified JSON document and an Excel template sheet.",
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
                "200": {
                    "description": "Mapping created successfully",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": True},
                            "data": {
                                "type": "object",
                                "properties": {
                                    "mapping_id": {"type": "string", "example": "651a2b3c4d5e6f7a8b9c0d99"},
                                    "document_id": {"type": "string", "example": "651a2b3c4d5e6f7a8b9c0d1e"},
                                    "template_file": {"type": "string", "example": "SalesTemplate.xlsx"},
                                    "sheet_name": {"type": "string", "example": "Sheet1"},
                                    "mapping_count": {"type": "integer", "example": 15},
                                    "mappings": {"type": "array", "items": {"type": "object"}}
                                }
                            }
                        }
                    }
                },
                "400": {
                    "description": "JSON/Template not found or no fields mapped",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "success": {"type": "boolean", "example": False},
                            "message": {"type": "string", "example": "No JSON fields could be mapped to the Excel template."}
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
            "description": "Populates an Excel template sheet with data from a JSON document using existing or newly generated mappings.",
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
                            "sheet_name": {"type": "string", "example": "Sheet1"},
                            "mapping_id": {"type": "string", "example": "651a2b3c4d5e6f7a8b9c0d99", "description": "Optional pre-existing mapping ID"}
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
                                    "file_name": {"type": "string", "example": "1a2b3c4d5e6f7a8b9c0d.xlsx"},
                                    "mapping_id": {"type": "string", "example": "651a2b3c4d5e6f7a8b9c0d99"},
                                    "mapping_count": {"type": "integer", "example": 15},
                                    "written": {
                                        "type": "object",
                                        "properties": {
                                            "scalar_values_written": {"type": "integer", "example": 10},
                                            "repeating_values_written": {"type": "integer", "example": 25},
                                            "empty_values": {"type": "integer", "example": 2},
                                            "total_mappings": {"type": "integer", "example": 37},
                                            "truncated_values": {"type": "integer", "example": 0}
                                        }
                                    }
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
                            "message": {"type": "string", "example": "No JSON fields could be mapped."}
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
                            "message": {"type": "string", "example": "Export not found."}
                        }
                    }
                }
            }
        }
