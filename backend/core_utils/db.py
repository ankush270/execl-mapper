from pymongo import MongoClient
from core_utils.config import Config

class Database:
    def __init__(self):
        self.client = MongoClient(Config.MONGO_URI)
        self.db = self.client[Config.DATABASE_NAME]
        self.json_collection = self.db["json_documents"]
        self.template_collection = self.db["excel_templates"]
        self.mapping_collection = self.db["excel_mappings"]
        self.export_collection = self.db["excel_exports"]