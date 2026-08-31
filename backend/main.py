import os
import sys
from pathlib import Path
from flask import Flask
from flasgger import Swagger
from dotenv import load_dotenv

# Ensure backend directory is in sys.path
backend_path = Path(__file__).resolve().parent
if str(backend_path) not in sys.path:
    sys.path.insert(0, str(backend_path))

from excel_mapping.excel_manage_services import excel_mapping_bp
from core_utils.config import Config
from core_utils.logging import logger

def create_app():
    load_dotenv()
    
    app = Flask(__name__)
    app.config.from_object(Config)
    
    # Configure Swagger
    app.config['SWAGGER'] = {
        'title': 'Excel Mapping API',
        'uiversion': 3,
        'description': 'API documentation for uploading JSON files and mapping them to Excel templates.',
        'specs_route': '/apidocs/'
    }
    
    # Initialize Swagger
    Swagger(app)
    
    # Register blueprint
    app.register_blueprint(excel_mapping_bp, url_prefix="/api")
    
    # Create templates directory if it doesn't exist
    templates_dir = Path(Config.TEMPLATE_FOLDER)
    templates_dir.mkdir(parents=True, exist_ok=True)
    
    # Create generated excels directory if it doesn't exist
    generated_dir = Path(Config.GENERATED_FOLDER)
    generated_dir.mkdir(parents=True, exist_ok=True)
    
    logger.info("Application initialized successfully.")
    return app

if __name__ == '__main__':
    app = create_app()
    # Run the server
    host = os.getenv("FLASK_HOST", "0.0.0.0")
    port = int(os.getenv("FLASK_PORT", 5000))
    debug = os.getenv("FLASK_DEBUG", "True").lower() == "true"
    
    logger.info(f"Starting server on {host}:{port} with debug={debug}")
    app.run(host=host, port=port, debug=debug)
