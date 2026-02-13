"""
Configuration Module
Loads environment variables and provides configuration settings for the application.
"""

import os
from pathlib import Path
from dotenv import load_dotenv
from typing import List

# Load environment variables from .env file
load_dotenv()

class Config:
    """Application configuration settings loaded from environment variables."""
    
    # Azure AD App Registration
    CLIENT_ID: str = os.getenv("CLIENT_ID", "")
    TENANT_ID: str = os.getenv("TENANT_ID", "")
    CLIENT_SECRET: str = os.getenv("CLIENT_SECRET", "")  # Optional
    
    # Microsoft Graph API
    GRAPH_API_ENDPOINT: str = os.getenv("GRAPH_API_ENDPOINT", "https://graph.microsoft.com/v1.0")
    AUTHORITY: str = os.getenv("AUTHORITY", "https://login.microsoftonline.com/common")
    
    # Scopes (permissions)
    SCOPES: List[str] = os.getenv("SCOPES", "Files.Read.All User.Read").split()
    
    # Local storage paths
    BASE_DIR: Path = Path(__file__).parent
    DOWNLOAD_DIR: Path = BASE_DIR / os.getenv("DOWNLOAD_DIR", "downloaded_files")
    EXTRACTED_TEXT_DIR: Path = BASE_DIR / os.getenv("EXTRACTED_TEXT_DIR", "extracted_text")
    LOGS_DIR: Path = BASE_DIR / os.getenv("LOGS_DIR", "logs")
    
    # File types
    SUPPORTED_FILE_TYPES: List[str] = os.getenv("SUPPORTED_FILE_TYPES", "docx,pdf").split(",")
    
    # Pagination
    MAX_ITEMS_PER_PAGE: int = int(os.getenv("MAX_ITEMS_PER_PAGE", "200"))
    
    # Logging
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "INFO")
    
    # Token cache
    TOKEN_CACHE_FILE: Path = BASE_DIR / ".msal_token_cache.json"
    
    @classmethod
    def validate(cls) -> bool:
        """
        Validate that required configuration values are set.
        
        Returns:
            bool: True if configuration is valid, False otherwise
        """
        if not cls.CLIENT_ID:
            print("❌ ERROR: CLIENT_ID not set in .env file")
            return False
        
        if not cls.TENANT_ID:
            print("❌ ERROR: TENANT_ID not set in .env file")
            return False
        
        return True
    
    @classmethod
    def create_directories(cls) -> None:
        """Create necessary directories if they don't exist."""
        cls.DOWNLOAD_DIR.mkdir(exist_ok=True)
        cls.EXTRACTED_TEXT_DIR.mkdir(exist_ok=True)
        cls.LOGS_DIR.mkdir(exist_ok=True)
    
    @classmethod
    def print_config(cls) -> None:
        """Print current configuration (without secrets)."""
        print("\n📋 Configuration:")
        print(f"  Client ID: {cls.CLIENT_ID[:8]}..." if cls.CLIENT_ID else "  Client ID: NOT SET")
        print(f"  Tenant ID: {cls.TENANT_ID[:8]}..." if cls.TENANT_ID else "  Tenant ID: NOT SET")
        print(f"  Graph API: {cls.GRAPH_API_ENDPOINT}")
        print(f"  Scopes: {', '.join(cls.SCOPES)}")
        print(f"  Download Dir: {cls.DOWNLOAD_DIR}")
        print(f"  Extracted Text Dir: {cls.EXTRACTED_TEXT_DIR}")
        print(f"  Logs Dir: {cls.LOGS_DIR}")
        print(f"  Supported File Types: {', '.join(cls.SUPPORTED_FILE_TYPES)}")
        print(f"  Log Level: {cls.LOG_LEVEL}\n")


if __name__ == "__main__":
    # Test configuration
    Config.print_config()
    if Config.validate():
        print("✅ Configuration is valid!")
        Config.create_directories()
        print("✅ Directories created!")
    else:
        print("❌ Configuration is invalid. Please check your .env file.")
