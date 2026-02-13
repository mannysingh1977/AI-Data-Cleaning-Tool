"""
Main Orchestration Script
Coordinates authentication, file downloading, and text extraction.
"""

import logging
import json
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any

from config import Config
from auth import AuthenticationManager
from sharepoint_api import SharePointClient
from text_extraction import TextExtractor


# Set up logging
def setup_logging():
    """Configure logging for the application."""
    Config.create_directories()
    
    log_file = Config.LOGS_DIR / f"sharepoint_extraction_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    logging.basicConfig(
        level=getattr(logging, Config.LOG_LEVEL),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info(f"Logging initialized. Log file: {log_file}")
    return logger


class SharePointExtractor:
    """Main orchestrator for SharePoint document extraction and processing."""
    
    def __init__(self):
        """Initialize the extractor."""
        self.logger = logging.getLogger(__name__)
        self.auth_manager = None
        self.sp_client = None
        self.access_token = None
        
        # Statistics
        self.stats = {
            "files_found": 0,
            "files_downloaded": 0,
            "files_extracted": 0,
            "errors": []
        }
    
    def authenticate(self) -> bool:
        """
        Authenticate with Microsoft Graph API.
        
        Returns:
            bool: True if authentication successful, False otherwise
        """
        self.logger.info("Starting authentication...")
        
        self.auth_manager = AuthenticationManager()
        self.access_token = self.auth_manager.get_access_token()
        
        if self.access_token:
            self.sp_client = SharePointClient(self.access_token)
            self.logger.info("✅ Authentication successful")
            return True
        else:
            self.logger.error("❌ Authentication failed")
            return False
    
    def list_files(self, limit: int = None) -> List[Dict[str, Any]]:
        """
        List files from SharePoint/OneDrive.
        
        Args:
            limit: Maximum number of files to list (None for all)
        
        Returns:
            List[Dict[str, Any]]: List of file metadata
        """
        if not self.sp_client:
            self.logger.error("Not authenticated. Call authenticate() first.")
            return []
        
        self.logger.info(f"Listing files (limit: {limit if limit else 'all'})...")
        
        files = self.sp_client.get_all_files_recursive(
            file_types=Config.SUPPORTED_FILE_TYPES,
            limit=limit
        )
        
        self.stats["files_found"] = len(files)
        self.logger.info(f"Found {len(files)} files")
        
        return files
    
    def download_files(
        self,
        file_items: List[Dict[str, Any]],
        download_dir: Path = None
    ) -> List[Path]:
        """
        Download files from SharePoint/OneDrive.
        
        Args:
            file_items: List of file metadata
            download_dir: Directory to save files
        
        Returns:
            List[Path]: List of downloaded file paths
        """
        if not self.sp_client:
            self.logger.error("Not authenticated. Call authenticate() first.")
            return []
        
        if download_dir is None:
            download_dir = Config.DOWNLOAD_DIR
        
        self.logger.info(f"Downloading {len(file_items)} files to {download_dir}...")
        
        downloaded = self.sp_client.download_multiple_files(file_items, download_dir)
        
        self.stats["files_downloaded"] = len(downloaded)
        self.logger.info(f"Successfully downloaded {len(downloaded)}/{len(file_items)} files")
        
        return downloaded
    
    def extract_text_from_files(
        self,
        file_paths: List[Path],
        output_dir: Path = None
    ) -> List[Path]:
        """
        Extract text from downloaded files.
        
        Args:
            file_paths: List of file paths to extract text from
            output_dir: Directory to save extracted text
        
        Returns:
            List[Path]: List of successfully extracted text file paths
        """
        if output_dir is None:
            output_dir = Config.EXTRACTED_TEXT_DIR
        
        self.logger.info(f"Extracting text from {len(file_paths)} files...")
        
        extracted = []
        
        for file_path in file_paths:
            try:
                output_path = TextExtractor.extract_and_save(file_path, output_dir)
                
                if output_path:
                    extracted.append(output_path)
                else:
                    self.stats["errors"].append({
                        "file": file_path.name,
                        "error": "Text extraction failed"
                    })
            
            except Exception as e:
                self.logger.error(f"Error processing {file_path.name}: {e}")
                self.stats["errors"].append({
                    "file": file_path.name,
                    "error": str(e)
                })
        
        self.stats["files_extracted"] = len(extracted)
        self.logger.info(f"Successfully extracted text from {len(extracted)}/{len(file_paths)} files")
        
        return extracted
    
    def run_full_pipeline(
        self,
        limit: int = None,
        download: bool = True,
        extract: bool = True
    ) -> Dict[str, Any]:
        """
        Run the complete pipeline: authenticate, list, download, and extract.
        
        Args:
            limit: Maximum number of files to process
            download: Whether to download files
            extract: Whether to extract text
        
        Returns:
            Dict[str, Any]: Summary statistics
        """
        print("\n" + "="*60)
        print("📚 SHAREPOINT DOCUMENT EXTRACTION PIPELINE")
        print("="*60 + "\n")
        
        # Step 1: Authenticate
        print("Step 1: Authentication")
        print("-" * 40)
        if not self.authenticate():
            print("\n❌ Pipeline failed: Authentication error\n")
            return self.stats
        
        # Step 2: List files
        print("\nStep 2: Listing Files")
        print("-" * 40)
        files = self.list_files(limit=limit)
        
        if not files:
            print("\n⚠️  No files found or error occurred\n")
            return self.stats
        
        print(f"\n✅ Found {len(files)} files:\n")
        for i, file_item in enumerate(files[:10], 1):  # Show first 10
            name = file_item.get("name", "Unknown")
            size = file_item.get("size", 0)
            size_kb = size / 1024
            print(f"  {i}. {name} ({size_kb:.1f} KB)")
        
        if len(files) > 10:
            print(f"  ... and {len(files) - 10} more")
        
        # Step 3: Download files (optional)
        downloaded_files = []
        
        if download:
            print(f"\nStep 3: Downloading Files")
            print("-" * 40)
            downloaded_files = self.download_files(files)
            print(f"\n✅ Downloaded {len(downloaded_files)} files")
        else:
            print("\nStep 3: Skipping download (download=False)")
        
        # Step 4: Extract text (optional)
        if extract and downloaded_files:
            print(f"\nStep 4: Extracting Text")
            print("-" * 40)
            extracted_files = self.extract_text_from_files(downloaded_files)
            print(f"\n✅ Extracted text from {len(extracted_files)} files")
        else:
            if not extract:
                print("\nStep 4: Skipping text extraction (extract=False)")
            elif not downloaded_files:
                print("\nStep 4: Skipping text extraction (no downloaded files)")
        
        # Print summary
        self.print_summary()
        
        return self.stats
    
    def print_summary(self):
        """Print a summary of the extraction process."""
        print("\n" + "="*60)
        print("📊 EXTRACTION SUMMARY")
        print("="*60)
        print(f"Files found:      {self.stats['files_found']}")
        print(f"Files downloaded: {self.stats['files_downloaded']}")
        print(f"Files extracted:  {self.stats['files_extracted']}")
        print(f"Errors:           {len(self.stats['errors'])}")
        
        if self.stats['errors']:
            print("\n⚠️  Errors encountered:")
            for error in self.stats['errors'][:5]:  # Show first 5 errors
                print(f"  • {error['file']}: {error['error']}")
            
            if len(self.stats['errors']) > 5:
                print(f"  ... and {len(self.stats['errors']) - 5} more errors")
        
        print("="*60 + "\n")
    
    def save_summary_report(self, output_path: Path = None):
        """
        Save a summary report to a JSON file.
        
        Args:
            output_path: Path to save report (defaults to logs directory)
        """
        if output_path is None:
            output_path = Config.LOGS_DIR / f"extraction_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        try:
            with open(output_path, 'w') as f:
                json.dump(self.stats, f, indent=2)
            
            self.logger.info(f"Summary report saved to: {output_path}")
            print(f"📄 Summary report saved to: {output_path}\n")
        
        except Exception as e:
            self.logger.error(f"Failed to save summary report: {e}")


def main():
    """Main entry point."""
    # Set up logging
    logger = setup_logging()
    
    # Print configuration
    Config.print_config()
    
    # Validate configuration
    if not Config.validate():
        print("\n❌ Configuration validation failed!")
        print("Please create a .env file based on .env.template and fill in your credentials.\n")
        return
    
    # Create necessary directories
    Config.create_directories()
    
    # Create and run extractor
    extractor = SharePointExtractor()
    
    # Run the pipeline
    # For initial testing, limit to 5 files
    stats = extractor.run_full_pipeline(
        limit=5,  # Change to None to process all files
        download=True,
        extract=True
    )
    
    # Save summary report
    extractor.save_summary_report()
    
    logger.info("Pipeline execution completed")


if __name__ == "__main__":
    main()
