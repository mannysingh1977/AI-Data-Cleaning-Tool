"""
SharePoint API Module
Handles interactions with Microsoft Graph API for SharePoint and OneDrive operations.
"""

import logging
import requests
from pathlib import Path
from typing import List, Dict, Any, Optional
from tqdm import tqdm

from config import Config

# Set up logging
logger = logging.getLogger(__name__)


class SharePointClient:
    """Client for interacting with SharePoint and OneDrive via Microsoft Graph API."""
    
    def __init__(self, access_token: str):
        """
        Initialize SharePoint client.
        
        Args:
            access_token: Valid access token for Microsoft Graph API
        """
        self.access_token = access_token
        self.base_url = Config.GRAPH_API_ENDPOINT
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        logger.info("SharePoint client initialized")
    
    def _make_request(self, url: str, method: str = "GET", **kwargs) -> Optional[Dict[str, Any]]:
        """
        Make an HTTP request to Microsoft Graph API.
        
        Args:
            url: API endpoint URL
            method: HTTP method (GET, POST, etc.)
            **kwargs: Additional arguments for requests
        
        Returns:
            Optional[Dict[str, Any]]: JSON response if successful, None otherwise
        """
        try:
            response = requests.request(
                method=method,
                url=url,
                headers=self.headers,
                **kwargs
            )
            response.raise_for_status()
            return response.json() if response.content else {}
        
        except requests.exceptions.HTTPError as e:
            logger.error(f"HTTP error: {e}")
            if hasattr(e.response, 'text'):
                logger.error(f"Response: {e.response.text}")
            return None
        
        except Exception as e:
            logger.error(f"Request failed: {e}")
            return None
    
    def list_my_drive_files(
        self,
        file_types: Optional[List[str]] = None,
        limit: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """
        List files from the user's OneDrive, recursing into all folders.

        Args:
            file_types: List of file extensions to filter (e.g., ['docx', 'pdf'])
            limit: Maximum number of files to return (None for all)

        Returns:
            List[Dict[str, Any]]: List of file metadata
        """
        logger.info("Listing files from OneDrive (recursive)...")

        if file_types is None:
            file_types = Config.SUPPORTED_FILE_TYPES

        all_files = []
        # Queue of folder URLs to process (start with root)
        folders_to_scan = [f"{self.base_url}/me/drive/root/children"]

        while folders_to_scan and (limit is None or len(all_files) < limit):
            url = folders_to_scan.pop(0)

            # Handle pagination within each folder
            while url and (limit is None or len(all_files) < limit):
                logger.debug(f"Fetching: {url}")

                response = self._make_request(url)
                if not response:
                    break

                items = response.get("value", [])

                for item in items:
                    if "folder" in item:
                        # It's a folder — add to scan queue
                        folder_id = item.get("id")
                        folder_name = item.get("name", "Unknown")
                        logger.info(f"Found folder: {folder_name}")
                        folders_to_scan.append(
                            f"{self.base_url}/me/drive/items/{folder_id}/children"
                        )
                    elif "file" in item:
                        # It's a file — check extension
                        file_name = item.get("name", "")
                        file_ext = file_name.split(".")[-1].lower() if "." in file_name else ""

                        if not file_types or file_ext in file_types:
                            # Add folder path info for display
                            parent = item.get("parentReference", {})
                            item["_folder_path"] = parent.get("path", "").replace("/drive/root:", "") or "/"
                            all_files.append(item)

                            if limit and len(all_files) >= limit:
                                break

                # Check for next page within this folder
                url = response.get("@odata.nextLink")

        logger.info(f"Found {len(all_files)} files in OneDrive")
        return all_files[:limit] if limit else all_files
    
    def search_files(
        self,
        query: str = "",
        file_types: Optional[List[str]] = None,
        limit: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """
        Search for files across OneDrive and SharePoint.
        
        Args:
            query: Search query (empty string returns all files)
            file_types: List of file extensions to filter
            limit: Maximum number of files to return
        
        Returns:
            List[Dict[str, Any]]: List of file metadata
        """
        logger.info(f"Searching for files with query: '{query}'")
        
        if file_types is None:
            file_types = Config.SUPPORTED_FILE_TYPES
        
        all_files = []
        
        # Search endpoint
        if query:
            url = f"{self.base_url}/me/drive/root/search(q='{query}')"
        else:
            # If no query, just list all files recursively
            url = f"{self.base_url}/me/drive/root/children"
        
        while url and (limit is None or len(all_files) < limit):
            response = self._make_request(url)
            if not response:
                break
            
            files = response.get("value", [])
            
            for item in files:
                if "file" in item:
                    file_name = item.get("name", "")
                    file_ext = file_name.split(".")[-1].lower() if "." in file_name else ""
                    
                    if not file_types or file_ext in file_types:
                        all_files.append(item)
                        
                        if limit and len(all_files) >= limit:
                            break
            
            url = response.get("@odata.nextLink")
        
        logger.info(f"Found {len(all_files)} files matching search")
        return all_files[:limit] if limit else all_files
    
    def get_all_files_recursive(
        self,
        file_types: Optional[List[str]] = None,
        limit: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """
        Get all files from OneDrive recursively (including subdirectories).
        
        Args:
            file_types: List of file extensions to filter
            limit: Maximum number of files to return
        
        Returns:
            List[Dict[str, Any]]: List of file metadata
        """
        logger.info("Fetching all files recursively from OneDrive...")
        
        if file_types is None:
            file_types = Config.SUPPORTED_FILE_TYPES
        
        # Using search with empty query is the easiest way to get all files recursively
        return self.search_files(query="", file_types=file_types, limit=limit)
    
    def download_file(
        self,
        file_item: Dict[str, Any],
        download_dir: Optional[Path] = None
    ) -> Optional[Path]:
        """
        Download a file from OneDrive/SharePoint.
        
        Args:
            file_item: File metadata from Graph API
            download_dir: Directory to save file (defaults to Config.DOWNLOAD_DIR)
        
        Returns:
            Optional[Path]: Path to downloaded file if successful, None otherwise
        """
        if download_dir is None:
            download_dir = Config.DOWNLOAD_DIR
        
        file_name = file_item.get("name", "unknown_file")
        file_id = file_item.get("id")
        
        if not file_id:
            logger.error(f"No file ID found for {file_name}")
            return None
        
        # Get download URL
        download_url = file_item.get("@microsoft.graph.downloadUrl")
        
        if not download_url:
            # If download URL not in item, fetch it
            url = f"{self.base_url}/me/drive/items/{file_id}"
            response = self._make_request(url)
            if response:
                download_url = response.get("@microsoft.graph.downloadUrl")
        
        if not download_url:
            logger.error(f"Could not get download URL for {file_name}")
            return None
        
        # Download the file
        try:
            logger.info(f"Downloading: {file_name}")
            
            # Create download directory if it doesn't exist
            download_dir.mkdir(parents=True, exist_ok=True)
            
            # Download with progress bar
            response = requests.get(download_url, stream=True)
            response.raise_for_status()
            
            file_path = download_dir / file_name
            total_size = int(response.headers.get('content-length', 0))
            
            with open(file_path, 'wb') as f, tqdm(
                desc=file_name,
                total=total_size,
                unit='B',
                unit_scale=True,
                unit_divisor=1024,
            ) as progress_bar:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
                    progress_bar.update(len(chunk))
            
            logger.info(f"✅ Downloaded: {file_path}")
            return file_path
        
        except Exception as e:
            logger.error(f"Failed to download {file_name}: {e}")
            return None
    
    def download_multiple_files(
        self,
        file_items: List[Dict[str, Any]],
        download_dir: Optional[Path] = None
    ) -> List[Path]:
        """
        Download multiple files.
        
        Args:
            file_items: List of file metadata from Graph API
            download_dir: Directory to save files
        
        Returns:
            List[Path]: List of paths to successfully downloaded files
        """
        logger.info(f"Downloading {len(file_items)} files...")
        
        downloaded_files = []
        
        for file_item in file_items:
            file_path = self.download_file(file_item, download_dir)
            if file_path:
                downloaded_files.append(file_path)
        
        logger.info(f"Successfully downloaded {len(downloaded_files)}/{len(file_items)} files")
        return downloaded_files
    
    def get_file_metadata(self, file_id: str) -> Optional[Dict[str, Any]]:
        """
        Get detailed metadata for a specific file.
        
        Args:
            file_id: File ID from Graph API
        
        Returns:
            Optional[Dict[str, Any]]: File metadata if successful, None otherwise
        """
        url = f"{self.base_url}/me/drive/items/{file_id}"
        return self._make_request(url)
    
    @staticmethod
    def print_file_info(file_item: Dict[str, Any]) -> None:
        """
        Print formatted information about a file.
        
        Args:
            file_item: File metadata from Graph API
        """
        name = file_item.get("name", "Unknown")
        size = file_item.get("size", 0)
        created = file_item.get("createdDateTime", "Unknown")
        modified = file_item.get("lastModifiedDateTime", "Unknown")
        
        # Convert size to human-readable format
        size_kb = size / 1024
        size_mb = size_kb / 1024
        
        if size_mb >= 1:
            size_str = f"{size_mb:.2f} MB"
        else:
            size_str = f"{size_kb:.2f} KB"
        
        print(f"📄 {name}")
        print(f"   Size: {size_str}")
        print(f"   Created: {created}")
        print(f"   Modified: {modified}")


def test_sharepoint_api():
    """Test the SharePoint API module - lists all Word and PDF files."""
    from auth import AuthenticationManager

    # Set up logging (suppress info noise, only show warnings+)
    logging.basicConfig(
        level=logging.WARNING,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    print("\n" + "="*60)
    print("📂 ONEDRIVE FILE LISTING")
    print("="*60)

    # Authenticate
    auth_manager = AuthenticationManager()
    token = auth_manager.get_access_token()

    if not token:
        print("❌ Authentication failed")
        return

    # Create SharePoint client
    sp_client = SharePointClient(token)

    # Get ALL docx and pdf files (no limit)
    files = sp_client.list_my_drive_files(file_types=["docx", "pdf"], limit=None)

    if not files:
        print("\n❌ No .docx or .pdf files found in your OneDrive")
        return

    # Separate by file type
    pdf_files = [f for f in files if f.get("name", "").lower().endswith(".pdf")]
    docx_files = [f for f in files if f.get("name", "").lower().endswith(".docx")]

    # Sort each group alphabetically
    pdf_files.sort(key=lambda f: f.get("name", "").lower())
    docx_files.sort(key=lambda f: f.get("name", "").lower())

    # Print PDF files
    if pdf_files:
        print(f"\n📕 PDF FILES ({len(pdf_files)})")
        print("-" * 40)
        for i, f in enumerate(pdf_files, 1):
            name = f.get("name", "Unknown")
            folder = f.get("_folder_path", "/")
            size_kb = f.get("size", 0) / 1024
            print(f"  {i:3}. {name} ({size_kb:.1f} KB)  [{folder}]")

    # Print Word files
    if docx_files:
        print(f"\n📘 WORD FILES ({len(docx_files)})")
        print("-" * 40)
        for i, f in enumerate(docx_files, 1):
            name = f.get("name", "Unknown")
            folder = f.get("_folder_path", "/")
            size_kb = f.get("size", 0) / 1024
            print(f"  {i:3}. {name} ({size_kb:.1f} KB)  [{folder}]")

    # Summary
    print(f"\n{'='*60}")
    print(f"TOTAL: {len(files)} files ({len(pdf_files)} PDF, {len(docx_files)} Word)")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    test_sharepoint_api()