"""
Authentication Module
Handles Microsoft Azure AD authentication using MSAL (Microsoft Authentication Library).
Supports device code flow for easy testing without a web server.
"""

import json
import logging
from pathlib import Path
from typing import Optional, Dict, Any

import msal

from config import Config

# Set up logging
logger = logging.getLogger(__name__)


class AuthenticationManager:
    """Manages authentication with Microsoft Graph API using Azure AD."""
    
    def __init__(self):
        """Initialize the authentication manager."""
        self.client_id = Config.CLIENT_ID
        self.tenant_id = Config.TENANT_ID
        self.authority = Config.AUTHORITY
        self.scopes = Config.SCOPES
        self.token_cache_file = Config.TOKEN_CACHE_FILE
        
        # Initialize token cache
        self.cache = self._load_cache()
        
        # Create MSAL public client application
        self.app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            token_cache=self.cache
        )
        
        logger.info("Authentication manager initialized")
    
    def _load_cache(self) -> msal.SerializableTokenCache:
        """
        Load token cache from file if it exists.
        
        Returns:
            msal.SerializableTokenCache: Token cache object
        """
        cache = msal.SerializableTokenCache()
        
        if self.token_cache_file.exists():
            try:
                with open(self.token_cache_file, 'r') as f:
                    cache.deserialize(f.read())
                logger.info("Token cache loaded from file")
            except Exception as e:
                logger.warning(f"Failed to load token cache: {e}")
        
        return cache
    
    def _save_cache(self) -> None:
        """Save token cache to file if it has changed."""
        if self.cache.has_state_changed:
            try:
                with open(self.token_cache_file, 'w') as f:
                    f.write(self.cache.serialize())
                logger.info("Token cache saved to file")
            except Exception as e:
                logger.error(f"Failed to save token cache: {e}")
    
    def get_access_token(self, use_device_code: bool = True) -> Optional[str]:
        """
        Get an access token for Microsoft Graph API.
        
        Args:
            use_device_code: If True, use device code flow (interactive).
                           If False, try to get token silently from cache.
        
        Returns:
            Optional[str]: Access token if successful, None otherwise
        """
        # Try to get token silently from cache first
        accounts = self.app.get_accounts()
        
        if accounts:
            logger.info(f"Found {len(accounts)} account(s) in cache")
            result = self.app.acquire_token_silent(
                scopes=self.scopes,
                account=accounts[0]
            )
            
            if result and "access_token" in result:
                logger.info("✅ Token acquired silently from cache")
                self._save_cache()
                return result["access_token"]
            else:
                logger.info("Silent token acquisition failed, need to authenticate")
        
        # If silent acquisition failed and device code is allowed, use device code flow
        if use_device_code:
            return self._authenticate_with_device_code()
        else:
            logger.error("No valid token in cache and device code flow not allowed")
            return None
    
    def _authenticate_with_device_code(self) -> Optional[str]:
        """
        Authenticate using device code flow (user-friendly for testing).
        
        Returns:
            Optional[str]: Access token if successful, None otherwise
        """
        logger.info("Starting device code authentication flow...")
        
        # Initiate device code flow
        flow = self.app.initiate_device_flow(scopes=self.scopes)
        
        if "user_code" not in flow:
            logger.error(f"Failed to create device flow: {flow.get('error_description', 'Unknown error')}")
            return None
        
        # Display instructions to user
        print("\n" + "="*60)
        print("🔐 AUTHENTICATION REQUIRED")
        print("="*60)
        print(flow["message"])
        print("="*60)
        print("\nWaiting for you to complete sign-in in the browser...")
        print("(After approving, close the browser tab and come back here)\n")

        # Wait for the user to authenticate
        try:
            result = self.app.acquire_token_by_device_flow(
                flow,
                exit_condition=lambda flow: print(".", end="", flush=True) or False
            )
            print()  # newline after the dots
            print(f"\nDEBUG - Response keys: {list(result.keys())}")
            if "error" in result:
                print(f"DEBUG - Error: {result.get('error')}")
                print(f"DEBUG - Description: {result.get('error_description')}")

            if "access_token" in result:
                logger.info("✅ Authentication successful!")
                self._save_cache()
                
                # Display user info
                if "id_token_claims" in result:
                    claims = result["id_token_claims"]
                    print(f"✅ Logged in as: {claims.get('name', 'Unknown')} ({claims.get('preferred_username', 'Unknown')})\n")
                
                return result["access_token"]
            else:
                error = result.get("error_description", result.get("error", "Unknown error"))
                logger.error(f"Authentication failed: {error}")
                print(f"\n❌ Authentication failed: {error}\n")
                return None
        
        except Exception as e:
            logger.error(f"Exception during authentication: {e}")
            print(f"\n❌ Authentication error: {e}\n")
            return None
    
    def get_user_info(self, access_token: str) -> Optional[Dict[str, Any]]:
        """
        Get information about the authenticated user.
        
        Args:
            access_token: Valid access token
        
        Returns:
            Optional[Dict[str, Any]]: User information if successful, None otherwise
        """
        import requests
        
        headers = {"Authorization": f"Bearer {access_token}"}
        
        try:
            response = requests.get(
                f"{Config.GRAPH_API_ENDPOINT}/me",
                headers=headers
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Failed to get user info: {e}")
            return None
    
    def clear_cache(self) -> None:
        """Clear the token cache (forces re-authentication)."""
        try:
            if self.token_cache_file.exists():
                self.token_cache_file.unlink()
                logger.info("Token cache cleared")
                print("✅ Token cache cleared. You'll need to authenticate again.")
            else:
                logger.info("No token cache to clear")
                print("ℹ️  No token cache found.")
        except Exception as e:
            logger.error(f"Failed to clear cache: {e}")
            print(f"❌ Failed to clear cache: {e}")


def test_authentication():
    """Test the authentication module."""
    # Set up basic logging for testing
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    print("\n🧪 Testing Authentication Module\n")
    
    # Validate config
    if not Config.validate():
        print("\n❌ Configuration validation failed. Please check your .env file.")
        return
    
    # Create auth manager
    auth_manager = AuthenticationManager()
    
    # Get access token
    print("Attempting to get access token...\n")
    token = auth_manager.get_access_token()
    
    if token:
        print(f"✅ Access token acquired (length: {len(token)} characters)")
        print(f"   Token preview: {token[:20]}...{token[-20:]}\n")
        
        # Get user info
        print("Fetching user information...\n")
        user_info = auth_manager.get_user_info(token)
        
        if user_info:
            print("✅ User Information:")
            print(f"   Name: {user_info.get('displayName', 'N/A')}")
            print(f"   Email: {user_info.get('userPrincipalName', 'N/A')}")
            print(f"   Job Title: {user_info.get('jobTitle', 'N/A')}")
            print(f"   Office: {user_info.get('officeLocation', 'N/A')}")
        else:
            print("❌ Failed to get user information")
    else:
        print("❌ Failed to acquire access token")
    
    print("\n✅ Authentication test completed!\n")


if __name__ == "__main__":
    test_authentication()
