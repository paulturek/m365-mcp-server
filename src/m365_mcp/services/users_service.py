"""Microsoft 365 Users service.

Provides operations for:
- Getting current user profile
- Searching/listing organization users
- Getting user details by ID or email
- Getting user's manager and direct reports
- Getting user's photo

Graph API Reference:
    https://learn.microsoft.com/graph/api/resources/user

Required Scopes:
    - User.Read (current user profile)
    - User.ReadBasic.All (basic profiles of all users)
    - User.Read.All (full profiles - requires admin consent)
    - Directory.Read.All (organizational hierarchy)
"""

from typing import Any, Optional

from ..clients.graph_client import GraphClient


class UsersService:
    """Microsoft 365 Users management via Microsoft Graph.
    
    Example:
        >>> service = UsersService(graph_client)
        >>> me = await service.get_current_user()
        >>> users = await service.search_users("john")
        >>> manager = await service.get_manager("user-id")
    """
    
    # Default fields for user queries
    DEFAULT_SELECT = [
        "id", "displayName", "givenName", "surname", "mail",
        "userPrincipalName", "jobTitle", "department", "officeLocation",
        "mobilePhone", "businessPhones", "employeeId"
    ]
    
    # Extended fields (may require additional permissions)
    EXTENDED_SELECT = DEFAULT_SELECT + [
        "accountEnabled", "createdDateTime", "employeeType",
        "companyName", "streetAddress", "city", "state", "postalCode", "country"
    ]
    
    def __init__(self, client: GraphClient) -> None:
        """Initialize Users service."""
        self.client = client
    
    async def get_current_user(self) -> dict[str, Any]:
        """Get the current authenticated user's profile.
        
        Returns:
            User profile object
        """
        params = {"$select": ",".join(self.DEFAULT_SELECT)}
        return await self.client.get("/me", params=params)
    
    async def get_user(
        self,
        user_id: str,
        extended: bool = False
    ) -> dict[str, Any]:
        """Get a user by ID or userPrincipalName (email).
        
        Args:
            user_id: User ID (GUID) or userPrincipalName (email)
            extended: Whether to include extended profile fields
            
        Returns:
            User profile object
        """
        select = self.EXTENDED_SELECT if extended else self.DEFAULT_SELECT
        params = {"$select": ",".join(select)}
        return await self.client.get(f"/users/{user_id}", params=params)
    
    async def list_users(
        self,
        count: int = 100,
        filter_query: Optional[str] = None,
        order_by: str = "displayName",
    ) -> list[dict[str, Any]]:
        """List users in the organization.
        
        Args:
            count: Maximum number of users to return
            filter_query: OData filter (e.g., "department eq 'Sales'")
            order_by: Sort order (default: displayName)
            
        Returns:
            List of user objects
        """
        params: dict[str, Any] = {
            "$top": count,
            "$orderby": order_by,
            "$select": ",".join(self.DEFAULT_SELECT),
        }
        
        if filter_query:
            params["$filter"] = filter_query
        
        result = await self.client.get("/users", params=params)
        return result.get("value", [])
    
    async def search_users(
        self,
        query: str,
        count: int = 25,
    ) -> list[dict[str, Any]]:
        """Search for users by name or email.
        
        Args:
            query: Search query (searches displayName, mail, userPrincipalName)
            count: Maximum results to return
            
        Returns:
            List of matching user objects
        """
        # Use $filter with startswith for searching
        # Graph API doesn't have a simple search for users, so we use filter
        filter_query = (
            f"startswith(displayName, '{query}') or "
            f"startswith(mail, '{query}') or "
            f"startswith(userPrincipalName, '{query}') or "
            f"startswith(givenName, '{query}') or "
            f"startswith(surname, '{query}')"
        )
        
        params: dict[str, Any] = {
            "$top": count,
            "$filter": filter_query,
            "$select": ",".join(self.DEFAULT_SELECT),
            "$orderby": "displayName",
        }
        
        # Need ConsistencyLevel header for advanced queries
        result = await self.client.get(
            "/users",
            params=params,
            headers={"ConsistencyLevel": "eventual"}
        )
        return result.get("value", [])
    
    async def get_manager(self, user_id: Optional[str] = None) -> Optional[dict[str, Any]]:
        """Get a user's manager.
        
        Args:
            user_id: User ID or email. If None, gets current user's manager.
            
        Returns:
            Manager's user profile, or None if no manager
        """
        endpoint = "/me/manager" if user_id is None else f"/users/{user_id}/manager"
        
        try:
            result = await self.client.get(endpoint)
            return result
        except Exception as e:
            # No manager assigned returns 404
            if "404" in str(e) or "ResourceNotFound" in str(e):
                return None
            raise
    
    async def get_direct_reports(
        self,
        user_id: Optional[str] = None,
        count: int = 100,
    ) -> list[dict[str, Any]]:
        """Get a user's direct reports.
        
        Args:
            user_id: User ID or email. If None, gets current user's direct reports.
            count: Maximum results to return
            
        Returns:
            List of user profiles who report to this user
        """
        endpoint = "/me/directReports" if user_id is None else f"/users/{user_id}/directReports"
        
        params = {
            "$top": count,
            "$select": ",".join(self.DEFAULT_SELECT),
        }
        
        result = await self.client.get(endpoint, params=params)
        return result.get("value", [])
    
    async def get_user_photo(
        self,
        user_id: Optional[str] = None,
        size: str = "48x48"
    ) -> Optional[bytes]:
        """Get a user's profile photo.
        
        Args:
            user_id: User ID or email. If None, gets current user's photo.
            size: Photo size ('48x48', '64x64', '96x96', '120x120', '240x240',
                  '360x360', '432x432', '504x504', '648x648')
            
        Returns:
            Photo bytes, or None if no photo
        """
        if user_id:
            endpoint = f"/users/{user_id}/photos/{size}/$value"
        else:
            endpoint = f"/me/photos/{size}/$value"
        
        try:
            return await self.client.download_file(endpoint)
        except Exception as e:
            if "404" in str(e) or "ImageNotFound" in str(e):
                return None
            raise
    
    async def get_people(
        self,
        query: Optional[str] = None,
        count: int = 25,
    ) -> list[dict[str, Any]]:
        """Get people relevant to the current user.
        
        This uses the People API which returns people ordered by relevance
        to the current user (based on communication patterns, etc.)
        
        Args:
            query: Optional search query
            count: Maximum results
            
        Returns:
            List of person objects (includes relevance ranking)
        """
        params: dict[str, Any] = {"$top": count}
        
        if query:
            params["$search"] = f'"{query}"'
        
        result = await self.client.get("/me/people", params=params)
        return result.get("value", [])
    
    async def get_users_by_department(
        self,
        department: str,
        count: int = 100,
    ) -> list[dict[str, Any]]:
        """Get all users in a specific department.
        
        Args:
            department: Department name
            count: Maximum results
            
        Returns:
            List of users in the department
        """
        return await self.list_users(
            count=count,
            filter_query=f"department eq '{department}'"
        )
    
    async def get_users_by_job_title(
        self,
        job_title: str,
        count: int = 100,
    ) -> list[dict[str, Any]]:
        """Get all users with a specific job title.
        
        Args:
            job_title: Job title to search for
            count: Maximum results
            
        Returns:
            List of users with the job title
        """
        return await self.list_users(
            count=count,
            filter_query=f"jobTitle eq '{job_title}'"
        )
