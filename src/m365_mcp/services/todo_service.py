"""Microsoft To Do service.

Provides operations for:
- Listing task lists
- Creating/updating/deleting task lists
- Listing tasks within lists
- Creating/updating/completing/deleting tasks

Graph API Reference:
    https://learn.microsoft.com/graph/api/resources/todo-overview

Required Scopes:
    - Tasks.ReadWrite (read and write tasks)
    - Tasks.Read (read-only access)
"""

from typing import Any, Optional
from datetime import datetime, date

from ..clients.graph_client import GraphClient


class TodoService:
    """Microsoft To Do task management via Microsoft Graph.
    
    Example:
        >>> service = TodoService(graph_client)
        >>> lists = await service.list_task_lists()
        >>> tasks = await service.list_tasks(list_id="abc123")
        >>> await service.create_task(list_id="abc123", title="Buy groceries")
    """
    
    def __init__(self, client: GraphClient) -> None:
        """Initialize To Do service."""
        self.client = client
    
    # =========================================================================
    # Task Lists
    # =========================================================================
    
    async def list_task_lists(self) -> list[dict[str, Any]]:
        """List all task lists.
        
        Returns:
            List of task list objects with id, displayName, etc.
        """
        result = await self.client.get("/me/todo/lists")
        return result.get("value", [])
    
    async def get_task_list(self, list_id: str) -> dict[str, Any]:
        """Get a specific task list by ID.
        
        Args:
            list_id: The task list ID
            
        Returns:
            Task list object
        """
        return await self.client.get(f"/me/todo/lists/{list_id}")
    
    async def create_task_list(self, display_name: str) -> dict[str, Any]:
        """Create a new task list.
        
        Args:
            display_name: Name for the new list
            
        Returns:
            Created task list object
        """
        return await self.client.post(
            "/me/todo/lists",
            json={"displayName": display_name}
        )
    
    async def update_task_list(
        self,
        list_id: str,
        display_name: str
    ) -> dict[str, Any]:
        """Update a task list name.
        
        Args:
            list_id: The task list ID
            display_name: New name for the list
            
        Returns:
            Updated task list object
        """
        return await self.client.patch(
            f"/me/todo/lists/{list_id}",
            json={"displayName": display_name}
        )
    
    async def delete_task_list(self, list_id: str) -> None:
        """Delete a task list.
        
        Args:
            list_id: The task list ID to delete
        """
        await self.client.delete(f"/me/todo/lists/{list_id}")
    
    # =========================================================================
    # Tasks
    # =========================================================================
    
    async def list_tasks(
        self,
        list_id: str,
        include_completed: bool = True,
        count: int = 100,
    ) -> list[dict[str, Any]]:
        """List tasks in a task list.
        
        Args:
            list_id: The task list ID
            include_completed: Whether to include completed tasks
            count: Maximum number of tasks to return
            
        Returns:
            List of task objects
        """
        params: dict[str, Any] = {"$top": count}
        
        if not include_completed:
            params["$filter"] = "status ne 'completed'"
        
        result = await self.client.get(
            f"/me/todo/lists/{list_id}/tasks",
            params=params
        )
        return result.get("value", [])
    
    async def get_task(self, list_id: str, task_id: str) -> dict[str, Any]:
        """Get a specific task.
        
        Args:
            list_id: The task list ID
            task_id: The task ID
            
        Returns:
            Task object
        """
        return await self.client.get(
            f"/me/todo/lists/{list_id}/tasks/{task_id}"
        )
    
    async def create_task(
        self,
        list_id: str,
        title: str,
        body: Optional[str] = None,
        due_date: Optional[str] = None,
        importance: str = "normal",
        reminder_datetime: Optional[str] = None,
    ) -> dict[str, Any]:
        """Create a new task.
        
        Args:
            list_id: The task list ID to add task to
            title: Task title
            body: Optional task body/notes
            due_date: Optional due date (YYYY-MM-DD format)
            importance: 'low', 'normal', or 'high'
            reminder_datetime: Optional reminder ISO datetime
            
        Returns:
            Created task object
        """
        task_data: dict[str, Any] = {
            "title": title,
            "importance": importance,
        }
        
        if body:
            task_data["body"] = {
                "content": body,
                "contentType": "text"
            }
        
        if due_date:
            # Graph API expects dueDateTime as dateTimeTimeZone object
            task_data["dueDateTime"] = {
                "dateTime": f"{due_date}T00:00:00",
                "timeZone": "UTC"
            }
        
        if reminder_datetime:
            task_data["reminderDateTime"] = {
                "dateTime": reminder_datetime,
                "timeZone": "UTC"
            }
            task_data["isReminderOn"] = True
        
        return await self.client.post(
            f"/me/todo/lists/{list_id}/tasks",
            json=task_data
        )
    
    async def update_task(
        self,
        list_id: str,
        task_id: str,
        title: Optional[str] = None,
        body: Optional[str] = None,
        due_date: Optional[str] = None,
        importance: Optional[str] = None,
        status: Optional[str] = None,
    ) -> dict[str, Any]:
        """Update an existing task.
        
        Args:
            list_id: The task list ID
            task_id: The task ID
            title: New title (optional)
            body: New body/notes (optional)
            due_date: New due date YYYY-MM-DD (optional)
            importance: 'low', 'normal', or 'high' (optional)
            status: 'notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred' (optional)
            
        Returns:
            Updated task object
        """
        update_data: dict[str, Any] = {}
        
        if title is not None:
            update_data["title"] = title
        
        if body is not None:
            update_data["body"] = {
                "content": body,
                "contentType": "text"
            }
        
        if due_date is not None:
            update_data["dueDateTime"] = {
                "dateTime": f"{due_date}T00:00:00",
                "timeZone": "UTC"
            }
        
        if importance is not None:
            update_data["importance"] = importance
        
        if status is not None:
            update_data["status"] = status
            # If marking as completed, set completedDateTime
            if status == "completed":
                update_data["completedDateTime"] = {
                    "dateTime": datetime.utcnow().isoformat(),
                    "timeZone": "UTC"
                }
        
        return await self.client.patch(
            f"/me/todo/lists/{list_id}/tasks/{task_id}",
            json=update_data
        )
    
    async def complete_task(self, list_id: str, task_id: str) -> dict[str, Any]:
        """Mark a task as completed.
        
        Args:
            list_id: The task list ID
            task_id: The task ID
            
        Returns:
            Updated task object
        """
        return await self.update_task(
            list_id=list_id,
            task_id=task_id,
            status="completed"
        )
    
    async def delete_task(self, list_id: str, task_id: str) -> None:
        """Delete a task.
        
        Args:
            list_id: The task list ID
            task_id: The task ID to delete
        """
        await self.client.delete(
            f"/me/todo/lists/{list_id}/tasks/{task_id}"
        )
