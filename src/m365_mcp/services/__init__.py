"""M365 services package.

Each service wraps a specific Microsoft 365 API.
"""

from .outlook import OutlookService
from .onedrive import OneDriveService
from .sharepoint import SharePointService
from .excel import ExcelService
from .office_docs import OfficeDocsService
from .teams import TeamsService
from .todo_service import TodoService
from .users_service import UsersService

__all__ = [
    "OutlookService",
    "OneDriveService",
    "SharePointService",
    "ExcelService",
    "OfficeDocsService",
    "TeamsService",
    "TodoService",
    "UsersService",
]
