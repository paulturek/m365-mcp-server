"""M365 service modules.

Each service module provides high-level operations for a specific
Microsoft 365 service, abstracting the underlying Graph API calls.

Services:
- OutlookService: Mail and Calendar
- OneDriveService: Personal file storage
- SharePointService: Sites, libraries, and lists
- ExcelService: Workbook operations
- OfficeDocsService: Word and PowerPoint
- TeamsService: Teams, channels, and chats
- PowerBIService: Reports and datasets
"""

from .outlook import OutlookService
from .onedrive import OneDriveService
from .sharepoint import SharePointService
from .excel import ExcelService
from .office_docs import OfficeDocsService
from .teams import TeamsService
from .powerbi import PowerBIService

__all__ = [
    "OutlookService",
    "OneDriveService", 
    "SharePointService",
    "ExcelService",
    "OfficeDocsService",
    "TeamsService",
    "PowerBIService",
]
