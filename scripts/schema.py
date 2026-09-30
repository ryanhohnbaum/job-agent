"""Workbook schema. Column names are a contract shared by every script."""

APPLICATION_HEADERS = [
    "Application ID", "Date Found", "Date Applied", "Company", "Role", "Requisition ID",
    "Location", "Work Model", "Job URL", "Salary Range", "Fit Score", "Resume Filename",
    "Application Folder", "Status", "Confirmation Received", "Latest Correspondence Date",
    "Latest Correspondence Type", "Interview Date", "Rejection Date", "Evidence Complete", "Notes",
]

CORRESPONDENCE_HEADERS = [
    "Application ID", "Date", "Company", "Role", "Direction", "From / To", "Subject",
    "Correspondence Type", "PDF Filename", "Application Folder", "Notes", "Gmail Message ID",
    "Classification Reason", "Classification Confidence",
]

SEARCH_ACTIVITY_HEADERS = [
    "Date", "Search / Source", "Roles Reviewed", "Qualified Roles", "Applications Started",
    "Applications Submitted", "Notes",
]

STATUSES = [
    "Qualified", "On Hold", "Applied", "Interview", "Rejected", "Withdrawn", "Skipped",
    "Closed", "Do Not Reapply", "Offer",
]

ACTIONABLE = {"qualified", "on hold"}
