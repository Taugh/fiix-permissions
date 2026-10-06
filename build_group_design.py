#!/usr/bin/env python3
"""
Build Fiix_v5_Group_Build_Sheet.xlsx: the eight management-requested user groups
translated into Fiix v5 permission settings.

    python build_group_design.py <output.xlsx> <management_model.xlsx>

Inputs
  data/*.txt|tsv           Fiix default permission tables (verified against the Help Center)
  <management_model.xlsx>  IREN_Fiix_Enterprise_Permission_Model_Updated.xlsx (Permission Matrix sheet)

How the design works
  Every permission is assigned an ENTRY group, the lowest group that needs it. Because the
  management model is cumulative, every group above the entry group also gets it.
  Fiix has no inheritance, so each group is built as a full copy; see the Build Deltas sheet.
"""
from __future__ import annotations

import re
import sys
from collections import OrderedDict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).parent
DATA = HERE / "data"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "Fiix_v5_Group_Build_Sheet.xlsx"
MGMT = Path(sys.argv[2]) if len(sys.argv) > 2 else None
CHECK = "✓"
HELP = "https://helpdesk.fiixsoftware.com/hc/en-us/"

# ----------------------------------------------------------------- groups ----
GROUPS = [
    ("G", "General / View-Only"), ("T", "Technician"), ("P", "Maintenance Planner"),
    ("M", "Maintenance Planning Manager"), ("I", "Inventory Team"), ("A", "CMMS Administrator"),
    ("S", "Asset Management"), ("O", "Owner"),
]
KEYS = [k for k, _ in GROUPS]
NAME = dict(GROUPS)
# Management's chain. CMMS Administrator also takes the Inventory rights, because the matrix gives it
# inventory authority (M/TA) even though the Bucket Definitions chain does not list Inventory.
PARENTS = {"G": [], "T": ["G"], "P": ["T"], "M": ["P"], "I": ["T"], "A": ["M", "I"], "S": ["A"], "O": ["S"]}
COPY_FROM = {"G": None, "T": "G", "P": "T", "M": "P", "I": "T", "A": "M", "S": "A", "O": "S"}
DEFAULTS = ["Guests", "Logistics", "Manager", "Purchasing", "Technician", "Admin"]


def ancestors(g):
    out = {g}
    for p in PARENTS[g]:
        out |= ancestors(p)
    return out


ANC = {g: ancestors(g) for g in KEYS}


def holders(spec):
    """spec: entry group(s) like 'P' or 'M+I'. Returns the set of groups that hold the permission."""
    if spec is None:
        return {"O"}
    entries = spec.split("+")
    got = {g for g in KEYS if any(e in ANC[g] for e in entries)}
    got.add("O")  # Owner holds all functional permissions
    return got


TIER_TEXT = {
    "G": "all groups", "T": "Technician and above", "P": "Planner and above", "M": "Planning Manager and above",
    "I": "Inventory Team and above", "A": "CMMS Administrator and above", "S": "Asset Management and above",
    "O": "Owner only",
}


def tier_text(spec):
    if spec is None:
        return "Owner only"
    return " + ".join(TIER_TEXT[e] if e in TIER_TEXT else e for e in spec.split("+"))


# ------------------------------------------------------------------ styles ----
F = lambda **k: Font(name="Arial", size=k.pop("size", 10), **k)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = F(bold=True, color="FFFFFF")
YES_FILL = PatternFill("solid", fgColor="C6EFCE")
YES_FONT = F(bold=True, color="006100")
DEF_FILL = PatternFill("solid", fgColor="EDEDED")
DEF_FONT = F(color="595959")
BAND = PatternFill("solid", fgColor="DDEBF7")
AMBER = PatternFill("solid", fgColor="FFEB9C")
RED = PatternFill("solid", fgColor="FFC7CE")
GREEN = PatternFill("solid", fgColor="C6EFCE")
NONE_FILL = PatternFill("solid", fgColor="F2F2F2")
TITLE = F(size=14, bold=True, color="1F3864")
BOLD = F(bold=True)
LINK = F(color="0563C1", underline="single")
THIN = Side(style="thin", color="D9D9D9")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(wrap_text=True, vertical="top")


def header(ws, row, labels, widths=None, height=42):
    for i, t in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=t)
        c.fill, c.font, c.border, c.alignment = HDR_FILL, HDR_FONT, BOX, CENTER
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = height


def flag(cell, on, default=False):
    cell.border, cell.alignment = BOX, CENTER
    if on:
        cell.value = CHECK
        cell.fill, cell.font = (DEF_FILL, DEF_FONT) if default else (YES_FILL, YES_FONT)
    elif default:
        cell.fill = DEF_FILL


# --------------------------------------------------------------- load data ----
def lines(name):
    return [l for l in (DATA / name).read_text(encoding="utf-8").split("\n") if l.strip()]


def load_page_defaults():
    rows, page = [], None
    for ln in lines("page_perms.txt"):
        if ln.startswith("#"):
            page = ln[1:]
        else:
            perm, bits = ln.rsplit("\t", 1)
            rows.append((page, perm, [b == "1" for b in bits]))
    return rows


def load_menu_defaults():
    return [(ln.split("\t")[0], [b == "1" for b in ln.split("\t")[1]]) for ln in lines("menu_perms.tsv")]


def load_purch_defaults():
    out = []
    for ln in lines("purchasing.tsv"):
        s, p, b = ln.split("\t")
        out.append((s, p, [x == "1" for x in b]))
    return out


def load_api_defaults():
    sec, cur = {}, None
    for ln in lines("api_perms.txt"):
        if ln.startswith("#"):
            cur = ln[1:]
            sec[cur] = OrderedDict()
        else:
            c, b = ln.split("\t")
            sec[cur][c] = b
    return sec


# ------------------------------------------------------------ page design ----
# (view, write, delete, export, import) entry group per page.
# view = Can Access / Print / viewing; write = Save / New / Edit / Update-Close; import = Import button.
PAGES = {
    # ---- Work orders
    "Work Orders": ("G", "P", "M", "P", "A"),
    "Work Order Administration": ("P", "P", None, None, None),
    "Work Order Business": ("G", "P", None, None, "A"),
    "Work Order File": ("G", "T", None, None, "A"),
    "Work Order Parts": ("G", "T", "P", None, "A"),
    "Work Order Task": ("G", "T", None, None, "A"),
    "Work Order User": ("G", "P", None, None, "A"),
    "Assigned Work Orders": ("T", "T", "A", "P", "A"),
    "Misc Cost": ("T", "P", None, None, "A"),
    "Failure Code Actions": ("T", "A", "A", "P", "A"),
    "Failure Code Causes": ("T", "A", "A", "P", "A"),
    "Failure Code Problems": ("T", "A", "A", "P", "A"),
    "Failure Code Similar Work Orders": ("T", "A", "A", "P", "A"),
    "Submit Work Request": ("G", "G", None, None, "A"),
    "Calendar": ("G", None, None, None, None),
    "Calendar Scheduler": ("P", None, None, None, None),
    "Scheduler": ("T", "P", None, None, None),
    "Task": ("G", "P", None, None, "A"),
    "Task Group": ("G", "P", None, None, "A"),
    "Task Groups": ("G", "P", "A", "P", "A"),
    # ---- Preventive / scheduled maintenance
    "Scheduled Maintenance List": ("G", "P", "A", "P", "A"),
    "Scheduled Maintenance Details": ("G", "P", None, None, None),
    "Scheduled Maintenance Business": ("G", "P", None, None, "A"),
    "Scheduled Maintenance File": ("G", "P", None, None, "A"),
    "Scheduled Maintenance Trigger": ("G", "P", None, None, "A"),
    "Scheduled Parts": ("G", "P", None, None, "A"),
    "Scheduled Task": ("G", "P", None, None, "A"),
    "Scheduled Report": ("P", "P", None, None, "A"),
    "Scheduled Report User": ("P", "P", None, None, "A"),
    "Nested PMs": ("G", "P", "A", "P", "A"),
    "Nested PMs Configuration": ("G", "P", None, None, "A"),
    "Nested PMView Modal": ("G", "P", None, None, "A"),
    # ---- Assets
    "Assets": ("G", "A", "A", "P", "A"),
    "Asset": ("G", "A", None, None, "A"),
    "Asset Business": ("G", "A", None, None, "A"),
    "Asset Categories": ("A", "A", "A", "A", "A"),
    "Asset Category": ("A", "A", None, None, "A"),
    "Asset Event": ("G", "T", None, None, "A"),
    "Asset Events History": ("G", "P", None, None, "A"),
    "Asset File": ("G", "P", None, None, "A"),
    "Asset Parts Supplies": ("G", "I", None, None, "A"),
    "Assigned Assets": ("T", "A", "A", "A", "A"),
    "Warranty Certificate": ("G", "A", None, None, "A"),
    "Meter Reading": ("G", "T", None, None, "A"),
    "Meter Reading History": ("G", "P", None, None, "A"),
    # ---- Inventory
    "Current Stock": ("G", "I", "A", "I", "I"),
    "Stock": ("G", "I", None, None, "I"),
    "Stock (page name blank in source)": ("I", "I", None, None, None),
    "Stock Adjustment": ("M+I", "I", "A", "I", "A"),
    "Cycle Count": ("M+I", "I", None, None, None),
    "Cycle Counts": ("M+I", "I", "A", "I", "A"),
    "Select Classes for Cycle Count": ("M+I", "I", None, None, "A"),
    "BOM Groups Manager": ("G", "I", "A", "I", "A"),
    "BOMGroup": ("G", "I", None, None, "A"),
    "Select Parts For Bill of Materials": ("I", "I", "A", "I", "A"),
    "Part Supply Classification": ("G", "I", None, None, "A"),
    "Parts and Supplies Classification": ("M+I", "I", None, None, None),
    "Move": ("I", "I", None, None, None),
    "Move Back": ("I", "I", None, None, None),
    "Business": ("P+I", "I", None, None, "A"),
    "Business Contact": ("P+I", "I", None, None, "A"),
    "Business File": ("P+I", "I", None, None, "A"),
    "Businesses": ("P+I", "I", "A", "I", "A"),
    # ---- Purchasing and requisitions
    "My Purchase Requests": ("T", "T", "P", "P", "A"),
    "Purchase Request": ("T", "T", None, None, "A"),
    "Add To Open Purchase Order": ("I", "I", "A", "I", "A"),
    "Purchase Order Comment": ("I", "I", None, None, "A"),
    "Purchase Order": ("I", "I", None, None, "A"),
    "Purchase Order File": ("I", "I", None, None, "A"),
    "Purchase Order Line Item": ("I", "I", None, None, "A"),
    "Purchase Order Line Items": ("I", "I", "A", "I", "A"),
    "Purchase Order User": ("I", "I", None, None, "A"),
    "Purchase Order, Comment Email Or Print": ("I", "I", None, None, "A"),
    "Purchase Orders": ("I", "I", "A", "I", "A"),
    "Receipt": ("I", "I", None, None, "A"),
    "Receipt Line Item": ("I", "I", None, None, "A"),
    "Receipts": ("I", "I", "A", "I", "A"),
    "Receive Part To Multiple Locations": ("I", "I", "A", "I", "A"),
    "Select From Multiple Suppliers": ("I", "I", None, None, "A"),
    "Select Purchase Order Cost Type": ("I", "I", None, None, "A"),
    "Select Supplier and Confirm Information": ("I", "I", None, None, "A"),
    "Add to Open RFQ": ("A", "A", "A", "A", "A"),
    "Email or Print RFQ": ("A", "A", None, None, "A"),
    "RFQ": ("A", "A", None, None, "A"),
    "RFQ File": ("A", "A", None, None, "A"),
    "RFQ Line Item": ("A", "A", None, None, "A"),
    "RFQs": ("A", "A", "A", "A", "A"),
    # ---- Reporting, dashboards, messages
    "Report": ("P", "P", "M", "P", "A"),
    "Reports": ("P", "P", "M", "P", "A"),
    "Dashboard": ("M", None, None, None, None),
    "Message": ("G", "G", None, None, "A"),
    "Messages": ("G", "G", "G", "P", "A"),
    "Project": ("M", "M", None, None, "A"),
    "Project File": ("M", "M", None, None, "A"),
    "Projects": ("M", "M", "A", "M", "A"),
    # ---- Users
    "Personnel": ("P+I", "A", None, None, "A"),
    "Reports To User": ("A", "A", None, None, "A"),
    "User": ("A", "A", None, None, "A"),
    "User Certification": ("A", "A", None, None, "A"),
    "User File": ("A", "A", None, None, "A"),
}

# Per-permission overrides: (page, permission) -> entry group
OVERRIDE = {
    ("Work Order Administration", "Can See 'Log Work' Info"): "T",
    ("Work Order Administration", "Can Toggle 'Requires Signature'"): "A",
    ("Assigned Work Orders", "Has 'New' Button"): "P",
    ("Receipts", "Has 'Edit' Menu Item"): "A",
    ("Dashboard", "Can See Timeline"): "G",
    ("Scheduler", "Can Access Filter"): "P",
}

# Pages documented in Fiix feature articles but missing from the default tables.
FULL7 = ["Can Access", "Has 'New' Button", "Has 'Edit' Menu Item", "Has 'Print' Button", "Has 'Delete' Button",
         "Has 'Export' Button", "Has 'Import' Button"]
DETAIL4 = ["Can Access", "Has 'Save' Button", "Has 'Print' Button", "Has 'Import' Button"]
EXTRAS = [
    # page, permissions, spec, article
    ("Work Order Assets", FULL7, ("G", "P", "P", "P", "A"), "Work orders permissions (v5)"),
    ("Work Order Filters", FULL7, ("G", "P", "P", "P", "A"), "Work orders permissions (v5)"),
    ("Work Order Status", FULL7, ("A", "A", "A", "A", "A"), "Work orders permissions (v5)"),
    ("Add Labor to Task", DETAIL4, ("T", "T", None, None, "A"), "Work orders permissions (v5)"),
    ("Add Task Group to WO", FULL7, ("P", "P", "P", "P", "A"), "Work orders permissions (v5)"),
    ("Import Work Orders", ["Can Access"], ("A", None, None, None, None), "Work orders permissions (v5)"),
    ("Asset Event Types", FULL7, ("A", "A", "A", "A", "A"), "Asset permissions (v5)"),
    ("Asset Filter", DETAIL4, ("G", "P", None, None, "A"), "Asset permissions (v5)"),
    ("Asset Tag List", DETAIL4, ("P", "P", None, None, "A"), "Asset permissions (v5)"),
    ("Import Asset Event Types", ["Can Access"], ("A", None, None, None, None), "Asset permissions (v5)"),
    ("Import Asset Events", ["Can Access"], ("A", None, None, None, None), "Asset permissions (v5)"),
    ("Import Assets", ["Can Access"], ("A", None, None, None, None), "Asset permissions (v5)"),
    ("CMMS Settings", ["Can Access"], ("A", None, None, None, None), "Asset permissions (v5), prerequisite page"),
    ("Schedule Maintenance Upcoming", DETAIL4, ("G", "P", None, None, "A"), "Scheduled maintenance permissions (v5)"),
    ("Scheduled Maintenance Filter", FULL7, ("G", "P", "P", "P", "A"), "Scheduled maintenance permissions (v5)"),
    ("Scheduled Trigger Change Threshold", DETAIL4, ("G", "M", None, None, "A"), "Scheduled maintenance permissions (v5)"),
    ("Import Scheduled Maintenance", ["Can Access"], ("A", None, None, None, None), "Scheduled maintenance permissions (v5)"),
    ("Import Task Groups", ["Can Access"], ("A", None, None, None, None), "Scheduled maintenance permissions (v5)"),
    ("Users", FULL7, ("A", "A", "S", "A", "A"), "User management permissions (v5)"),
    ("User Groups", FULL7, ("A", "A", "S", "A", "A"), "User management permissions (v5)"),
    ("User Group", DETAIL4, ("A", "A", None, None, "A"), "User management permissions (v5)"),
    ("User Status", FULL7, ("A", "A", "S", "A", "A"), "User management permissions (v5)"),
    ("Business Contact for User", DETAIL4, ("A", "A", None, None, "A"), "User management permissions (v5)"),
    ("User Interface Customization Settings", ["Can Access", "Has 'New' Button", "Has 'Save' Button", "Has 'Print' Button",
                                               "Has 'Import' Button"], ("A", "A", None, None, "A"), "User management permissions (v5)"),
    ("Import Users From CSV File", ["Can Access"], ("A", None, None, None, None), "Compare user management v5 and v6"),
    ("Languages", ["Can Access"], ("A", None, None, None, None), "Compare user management v5 and v6 (sub-permissions not listed)"),
    ("User Increase Tenant Seats - UI handler", ["Can Access"], ("A", None, None, None, None),
     "Compare user management v5 and v6 (called defunct in User management permissions)"),
]

NOTES = {
    "Work Orders": "List view for all. New, Edit and Update/Close from Planner up. Delete from Planning Manager up as the matrix requires.",
    "Work Order Administration": "Manager-level WO access, including all work orders in the mobile app. Withheld from Technician to keep access to assigned work.",
    "Assigned Work Orders": "Technician's working list. New is Planner and above. Edit is given to Technician so assigned work can be updated; Fiix cannot limit it to execution fields.",
    "Work Order Task": "Technician completes checklist items and enters readings here.",
    "Work Order Parts": "Technician records parts consumed. Removing a part line is Planner and above.",
    "Work Order File": "Technician attaches photos and documents.",
    "Failure Code Actions": "Technician can view codes to select them. Editing the lookup lists is CMMS Administrator and above.",
    "Failure Code Causes": "Same as Failure Code Actions.", "Failure Code Problems": "Same as Failure Code Actions.",
    "Failure Code Similar Work Orders": "Same as Failure Code Actions.",
    "Scheduled Trigger Change Threshold": "Changing the next-due threshold changes the effective frequency, so Planning Manager and above.",
    "Receipts": "Edit menu item held back to CMMS Administrator to approximate 'no edit after posting'. Receiving still needs Save on Receipt.",
    "Receipt": "Inventory needs Save to receive. Fiix cannot make posted receipts read-only for one group.",
    "Stock": "Inventory maintains bins, min, max and reorder fields here. Planning Manager views only.",
    "Stock Adjustment": "Inventory performs adjustments. Reason and evidence are a process control.",
    "Cycle Counts": "Planning Manager is view-only, Inventory manages.",
    "Dashboard": "Controls the legacy dashboard page. Timeline is visible to all.",
    "Reports": "Legacy reports reserved for Planner and above. General and Technician use Analytics instead.",
    "Personnel": "Limited directory view for Planner, Planning Manager and Inventory. Technician does not get it.",
    "Users": "Delete reserved for Asset Management and Owner. Prefer deactivation, as the matrix states.",
    "User Groups": "Delete reserved for Asset Management and Owner. Create and edit are CMMS Administrator, which implements approved changes.",
    "Assets": "Create and edit from CMMS Administrator. Planners and Planning Managers request corrections through process.",
    "Asset": "One Save covers description, hierarchy, location and criticality. Field-level control does not exist in Fiix.",
    "Asset Parts Supplies": "Links spares and BOMs to assets. Inventory Team manages these links.",
    "Import Assets": "Bulk asset import is a technical task for CMMS Administrator.",
    "Import Scheduled Maintenance": "Technical upload by CMMS Administrator. Planning Manager authorizes content by process.",
    "My Purchase Requests": "Technician and above can raise and edit their own requisitions. Delete is Planner and above.",
    "Purchase Orders": "Inventory processes requisitions into purchase orders. Approval is in the Purchasing tab.",
    "Stock (page name blank in source)": "Transfer stock button. Source table left the page name blank; it follows Stock.",
    "User Increase Tenant Seats - UI handler": "Fiix calls this defunct in one article and lists it for user status changes in another. Included to be safe.",
}
LOW_CONF = {
    "Assigned Work Orders", "Work Order Administration", "Receipts", "Receipt", "Personnel", "Dashboard", "Misc Cost",
    "Businesses", "Business", "Business Contact", "Business File", "Message", "Messages", "Move", "Move Back",
    "Failure Code Actions", "Failure Code Causes", "Failure Code Problems", "Failure Code Similar Work Orders",
    "Parts and Supplies Classification", "Part Supply Classification", "Languages", "Add Task Group to WO",
    "User Increase Tenant Seats - UI handler", "CMMS Settings", "Project", "Project File", "Projects",
    "Scheduled Report", "Scheduled Report User", "Reports", "Report", "Calendar Scheduler", "Scheduler",
}


def category(perm: str) -> str:
    if "Delete" in perm:
        return "d"
    if "Export" in perm:
        return "x"
    if "Import" in perm:
        return "i"
    if perm == "Can Access" or "Print" in perm or perm.startswith("Can See Timeline") or perm == "Can Access Filter" \
            or "Log Work" in perm:
        return "v"
    return "w"


CAT_IDX = {"v": 0, "w": 1, "d": 2, "x": 3, "i": 4}
CAT_LABEL = {"v": "view", "w": "create/edit", "d": "delete", "x": "export", "i": "import"}


def spec_for(page, perm, spec_tuple):
    if (page, perm) in OVERRIDE:
        return OVERRIDE[(page, perm)]
    return spec_tuple[CAT_IDX[category(perm)]]


# ---------------------------------------------------------------- menus ----
MENU = {
    "Can See 'Settings' Top Menu": "A", "Can See 'Users' Menu Item": "A", "Can See 'User Groups' Menu Item": "A",
    "Can See 'Business' Menu Item": "A", "Can See 'Import' Menu Item": "A", "Can See 'Cmms Settings' Menu Item": "A",
    "Can See 'Account Settings' Menu Item": "O", "Can See 'Notification Templates' Menu Item": "A",
    "Can See 'Connect Management' Menu Item": "A", "Can See 'Data Dictionary' Menu Item": "A",
    "Can See 'Localization' Menu Item": "A", "Can See 'Audit Log' Menu Item": "A", "Can See 'Workflows' Menu Item": "A",
    "Can See 'Security And Infrastructure' Menu Item": "A",
    "Can See 'Help' Top Menu": "G", "Can See 'Log Off' Top Menu": "G", "Can See 'Dashboard' Top Menu": "G",
    "Can See 'Dashboard' Menu Item": "G", "Can See 'Calendar' Menu Item": "G",
    "Can See 'Assigned Assets' Menu Item": "T", "Can See 'Assigned Work Orders' Menu Item": "T",
    "Can See 'Submit Work Request' Menu Item": "G", "Can See 'My Purchase Requests' Menu Item": "T",
    "Can See 'Submit Purchase Request' Menu Item": "T",
    "Can See 'Maintenance' Top Menu": "G", "Can See 'Work Orders' Menu Item": "G",
    "Can See 'Scheduled Maintenance' Menu Item": "G", "Can See 'Task Groups' Menu Item": "G",
    "Can See 'Projects' Menu Item": "M",
    "Can See 'Notifications' Top Menu": "G", "Can See 'Inbox' Menu Item": "G", "Can See 'Sent' Menu Item": "G",
    "Can See 'System Messages' Menu Item": "G", "Can See 'Trash' Menu Item": "G",
    "Can See 'Assets' Top Menu": "G", "Can See 'All Assets' Menu Item": "G", "Can See 'Facilities' Menu Item": "G",
    "Can See 'Equipment' Menu Item": "G", "Can See 'Tools' Menu Item": "G",
    "Can See 'Supplies' Top Menu": "G", "Can See 'Parts And Supplies' Menu Item": "G",
    "Can See 'Current Stock' Menu Item": "G", "Can See 'Batch Stock Adjustment' Menu Item": "I",
    "Can See 'Inventory Cycle Count' Menu Item": "M+I", "Can See 'Bill Of Materials Groups' Menu Item": "G",
    "Can See 'Businesses' Menu Item": "P+I",
    "Can See 'Purchasing' Top Menu": "I", "Can See 'Purchase Planning Board' Menu Item": "I",
    "Can See 'Purchase Orders' Menu Item": "I", "Can See 'Receipts' Menu Item": "I", "Can See 'RFQs' Menu Item": "A",
    "Can See 'Reports' Top Menu": "P", "Can See 'Reports' Menu Item": "P",
}
MENU_EXTRA = [
    ("Can See 'Analytics' Top Menu", "G", "Enable the Analytics feature"),
    ("Can See 'Analytics' Menu item", "G", "Enable the Analytics feature"),
]
MENU_NOTE = {
    "Can See 'Account Settings' Menu Item": "Licence and seat counts. Not in the matrix, so held back for Owner.",
    "Can See 'Audit Log' Menu Item": "Settings audit log. Row 121 'view audit history' for Technician up is read as record history, not this menu.",
    "Can See 'Maintenance' Top Menu": "Needed so every group can reach the work order and PM lists the matrix lets them view.",
    "Can See 'Reports' Top Menu": "Legacy reports are Planner and above.",
    "Can See 'Connect Management' Menu Item": "API application keys and integrations.",
    "Can See 'RFQs' Menu Item": "RFQs are not in the matrix; kept at CMMS Administrator, as in Fiix defaults.",
    "Can See 'Businesses' Menu Item": "Suppliers and manufacturers.",
}

PURCH_TIER = {
    ("Draft", "Can Submit for Approval"): "I", ("Draft", "Can Send to Supplier"): "S",
    ("Draft", "Can Cancel Purchase Order"): "I",
    ("Waiting for Approval", "Can Approve Purchase Order"): "S", ("Waiting for Approval", "Can Reject Purchase Order"): "S",
    ("Waiting for Approval", "Can Cancel Purchase Order"): "I",
    ("Approved", "Can Send to Supplier"): "I", ("Approved", "Can Cancel Purchase Order"): "I",
    ("On Order", "Can Send to Supplier"): "I", ("On Order", "Can Set to Order Not Fulfilled"): "I",
    ("On Order", "Can Cancel Purchase Order"): "I",
    ("Order Fulfilled", "Can Back to Draft"): "I", ("Order Not Fulfilled", "Can Back to Draft"): "I",
}
PURCH_NOTE = {
    ("Draft", "Can Send to Supplier"): "Held back so an unapproved PO cannot be sent. Matrix row 86: PO and financial approval follows a separate Delegation of Authority.",
    ("Waiting for Approval", "Can Approve Purchase Order"): "Matrix row 86 gives approval to Asset Management only. Delegation-of-Authority approvers outside that bucket need a separate decision.",
    ("Waiting for Approval", "Can Reject Purchase Order"): "Same as Approve.",
}

# ------------------------------------------------------------------ API ----
# class: (Add, Find, Change, Remove) entry group
API = {
    "Account": (None, "G", None, None), "Asset": ("A", "G", "A", "A"), "AssetBusiness": ("A", "G", "A", "A"),
    "AssetCategory": ("A", "G", "A", "A"), "AssetConsumingReference": ("I", "G", "I", "I"),
    "AssetEvent": ("T", "G", "T", "A"), "AssetEventType": ("A", "G", "A", "A"), "AssetEventTypeAsset": ("A", "G", "A", "A"),
    "AssetOfflineTracker": ("A", "A", "A", "A"), "AssetUser": ("P", "G", "P", "A"), "BillingTerm": ("A", "A", "A", "A"),
    "Business": ("I", "G", "I", "A"), "BusinessClassification": ("A", "G", "A", "A"), "BusinessGroup": ("A", "A", "A", "A"),
    "BusinessRole": ("A", "A", "A", "A"), "CalendarEvent": (None, "G", None, None), "ChargeDepartment": (None, "G", None, None),
    "Country": (None, "G", None, None), "Currency": (None, "G", None, None),
    "CycleCount": ("I", "M+I", "I", "A"), "File": ("T", "G", "P", "P"), "FileContents": (None, "G", None, "P"),
    "MaintenanceType": (None, "G", None, None), "MeterReading": ("T", "G", "T", "P"), "MeterReadingUnit": (None, "G", None, None),
    "Move": ("I", "I", "I", "A"), "MoveAsset": ("I", "I", "I", "A"), "MoveBack": ("I", "I", "I", "A"),
    "MoveBackAsset": ("I", "I", "I", "A"), "MoveSiteManager": (None, "G", None, None), "MoveStatus": (None, "G", None, None),
    "Priority": (None, "G", None, None), "Project": ("M", "G", "M", "M"),
    "PurchaseOrder": ("I", "I", "I", "A"), "PurchaseOrderLineItem": ("I", "I", "I", "A"), "PurchaseOrderStatus": ("A", "A", "A", "A"),
    "RCAAction": ("A", "G", "A", "A"), "RCACause": ("A", "G", "A", "A"), "RCAGrouping": ("A", "G", "A", "A"),
    "RCAGroupingAction": ("A", "G", "A", "A"), "RCAGroupingCause": ("A", "G", "A", "A"), "RCAProblem": ("A", "G", "A", "A"),
    "ReasonToSetAssetOffline": ("A", "A", "A", "A"), "ReasonToSetAssetOnline": ("A", "A", "A", "A"),
    "Receipt": ("I", "I", "I", "A"), "ReceiptLineItem": ("I", "I", "I", "A"), "ReceiptStatus": (None, "T", None, None),
    "ReportsToResolved": (None, "G", None, None), "ScheduleTrigger": ("P", "G", "P", "P"),
    "ScheduledMaintenance": ("P", "G", "P", "A"), "ScheduledMaintenanceAsset": ("P", "G", "P", "P"),
    "ScheduledMaintenancePart": ("P", "G", "P", "P"), "ScheduledMaintenanceUser": ("P", "G", "P", "P"),
    "ScheduledTask": ("P", "G", "P", "P"), "SiteUser": ("A", "A", "A", "A"), "SiteUserGroup": ("A", "A", "A", "A"),
    "Stock": ("I", "G", "I", "A"), "StockCycleCount": ("I", "M+I", "I", "A"), "StockHistory": ("A", "G", "A", "A"),
    "StockTxType": (None, "G", None, None), "TaskGroup": ("P", "G", "P", "A"), "TaskResultValue": ("A", "G", "A", "A"),
    "User": ("A", "G", "A", "A"), "Warranty": ("A", "G", "A", "A"),
    "WorkOrder": ("P", "G", "T", "M"), "WorkOrderAsset": ("P", "G", "P", "P"), "WorkOrderBusiness": ("P", "G", "P", "P"),
    "WorkOrderPart": ("T", "G", "T", "P"), "WorkOrderStatus": ("A", "G", "A", "A"), "WorkOrderTask": ("P", "G", "T", "P"),
    "WorkOrderUser": ("P", "G", "P", "P"),
}
API_NOTE = {
    "WorkOrder": "Technician gets Change so the mobile app can complete work. That also allows other field edits through the app; Fiix cannot split them. Add is Planner and above, so Technician cannot create formal work orders.",
    "WorkOrderTask": "Technician changes tasks to complete checklists and record readings.",
    "WorkOrderPart": "Technician records parts consumed.",
    "File": "Technician uploads photos and documents.",
    "AssetEvent": "Technician records asset condition events.", "MeterReading": "Technician enters meter readings.",
    "User": "Find only for everyone below CMMS Administrator, as Fiix defaults do.",
    "SiteUserGroup": "User-to-group links. The weekly audit idea in the Gaps sheet reads this class.",
    "SiteUser": "Admin-only in Fiix defaults. Kept at CMMS Administrator.",
}


# ------------------------------------------------------ assemble rows ----
def build_rows():
    page_rows = []   # dict(page, perm, spec, defaults(list or None), source, basis, conf)
    for page, perm, dbits in load_page_defaults():
        sp = PAGES[page]
        spec = spec_for(page, perm, sp)
        page_rows.append(dict(page=page, perm=perm, spec=spec, defaults=dbits, source="Fiix default table",
                              note=NOTES.get(page, ""), conf="Low" if page in LOW_CONF else "High"))
    for page, perms, sp, art in EXTRAS:
        for perm in perms:
            spec = spec_for(page, perm, sp)
            page_rows.append(dict(page=page, perm=perm, spec=spec, defaults=None, source="Feature article: " + art,
                                  note=NOTES.get(page, ""), conf="Low" if page in LOW_CONF or perm != "Can Access" and page in
                                  ("Languages", "CMMS Settings") else "Medium"))
    menu_rows = []
    for perm, dbits in load_menu_defaults():
        menu_rows.append(dict(perm=perm, spec=MENU[perm], defaults=dbits, source="Fiix default table",
                              note=MENU_NOTE.get(perm, "")))
    for perm, spec, art in MENU_EXTRA:
        menu_rows.append(dict(perm=perm, spec=spec, defaults=None, source="Feature article: " + art, note=
                              "Shows Analytics in the menu. Each user also needs an analytics seat (Per-User Settings)."))
    purch_rows = []
    for status, perm, dbits in load_purch_defaults():
        purch_rows.append(dict(status=status, perm=perm, spec=PURCH_TIER[(status, perm)], defaults=dbits,
                               note=PURCH_NOTE.get((status, perm), "")))
    api_def = load_api_defaults()
    assert set(API) == set(api_def["Administrators"]), f"API classes differ: {set(API) ^ set(api_def['Administrators'])}"
    api_rows = [dict(cls=cls, spec=API[cls], note=API_NOTE.get(cls, "")) for cls in api_def["Administrators"]]
    # no permission may silently fall to Owner-only
    for r in page_rows:
        if r["spec"] is None:
            raise AssertionError(f"page permission has no entry group: {r['page']} > {r['perm']}")
    return page_rows, menu_rows, purch_rows, api_rows


def api_letters(spec, g):
    out = ""
    for ch, e in zip("AFCR", spec):
        if e is not None and g in holders(e):
            out += ch
    return out or "-"


# --------------------------------------------------------------- sheets ----
def ws_new(wb, title, tab=None):
    ws = wb.create_sheet(title)
    if tab:
        ws.sheet_properties.tabColor = tab
    return ws


def group_header(ws, row, first_label_cols, default_cols=True):
    labels = list(first_label_cols) + [NAME[k] for k in KEYS]
    if default_cols:
        labels += ["Fiix default: " + d for d in DEFAULTS]
    return labels


def sheet_menu(wb, rows):
    ws = ws_new(wb, "Menu Settings")
    labels = ["Menu permission"] + [NAME[k] for k in KEYS] + ["Fiix default: " + d for d in DEFAULTS] + ["Granted from", "Source", "Note"]
    header(ws, 1, labels, [46] + [13] * 8 + [10] * 6 + [26, 28, 60], 56)
    for i, r in enumerate(rows, 2):
        ws.cell(row=i, column=1, value=r["perm"]).border = BOX
        h = holders(r["spec"])
        for j, k in enumerate(KEYS):
            flag(ws.cell(row=i, column=2 + j), k in h)
        for j in range(6):
            flag(ws.cell(row=i, column=10 + j), bool(r["defaults"] and r["defaults"][j]), True)
        for c, v in ((16, tier_text(r["spec"])), (17, r["source"]), (18, r["note"] or None)):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border, cell.alignment = BOX, WRAP
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:R{len(rows) + 1}"
    return len(rows) + 1


def sheet_page(wb, rows):
    ws = ws_new(wb, "Page Settings")
    labels = ["Page", "Permission"] + [NAME[k] for k in KEYS] + ["Fiix default: " + d for d in DEFAULTS] + \
             ["Granted from", "Source", "Confidence", "Note"]
    header(ws, 1, labels, [34, 34] + [13] * 8 + [10] * 6 + [26, 34, 11, 70], 56)
    n_default = 0
    for i, r in enumerate(rows, 2):
        ws.cell(row=i, column=1, value=r["page"]).border = BOX
        ws.cell(row=i, column=2, value=r["perm"]).border = BOX
        h = holders(r["spec"])
        for j, k in enumerate(KEYS):
            flag(ws.cell(row=i, column=3 + j), k in h)
        for j in range(6):
            flag(ws.cell(row=i, column=11 + j), bool(r["defaults"] and r["defaults"][j]), True)
        vals = (tier_text(r["spec"]), r["source"], r["conf"], r["note"] or None)
        for c, v in zip((17, 18, 19, 20), vals):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border, cell.alignment = BOX, WRAP
        if r["defaults"] is not None:
            n_default += 1
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:T{len(rows) + 1}"
    return len(rows) + 1, n_default + 1


def sheet_purch(wb, rows):
    ws = ws_new(wb, "Purchasing Settings")
    labels = ["PO status", "Permission"] + [NAME[k] for k in KEYS] + ["Fiix default: " + d for d in DEFAULTS] + ["Granted from", "Note"]
    header(ws, 1, labels, [22, 32] + [13] * 8 + [10] * 6 + [26, 70], 56)
    for i, r in enumerate(rows, 2):
        ws.cell(row=i, column=1, value=r["status"]).border = BOX
        ws.cell(row=i, column=2, value=r["perm"]).border = BOX
        h = holders(r["spec"])
        for j, k in enumerate(KEYS):
            flag(ws.cell(row=i, column=3 + j), k in h)
        for j in range(6):
            flag(ws.cell(row=i, column=11 + j), bool(r["defaults"][j]), True)
        for c, v in ((17, tier_text(r["spec"])), (18, r["note"] or None)):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border, cell.alignment = BOX, WRAP
    ws.freeze_panes = "C2"
    return len(rows) + 1


def sheet_api(wb, rows):
    ws = ws_new(wb, "API Settings")
    labels = ["Fiix class"] + [NAME[k] for k in KEYS] + ["Note"]
    header(ws, 1, labels, [28] + [14] * 8 + [80], 56)
    for i, r in enumerate(rows, 2):
        ws.cell(row=i, column=1, value=r["cls"]).border = BOX
        for j, k in enumerate(KEYS):
            t = api_letters(r["spec"], k)
            c = ws.cell(row=i, column=2 + j, value=t)
            c.border, c.alignment = BOX, CENTER
            if t == "AFCR":
                c.fill, c.font = GREEN, YES_FONT
            elif t == "-":
                c.fill, c.font = NONE_FILL, F(color="A6A6A6")
            else:
                c.fill = AMBER
        n = ws.cell(row=i, column=10, value=r["note"] or None)
        n.border, n.alignment = BOX, WRAP
    last = len(rows) + 1
    ws.cell(row=last + 2, column=1, value="Key").font = BOLD
    ws.cell(row=last + 3, column=1, value="A = Add   F = Find   C = Change   R = Remove   - = none. These are the checkboxes on the API Permissions tab.")
    ws.cell(row=last + 4, column=1, value="Green = full access, amber = partial, grey = none. Fiix applies API permissions to the mobile app and to connected apps.")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:J{last}"
    return last


def sheet_user_settings(wb):
    ws = ws_new(wb, "Per-User Settings")
    header(ws, 1, ["Setting", "Where in Fiix"] + [NAME[k] for k in KEYS] + ["Why"], [30, 40] + [16] * 8 + [70], 48)
    rows = [
        ("Administrator checkbox", "Settings > Users > user > Access Control tab (per Fiix's assign-an-administrator article)",
         ["No", "No", "No", "No", "No", "No", "No", "Yes"],
         "Fiix says only an Administrator can make another user an Administrator, and Administrators bypass group settings. So Owner users carry this flag and no one else does. Fiix recommends two."),
        ("Analytics seat", "Settings > Users > user > Analytics tab",
         ["Viewer / User", "Viewer / User", "Creator", "Creator", "Creator", "Creator", "Creator", "Creator"],
         "Viewer/User can view, filter, download and share but not build. Creator can build looks and dashboards. Count Creators before rollout, because seats are limited by plan."),
        ("Analytics menu in group", "Group > Menu Permissions", ["Yes"] * 8, "Both Analytics menu items are on in every group. Without a seat the user sees nothing."),
        ("Legacy dashboard assignment", "Settings > CMMS Settings > Dashboard Permissions",
         ["Basic", "Basic", "Basic", "Manager", "Basic", "Administrator", "Administrator", "Administrator"],
         "Fiix only lets groups on the Manager or Administrator dashboards create and delete shared dashboards. 'Basic' means any non-Manager standard dashboard. The exact dashboard names are not in the public docs; pick them in the tenant."),
        ("Groups per user", "Settings > Users > user > Groups", ["One"] * 8,
         "Management wants one bucket per user. Fiix allows several, so this is an audit control (see Gaps)."),
        ("Site assignment", "Settings > Users > user", ["Per site"] * 8,
         "User groups are site-specific in a multi-site tenant. Repeat the group assignment at each site."),
        ("Mobile app settings", "Settings > CMMS Settings > Mobile app settings", ["Not specified"] * 8,
         "The management model does not cover mobile feature switches. API Settings controls what the app can do."),
    ]
    for i, (a, b, vals, why) in enumerate(rows, 2):
        ws.cell(row=i, column=1, value=a).font = BOLD
        ws.cell(row=i, column=2, value=b)
        for j, v in enumerate(vals):
            c = ws.cell(row=i, column=3 + j, value=v)
            c.alignment = CENTER
            if v in ("Creator", "Yes", "Manager", "Administrator"):
                c.fill = YES_FILL
        ws.cell(row=i, column=11, value=why)
        for c in range(1, 12):
            ws.cell(row=i, column=c).border = BOX
            if c in (1, 2, 11):
                ws.cell(row=i, column=c).alignment = WRAP
    ws.freeze_panes = "C2"


def build_deltas(page_rows, menu_rows, purch_rows, api_rows):
    out = []
    for g in KEYS:
        src = COPY_FROM[g]
        for tab, rows, label in (("Menu", menu_rows, lambda r: r["perm"]),
                                 ("Page", page_rows, lambda r: f"{r['page']} > {r['perm']}"),
                                 ("Purchasing", purch_rows, lambda r: f"{r['status']} > {r['perm']}")):
            for r in rows:
                has = g in holders(r["spec"])
                had = src is not None and src in holders(r["spec"])
                if src is not None and had and not has:
                    raise AssertionError(f"not cumulative: {g} lost {label(r)}")
                if has and not had:
                    out.append((g, src, tab, label(r), "Turn on"))
        for r in api_rows:
            now = api_letters(r["spec"], g)
            before = api_letters(r["spec"], src) if src else "-"
            add = "".join(ch for ch in now if ch != "-" and ch not in before)
            if add:
                out.append((g, src, "API", r["cls"], "Tick " + " ".join({"A": "Add", "F": "Find", "C": "Change", "R": "Remove"}[c] for c in add)))
    return out


def sheet_deltas(wb, deltas):
    ws = ws_new(wb, "Build Deltas")
    header(ws, 1, ["Build step", "Group", "Start by copying", "Tab in Fiix", "Permission", "Action"], [11, 30, 30, 13, 70, 26], 34)
    order = {k: i for i, k in enumerate(KEYS, 1)}
    for i, (g, src, tab, lab, act) in enumerate(deltas, 2):
        vals = [order[g], NAME[g], NAME[src] if src else "New blank group", tab, lab, act]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=i, column=j, value=v)
            c.border = BOX
            if j in (1, 4):
                c.alignment = CENTER
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{len(deltas) + 1}"
    return len(deltas) + 1


def sheet_comparison(wb, page_def_last, menu_last, purch_last):
    """Compare each requested group against each Fiix default group, with formulas."""
    ws = ws_new(wb, "Comparison to Defaults")
    ws["A1"] = "How close each requested group is to Fiix's own default groups"
    ws["A1"].font = TITLE
    ws["A2"] = ("Counts only permissions listed in Fiix's default tables (page, menu and purchasing). Pages that Fiix defaults do not list are excluded, "
                "so these numbers understate what the new groups add. Management's groups are stricter than Fiix's defaults in places, for example Fiix's "
                "Technician can create and delete on many pages, so a new group can be 'closest' to a small default without being like it. "
                "All values are formulas over the settings sheets.")
    ws["A2"].font = F(italic=True)
    ws["A2"].alignment = WRAP
    ws.merge_cells("A2:I2")
    ws.row_dimensions[2].height = 44
    ws.column_dimensions["A"].width = 34
    for c in range(2, 10):
        ws.column_dimensions[get_column_letter(c)].width = 15
    # settings-sheet column letters: requested groups start at C (page/purch) or B (menu)
    def req_col(sheet, i):
        return get_column_letter((2 if sheet == "Menu Settings" else 3) + i)

    def def_col(sheet, j):
        return get_column_letter((10 if sheet == "Menu Settings" else 11) + j)

    specs = [("Page Settings", page_def_last), ("Menu Settings", menu_last), ("Purchasing Settings", purch_last)]
    blocks = {}
    row = 4
    ws.cell(row=row, column=1, value="Shared permissions: requested group has it AND the Fiix default group has it").font = BOLD
    row += 1
    for sheet, last in specs:
        ws.cell(row=row, column=1, value=sheet).font = F(bold=True, color="1F3864")
        row += 1
        header(ws, row, ["Requested group"] + DEFAULTS, None, 30)
        top = row
        for i, k in enumerate(KEYS):
            r = row + 1 + i
            ws.cell(row=r, column=1, value=NAME[k]).border = BOX
            for j in range(len(DEFAULTS)):
                rc, dc = req_col(sheet, i), def_col(sheet, j)
                f = (f"=SUMPRODUCT(('{sheet}'!${rc}$2:${rc}${last}=\"{CHECK}\")*('{sheet}'!{dc}$2:{dc}${last}=\"{CHECK}\"))")
                c = ws.cell(row=r, column=2 + j, value=f)
                c.border, c.alignment = BOX, CENTER
        blocks[sheet] = top
        row += len(KEYS) + 2
    # totals per requested group and default group
    ws.cell(row=row, column=1, value="Size of each set (default-table rows only)").font = BOLD
    row += 1
    header(ws, row, ["Group"] + ["Page", "Menu", "Purchasing", "Total"], None, 30)
    size_top = row
    for i, k in enumerate(KEYS):
        r = row + 1 + i
        ws.cell(row=r, column=1, value="Requested: " + NAME[k]).border = BOX
        for c, (sheet, last) in enumerate(specs):
            rc = req_col(sheet, i)
            ws.cell(row=r, column=2 + c, value=f"=COUNTIF('{sheet}'!${rc}$2:${rc}${last},\"{CHECK}\")").border = BOX
        ws.cell(row=r, column=5, value=f"=SUM(B{r}:D{r})").border = BOX
    base = row + 1 + len(KEYS)
    for j, d in enumerate(DEFAULTS):
        r = base + j
        ws.cell(row=r, column=1, value="Fiix default: " + d).border = BOX
        for c, (sheet, last) in enumerate(specs):
            dc = def_col(sheet, j)
            ws.cell(row=r, column=2 + c, value=f"=COUNTIF('{sheet}'!{dc}$2:{dc}${last},\"{CHECK}\")").border = BOX
        ws.cell(row=r, column=5, value=f"=SUM(B{r}:D{r})").border = BOX
    row = base + len(DEFAULTS) + 1

    def cell_ref(sheet, i, j):
        return f"{get_column_letter(2 + j)}{blocks[sheet] + 1 + i}"

    sim_top = row
    ws.cell(row=row, column=1, value=("Distance: permissions to add or remove to turn a Fiix default group into the requested group "
                                      "(smaller is closer), all three tabs combined")).font = BOLD
    row += 1
    header(ws, row, ["Requested group"] + DEFAULTS + ["Closest Fiix default"], None, 30)
    ws.column_dimensions["H"].width = 22
    sim_hdr = row
    for i, k in enumerate(KEYS):
        r = row + 1 + i
        ws.cell(row=r, column=1, value=NAME[k]).border = BOX
        for j in range(len(DEFAULTS)):
            shared = "+".join(cell_ref(s, i, j) for s, _ in specs)
            req_total = f"$E${size_top + 1 + i}"
            def_total = f"$E${base + j}"
            c = ws.cell(row=r, column=2 + j, value=f"={req_total}+{def_total}-2*({shared})")
            c.number_format, c.border, c.alignment = "0", BOX, CENTER
        rng = f"B{r}:G{r}"
        c = ws.cell(row=r, column=8, value=f"=INDEX($B${sim_hdr}:$G${sim_hdr},MATCH(MIN({rng}),{rng},0))")
        c.border, c.font = BOX, BOLD
    row = row + len(KEYS) + 2
    added_top = row
    ws.cell(row=row, column=1, value="Added vs closest default: permissions the requested group has that the closest default does not").font = BOLD
    row += 1
    header(ws, row, ["Requested group", "Closest default", "Added", "Removed"], None, 30)
    for i, k in enumerate(KEYS):
        r = row + 1 + i
        sim_row = sim_hdr + 1 + i
        ws.cell(row=r, column=1, value=NAME[k]).border = BOX
        ws.cell(row=r, column=2, value=f"=H{sim_row}").border = BOX
        j_expr = f"MATCH(H{sim_row},$B${sim_hdr}:$G${sim_hdr},0)"
        shared = "+".join(f"INDEX({get_column_letter(2)}{blocks[s] + 1 + i}:{get_column_letter(7)}{blocks[s] + 1 + i},{j_expr})"
                          for s, _ in specs)
        ws.cell(row=r, column=3, value=f"=$E${size_top + 1 + i}-({shared})").border = BOX
        ws.cell(row=r, column=4, value=f"=INDEX($E${base}:$E${base + 5},{j_expr})-({shared})").border = BOX
    ws.freeze_panes = "B4"
    return sim_hdr  # row index of the similarity header


# ------------------------------------------------------- requirement map ----
FIT_ENF, FIT_APP, FIT_PRO = "Enforced", "Approximate", "Process only"
# matrix row number -> (fit, how Fiix implements it / what is lost)
MAP = {
    5: (FIT_ENF, "Work Orders list: Can Access and Print for all groups, with the Maintenance menu items."),
    6: (FIT_ENF, "Submit Work Request page and menu item for all groups."),
    7: (FIT_ENF, "Work Orders: Has New Button, from Planner up."),
    8: (FIT_APP, "Work Orders Edit plus Work Order Administration Save, from Planner up. Fiix cannot limit editing to planning fields."),
    9: (FIT_APP, "Technician and Inventory have no Edit on Work Orders. Their execution comments go in task and note fields, which Fiix cannot separate from scope fields."),
    10: (FIT_ENF, "Work Order User: Save, from Planner up."),
    11: (FIT_ENF, "Scheduler: Can Edit and Calendar Scheduler, from Planner up."),
    12: (FIT_APP, "Part of Work Orders Edit. Planning Manager approval of priority changes is a process control."),
    13: (FIT_APP, "Add Labor to Task from Technician up. 'Assigned work only' is not a permission; the Technician's list is Assigned Work Orders."),
    14: (FIT_APP, "Work Order Task Save from Technician up. Assigned-only is not enforceable by permission."),
    15: (FIT_ENF, "Work Order File: Save, from Technician up."),
    16: (FIT_ENF, "Work Order Task Save and Meter Reading Save, from Technician up."),
    17: (FIT_APP, "Technician can view and select failure codes; Fiix has no 'limited' versus 'manage' split for entering failure data."),
    18: (FIT_ENF, "Work Order Parts: Save from Technician up. Removing lines needs Planner."),
    19: (FIT_APP, "Technician completes tasks and uses the mobile app (WorkOrder Change). Web 'Update/Close Work Order' is kept for Planner and above."),
    20: (FIT_ENF, "Work Orders: Update/Close menu item, from Planner up."),
    21: (FIT_PRO, "Fiix has one close permission. Escalation and Planning Manager approval for critical work is procedural."),
    22: (FIT_PRO, "Same Update/Close permission. 'Controlled' reopening is procedural."),
    23: (FIT_ENF, "Work Orders: Delete from Planning Manager up, as the matrix says. Governance sheet prefers exception-only, which Fiix cannot enforce."),
    24: (FIT_PRO, "Edit permission cannot tell completed from open records. Controlled corrections are procedural."),
    26: (FIT_ENF, "Scheduled Maintenance pages: Can Access for all groups."),
    27: (FIT_APP, "Scheduled Maintenance List New and Details Save from Planner up. Fiix cannot hold a PM in 'draft'."),
    28: (FIT_APP, "Same pages. Planning Manager's primary ownership is organisational."),
    29: (FIT_APP, "Import Scheduled Maintenance and Import buttons are CMMS Administrator and above. Planner and Planning Manager prepare and authorize outside Fiix, which matches the governance sheet."),
    30: (FIT_APP, "Same Save permission as active PMs."),
    31: (FIT_PRO, "Fiix cannot tell draft from active PMs. Planning Manager approval of changes is procedural."),
    32: (FIT_ENF, "Task Groups, Task Group, Task and Scheduled Task pages: Save from Planner up."),
    33: (FIT_PRO, "No approval permission exists for task standards."),
    34: (FIT_PRO, "A proposal is not a Fiix action. Planners already hold the edit permission."),
    35: (FIT_PRO, "Approval is procedural."),
    36: (FIT_PRO, "Approval is procedural."),
    37: (FIT_PRO, "Activation is part of the Scheduled Maintenance Save."),
    38: (FIT_PRO, "Deactivation is part of the Scheduled Maintenance Save."),
    39: (FIT_APP, "Scheduled Maintenance Trigger Save from Planner up. Planner 'propose only' is procedural."),
    40: (FIT_APP, "Scheduled Trigger Change Threshold Save held at Planning Manager and above."),
    41: (FIT_ENF, "Scheduled Maintenance List, Nested PMs and Task Groups Delete: CMMS Administrator and above."),
    43: (FIT_ENF, "Assets and Asset pages: Can Access for all groups."),
    44: (FIT_ENF, "Asset Events History and Meter Reading History view for all groups."),
    45: (FIT_ENF, "Assets: Has New Button, CMMS Administrator and above. Prior approval is procedural."),
    46: (FIT_APP, "Asset Save is CMMS Administrator and above. Planners request corrections by process."),
    47: (FIT_APP, "Same Asset Save."), 48: (FIT_APP, "Same Asset Save."),
    49: (FIT_APP, "Asset Save plus Asset Categories pages, CMMS Administrator and above."),
    50: (FIT_APP, "Asset Categories pages and Asset Save, CMMS Administrator and above."),
    51: (FIT_APP, "Same as hierarchy."), 52: (FIT_APP, "Asset Save and Asset Category Save."),
    53: (FIT_PRO, "Criticality is a field on the asset. Fiix has no field-level permission, so Asset Save covers it. Asset Management approval is procedural."),
    54: (FIT_ENF, "Asset Event Save, from Technician up."),
    55: (FIT_ENF, "Meter Reading Save, from Technician up."),
    56: (FIT_APP, "Setting an asset offline or retiring it is part of Asset Save."),
    57: (FIT_ENF, "Assets Delete, CMMS Administrator and above. 'Exceptional only' is procedural."),
    59: (FIT_ENF, "Current Stock and Stock: Can Access for all groups."),
    60: (FIT_ENF, "Same pages."),
    61: (FIT_ENF, "Asset Parts Supplies and BOM Groups: view for all groups."),
    62: (FIT_ENF, "Current Stock New and Stock Save, Inventory Team and above."),
    63: (FIT_APP, "Receipts New and Receipt Save, Inventory Team and above. Fiix receipts are tied to the purchase order flow; receiving without a PO may need Stock Adjustment."),
    64: (FIT_APP, "Same as receiving."),
    65: (FIT_APP, "Receipts Edit menu item is held at CMMS Administrator. Inventory still needs Receipt Save to receive, so Fiix cannot make a posted receipt read-only for them."),
    66: (FIT_ENF, "Work Order Parts Save from Technician up, plus Stock pages for Inventory."),
    67: (FIT_APP, "Technician edits the part line. Removing a line needs Planner. Return-to-stock posting for Inventory is via Stock Adjustment."),
    68: (FIT_ENF, "Stock Adjustment New and Edit, Inventory Team and above."),
    69: (FIT_ENF, "Cycle Count pages: manage for Inventory, view for Planning Manager."),
    70: (FIT_ENF, "Stock 'Transfer stock' button and Move pages, Inventory Team and above."),
    71: (FIT_ENF, "Stock Save, Inventory Team and above."),
    72: (FIT_ENF, "Stock Save for Inventory; Planning Manager views."),
    73: (FIT_ENF, "Same."), 74: (FIT_ENF, "Same."),
    75: (FIT_ENF, "Work Order Parts Save from Technician up."),
    76: (FIT_ENF, "Asset Parts Supplies Save and BOM Groups, Inventory Team and above."),
    77: (FIT_PRO, "No permission for critical-spare designation. Asset Management governs it by process."),
    78: (FIT_ENF, "Current Stock Delete, CMMS Administrator and above. Inventory requests."),
    80: (FIT_ENF, "Purchase Request and My Purchase Requests, from Technician up."),
    81: (FIT_APP, "My Purchase Requests shows the user's own. A broader view for Planner and Planning Manager cannot be separated."),
    82: (FIT_ENF, "Purchasing menu, Purchase Planning Board, Purchase Orders and Receipts pages, Inventory Team and above."),
    83: (FIT_APP, "My Purchase Requests Edit. 'Own' is how Fiix scopes the page."),
    84: (FIT_PRO, "Fiix cannot tell approved requisitions from drafts."),
    85: (FIT_PRO, "No approval permission for material requirements."),
    86: (FIT_APP, "Purchasing tab: Approve and Reject at Asset Management and above. Delegation-of-Authority thresholds are not modelled in Fiix."),
    88: (FIT_ENF, "Analytics menu items in every group plus a Viewer seat."),
    89: (FIT_ENF, "Viewer seat allows filtering."), 90: (FIT_ENF, "Viewer seat."),
    91: (FIT_ENF, "Creator seat for Planner, Planning Manager, Inventory and above."),
    92: (FIT_APP, "Creator seat, plus Dashboard Permissions: Manager dashboard for Planning Manager. 'Limited' scopes for Planner and Inventory are not separable."),
    93: (FIT_APP, "Dashboard Permissions assignment (Manager or Administrator) controls shared dashboard editing."),
    94: (FIT_APP, "Same as editing."),
    95: (FIT_PRO, "No per-group control over KPI formulas."),
    96: (FIT_ENF, "Viewer seats can download."),
    97: (FIT_PRO, "Reporting logic changes are procedural."),
    99: (FIT_APP, "Personnel page for Planner, Planning Manager and Inventory; Users page for CMMS Administrator and above."),
    100: (FIT_ENF, "Users: Has New Button, CMMS Administrator and above."),
    101: (FIT_ENF, "User Status page, CMMS Administrator and above."),
    102: (FIT_ENF, "User Save, CMMS Administrator and above."),
    103: (FIT_ENF, "User Save (Groups) for CMMS Administrator and above. 'One bucket only' is not enforced by Fiix."),
    104: (FIT_ENF, "Same."),
    105: (FIT_ENF, "User Groups: Has New Button, CMMS Administrator and above."),
    106: (FIT_ENF, "User Group Save and its permission tabs, CMMS Administrator and above."),
    107: (FIT_APP, "Only a Fiix Administrator can grant Administrator. That makes it Owner only, which differs from the matrix (Asset Management BG)."),
    108: (FIT_APP, "Owner is the Administrator flag, so Owner only."),
    110: (FIT_ENF, "Settings menus plus CMMS Settings page, CMMS Administrator and above."),
    111: (FIT_ENF, "CMMS Settings and lookup pages, CMMS Administrator and above. Planner and Planning Manager 'recommend' by process."),
    112: (FIT_ENF, "Work Order Status page, CMMS Administrator and above."),
    113: (FIT_ENF, "CMMS Settings, CMMS Administrator and above."),
    114: (FIT_ENF, "Asset Categories pages, CMMS Administrator and above."),
    115: (FIT_ENF, "Failure Code pages New and Edit, CMMS Administrator and above."),
    116: (FIT_ENF, "Import Assets page and Assets Import button, CMMS Administrator and above."),
    117: (FIT_ENF, "Import Scheduled Maintenance, CMMS Administrator and above."),
    118: (FIT_APP, "Import buttons on inventory pages for Inventory Team and above, as the matrix says. The governance sheet says CMMS Administrator implements; the two sheets conflict."),
    119: (FIT_ENF, "Connect Management menu, CMMS Administrator and above."),
    120: (FIT_APP, "Delete buttons sit at CMMS Administrator and above on the controlled pages. 'Exception only' is procedural."),
    121: (FIT_APP, "Record Log tabs follow normal page access. Settings > Audit Log is CMMS Administrator and above."),
    122: (FIT_APP, "Audit Log menu plus read access to User Groups. A weekly export of SiteUserGroup through the API is a stronger check."),
}


def load_matrix():
    if MGMT is None:
        sys.exit("Pass the management workbook path as the second argument.")
    wb = load_workbook(MGMT, data_only=True)
    ws = wb["Permission Matrix"]
    rows = []
    for r in range(4, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, 10)]
        if vals[0] is None:
            continue
        rows.append((r, vals))
    return rows


def sheet_reqmap(wb, matrix):
    ws = ws_new(wb, "Requirement Map")
    ws["A1"] = "Management matrix, row by row, against what Fiix v5 can enforce"
    ws["A1"].font = TITLE
    labels = ["Matrix row", "Fiix function / permission"] + [NAME[k] for k in KEYS] + ["Fit in Fiix v5", "How it is implemented, or what is lost"]
    # summary counts first, then table
    ws["A3"] = "Fit"
    ws["A3"].font = BOLD
    ws["B3"] = "Rows"
    ws["B3"].font = BOLD
    first = 9
    last = first
    body = []
    section = None
    for r, vals in matrix:
        if all(v is None for v in vals[1:]):
            body.append(("section", r, vals[0]))
        else:
            body.append(("row", r, vals))
    data_rows = [b for b in body if b[0] == "row"]
    last = first + len(body) - 1
    for i, fit in enumerate((FIT_ENF, FIT_APP, FIT_PRO)):
        ws.cell(row=4 + i, column=1, value=fit).fill = (GREEN, AMBER, RED)[i]
        ws.cell(row=4 + i, column=2, value=f'=COUNTIF($K${first}:$K${last},A{4 + i})')
    ws.cell(row=7, column=1, value="Total").font = BOLD
    ws.cell(row=7, column=2, value="=SUM(B4:B6)").font = BOLD
    ws.cell(row=8, column=1)
    header(ws, first - 1, labels, [10, 36] + [17] * 8 + [15, 90], 48)
    i = first
    missing = []
    for kind, r, vals in body:
        if kind == "section":
            c = ws.cell(row=i, column=1, value=vals)
            c.font = BOLD
            for col in range(1, 12):
                ws.cell(row=i, column=col).fill = BAND
        else:
            ws.cell(row=i, column=1, value=r).alignment = CENTER
            ws.cell(row=i, column=2, value=vals[0]).alignment = WRAP
            for j in range(8):
                c = ws.cell(row=i, column=3 + j, value=vals[1 + j])
                c.alignment = CENTER
            fit, how = MAP.get(r, (None, None))
            if fit is None:
                missing.append(r)
            c = ws.cell(row=i, column=11, value=fit)
            c.alignment = CENTER
            c.fill = {FIT_ENF: GREEN, FIT_APP: AMBER, FIT_PRO: RED}.get(fit, NONE_FILL)
            ws.cell(row=i, column=12, value=how).alignment = WRAP
            for col in range(1, 13):
                ws.cell(row=i, column=col).border = BOX
        i += 1
    # fix header width: the table has 12 columns (A..L)
    ws.cell(row=first - 1, column=12).value = "How it is implemented, or what is lost"
    ws.column_dimensions["K"].width = 15
    ws.column_dimensions["L"].width = 90
    ws.freeze_panes = ws.cell(row=first, column=3)
    ws.auto_filter.ref = f"A{first - 1}:L{last}"
    if missing:
        raise AssertionError(f"matrix rows without a mapping: {missing}")
    extra = [k for k in MAP if k not in {r for _, r, _ in [b for b in body if b[0] == 'row']}]
    if extra:
        raise AssertionError(f"mappings for unknown rows: {extra}")
    return len(data_rows)


# ------------------------------------------------------------- gaps text ----
GAPS = [
    ("Decision", "High", "Who grants Administrator and Owner rights",
     "The matrix shows Asset Management with break-glass on 'Grant Administrator rights', and the governance sheet says the CMMS Administrator implements it. "
     "Fiix says only an existing Administrator can make another user an Administrator.",
     "Only Owner users hold the Administrator flag, so only they can grant it. CMMS Administrator and Asset Management recommend or prepare. Confirm this is acceptable."),
    ("Decision", "High", "CMMS Administrator also needs inventory rights",
     "Bucket Definitions say CMMS Administrator inherits Planning Manager only. The matrix gives it inventory authority (M/TA) on receiving, adjustments, transfers and imports.",
     "CMMS Administrator is built from Planning Manager plus all Inventory Team permissions. Confirm the inheritance text should be corrected."),
    ("Decision", "High", "Bulk inventory import: Inventory Team or CMMS Administrator",
     "Matrix row 118 gives Inventory Team 'M'. The governance sheet says Inventory prepares and CMMS Administrator implements.",
     "Built to the matrix: the Import buttons on inventory pages are on for Inventory Team. Move them to CMMS Administrator if governance should win."),
    ("Decision", "High", "Delete rights",
     "Planning Manager has Delete Work Order 'Yes'. CMMS Administrator and Asset Management have 'Exceptional only'. Governance says delete only by exception.",
     "Delete buttons are on exactly where the matrix gives technical ability. Fiix cannot limit a Delete to exceptions, so review the audit log. Stricter option: Owner only."),
    ("Decision", "Medium", "Purchase order approval authority",
     "Planning Manager, Inventory Team and CMMS Administrator are 'per separate DoA'. Asset Management has A.",
     "Approve and Reject on the Purchasing tab are Asset Management and Owner only. Anyone approving under a Delegation of Authority outside those groups needs a separate decision."),
    ("Decision", "Medium", "Account Settings (licences and seats)",
     "Not mentioned in the matrix.", "Held back for Owner. Say if CMMS Administrator should manage seats."),
    ("Decision", "Medium", "Analytics Creator seats",
     "Planner, Planning Manager, Inventory, CMMS Administrator, Asset Management and Owner all need to create analyses, which uses Creator seats.",
     "Count the people in those six groups against the Creator seats on the Enterprise plan before rollout."),
    ("Decision", "Medium", "Submitting maintenance requests from the mobile app",
     "All groups can submit a request. In Fiix, the web request form works for General, but mobile creation needs WorkOrder Add through the API, which would also let General create formal work orders.",
     "General has no WorkOrder Add. Requests are submitted on the web only. Confirm this is acceptable."),
    ("Limit", "High", "Approvals cannot be enforced by permissions",
     "Rows marked 'Process only' on the Requirement Map (approve, recommend, activate, escalate, controlled correction).",
     "Fiix permissions only allow or deny an action. Keep the approval trail in the governance process. Ask Fiix whether Workflows (a Settings menu item) can add approval steps; the public documentation does not say."),
    ("Limit", "High", "No inheritance between groups",
     "Fiix groups are independent. Copying a group copies its permissions once.",
     "Use the Build Deltas sheet in order. Any later change to a lower group must be repeated in every group above it, so treat this workbook as the source of truth."),
    ("Limit", "High", "No field-level permissions",
     "Asset criticality, hierarchy, description, PM draft versus active, WO description versus comments are not separately controllable.",
     "Control with the page-level Save and by approval process. The matrix distinctions between these rows are procedural."),
    ("Limit", "Medium", "Close routine versus critical work orders",
     "One Update/Close permission covers both.", "Planner can close. Critical or high-risk closure approval by the Planning Manager is a process control."),
    ("Limit", "Medium", "Posted receipts",
     "Matrix: Inventory cannot edit a receipt after input.", "Edit is held at CMMS Administrator, but Inventory needs Save to receive. Reconcile through the audit log."),
    ("Limit", "Medium", "Technician completes work, Planner closes it",
     "Matrix: Technician marks assigned work complete; Planner closes.",
     "Technician completes tasks and uses the mobile app (WorkOrder Change). The web close menu is Planner and above. Test this flow end to end."),
    ("Limit", "Medium", "One bucket per user",
     "Fiix allows several groups per user.",
     "Audit it: a weekly export of the SiteUserGroup class through the API (the fiix_query.py tool) flags users in more than one group."),
    ("Limit", "Medium", "CMMS Administrator, Asset Management and Owner are almost the same in Fiix",
     "The three differ by approval authority, not by permission.",
     "In Fiix they differ only in: Delete of users and groups, Approve and Reject purchase orders (Asset Management and Owner), Account Settings and the Administrator flag (Owner). Everything else is procedural."),
    ("Unverified", "Medium", "Pages Fiix does not list publicly",
     "Connect Management, Audit Log, Workflows, Notification Templates, Data Dictionary and lookup pages exist as menu items. Their page permission names are not in the public tables.",
     "Menu items are set here. Open Settings > User Groups > Page Permissions in the tenant and apply the matching pages at CMMS Administrator and above."),
    ("Unverified", "Medium", "Pages with low confidence",
     "Rows marked Low in the Confidence column of Page Settings, for example Assigned Work Orders, Work Order Administration, Receipts and Personnel.",
     "Behaviour is inferred from Fiix's descriptions. Test each in the Enterprise sandbox before production."),
    ("Unverified", "Medium", "API permissions also affect the web app",
     "Fiix documents API permissions for the mobile app and connected apps. Whether web screens depend on them is not stated.",
     "After building, log in as one user per group in the sandbox and walk through the main screens."),
    ("Unverified", "Low", "Administrator flag location in v5",
     "The assign-an-administrator article is labelled v6.", "Check the user's Access Control tab in your tenant before relying on it."),
    ("Unverified", "Low", "Legacy dashboard names",
     "Dashboard Permissions offers standard dashboards by name. The public docs name only Manager and Administrator.",
     "Pick the closest basic dashboard in the tenant for General, Technician, Planner and Inventory."),
]


def sheet_gaps(wb):
    ws = ws_new(wb, "Gaps & Decisions")
    header(ws, 1, ["Type", "Priority", "Topic", "What the models say", "What this build does / what to decide"],
           [12, 10, 40, 70, 80], 32)
    fills = {"Decision": AMBER, "Limit": RED, "Unverified": NONE_FILL}
    for i, row in enumerate(GAPS, 2):
        for j, v in enumerate(row, 1):
            c = ws.cell(row=i, column=j, value=v)
            c.border, c.alignment = BOX, WRAP
        ws.cell(row=i, column=1).fill = fills[row[0]]
        ws.cell(row=i, column=1).alignment = CENTER
        ws.cell(row=i, column=2).alignment = CENTER
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{len(GAPS) + 1}"


# --------------------------------------------------------------- summary ----
PURPOSE = {
    "G": "Leadership, Engineering, EHS, Operations, Finance. Read-only, can submit requests and use analytics.",
    "T": "Executes assigned work: labor, readings, parts, photos, completion.",
    "P": "Plans, schedules and closes routine work. Drafts PM content.",
    "M": "Owns planning quality, approves plans by process, deletes and manages controlled work.",
    "I": "Stores: parts, receiving, adjustments, cycle counts, requisition processing.",
    "A": "Technical owner: users, groups, settings, imports, API, controlled corrections.",
    "S": "Enterprise governance of asset master, standards and analytics. Same permissions as CMMS Administrator plus PO approval.",
    "O": "Ultimate authority and break-glass. Fiix Administrator flag.",
}
INHERITS = {"G": "Base", "T": "General", "P": "Technician", "M": "Planner", "I": "Technician",
            "A": "Planning Manager + Inventory Team", "S": "CMMS Administrator", "O": "Asset Management + all"}


def sheet_summary(wb, menu_last, page_last, purch_last, api_last, sim_hdr, n_matrix):
    ws = ws_new(wb, "Group Summary")
    wb.move_sheet(ws, offset=-(len(wb.sheetnames) - 2))
    ws["A1"] = "Eight groups, built for Fiix v5"
    ws["A1"].font = TITLE
    ws["A2"] = "Counts are formulas over the settings sheets."
    ws["A2"].font = F(italic=True)
    header(ws, 4, ["#", "Group", "Purpose", "Inherits (build copy of)", "Fiix default needing fewest changes", "Menu permissions",
                   "Page permissions", "Purchasing permissions", "API classes with any access", "API classes with write access",
                   "Analytics seat", "Administrator flag"],
           [5, 30, 60, 30, 20, 13, 13, 14, 15, 15, 16, 14], 56)
    seat = {"G": "Viewer / User", "T": "Viewer / User"}
    for i, k in enumerate(KEYS):
        r = 5 + i
        mc = get_column_letter(2 + i)    # Menu Settings
        pc = get_column_letter(3 + i)    # Page / Purchasing / API(+1)
        ac = get_column_letter(2 + i)    # API Settings
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=NAME[k]).font = BOLD
        ws.cell(row=r, column=3, value=PURPOSE[k])
        ws.cell(row=r, column=4, value=INHERITS[k])
        ws.cell(row=r, column=5, value=f"='Comparison to Defaults'!H{sim_hdr + 1 + i}")
        ws.cell(row=r, column=6, value=f"=COUNTIF('Menu Settings'!{mc}$2:{mc}${menu_last},\"{CHECK}\")")
        ws.cell(row=r, column=7, value=f"=COUNTIF('Page Settings'!{pc}$2:{pc}${page_last},\"{CHECK}\")")
        ws.cell(row=r, column=8, value=f"=COUNTIF('Purchasing Settings'!{pc}$2:{pc}${purch_last},\"{CHECK}\")")
        ws.cell(row=r, column=9, value=f"=COUNTIF('API Settings'!{ac}$2:{ac}${api_last},\"<>-\")")
        rng = f"'API Settings'!{ac}$2:{ac}${api_last}"
        ws.cell(row=r, column=10, value=(f'=SUMPRODUCT(--((ISNUMBER(SEARCH("A",{rng}))+ISNUMBER(SEARCH("C",{rng}))'
                                         f'+ISNUMBER(SEARCH("R",{rng})))>0))'))
        ws.cell(row=r, column=11, value=seat.get(k, "Creator"))
        ws.cell(row=r, column=12, value="Yes" if k == "O" else "No")
        for c in range(1, 13):
            cell = ws.cell(row=r, column=c)
            cell.border = BOX
            cell.alignment = WRAP if c in (3, 4) else CENTER if c != 2 else Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 42
    r = 14
    ws.cell(row=r, column=2, value="Total available").font = BOLD
    ws.cell(row=r, column=6, value=f"=COUNTA('Menu Settings'!A2:A{menu_last})")
    ws.cell(row=r, column=7, value=f"=COUNTA('Page Settings'!B2:B{page_last})")
    ws.cell(row=r, column=8, value=f"=COUNTA('Purchasing Settings'!B2:B{purch_last})")
    ws.cell(row=r, column=9, value=f"=COUNTA('API Settings'!A2:A{api_last})")
    for c in range(1, 13):
        ws.cell(row=r, column=c).fill = BAND
        ws.cell(row=r, column=c).border = BOX
        ws.cell(row=r, column=c).alignment = CENTER
    ws.cell(row=14, column=2).alignment = Alignment(vertical="center")
    ws["A16"] = "Requirement coverage"
    ws["A16"].font = F(size=12, bold=True, color="1F3864")
    for i, (lab, col) in enumerate((("Enforced by Fiix permissions", GREEN), ("Approximate, qualifier cannot be enforced", AMBER),
                                    ("Process control only", RED))):
        ws.cell(row=17 + i, column=2, value=lab).fill = col
        ws.cell(row=17 + i, column=3, value=f"='Requirement Map'!B{4 + i}")
        ws.cell(row=17 + i, column=3).alignment = Alignment(horizontal="left")
    ws.cell(row=20, column=2, value="Total matrix rows").font = BOLD
    ws.cell(row=20, column=3, value="='Requirement Map'!B7").alignment = Alignment(horizontal="left")
    ws.freeze_panes = "C5"
    ws.sheet_view.showGridLines = False


def sheet_readme(wb):
    ws = wb.active
    ws.title = "Read Me"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Fiix v5 group build sheet"
    ws["A1"].font = TITLE
    ws["A2"] = "IREN management's eight-group permission model translated into Fiix v5 settings"
    ws["A2"].font = F(italic=True)
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 120
    rows = [
        ("Inputs", "IREN_Fiix_Enterprise_Permission_Model_Updated.xlsx (Permission Matrix, Bucket Definitions, Governance Fiix Mapping) and the Fiix default permission tables in Fiix_Permissions_Reference.xlsx."),
        ("Method", "Each Fiix permission is given an entry group, the lowest group that needs it. The management model is cumulative, so every group above the entry group also gets it. "
                   "CMMS Administrator also receives the Inventory Team permissions, because the matrix gives it inventory authority."),
        ("Reading the sheets", ""),
        ("  Group Summary", "One line per group, with live counts and coverage of the management matrix."),
        ("  Menu / Page / Purchasing Settings", "The checkboxes to set on each tab of a Fiix user group. Grey columns show Fiix's own default groups for comparison."),
        ("  API Settings", "The Add, Find, Change, Remove boxes for each Fiix class."),
        ("  Per-User Settings", "Things set on the user, not the group: Administrator flag and analytics seats."),
        ("  Build Deltas", "Build order. Copy the previous group, then turn on only what is listed. This is the fastest way to build in Fiix."),
        ("  Comparison to Defaults", "How close each new group is to Fiix's default groups, and what was added or removed."),
        ("  Requirement Map", "All matrix rows with how Fiix implements each, or what is lost."),
        ("  Gaps & Decisions", "Conflicts between the management sheets, things Fiix cannot enforce, and items to confirm."),
        ("How to build in Fiix", "Settings > User Groups > Create > Create New for General, then Copy Existing for each next group as shown in Build Deltas. "
                                 "Open a group > Menu Permissions, Page Permissions, Purchasing Permissions and API Permissions tabs. Group names are suggestions."),
        ("Confidence", "Default-table rows were verified against the Help Center. Rows from feature articles are Medium confidence. Rows marked Low are inferred from descriptions. "
                       "Test with one user per group in the Enterprise sandbox before production."),
        ("Scope", "Fiix v5 only. Fiix's v6 system roles and v6 work order permissions do not apply and are not used."),
        ("Not covered", "Mobile app feature switches, notification setup, site and region setup, and the exact permission names for Settings pages that Fiix does not list publicly."),
    ]
    r = 4
    for a, b in rows:
        ws.cell(row=r, column=1, value=a).font = BOLD if not a.startswith("  ") else F()
        c = ws.cell(row=r, column=2, value=b)
        c.alignment = WRAP
        ws.cell(row=r, column=1).alignment = Alignment(vertical="top")
        r += 1
    return ws


def normalize_fonts(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                f = cell.font
                if f.name != "Arial":
                    cell.font = Font(name="Arial", size=10 if f.size in (None, 11) else f.size, bold=f.bold, italic=f.italic,
                                     color=f.color, underline=f.underline)


def main():
    page_rows, menu_rows, purch_rows, api_rows = build_rows()
    # sanity: cumulative design must hold everywhere
    deltas = build_deltas(page_rows, menu_rows, purch_rows, api_rows)
    wb = Workbook()
    sheet_readme(wb)
    menu_last = sheet_menu(wb, menu_rows)
    page_last, page_def_last = sheet_page(wb, page_rows)
    purch_last = sheet_purch(wb, purch_rows)
    api_last = sheet_api(wb, api_rows)
    sheet_user_settings(wb)
    sheet_deltas(wb, deltas)
    menu_def_last = 1 + sum(1 for r in menu_rows if r["defaults"] is not None)
    sim_hdr = sheet_comparison(wb, page_def_last, menu_def_last, purch_last)
    n_matrix = sheet_reqmap(wb, load_matrix())
    sheet_gaps(wb)
    sheet_summary(wb, menu_last, page_last, purch_last, api_last, sim_hdr, n_matrix)
    normalize_fonts(wb)
    wb.save(OUT)
    print(f"saved {OUT}")
    print(f"menu {len(menu_rows)} | page {len(page_rows)} ({page_def_last - 1} default-table rows) | purchasing {len(purch_rows)} | "
          f"api {len(api_rows)} | deltas {len(deltas)} | matrix rows {n_matrix}")


if __name__ == "__main__":
    main()
