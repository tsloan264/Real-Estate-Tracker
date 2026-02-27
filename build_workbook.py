"""
HH Realty Transaction Tracker - Redesigned Workbook Builder
Architect: Claude (Senior Excel Systems Architect)
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.protection import SheetProtection
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.styles.differential import DifferentialStyle
from datetime import datetime, date
import copy

# ─────────────────────────────────────────────────────────────────────────────
# COLOR PALETTE
# ─────────────────────────────────────────────────────────────────────────────
C_NAVY       = "1F3864"   # Sheet headers
C_DARK_BLUE  = "2F5496"   # Section headers
C_TEAL       = "17375E"   # Alt headers
C_GOLD       = "C9A227"   # Accents
C_WHITE      = "FFFFFF"
C_INPUT_YEL  = "FFF2CC"   # User input cells
C_INPUT_BDR  = "F4B942"   # Input cell border
C_CALC_BLUE  = "D9E1F2"   # Calculated cells
C_CALC_BDR   = "9DC3E6"   # Calc border
C_LOCK_GRAY  = "F2F2F2"   # Locked / system cells
C_LOCK_BDR   = "BFBFBF"   # Locked border
C_GREEN_KPI  = "E2EFDA"   # Dashboard positive KPI
C_RED_KPI    = "FFDBD6"   # Dashboard warning KPI
C_DASH_BG    = "1F3864"   # Dashboard background
C_ROW_ALT    = "EBF3FF"   # Alternate row in tables
C_ERROR      = "FF0000"   # Error highlight
C_HEADER_ROW = "2E75B6"   # Table column headers

# ─────────────────────────────────────────────────────────────────────────────
# STYLE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, size=11, color=C_WHITE, italic=False, name="Calibri"):
    return Font(bold=bold, size=size, color=color, italic=italic, name=name)

def border(style="thin", color="BFBFBF"):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def apply(cell, fill_c=None, font_c=None, align_c=None, border_c=None,
          bold=False, size=11, color=C_WHITE, h="left", v="center",
          wrap=False, num_fmt=None, locked=True):
    if fill_c:
        cell.fill = fill(fill_c)
    if font_c:
        cell.font = font_c
    else:
        cell.font = Font(bold=bold, size=size, color=color, name="Calibri")
    if align_c:
        cell.alignment = align_c
    else:
        cell.alignment = align(h, v, wrap)
    if border_c:
        cell.border = border_c
    if num_fmt:
        cell.number_format = num_fmt
    cell.protection = openpyxl.styles.Protection(locked=locked)

def set_cell(ws, row, col, value=None, fill_c=None, bold=False, size=11,
             color="000000", h="left", v="center", wrap=False, num_fmt=None,
             locked=True, border_c=None, italic=False):
    cell = ws.cell(row=row, column=col)
    if value is not None:
        cell.value = value
    cell.fill = fill(fill_c) if fill_c else PatternFill()
    cell.font = Font(bold=bold, size=size, color=color, name="Calibri", italic=italic)
    cell.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    if border_c:
        cell.border = border_c
    if num_fmt:
        cell.number_format = num_fmt
    cell.protection = openpyxl.styles.Protection(locked=locked)
    return cell

def header_cell(ws, row, col, value, bg=C_NAVY, fg=C_WHITE, size=12,
                bold=True, h="center", v="center", wrap=False, span=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = fill(bg)
    cell.font = Font(bold=bold, size=size, color=fg, name="Calibri")
    cell.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    cell.protection = openpyxl.styles.Protection(locked=True)
    if span and span > 1:
        end_col = get_column_letter(col + span - 1)
        ws.merge_cells(f"{get_column_letter(col)}{row}:{end_col}{row}")
    return cell

def section_header(ws, row, col, value, span, bg=C_DARK_BLUE):
    header_cell(ws, row, col, value, bg=bg, size=11, span=span)

def input_cell(ws, row, col, value=None, num_fmt=None, h="left"):
    cell = ws.cell(row=row, column=col)
    if value is not None:
        cell.value = value
    cell.fill = fill(C_INPUT_YEL)
    cell.font = Font(size=11, color="000000", name="Calibri")
    cell.alignment = Alignment(horizontal=h, vertical="center")
    s = Side(style="thin", color=C_INPUT_BDR)
    cell.border = Border(left=s, right=s, top=s, bottom=s)
    if num_fmt:
        cell.number_format = num_fmt
    cell.protection = openpyxl.styles.Protection(locked=False)
    return cell

def calc_cell(ws, row, col, value=None, num_fmt=None, h="center"):
    cell = ws.cell(row=row, column=col)
    if value is not None:
        cell.value = value
    cell.fill = fill(C_CALC_BLUE)
    cell.font = Font(size=11, color="000000", name="Calibri")
    cell.alignment = Alignment(horizontal=h, vertical="center")
    s = Side(style="thin", color=C_CALC_BDR)
    cell.border = Border(left=s, right=s, top=s, bottom=s)
    if num_fmt:
        cell.number_format = num_fmt
    cell.protection = openpyxl.styles.Protection(locked=True)
    return cell

def lock_cell(ws, row, col, value=None, num_fmt=None, h="left"):
    cell = ws.cell(row=row, column=col)
    if value is not None:
        cell.value = value
    cell.fill = fill(C_LOCK_GRAY)
    cell.font = Font(size=11, color="595959", name="Calibri", italic=True)
    cell.alignment = Alignment(horizontal=h, vertical="center")
    s = Side(style="thin", color=C_LOCK_BDR)
    cell.border = Border(left=s, right=s, top=s, bottom=s)
    if num_fmt:
        cell.number_format = num_fmt
    cell.protection = openpyxl.styles.Protection(locked=True)
    return cell

def protect_sheet(ws, password="HHRealty2026"):
    ws.protection = SheetProtection(
        sheet=True,
        password=password,
        formatCells=False,
        formatColumns=False,
        formatRows=False,
        insertColumns=False,
        insertRows=True,    # allow inserting rows
        insertHyperlinks=False,
        deleteColumns=False,
        deleteRows=True,    # allow deleting rows
        selectLockedCells=True,
        selectUnlockedCells=True,
        sort=True,
        autoFilter=True,
    )

# ─────────────────────────────────────────────────────────────────────────────
# EXISTING DATA (extracted from V13)
# ─────────────────────────────────────────────────────────────────────────────
AGENTS = [
    ("Pulte Homes",           "Pulte Homes"),
    ("Barb Lesure",           "HH"),
    ("Ivana Carson",          "Keller"),
    ("Rachel Eastwood",       "HH"),
    ("Eric Yetzer",           "HH"),
    ("Karen/Milka",           "HH"),
    ("Kelly/Heidi",           "HH"),
    ("Matthew Kalogeras",     "Trelora"),
    ("Karen/Annette",         "HH"),
    ("Mark Metz",             "Keller"),
    ("Rebbecca Tessanne",     "HH"),
    ("Ron Ho",                "HH"),
    ("Vicky & Beth",          "HH"),
    ("Lenny Herbert",         "Remax"),
    ("Vicky & Beth (Shelly)", "HH"),
    ("Debbie Ferrante",       "Remax"),
    ("Jennifer Ham",          "BHHS Stouffer"),
    ("Kelly Watts",           "HH"),
    ("Amy Frary",             "HH"),
    ("Katrina Heath",         "McDowell"),
    ("Andrew Jenkins",        "Keller"),
    ("Jerry Lesac",           "McDowell"),
    ("Tim Vanderlaan",        "EXP"),
    ("Oleg Bosovic",          "Keller"),
    ("Shelly Booth",          "Coldwell"),
    ("Autumn Moore",          "Keller"),
    ("Brian Leatherman",      "Milton"),
    ("Cindy McCory",          "Remax"),
    ("Lila Wohlwend",         "Clear Sky"),
    ("Amy Marinello",         "BHHS"),
    ("Melissa Steiner",       "Keller"),
    ("Deb Shreiner",          "Helen Scott"),
    ("Brenda Verbeka",        "HH Medina"),
    ("Christine Conkle",      "ReMax Edge"),
    ("Sonja Halstead",        "Keller Williams"),
]

LOAN_TYPES = ["CONV", "FHA", "VA", "CASH", "New Const", "Eq. Ln", "Other"]
STATUSES   = ["Not Started", "In Progress", "Waiting", "Complete", "N/A"]
EM_HOLD    = ["HH", "Title", "CoBroke", "Buyer", "Seller", "N/A", "Other"]
HW_PAY     = ["Seller", "Buyer", "Split", "Other", "N/A"]
YES_NO     = ["Yes", "No"]
TRANS_STATUS = ["Active", "Closed", "BOMB"]

# fmt: off
TRANSACTIONS = [
    # (Address, SignNum, ListPrice, SalePrice, OfferAcceptDt, ClosingDt, LoanType, Concessions,
    #  ListAgent, SellAgent, CommPct, ListingCommPct, CommFlatAmt,
    #  EMAmt, EMHold, HomeWarranty, HSA, HWPay, FlatFee, TitleCo, TitleContact, Notes, Status, CreatedDt)
    ("998 Barn Swallow, Wadsworth 44281",    None,     None,   617785, None,                       datetime(2026,3,15),  "New Const", 0,      "Pulte Homes",           "Kelly Watts",            0.02,  0,     0,      0,    None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("133 Beck, Wadsworth 44281",            None,     None,    85500, None,                       datetime(2026,3,6),   "CASH",      0,      "Barb Lesure",           "Ron Ho",                 0,     0,     2500,   0,    None,  None,  None,  None,  "No",   None,               None, None, "Active", datetime(2026,2,26)),
    ("743 Bent Creek, Wadsworth 44281",      None,     None,   245000, None,                       datetime(2026,3,7),   "CASH",      0,      "Ivana Carson",          "Amy Frary",              0.02,  0,     0,      0,    None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("2644 Harpster, Rittman 44270",         None,     None,   275000, None,                       datetime(2026,3,13),  "FHA",       0,      "Karen/Milka",           "Jerry Lesac",            0.06,  0.03,  0,      0,    None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("1774 Hila Way, Wooster 44691",         None,     None,   310000, None,                       datetime(2026,2,27),  "CONV",      600,    "Kelly/Heidi",           "Cindy McCory",           0.055, 0.03,  0,      0,    None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("744 Lawrence, Wadsworth 44281",        None,     None,   360000, None,                       datetime(2026,2,16),  "CONV",      0,      "Matthew Kalogeras",     "Rachel Eastwood",        0.03,  0,     0,      1000, None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("281 Park Pl, Wadsworth 44281",         None,     None,   266000, None,                       datetime(2026,3,5),   "Eq. Ln",    0,      "Mark Metz",             "Karen/Milka",            0.03,  0,     0,      2000, None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("92 30th, Barberton 44203",             None,     None,   193000, None,                       datetime(2026,2,27),  "CONV",      0,      "Eric Yetzer",           "Mark Metz",              0.065, 0.035, 0,      0,    None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("455 Barrenwood, Wadsworth 44281",      None,     None,   315000, None,                       datetime(2026,3,11),  "CASH",      0.02,   "Rebbecca Tessanne",     "Oleg Bosovic",           0.046, 0.026, 0,      0,    None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("5052 Catawba, Seville 44273",          None,     None,   365000, None,                       datetime(2026,4,7),   "CONV",      0,      "Eric Yetzer",           "Shelly Booth",           0.06,  0.03,  0,      0,    None,  None,  None,  None,  "No",   None,               None, None, "Active", datetime(2026,2,26)),
    ("216 N. Lyman, Wadsworth 44281",        None,     None,   235000, None,                       datetime(2026,3,6),   "FHA",       5000,   "Ron Ho",                "Autumn Moore",           0.06,  0,     0,      0,    None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("6407 Southview, Clinton 44216",        None,     None,   256000, None,                       datetime(2026,3,16),  "CONV",      0,      "Vicky & Beth",          "Vicky & Beth",           0.04,  0.025, 0,      0,    None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("1821 Myersville, Akron 44312",         None,     None,   330000, None,                       datetime(2026,2,26),  "FHA",       0.03,   "Lenny Herbert",         "Vicky & Beth (Shelly)",  0.02,  0,     0,      1000, None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("96 W Ohio, Rittman 44270",             None,     None,   168000, None,                       datetime(2026,3,13),  "CONV",      2500,   "Eric Yetzer",           "Brian Leatherman",       0.065, 0.035, 0,      0,    None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("317 Ihrig, Wooster 44691",             None,     None,   230000, None,                       datetime(2026,3,20),  "CONV",      6000,   "Vicky & Beth (Shelly)", "Cindy McCory",           0.03,  0.01,  0,      0,    None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("6116 Boneta, Medina 44256",            None,     None,   541000, None,                       datetime(2026,3,4),   "CONV",      0,      "Debbie Ferrante",       "Eric Yetzer",            0.02,  0,     0,      3000, None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("2736 Applehill, Robertsville 44670",   None,     None,    67000, None,                       datetime(2026,2,27),  "CASH",      0,      "Vicky & Beth (Shelly)", "Lila Wohlwend",          0.04,  0,     0,      0,    None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("552 Main #H, Wadsworth 44281",         None,     None,   200000, None,                       datetime(2026,3,12),  "CASH",      0,      "Jennifer Ham",          "Barb Lesure",            0.025, 0,     0,      1000, None,  None,  None,  None,  "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("1047 Ashwood, Wooster 44691",          None,    335000,  325000, datetime(2026,2,13),        datetime(2026,3,20),  "VA",        0,      "Amy Marinello",         "Kelly/Heidi",            0.02,  0,     0,      0,    "Title","No",  "No",  "N/A", "Yes",  None,               None, None, "Active", datetime(2026,2,26)),
    ("18592 Edwards #169, Doylestown 44230", None,     34900,   30000, datetime(2026,2,18),        datetime(2026,3,3),   "CASH",      0,      "Vicky & Beth",          "Melissa Steiner",        0,     0,     4000,   0,    None,  "No",  "No",  "N/A", "No",   "NA",               None, None, "Active", datetime(2026,2,26)),
    ("536 Austin, Akron 44320",              None,    134900,  134900, datetime(2026,2,17),        datetime(2026,3,19),  "CONV",      0,      "Kelly Watts",           "Deb Shreiner",           0,     0,     7070.5, 0,    "Title","No",  "No",  "N/A", "Yes",  "First Security",   None, None, "Active", datetime(2026,2,26)),
    ("349 E Hopocan, Barberton 44203",       3203612, 105000,  108000, datetime(2026,2,22),        datetime(2026,3,27),  "FHA",       5500,   "Barb Lesure",           "Brenda Verbeka",         0.07,  0.045, 0,      500,  "HH",  "Yes", "Yes", "Buyer","Yes",  "Ohio Real Title",  None, None, "Active", datetime(2026,2,26)),
    ("12249 Whitman, Doylestown 44230",      3205909, 199900,  185500, datetime(2026,2,23),        datetime(2026,4,6),   "CONV",      5500,   "Vicky & Beth",          "Christine Conkle",       0.05,  0.03,  0,      1000, "Other","Yes","Yes", "Seller","Yes", "Kropf (Orrville Law)",None,None,"Active", datetime(2026,2,26)),
    ("3179 Vanderhoof, Clinton 44216",       None,    189900,  200000, datetime(2026,2,22),        datetime(2026,4,3),   "CONV",      0,      "Sonja Halstead",        "Rebbecca Tessanne",      0,     0,     0,      1000, "HH",  "No",  "No",  "N/A", "Yes",  "First Meridian",   None, None, "Active", datetime(2026,2,26)),
]

CHECKLIST_COLS = [
    'MLSPrintout','FHAVAAddendum','AgencyDisclosure','CondoAddendum','ConsumerGuide',
    'PurchaseAgreement','ABA','ABAPOD','PropertyDisclosure','HWAppReceived','EBRA',
    'LBPDisclosure','OtherDocs','HomeInspDisclosure','AntiFraudDisclosure','PreApprovalProof',
    'MLSStatusChange','ZipFormsSaleFile','ZipFormsEBRA','SentToDataEntry','PrestartToTitle',
    'ScanNoticeHHMS','LogSalesBook','FBPostWindowSheet','RecordRobSpreadsheet','RegisterBuyerEBRA',
    'HWOrdered','EMReceivedDeposited','EMDepositMethod','EMSentAccounting','CopyPDriveZipForms',
    'CommInvoiceToTitle','ManagerSalesSheet','CommInvoicePDrive','SocialMediaSoldPost',
    'MarkSoldMLS','OrderSignDown','ZipFormsClosed','RemoveWindowSheet',
    'BOMBMLSStatus','BOMBMutualRelease','BOMBTitleNotified','BOMBCodedPP',
    'BOMBHWCanceled','BOMBZipFormsStatus','BOMBWindowSheet',
]

CHECKLIST_LABELS = {
    'MLSPrintout':'MLS Printout','FHAVAAddendum':'FHA/VA Addendum','AgencyDisclosure':'Agency Disclosure',
    'CondoAddendum':'Condo Addendum','ConsumerGuide':'Consumer Guide','PurchaseAgreement':'Purchase Agreement',
    'ABA':'ABA','ABAPOD':'ABA (POD)','PropertyDisclosure':'Property Disclosure',
    'HWAppReceived':'Home Warranty App Received','EBRA':'EBRA','LBPDisclosure':'Lead-Based Paint Disclosure',
    'OtherDocs':'Other Documents','HomeInspDisclosure':'Home Inspection Disclosure',
    'AntiFraudDisclosure':'Anti-Fraud Disclosure','PreApprovalProof':'Pre-Approval / Proof of Funds',
    'MLSStatusChange':'Change MLS Status (Our Listing)','ZipFormsSaleFile':'Create Sale File — ZipForms',
    'ZipFormsEBRA':'Add EBRA to ZipForms (Our Buyer)','SentToDataEntry':'Sent to Data Entry',
    'PrestartToTitle':'Prestart Sent to Title Co.','ScanNoticeHHMS':'Scan Notice to HHMS (Our Buyer)',
    'LogSalesBook':'Log in Sales Book/Board','FBPostWindowSheet':'FB Post / Update Window Sheet (Our Listing)',
    'RecordRobSpreadsheet':"Record on Rob's Spreadsheet",'RegisterBuyerEBRA':'Register Buyer in EBRA Database',
    'HWOrdered':'Home Warranty Ordered','EMReceivedDeposited':'Earnest Money Received & Deposited',
    'EMDepositMethod':'EM Deposit Method','EMSentAccounting':'EM Sent to Accounting',
    'CopyPDriveZipForms':'Copy in P: Drive / ZipForms','CommInvoiceToTitle':'Send Commission Invoice to Title',
    'ManagerSalesSheet':'Manager Sales Sheet to E-Sign','CommInvoicePDrive':'Commission Invoice — P Drive',
    'SocialMediaSoldPost':'Social Media SOLD Post','MarkSoldMLS':'Mark Sold in MLS',
    'OrderSignDown':'Order Post Sign Down','ZipFormsClosed':'Move to Closed — ZipForms',
    'RemoveWindowSheet':'Take Sheet Out of Front Window','BOMBMLSStatus':'[BOMB] Change MLS Status',
    'BOMBMutualRelease':'[BOMB] Mutual Release to Susan Ipavec','BOMBTitleNotified':'[BOMB] Notify Title Co.',
    'BOMBCodedPP':'[BOMB] Coded in PP','BOMBHWCanceled':'[BOMB] Home Warranty Notified/Canceled',
    'BOMBZipFormsStatus':'[BOMB] Change Status in ZipForms','BOMBWindowSheet':'[BOMB] Change Window Sheet',
}

CHECKLIST_CATEGORY = {
    'MLSPrintout':'Check Paperwork','FHAVAAddendum':'Check Paperwork','AgencyDisclosure':'Check Paperwork',
    'CondoAddendum':'Check Paperwork','ConsumerGuide':'Check Paperwork','PurchaseAgreement':'Check Paperwork',
    'ABA':'Check Paperwork','ABAPOD':'Check Paperwork','PropertyDisclosure':'Check Paperwork',
    'HWAppReceived':'Check Paperwork','EBRA':'Check Paperwork','LBPDisclosure':'Check Paperwork',
    'OtherDocs':'Check Paperwork','HomeInspDisclosure':'Check Paperwork','AntiFraudDisclosure':'Check Paperwork',
    'PreApprovalProof':'Check Paperwork','MLSStatusChange':'Action Items','ZipFormsSaleFile':'Action Items',
    'ZipFormsEBRA':'Action Items','SentToDataEntry':'Action Items','PrestartToTitle':'Action Items',
    'ScanNoticeHHMS':'Action Items','LogSalesBook':'Action Items','FBPostWindowSheet':'Action Items',
    'RecordRobSpreadsheet':'Action Items','RegisterBuyerEBRA':'Action Items','HWOrdered':'Action Items',
    'EMReceivedDeposited':'EM / Commission / HW','EMDepositMethod':'EM / Commission / HW',
    'EMSentAccounting':'EM / Commission / HW','CopyPDriveZipForms':'EM / Commission / HW',
    'CommInvoiceToTitle':'EM / Commission / HW','ManagerSalesSheet':'EM / Commission / HW',
    'CommInvoicePDrive':'EM / Commission / HW','SocialMediaSoldPost':'Closed Sale',
    'MarkSoldMLS':'Closed Sale','OrderSignDown':'Closed Sale','ZipFormsClosed':'Closed Sale',
    'RemoveWindowSheet':'Closed Sale','BOMBMLSStatus':'Fallen Sale (BOMB)',
    'BOMBMutualRelease':'Fallen Sale (BOMB)','BOMBTitleNotified':'Fallen Sale (BOMB)',
    'BOMBCodedPP':'Fallen Sale (BOMB)','BOMBHWCanceled':'Fallen Sale (BOMB)',
    'BOMBZipFormsStatus':'Fallen Sale (BOMB)','BOMBWindowSheet':'Fallen Sale (BOMB)',
}

# Checklist data from original (partial - only available data)
CHECKLIST_DATA = {
    "2026-001": {"MLSPrintout":"Complete","FHAVAAddendum":"Complete","AgencyDisclosure":"Complete","CondoAddendum":"N/A","ConsumerGuide":"Complete","PurchaseAgreement":"Complete","ABA":"Complete","ABAPOD":"Complete","PropertyDisclosure":"Complete","HWAppReceived":"Complete","EBRA":"Complete","LBPDisclosure":"Not Started","HomeInspDisclosure":"Not Started","AntiFraudDisclosure":"Not Started","PreApprovalProof":"Not Started","MLSStatusChange":"Not Started","ZipFormsSaleFile":"Not Started"},
    "2026-002": {"MLSPrintout":"Complete","FHAVAAddendum":"N/A","AgencyDisclosure":"Complete","CondoAddendum":"N/A","ConsumerGuide":"Complete","PurchaseAgreement":"Complete","ABA":"Complete","ABAPOD":"N/A","PropertyDisclosure":"Complete","HWAppReceived":"N/A","EBRA":"Complete","LBPDisclosure":"Complete","HomeInspDisclosure":"N/A","AntiFraudDisclosure":"Complete","PreApprovalProof":"Complete","MLSStatusChange":"Complete","ZipFormsSaleFile":"Complete"},
    "2026-003": {"MLSPrintout":"Complete","FHAVAAddendum":"N/A","AgencyDisclosure":"Complete","CondoAddendum":"Complete","ConsumerGuide":"Complete","PurchaseAgreement":"Complete","ABA":"Complete","ABAPOD":"N/A","PropertyDisclosure":"Complete","HWAppReceived":"Complete","EBRA":"Complete","LBPDisclosure":"N/A","HomeInspDisclosure":"N/A","AntiFraudDisclosure":"Complete","PreApprovalProof":"Complete","MLSStatusChange":"N/A","ZipFormsSaleFile":"Complete"},
}

print("Data definitions loaded successfully.")
print(f"  Transactions: {len(TRANSACTIONS)}")
print(f"  Agents: {len(AGENTS)}")
print(f"  Checklist cols: {len(CHECKLIST_COLS)}")


# ─────────────────────────────────────────────────────────────────────────────
# CREATE WORKBOOK
# ─────────────────────────────────────────────────────────────────────────────
wb = Workbook()
wb.remove(wb.active)   # remove default sheet

# ─────────────────────────────────────────────────────────────────────────────
# SHEET 1: INSTRUCTIONS
# ─────────────────────────────────────────────────────────────────────────────
ws_inst = wb.create_sheet("📋 Instructions")
ws_inst.sheet_view.showGridLines = False
ws_inst.column_dimensions["A"].width = 3
ws_inst.column_dimensions["B"].width = 32
ws_inst.column_dimensions["C"].width = 60
ws_inst.column_dimensions["D"].width = 20

# Title
ws_inst.merge_cells("A1:D1")
c = ws_inst["A1"]
c.value = "📋  HH REALTY TRANSACTION TRACKER  —  How To Use This File"
c.fill = fill(C_NAVY)
c.font = Font(bold=True, size=18, color=C_WHITE, name="Calibri")
c.alignment = Alignment(horizontal="center", vertical="center")
ws_inst.row_dimensions[1].height = 40

ws_inst.merge_cells("A2:D2")
c = ws_inst["A2"]
c.value = "Designed for daily use by HH Realty staff  ·  Keep this file saved on the network drive"
c.fill = fill(C_DARK_BLUE)
c.font = Font(size=11, color=C_WHITE, name="Calibri", italic=True)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_inst.row_dimensions[2].height = 22

# Color legend
row = 4
ws_inst.merge_cells(f"B{row}:D{row}")
ws_inst[f"B{row}"].value = "COLOR GUIDE — What the cell colors mean"
ws_inst[f"B{row}"].fill = fill(C_DARK_BLUE)
ws_inst[f"B{row}"].font = Font(bold=True, size=12, color=C_WHITE, name="Calibri")
ws_inst[f"B{row}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws_inst.row_dimensions[row].height = 22

colors = [
    (C_INPUT_YEL, "000000", "🟡  YELLOW cell", "TYPE HERE — these are your entry fields"),
    (C_CALC_BLUE, "000000", "🔵  BLUE cell",   "AUTO-CALCULATED — do not type here, formula fills it in"),
    (C_LOCK_GRAY, "595959", "⚪  GREY cell",   "LOCKED — system field, cannot be edited"),
    ("E2EFDA",   "000000", "🟢  GREEN (Dashboard)", "Key performance number, updates automatically"),
    ("FFDBD6",   "000000", "🔴  RED (Dashboard)",   "Warning — needs attention"),
]
for i, (bg, fg, label, desc) in enumerate(colors):
    r = row + 1 + i
    ws_inst[f"B{r}"].fill = fill(bg)
    ws_inst[f"B{r}"].value = label
    ws_inst[f"B{r}"].font = Font(size=11, color=fg, name="Calibri", bold=True)
    ws_inst[f"B{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    s = Side(style="thin", color="BFBFBF")
    ws_inst[f"B{r}"].border = Border(left=s, right=s, top=s, bottom=s)
    ws_inst[f"C{r}"].value = desc
    ws_inst[f"C{r}"].font = Font(size=11, color="000000", name="Calibri")
    ws_inst[f"C{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_inst.row_dimensions[r].height = 20

# How to use steps
step_row = row + len(colors) + 2
ws_inst.merge_cells(f"B{step_row}:D{step_row}")
ws_inst[f"B{step_row}"].value = "STEP-BY-STEP GUIDE"
ws_inst[f"B{step_row}"].fill = fill(C_DARK_BLUE)
ws_inst[f"B{step_row}"].font = Font(bold=True, size=12, color=C_WHITE, name="Calibri")
ws_inst[f"B{step_row}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws_inst.row_dimensions[step_row].height = 22

steps = [
    ("STEP 1 — Add a New Transaction",
     "Go to the '🏠 INPUT' sheet. Fill in every YELLOW cell. Required fields are marked with ★. "
     "Click the 'Add Transaction' button when done. The row will appear in '🗄️ TRANSACTIONS'."),
    ("STEP 2 — Update Checklist Items",
     "Go to '✅ CHECKLIST'. Find the row for your transaction (search by address). "
     "Use the dropdown in each column to mark items as: Complete, In Progress, Waiting, N/A, or Not Started."),
    ("STEP 3 — Edit a Transaction",
     "Go to '🗄️ TRANSACTIONS'. Find the row and click any YELLOW cell to edit. "
     "BLUE and GREY cells update automatically — do not type in them."),
    ("STEP 4 — View the Dashboard",
     "Go to '📊 Dashboard'. It updates automatically every time you open the file. "
     "No need to press any buttons or refresh."),
    ("STEP 5 — Add a New Agent",
     "Go to '👤 Agents'. Scroll to the bottom of the agent list. "
     "Type the agent's name in column B and their office in column C. The ID fills automatically."),
    ("STEP 6 — Close or BOMB a Transaction",
     "In '🗄️ TRANSACTIONS', find the row and change the 'Status' column (col Y) to 'Closed' or 'BOMB'. "
     "The dashboard will update accordingly."),
    ("IMPORTANT — Do Not...",
     "• Do NOT delete the header rows (rows 1-2 in any sheet).\n"
     "• Do NOT sort or delete rows in the AGENTS sheet.\n"
     "• Do NOT type in BLUE or GREY cells — they are protected.\n"
     "• Do NOT copy/paste entire rows — paste 'Values Only' if needed."),
    ("TIPS",
     "• If a dropdown doesn't show your agent, first add them to the '👤 Agents' sheet.\n"
     "• Commission %: enter as a decimal (e.g., 6% = 0.06).\n"
     "• If you make a mistake, press Ctrl+Z to undo.\n"
     "• Save the file with Ctrl+S after each entry."),
]
for i, (title, body) in enumerate(steps):
    r = step_row + 1 + (i * 3)
    ws_inst.merge_cells(f"B{r}:D{r}")
    ws_inst[f"B{r}"].value = title
    ws_inst[f"B{r}"].fill = fill("D6E4F7")
    ws_inst[f"B{r}"].font = Font(bold=True, size=11, color=C_NAVY, name="Calibri")
    ws_inst[f"B{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_inst.row_dimensions[r].height = 20

    ws_inst.merge_cells(f"B{r+1}:D{r+1}")
    ws_inst[f"B{r+1}"].value = body
    ws_inst[f"B{r+1}"].fill = fill(C_WHITE) if fill else PatternFill()
    ws_inst[f"B{r+1}"].font = Font(size=10, color="000000", name="Calibri")
    ws_inst[f"B{r+1}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
    ws_inst.row_dimensions[r+1].height = 45

print("Sheet 1 (Instructions) built.")


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 2: LOOKUPS (hidden - must be created before sheets that reference it)
# ─────────────────────────────────────────────────────────────────────────────
ws_lkp = wb.create_sheet("⚙️ Lookups")
ws_lkp.sheet_state = "hidden"

def lkp_table(ws, start_col, start_row, title, headers, rows, tbl_name):
    """Write a lookup table."""
    ws.cell(start_row, start_col, title).font = Font(bold=True, size=10, color=C_NAVY, name="Calibri")
    for j, h in enumerate(headers):
        c = ws.cell(start_row + 1, start_col + j, h)
        c.fill = fill(C_HEADER_ROW)
        c.font = Font(bold=True, size=10, color=C_WHITE, name="Calibri")
        c.alignment = Alignment(horizontal="center")
    for i, row_data in enumerate(rows):
        for j, v in enumerate(row_data):
            c = ws.cell(start_row + 2 + i, start_col + j, v)
            c.font = Font(size=10, color="000000", name="Calibri")
            c.alignment = Alignment(horizontal="left")
    # create table
    end_row = start_row + 1 + len(rows)
    end_col = get_column_letter(start_col + len(headers) - 1)
    rng = f"{get_column_letter(start_col)}{start_row+1}:{end_col}{end_row}"
    try:
        tbl = Table(displayName=tbl_name, ref=rng)
        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight2", showRowStripes=True)
        ws.add_table(tbl)
    except Exception as e:
        print(f"  Warning: table {tbl_name}: {e}")
    return end_row

# Status table
lkp_table(ws_lkp, 1, 1, "CHECKLIST STATUS",
          ["Status", "Bucket"],
          [["Not Started","Active"],["In Progress","Active"],["Waiting","Active"],
           ["Complete","Complete"],["N/A","N/A"]],
          "tblStatus")

# Loan type
lkp_table(ws_lkp, 4, 1, "LOAN TYPE",
          ["LoanType"],
          [[lt] for lt in LOAN_TYPES],
          "tblLoanType")

# EM Hold
lkp_table(ws_lkp, 6, 1, "EARNEST MONEY HOLDER",
          ["EMHold"],
          [[v] for v in EM_HOLD],
          "tblEMHold")

# HW Pay
lkp_table(ws_lkp, 8, 1, "HOME WARRANTY PAYER",
          ["HWPay"],
          [[v] for v in HW_PAY],
          "tblHWPay")

# Yes/No
lkp_table(ws_lkp, 10, 1, "YES / NO",
          ["YN"],
          [["Yes"],["No"]],
          "tblYN")

# Trans Status
lkp_table(ws_lkp, 12, 1, "TRANSACTION STATUS",
          ["TransStatus"],
          [["Active"],["Closed"],["BOMB"]],
          "tblTransStatus")

# Flat fee constant
ws_lkp.cell(15, 1, "FlatFee").font = Font(bold=True, size=10, name="Calibri")
ws_lkp.cell(15, 2, "Amount")
ws_lkp.cell(16, 1, "Flat Fee Amount")
ws_lkp.cell(16, 2, 399)
ws_lkp.cell(16, 2).number_format = '"$"#,##0.00'

print("Sheet 2 (Lookups) built.")

# ─────────────────────────────────────────────────────────────────────────────
# SHEET 3: AGENTS
# ─────────────────────────────────────────────────────────────────────────────
ws_agent = wb.create_sheet("👤 Agents")
ws_agent.sheet_view.showGridLines = False
for col, width in [("A",6),("B",4),("C",30),("D",25),("E",3)]:
    ws_agent.column_dimensions[col].width = width

# Title
ws_agent.merge_cells("A1:D1")
c = ws_agent["A1"]
c.value = "👤  AGENT DIRECTORY  —  Add new agents at the bottom"
c.fill = fill(C_NAVY)
c.font = Font(bold=True, size=14, color=C_WHITE, name="Calibri")
c.alignment = Alignment(horizontal="center", vertical="center")
ws_agent.row_dimensions[1].height = 32

ws_agent.merge_cells("A2:D2")
c = ws_agent["A2"]
c.value = "⚠️  DO NOT delete rows or rearrange columns.  Add new agents in the YELLOW rows below."
c.fill = fill("FFF2CC")
c.font = Font(size=10, color="7F6000", name="Calibri", bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_agent.row_dimensions[2].height = 18

# Headers
for col, label in [(1,"#"),(2,"AgentID"),(3,"Agent Name"),(4,"Office")]:
    c = ws_agent.cell(3, col, label)
    c.fill = fill(C_HEADER_ROW)
    c.font = Font(bold=True, size=11, color=C_WHITE, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center")

ws_agent.row_dimensions[3].height = 20

# Agent data rows
for i, (name, office) in enumerate(AGENTS):
    r = 4 + i
    # Row number
    c_num = ws_agent.cell(r, 1, i + 1)
    c_num.fill = fill(C_LOCK_GRAY)
    c_num.font = Font(size=10, color="595959", name="Calibri")
    c_num.alignment = Alignment(horizontal="center")
    c_num.protection = openpyxl.styles.Protection(locked=True)
    
    # AgentID (calculated from name)
    initials = name[0] + name.split()[-1][0] if len(name.split()) > 1 else name[:2]
    c_id = ws_agent.cell(r, 2, initials.upper())
    c_id.fill = fill(C_LOCK_GRAY)
    c_id.font = Font(size=10, color="595959", name="Calibri", italic=True)
    c_id.alignment = Alignment(horizontal="center")
    c_id.protection = openpyxl.styles.Protection(locked=True)
    
    # Name (input)
    cn = ws_agent.cell(r, 3, name)
    cn.fill = fill(C_INPUT_YEL) if i >= len(AGENTS) - 2 else PatternFill()
    cn.font = Font(size=11, name="Calibri")
    cn.protection = openpyxl.styles.Protection(locked=False)
    
    # Office (input)
    co = ws_agent.cell(r, 4, office)
    co.fill = fill(C_INPUT_YEL) if i >= len(AGENTS) - 2 else PatternFill()
    co.font = Font(size=11, name="Calibri")
    co.protection = openpyxl.styles.Protection(locked=False)
    
    ws_agent.row_dimensions[r].height = 18
    
    # Alt row color
    if i % 2 == 0:
        for col in [3, 4]:
            if i < len(AGENTS) - 2:
                ws_agent.cell(r, col).fill = fill(C_ROW_ALT)

# Add 5 blank input rows for new agents
for i in range(5):
    r = 4 + len(AGENTS) + i
    for col in [3, 4]:
        c = ws_agent.cell(r, col)
        c.fill = fill(C_INPUT_YEL)
        c.protection = openpyxl.styles.Protection(locked=False)
        s = Side(style="thin", color=C_INPUT_BDR)
        c.border = Border(left=s, right=s, top=s, bottom=s)
    ws_agent.cell(r, 1, f"{len(AGENTS)+i+1}").fill = fill(C_LOCK_GRAY)
    ws_agent.row_dimensions[r].height = 18

# Create tblAgent table
last_agent_row = 4 + len(AGENTS) - 1
try:
    tbl_agent = Table(displayName="tblAgent", ref=f"C3:D{last_agent_row}")
    tbl_agent.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_agent.add_table(tbl_agent)
except Exception as e:
    print(f"  Agent table warning: {e}")

print("Sheet 3 (Agents) built.")


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 4: TRANSACTIONS (Main Data Table)
# ─────────────────────────────────────────────────────────────────────────────
ws_txn = wb.create_sheet("🗄️ Transactions")
ws_txn.sheet_view.showGridLines = False
ws_txn.freeze_panes = "D4"   # freeze ID + cols, keep headers visible

# Column layout:
# A=TxnID(lock), B=CreatedDt(lock), C=Address(input), D=SignNum(in), E=ListPrice(in),
# F=SalePrice(in), G=OfferAcceptDt(in), H=ClosingDt(in), I=LoanType(in), J=Concessions(in),
# K=ListAgent(in), L=SellAgent(in), M=CommPct(in), N=ListingCommPct(in), O=CommFlatAmt(in),
# P=EMAmt(in), Q=EMHold(in), R=HomeWarranty(in), S=HSA(in), T=HWPay(in), U=FlatFee(in),
# V=TitleCo(in), W=TitleContact(in), X=Notes(in), Y=TransStatus(in)
# Z=ListAgentOffice(calc), AA=SellAgentOffice(calc), AB=HHListing(calc), AC=HHSale(calc),
# AD=SaleCommPct(calc), AE=TotalComm(calc), AF=FlatFeeAmt(calc),
# AG=ListingSideAmt(calc), AH=SaleSideAmt(calc), AI=HHRevenue(calc), AJ=CommFlag(calc)

COL_WIDTHS = {
    "A": 11, "B": 12, "C": 38, "D": 10, "E": 13, "F": 13, "G": 14, "H": 14,
    "I": 12, "J": 13, "K": 24, "L": 24, "M": 11, "N": 14, "O": 13, "P": 11,
    "Q": 10, "R": 13, "S": 8, "T": 10, "U": 9, "V": 22, "W": 20, "X": 28,
    "Y": 10, "Z": 16, "AA": 16, "AB": 10, "AC": 9, "AD": 12, "AE": 14,
    "AF": 12, "AG": 14, "AH": 14, "AI": 12, "AJ": 28,
}
for col, w in COL_WIDTHS.items():
    ws_txn.column_dimensions[col].width = w

# --- Row 1: Sheet title ---
ws_txn.merge_cells("A1:AJ1")
c = ws_txn["A1"]
c.value = "🗄️  HH REALTY — TRANSACTION REGISTER  ·  All active, closed and BOMB transactions"
c.fill = fill(C_NAVY)
c.font = Font(bold=True, size=14, color=C_WHITE, name="Calibri")
c.alignment = Alignment(horizontal="center", vertical="center")
ws_txn.row_dimensions[1].height = 32

# --- Row 2: Color legend bar ---
ws_txn.merge_cells("A2:Y2")
c = ws_txn["A2"]
c.value = "🟡 YELLOW = Type here     🔵 BLUE = Auto-calculated (do not edit)     ⚪ GREY = Locked system field"
c.fill = fill("E9F0FB")
c.font = Font(size=10, color=C_NAVY, name="Calibri", italic=True)
c.alignment = Alignment(horizontal="center", vertical="center")

ws_txn.merge_cells("Z2:AJ2")
c = ws_txn["Z2"]
c.value = "◀  CALCULATED COLUMNS  —  Do Not Edit"
c.fill = fill(C_CALC_BLUE)
c.font = Font(size=10, color=C_DARK_BLUE, name="Calibri", italic=True, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_txn.row_dimensions[2].height = 18

# --- Row 3: Section headers ---
INPUT_SECTIONS = [
    (1,2,"SYSTEM",C_LOCK_GRAY,"595959"),
    (3,25,"◀  INPUT FIELDS — Type in the yellow cells  ▶", C_DARK_BLUE, C_WHITE),
    (26,36,"◀  AUTO-CALCULATED — Do not edit  ▶", C_TEAL, C_WHITE),
]
for s, e, label, bg, fg in INPUT_SECTIONS:
    if e > s:
        ws_txn.merge_cells(f"{get_column_letter(s)}3:{get_column_letter(e)}3")
    c = ws_txn.cell(3, s, label)
    c.fill = fill(bg)
    c.font = Font(bold=True, size=10, color=fg, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center")
ws_txn.row_dimensions[3].height = 16

# --- Row 4: Column headers ---
HEADERS = [
    # (col, label, is_input)
    (1,  "Txn ID",           False),
    (2,  "Created",          False),
    (3,  "★ Address",         True),
    (4,  "Sign #",           True),
    (5,  "List Price",       True),
    (6,  "Sale Price",       True),
    (7,  "Offer Accepted",   True),
    (8,  "★ Closing Date",    True),
    (9,  "★ Loan Type",       True),
    (10, "Concessions ($)",  True),
    (11, "★ List Agent",      True),
    (12, "★ Sell Agent",      True),
    (13, "Comm %",           True),
    (14, "List Side %",      True),
    (15, "Flat $ Amount",    True),
    (16, "EM Amount",        True),
    (17, "EM Holder",        True),
    (18, "Home Warranty",    True),
    (19, "HSA",              True),
    (20, "HW Payer",         True),
    (21, "Flat Fee?",        True),
    (22, "Title Co.",        True),
    (23, "Title Contact",    True),
    (24, "Notes",            True),
    (25, "Status",           True),
    (26, "List Office",      False),
    (27, "Sell Office",      False),
    (28, "HH Listing?",      False),
    (29, "HH Sale?",         False),
    (30, "Sale Comm %",      False),
    (31, "Total Comm",       False),
    (32, "Flat Fee Amt",     False),
    (33, "Listing Side",     False),
    (34, "Sale Side",        False),
    (35, "HH Revenue",       False),
    (36, "Comm Check",       False),
]

for col, label, is_input in HEADERS:
    c = ws_txn.cell(4, col, label)
    bg = C_HEADER_ROW if is_input else C_TEAL
    c.fill = fill(bg)
    c.font = Font(bold=True, size=10, color=C_WHITE, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.protection = openpyxl.styles.Protection(locked=True)
ws_txn.row_dimensions[4].height = 32

# ─── DATA ROWS ───────────────────────────────────────────────────────────────
DATA_START = 5
FLAT_FEE_AMOUNT = 399

def txn_id(idx):
    """Return stable Transaction ID string."""
    t = TRANSACTIONS[idx]
    yr = t[5].year if t[5] else 2026   # ClosingDt year
    return f"{yr}-{idx+1:03d}"

for i, t in enumerate(TRANSACTIONS):
    (address, sign_num, list_price, sale_price, offer_dt, close_dt,
     loan_type, concessions, list_agent, sell_agent, comm_pct, listing_comm_pct,
     comm_flat, em_amt, em_hold, home_warranty, hsa, hw_pay, flat_fee,
     title_co, title_contact, notes, trans_status, created_dt) = t

    r = DATA_START + i
    bg_alt = C_ROW_ALT if i % 2 == 0 else C_WHITE

    # Locked cells
    def lk(col, val, fmt=None):
        c = ws_txn.cell(r, col, val)
        c.fill = fill(C_LOCK_GRAY)
        c.font = Font(size=10, color="595959", name="Calibri", italic=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        if fmt: c.number_format = fmt
        c.protection = openpyxl.styles.Protection(locked=True)
        return c

    # Input cells
    def inp(col, val, fmt=None, h="left"):
        c = ws_txn.cell(r, col, val)
        c.fill = fill(C_INPUT_YEL)
        c.font = Font(size=10, color="000000", name="Calibri")
        c.alignment = Alignment(horizontal=h, vertical="center")
        s = Side(style="thin", color=C_INPUT_BDR)
        c.border = Border(left=s, right=s, top=s, bottom=s)
        if fmt: c.number_format = fmt
        c.protection = openpyxl.styles.Protection(locked=False)
        return c

    # Calc cells
    def cal(col, val, fmt=None, h="center"):
        c = ws_txn.cell(r, col, val)
        c.fill = fill(C_CALC_BLUE)
        c.font = Font(size=10, color="000000", name="Calibri")
        c.alignment = Alignment(horizontal=h, vertical="center")
        s = Side(style="thin", color=C_CALC_BDR)
        c.border = Border(left=s, right=s, top=s, bottom=s)
        if fmt: c.number_format = fmt
        c.protection = openpyxl.styles.Protection(locked=True)
        return c

    # Col A: TransactionID (locked, formula-computed but set as value for stability)
    lk(1, txn_id(i))

    # Col B: CreatedDt
    lk(2, created_dt, "mm/dd/yyyy")

    # Cols 3-25: Input fields
    inp(3, address)
    inp(4, sign_num, "#,##0", "center")
    inp(5, list_price, '"$"#,##0', "right")
    inp(6, sale_price, '"$"#,##0', "right")
    if offer_dt:
        inp(7, offer_dt.date() if hasattr(offer_dt,'date') else offer_dt, "mm/dd/yyyy", "center")
    else:
        inp(7, None, "mm/dd/yyyy", "center")
    if close_dt:
        inp(8, close_dt.date() if hasattr(close_dt,'date') else close_dt, "mm/dd/yyyy", "center")
    else:
        inp(8, None, "mm/dd/yyyy", "center")
    inp(9, loan_type, h="center")
    inp(10, concessions if isinstance(concessions, (int, float)) and concessions > 1 else
           (concessions * sale_price if isinstance(concessions, float) and 0 < concessions < 1 and sale_price else concessions),
        '"$"#,##0', "right")
    inp(11, list_agent)
    inp(12, sell_agent)
    inp(13, comm_pct, "0.00%", "center")
    inp(14, listing_comm_pct, "0.00%", "center")
    inp(15, comm_flat if comm_flat else None, '"$"#,##0.00', "right")
    inp(16, em_amt if em_amt else None, '"$"#,##0', "right")
    inp(17, em_hold, h="center")
    inp(18, home_warranty, h="center")
    inp(19, hsa, h="center")
    inp(20, hw_pay, h="center")
    inp(21, flat_fee, h="center")
    inp(22, title_co)
    inp(23, title_contact)
    inp(24, notes)
    inp(25, trans_status if trans_status else "Active", h="center")

    # Cols 26-36: Calculated
    # Col 26: List Agent Office  (XLOOKUP from tblAgent)
    row_ref = r
    cal(26, f'=IFERROR(VLOOKUP(K{row_ref},tblAgent[#All],2,FALSE),"")', h="left")
    # Col 27: Sell Agent Office
    cal(27, f'=IFERROR(VLOOKUP(L{row_ref},tblAgent[#All],2,FALSE),"")', h="left")
    # Col 28: HH Listing?
    cal(28, f'=IF(Z{row_ref}="HH","Yes","No")', h="center")
    # Col 29: HH Sale?
    cal(29, f'=IF(AA{row_ref}="HH","Yes","No")', h="center")
    # Col 30: Sale Comm % (total - listing side)
    cal(30, f'=IFERROR(IF(M{row_ref}>0,M{row_ref}-N{row_ref},0),0)', "0.00%", "center")
    # Col 31: Total Commission
    cal(31, f'=IFERROR(IF(M{row_ref}>0,F{row_ref}*M{row_ref},IF(O{row_ref}>0,O{row_ref},0)),0)',
        '"$"#,##0.00', "right")
    # Col 32: Flat Fee Amount
    cal(32, f'=IF(U{row_ref}="Yes",{FLAT_FEE_AMOUNT},0)', '"$"#,##0', "right")
    # Col 33: Listing Side Amount
    cal(33, f'=IFERROR(IF(N{row_ref}>0,F{row_ref}*N{row_ref},IF(AND(AB{row_ref}="Yes",O{row_ref}>0),O{row_ref}/2,0)),0)',
        '"$"#,##0.00', "right")
    # Col 34: Sale Side Amount
    cal(34, f'=IFERROR(IF(AD{row_ref}>0,F{row_ref}*AD{row_ref},IF(AND(AC{row_ref}="Yes",O{row_ref}>0),O{row_ref}/2,0)),0)',
        '"$"#,##0.00', "right")
    # Col 35: HH Revenue
    cal(35, f'=IF(AB{row_ref}="Yes",AG{row_ref},0)+IF(AC{row_ref}="Yes",AH{row_ref},0)+AF{row_ref}',
        '"$"#,##0.00', "right")
    # Col 36: Comm Check flag
    cal(36, f'=IF(AE{row_ref}=0,"",IF(ABS(AE{row_ref}-(AG{row_ref}+AH{row_ref}))<1,"✓ OK","⚠️ CHECK: Total="&TEXT(AE{row_ref},"$#,##0")&" vs Sides="&TEXT(AG{row_ref}+AH{row_ref},"$#,##0")))',
        h="left")

    ws_txn.row_dimensions[r].height = 18

print(f"  Wrote {len(TRANSACTIONS)} transaction rows.")

# ─── Data Validations ────────────────────────────────────────────────────────
def add_dv(ws, formula, col_letter, start_row, end_row=200, error_msg="Invalid value. Use the dropdown."):
    dv = DataValidation(
        type="list",
        formula1=formula,
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Invalid Entry",
        error=error_msg,
        showInputMessage=False,
    )
    dv.sqref = f"{col_letter}{start_row}:{col_letter}{end_row}"
    ws.add_data_validation(dv)

DS = DATA_START
add_dv(ws_txn, '=tblLoanType[LoanType]', "I", DS)
add_dv(ws_txn, '=tblAgent[AgentName]',   "K", DS)
add_dv(ws_txn, '=tblAgent[AgentName]',   "L", DS)
add_dv(ws_txn, '=tblEMHold[EMHold]',     "Q", DS)
add_dv(ws_txn, '=tblYN[YN]',             "R", DS)
add_dv(ws_txn, '=tblYN[YN]',             "S", DS)
add_dv(ws_txn, '=tblHWPay[HWPay]',       "T", DS)
add_dv(ws_txn, '=tblYN[YN]',             "U", DS)
add_dv(ws_txn, '=tblTransStatus[TransStatus]', "Y", DS)

# Date validation
dv_date = DataValidation(type="date", operator="greaterThan",
                         formula1="DATE(2000,1,1)", allow_blank=True,
                         showErrorMessage=True, errorTitle="Invalid Date",
                         error="Please enter a valid date (mm/dd/yyyy).")
dv_date.sqref = f"G{DS}:H200"
ws_txn.add_data_validation(dv_date)

# Positive number validation for prices
dv_price = DataValidation(type="decimal", operator="greaterThanOrEqual",
                          formula1="0", allow_blank=True,
                          showErrorMessage=True, errorTitle="Invalid Amount",
                          error="Please enter a number greater than or equal to 0.")
dv_price.sqref = f"E{DS}:F200"
ws_txn.add_data_validation(dv_price)

# ─── Conditional Formatting ──────────────────────────────────────────────────
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.styles.differential import DifferentialStyle

# Highlight BOMB rows red
ws_txn.conditional_formatting.add(
    f"C{DS}:Y200",
    FormulaRule(formula=[f'=$Y{DS}="BOMB"'],
                fill=PatternFill(fgColor="FFDBD6", fill_type="solid"),
                font=Font(color="C00000", name="Calibri"), stopIfTrue=True)
)
# Highlight Closed rows light green
ws_txn.conditional_formatting.add(
    f"C{DS}:Y200",
    FormulaRule(formula=[f'=$Y{DS}="Closed"'],
                fill=PatternFill(fgColor="E2EFDA", fill_type="solid"),
                font=Font(color="375623", name="Calibri"), stopIfTrue=True)
)
# Highlight Comm Check issues
ws_txn.conditional_formatting.add(
    f"AJ{DS}:AJ200",
    FormulaRule(formula=[f'=NOT(ISERROR(FIND("CHECK",AJ{DS})))'],
                fill=PatternFill(fgColor="FFDBD6", fill_type="solid"),
                font=Font(color="C00000", bold=True, name="Calibri"))
)
# Highlight overdue closings (closing date < today and not closed/BOMB)
ws_txn.conditional_formatting.add(
    f"H{DS}:H200",
    FormulaRule(formula=[f'=AND(H{DS}<TODAY(),H{DS}<>"",Y{DS}="Active")'],
                fill=PatternFill(fgColor="FFF2CC", fill_type="solid"),
                font=Font(color="7F6000", name="Calibri"))
)

protect_sheet(ws_txn)
print("Sheet 4 (Transactions) built.")


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 5: CHECKLIST
# ─────────────────────────────────────────────────────────────────────────────
ws_cl = wb.create_sheet("✅ Checklist")
ws_cl.sheet_view.showGridLines = False
ws_cl.freeze_panes = "C4"

# Wider first cols
ws_cl.column_dimensions["A"].width = 11
ws_cl.column_dimensions["B"].width = 36

# Checklist columns A-B are ID/address, then one col per step
# Title
ws_cl.merge_cells(f"A1:{get_column_letter(2 + len(CHECKLIST_COLS))}1")
c = ws_cl["A1"]
c.value = "✅  TRANSACTION CHECKLIST  —  Track paperwork and action items for each transaction"
c.fill = fill(C_NAVY)
c.font = Font(bold=True, size=14, color=C_WHITE, name="Calibri")
c.alignment = Alignment(horizontal="center", vertical="center")
ws_cl.row_dimensions[1].height = 32

# Legend
ws_cl.merge_cells(f"A2:{get_column_letter(2 + len(CHECKLIST_COLS))}2")
c = ws_cl["A2"]
c.value = "Use the dropdown in each cell: Not Started | In Progress | Waiting | Complete | N/A"
c.fill = fill("E9F0FB")
c.font = Font(size=10, color=C_NAVY, name="Calibri", italic=True)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_cl.row_dimensions[2].height = 16

# Category row (row 3) - group checklist items by category
CATEGORIES = ["Check Paperwork", "Action Items", "EM / Commission / HW", "Closed Sale", "Fallen Sale (BOMB)"]
CAT_COLS = {}  # cat -> list of col indices
for ci, col_name in enumerate(CHECKLIST_COLS):
    cat = CHECKLIST_CATEGORY.get(col_name, "Other")
    if cat not in CAT_COLS:
        CAT_COLS[cat] = []
    CAT_COLS[cat].append(3 + ci)

CAT_COLORS = {
    "Check Paperwork":    "2F5496",
    "Action Items":       "375623",
    "EM / Commission / HW": "7F4000",
    "Closed Sale":        "17375E",
    "Fallen Sale (BOMB)": "843C0C",
}
prev_cat = None
for cat in CATEGORIES:
    cols = CAT_COLS.get(cat, [])
    if not cols:
        continue
    start_col = cols[0]
    end_col = cols[-1]
    if end_col > start_col:
        ws_cl.merge_cells(f"{get_column_letter(start_col)}3:{get_column_letter(end_col)}3")
    c = ws_cl.cell(3, start_col, cat.upper())
    c.fill = fill(CAT_COLORS.get(cat, C_NAVY))
    c.font = Font(bold=True, size=9, color=C_WHITE, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center")
ws_cl.row_dimensions[3].height = 16

# Row 3 for col A and B
for col_idx in [1, 2]:
    c = ws_cl.cell(3, col_idx)
    c.fill = fill(C_NAVY)
    c.font = Font(bold=True, size=9, color=C_WHITE, name="Calibri")

# Header row (row 4)
for col_idx, label in [(1, "Txn ID"), (2, "Address")]:
    c = ws_cl.cell(4, col_idx, label)
    c.fill = fill(C_HEADER_ROW)
    c.font = Font(bold=True, size=9, color=C_WHITE, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.protection = openpyxl.styles.Protection(locked=True)

for ci, col_name in enumerate(CHECKLIST_COLS):
    col_idx = 3 + ci
    label = CHECKLIST_LABELS.get(col_name, col_name)
    c = ws_cl.cell(4, col_idx, label)
    cat = CHECKLIST_CATEGORY.get(col_name, "Other")
    c.fill = fill(CAT_COLORS.get(cat, C_HEADER_ROW))
    c.font = Font(bold=True, size=8, color=C_WHITE, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.protection = openpyxl.styles.Protection(locked=True)
    ws_cl.column_dimensions[get_column_letter(col_idx)].width = 13

ws_cl.row_dimensions[4].height = 48

# Add data validation for all checklist cells
dv_status = DataValidation(
    type="list",
    formula1='=tblStatus[Status]',
    allow_blank=True,
    showErrorMessage=True,
    errorTitle="Invalid Status",
    error="Choose: Not Started, In Progress, Waiting, Complete, or N/A",
    showInputMessage=True,
    promptTitle="Select Status",
    prompt="Choose the status for this checklist item."
)
dv_status.sqref = f"C5:{get_column_letter(2+len(CHECKLIST_COLS))}200"
ws_cl.add_data_validation(dv_status)

# Data rows
for i, t in enumerate(TRANSACTIONS):
    r = 5 + i
    tid = txn_id(i)

    # Locked: TxnID
    c_id = ws_cl.cell(r, 1, tid)
    c_id.fill = fill(C_LOCK_GRAY)
    c_id.font = Font(size=10, color="595959", name="Calibri", italic=True)
    c_id.alignment = Alignment(horizontal="center")
    c_id.protection = openpyxl.styles.Protection(locked=True)

    # Locked: Address (lookup from transactions sheet)
    c_addr = ws_cl.cell(r, 2)
    # Use VLOOKUP to get address from transactions sheet by TxnID
    c_addr.value = f'=IFERROR(VLOOKUP(A{r},\'🗄️ Transactions\'!$A:$C,3,FALSE),"")'
    c_addr.fill = fill(C_LOCK_GRAY)
    c_addr.font = Font(size=10, color="595959", name="Calibri", italic=True)
    c_addr.alignment = Alignment(horizontal="left", vertical="center")
    c_addr.protection = openpyxl.styles.Protection(locked=True)

    # Input cells for each checklist item
    existing = CHECKLIST_DATA.get(tid, {})
    for ci, col_name in enumerate(CHECKLIST_COLS):
        col_idx = 3 + ci
        val = existing.get(col_name, "Not Started")
        c = ws_cl.cell(r, col_idx, val)
        c.fill = fill(C_INPUT_YEL)
        c.font = Font(size=9, color="000000", name="Calibri")
        c.alignment = Alignment(horizontal="center", vertical="center")
        s = Side(style="thin", color=C_INPUT_BDR)
        c.border = Border(left=s, right=s, top=s, bottom=s)
        c.protection = openpyxl.styles.Protection(locked=False)

    ws_cl.row_dimensions[r].height = 18

# Conditional formatting on checklist - Complete = green, BOMB = red
ws_cl.conditional_formatting.add(
    f"C5:{get_column_letter(2+len(CHECKLIST_COLS))}200",
    FormulaRule(formula=['=C5="Complete"'],
                fill=PatternFill(fgColor="E2EFDA", fill_type="solid"),
                font=Font(color="375623", name="Calibri"))
)
ws_cl.conditional_formatting.add(
    f"C5:{get_column_letter(2+len(CHECKLIST_COLS))}200",
    FormulaRule(formula=['=C5="N/A"'],
                fill=PatternFill(fgColor="F2F2F2", fill_type="solid"),
                font=Font(color="808080", name="Calibri"))
)
ws_cl.conditional_formatting.add(
    f"C5:{get_column_letter(2+len(CHECKLIST_COLS))}200",
    FormulaRule(formula=['=C5="In Progress"'],
                fill=PatternFill(fgColor="FFF2CC", fill_type="solid"),
                font=Font(color="7F6000", name="Calibri"))
)
ws_cl.conditional_formatting.add(
    f"C5:{get_column_letter(2+len(CHECKLIST_COLS))}200",
    FormulaRule(formula=['=C5="Waiting"'],
                fill=PatternFill(fgColor="FCE4D6", fill_type="solid"),
                font=Font(color="843C0C", name="Calibri"))
)

protect_sheet(ws_cl)
print("Sheet 5 (Checklist) built.")


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 6: INPUT FORM
# ─────────────────────────────────────────────────────────────────────────────
ws_inp = wb.create_sheet("🏠 Input Form")
ws_inp.sheet_view.showGridLines = False

for col, width in [("A",3),("B",28),("C",32),("D",3),("E",28),("F",32),("G",3)]:
    ws_inp.column_dimensions[col].width = width

# Title
ws_inp.merge_cells("A1:G1")
c = ws_inp["A1"]
c.value = "🏠  NEW TRANSACTION ENTRY FORM  —  Fill in the yellow fields, then add to Transactions sheet"
c.fill = fill(C_NAVY)
c.font = Font(bold=True, size=14, color=C_WHITE, name="Calibri")
c.alignment = Alignment(horizontal="center", vertical="center")
ws_inp.row_dimensions[1].height = 36

ws_inp.merge_cells("A2:G2")
c = ws_inp["A2"]
c.value = "★ = Required field   |   Fill ALL starred fields before saving   |   Blue cells = auto-filled, do not edit"
c.fill = fill("FFF2CC")
c.font = Font(size=11, color="7F4000", name="Calibri", bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_inp.row_dimensions[2].height = 20

# Helper to write a label+input pair
def form_row(ws, row, left_label, left_col=2, right_label=None, right_col=5,
             left_locked=False, right_locked=False, height=24, bg_left=C_INPUT_YEL, bg_right=C_INPUT_YEL):
    ws.row_dimensions[row].height = height
    # Left label
    c = ws.cell(row, left_col, left_label)
    c.fill = fill(C_DARK_BLUE)
    c.font = Font(bold=True, size=11, color=C_WHITE, name="Calibri")
    c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    c.protection = openpyxl.styles.Protection(locked=True)
    # Left input
    c_inp = ws.cell(row, left_col + 1)
    c_inp.fill = fill(C_LOCK_GRAY if left_locked else bg_left)
    s = Side(style="medium", color=C_INPUT_BDR if not left_locked else C_LOCK_BDR)
    c_inp.border = Border(left=s, right=s, top=s, bottom=s)
    c_inp.font = Font(size=12, color="000000", name="Calibri")
    c_inp.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c_inp.protection = openpyxl.styles.Protection(locked=left_locked)
    if right_label:
        # Right label
        c2 = ws.cell(row, right_col, right_label)
        c2.fill = fill(C_DARK_BLUE)
        c2.font = Font(bold=True, size=11, color=C_WHITE, name="Calibri")
        c2.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        c2.protection = openpyxl.styles.Protection(locked=True)
        # Right input
        c2_inp = ws.cell(row, right_col + 1)
        c2_inp.fill = fill(C_LOCK_GRAY if right_locked else bg_right)
        s2 = Side(style="medium", color=C_INPUT_BDR if not right_locked else C_LOCK_BDR)
        c2_inp.border = Border(left=s2, right=s2, top=s2, bottom=s2)
        c2_inp.font = Font(size=12, color="000000", name="Calibri")
        c2_inp.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c2_inp.protection = openpyxl.styles.Protection(locked=right_locked)

def section_bar(ws, row, label, color=C_DARK_BLUE, height=22):
    ws.merge_cells(f"B{row}:F{row}")
    c = ws.cell(row, 2, label)
    c.fill = fill(color)
    c.font = Font(bold=True, size=12, color=C_WHITE, name="Calibri")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    c.protection = openpyxl.styles.Protection(locked=True)
    ws.row_dimensions[row].height = height

# Row 4: Auto ID display
section_bar(ws_inp, 4, "📋  SYSTEM INFORMATION  (auto-filled, do not edit)", C_TEAL, 22)
form_row(ws_inp, 5, "Transaction ID", left_locked=True, right_label="Entry Date", right_locked=True)
ws_inp.cell(5, 3).value = '=TEXT(YEAR(TODAY()),"0000")&"-"&TEXT(COUNTA(\'🗄️ Transactions\'!A:A)-3,"000")'
ws_inp.cell(5, 3).number_format = "@"
ws_inp.cell(5, 6).value = "=TODAY()"
ws_inp.cell(5, 6).number_format = "mm/dd/yyyy"
ws_inp.cell(5, 3).font = Font(size=12, color=C_DARK_BLUE, bold=True, name="Calibri")

# Section: Property
row = 7
section_bar(ws_inp, row,   "🏠  PROPERTY INFORMATION", C_DARK_BLUE, 22)
form_row(ws_inp, row+1, "★  Address (Street, City Zip)",    right_label="★  Closing Date")
ws_inp.cell(row+2, 2).value = "★  List Price ($)"
ws_inp.cell(row+2, 2).fill = fill(C_DARK_BLUE)
ws_inp.cell(row+2, 2).font = Font(bold=True, size=11, color=C_WHITE, name="Calibri")
ws_inp.cell(row+2, 2).alignment = Alignment(horizontal="right", vertical="center", indent=1)
form_row(ws_inp, row+2, "★  List Price ($)",                right_label="Sale Price ($)")
ws_inp.cell(row+2, 3).number_format = '"$"#,##0'
ws_inp.cell(row+2, 6).number_format = '"$"#,##0'
ws_inp.cell(row+1, 6).number_format = "mm/dd/yyyy"
form_row(ws_inp, row+3, "Offer Accepted Date",              right_label="Sign # (if our listing)")
ws_inp.cell(row+3, 3).number_format = "mm/dd/yyyy"
form_row(ws_inp, row+4, "★  Loan Type (dropdown ↓)",         right_label="Concessions ($)")
ws_inp.cell(row+4, 6).number_format = '"$"#,##0'

# Section: Agents
row = 13
section_bar(ws_inp, row,   "👤  AGENT & COMMISSION", C_DARK_BLUE, 22)
form_row(ws_inp, row+1, "★  List Agent (dropdown ↓)",         right_label="★  Sell Agent (dropdown ↓)")
form_row(ws_inp, row+2, "Total Comm % (e.g. 0.06)",          right_label="Listing Side % (e.g. 0.03)")
ws_inp.cell(row+2, 3).number_format = "0.00%"
ws_inp.cell(row+2, 6).number_format = "0.00%"
form_row(ws_inp, row+3, "Flat Comm Amount (if flat fee)",     right_label="Flat Fee Add-On?")
ws_inp.cell(row+3, 3).number_format = '"$"#,##0.00'

# Section: EM / HW
row = 19
section_bar(ws_inp, row,   "💰  EARNEST MONEY & HOME WARRANTY", C_DARK_BLUE, 22)
form_row(ws_inp, row+1, "Earnest Money Amount ($)",          right_label="EM Holder (dropdown ↓)")
ws_inp.cell(row+1, 3).number_format = '"$"#,##0'
form_row(ws_inp, row+2, "Home Warranty? (Yes/No)",           right_label="HW Payer (dropdown ↓)")
form_row(ws_inp, row+3, "HSA? (Yes/No)",                     right_label="Title Company")

# Section: Notes
row = 25
section_bar(ws_inp, row, "📝  ADDITIONAL NOTES", C_DARK_BLUE, 22)
ws_inp.merge_cells(f"B{row+1}:F{row+1}")
c = ws_inp.cell(row+1, 2)
c.fill = fill(C_INPUT_YEL)
s = Side(style="medium", color=C_INPUT_BDR)
c.border = Border(left=s, right=s, top=s, bottom=s)
c.protection = openpyxl.styles.Protection(locked=False)
ws_inp.row_dimensions[row+1].height = 48

# Section: Instructions
row = 28
section_bar(ws_inp, row, "📌  HOW TO ADD THIS TRANSACTION", color="375623", height=22)
ws_inp.merge_cells(f"B{row+1}:F{row+3}")
c = ws_inp.cell(row+1, 2)
c.value = ("1.  Fill in all yellow cells above (★ = required).\n"
           "2.  Go to the '🗄️ Transactions' sheet.\n"
           "3.  Scroll to the first EMPTY row at the bottom.\n"
           "4.  Copy each value from this form into the matching column.\n"
           "5.  Return here and press Ctrl+Z multiple times to clear the form for next use,\n"
           "    OR simply type over the values with the next transaction's information.")
c.fill = fill("E2EFDA")
c.font = Font(size=11, color="375623", name="Calibri")
c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
c.protection = openpyxl.styles.Protection(locked=True)
ws_inp.row_dimensions[row+1].height = 90

# Add dropdowns on Input form
dv_loan_inp = DataValidation(type="list", formula1='=tblLoanType[LoanType]', allow_blank=True)
dv_loan_inp.sqref = "C11"
ws_inp.add_data_validation(dv_loan_inp)

dv_agent_inp_l = DataValidation(type="list", formula1='=tblAgent[AgentName]', allow_blank=True)
dv_agent_inp_l.sqref = "C14"
ws_inp.add_data_validation(dv_agent_inp_l)

dv_agent_inp_s = DataValidation(type="list", formula1='=tblAgent[AgentName]', allow_blank=True)
dv_agent_inp_s.sqref = "F14"
ws_inp.add_data_validation(dv_agent_inp_s)

dv_yn_inp1 = DataValidation(type="list", formula1='=tblYN[YN]', allow_blank=True)
dv_yn_inp1.sqref = "C17"
ws_inp.add_data_validation(dv_yn_inp1)

dv_em_inp = DataValidation(type="list", formula1='=tblEMHold[EMHold]', allow_blank=True)
dv_em_inp.sqref = "F20"
ws_inp.add_data_validation(dv_em_inp)

dv_yn_inp2 = DataValidation(type="list", formula1='=tblYN[YN]', allow_blank=True)
dv_yn_inp2.sqref = "C21"
ws_inp.add_data_validation(dv_yn_inp2)

dv_hw_inp = DataValidation(type="list", formula1='=tblHWPay[HWPay]', allow_blank=True)
dv_hw_inp.sqref = "F21"
ws_inp.add_data_validation(dv_hw_inp)

dv_yn_inp3 = DataValidation(type="list", formula1='=tblYN[YN]', allow_blank=True)
dv_yn_inp3.sqref = "C22"
ws_inp.add_data_validation(dv_yn_inp3)

dv_yn_inp4 = DataValidation(type="list", formula1='=tblYN[YN]', allow_blank=True)
dv_yn_inp4.sqref = "C17"
ws_inp.add_data_validation(dv_yn_inp4)

protect_sheet(ws_inp)
print("Sheet 6 (Input Form) built.")


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 7: DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────
ws_dash = wb.create_sheet("📊 Dashboard")
ws_dash.sheet_view.showGridLines = False
ws_dash.sheet_view.tabSelected = False

for col, width in [("A",3),("B",22),("C",22),("D",22),("E",22),("F",3)]:
    ws_dash.column_dimensions[col].width = width

# ─── Title ───────────────────────────────────────────────────────────────────
ws_dash.merge_cells("A1:F1")
c = ws_dash["A1"]
c.value = "📊  HH REALTY — TRANSACTION DASHBOARD"
c.fill = fill(C_NAVY)
c.font = Font(bold=True, size=18, color=C_WHITE, name="Calibri")
c.alignment = Alignment(horizontal="center", vertical="center")
ws_dash.row_dimensions[1].height = 42

ws_dash.merge_cells("A2:F2")
c = ws_dash["A2"]
c.value = '=CONCATENATE("Last Updated: ",TEXT(NOW(),"mmmm d, yyyy  h:MM AM/PM"))'
c.fill = fill(C_DARK_BLUE)
c.font = Font(size=11, color=C_WHITE, name="Calibri", italic=True)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_dash.row_dimensions[2].height = 20

# ─── KPI Cards (row 4-9) ─────────────────────────────────────────────────────
def kpi_card(ws, row, col, title, formula, fmt=None, good_is_high=True, note=None):
    """Write a 2-row KPI card."""
    # Title
    c_title = ws.cell(row, col, title)
    c_title.fill = fill(C_DARK_BLUE)
    c_title.font = Font(bold=True, size=10, color=C_WHITE, name="Calibri")
    c_title.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c_title.protection = openpyxl.styles.Protection(locked=True)
    ws.row_dimensions[row].height = 28
    # Value
    c_val = ws.cell(row+1, col, formula)
    c_val.fill = fill(C_GREEN_KPI)
    c_val.font = Font(bold=True, size=22, color=C_NAVY, name="Calibri")
    c_val.alignment = Alignment(horizontal="center", vertical="center")
    if fmt: c_val.number_format = fmt
    c_val.protection = openpyxl.styles.Protection(locked=True)
    ws.row_dimensions[row+1].height = 42
    if note:
        c_note = ws.cell(row+2, col, note)
        c_note.fill = fill("F5F9FF")
        c_note.font = Font(size=9, color="595959", name="Calibri", italic=True)
        c_note.alignment = Alignment(horizontal="center", vertical="center")
        c_note.protection = openpyxl.styles.Protection(locked=True)
        ws.row_dimensions[row+2].height = 15
    return c_val

ws_dash.merge_cells("A3:F3")
c = ws_dash["A3"]
c.value = "KEY PERFORMANCE INDICATORS"
c.fill = fill(C_NAVY)
c.font = Font(bold=True, size=12, color=C_WHITE, name="Calibri")
c.alignment = Alignment(horizontal="center", vertical="center")
ws_dash.row_dimensions[3].height = 22

# KPI Row 1: 4 KPIs
kpi_card(ws_dash, 4, 2, "📈 ACTIVE TRANSACTIONS",
         "=COUNTIF('🗄️ Transactions'!Y5:Y200,\"Active\")")
kpi_card(ws_dash, 4, 3, "💰 TOTAL HH REVENUE",
         "=SUMIF('🗄️ Transactions'!Y5:Y200,\"Active\",'🗄️ Transactions'!AI5:AI200)",
         '"$"#,##0.00')
kpi_card(ws_dash, 4, 4, "📅 CLOSING ≤ 7 DAYS",
         "=COUNTIFS('🗄️ Transactions'!H5:H200,\">=\"&TODAY(),'🗄️ Transactions'!H5:H200,\"<=\"&(TODAY()+7),'🗄️ Transactions'!Y5:Y200,\"Active\")")
kpi_card(ws_dash, 4, 5, "⚠️ COMMISSION ISSUES",
         "=COUNTIF('🗄️ Transactions'!AJ5:AJ200,\"*CHECK*\")")

# KPI Row 2
kpi_card(ws_dash, 8, 2, "✅ CLOSED THIS YEAR",
         "=COUNTIF('🗄️ Transactions'!Y5:Y200,\"Closed\")")
kpi_card(ws_dash, 8, 3, "💼 CLOSINGS THIS MONTH",
         "=COUNTIFS('🗄️ Transactions'!H5:H200,\">=\"&EOMONTH(TODAY(),-1)+1,'🗄️ Transactions'!H5:H200,\"<=\"&EOMONTH(TODAY(),0),'🗄️ Transactions'!Y5:Y200,\"Active\")")
kpi_card(ws_dash, 8, 4, "🏦 MISSING TITLE CO.",
         "=COUNTIFS('🗄️ Transactions'!V5:V200,\"\",'🗄️ Transactions'!Y5:Y200,\"Active\")")
kpi_card(ws_dash, 8, 5, "💣 BOMB TRANSACTIONS",
         "=COUNTIF('🗄️ Transactions'!Y5:Y200,\"BOMB\")")

# ─── Closing Soon Table (row 12+) ────────────────────────────────────────────
ws_dash.row_dimensions[11].height = 8

ws_dash.merge_cells("B12:F12")
c = ws_dash["B12"]
c.value = "📅  CLOSING WITHIN 14 DAYS"
c.fill = fill(C_DARK_BLUE)
c.font = Font(bold=True, size=12, color=C_WHITE, name="Calibri")
c.alignment = Alignment(horizontal="left", vertical="center", indent=2)
ws_dash.row_dimensions[12].height = 24

for col_idx, label in [(2,"Txn ID"),(3,"Address"),(4,"Closing Date"),(5,"Days Left")]:
    c = ws_dash.cell(13, col_idx, label)
    c.fill = fill(C_HEADER_ROW)
    c.font = Font(bold=True, size=10, color=C_WHITE, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center")
ws_dash.row_dimensions[13].height = 20

# 10 rows for upcoming closings - use IFERROR with sorted FILTER approach
# Since we can't use dynamic FILTER (older Excel), use static SMALL/INDEX approach with helper
# For compatibility, write formula rows that look up into the transactions
for i in range(10):
    r = 14 + i
    # These use a simpler approach: sort transactions to show earliest closing first from active ones
    # Using INDEX/MATCH with SMALL to find the i-th closest closing date
    n = i + 1
    ws_dash.cell(r, 2,
        f'=IFERROR(INDEX(\'🗄️ Transactions\'!$A$5:$A$200,MATCH(SMALL(IF(\'🗄️ Transactions\'!$Y$5:$Y$200="Active",IF(\'🗄️ Transactions\'!$H$5:$H$200>=TODAY()-1,\'🗄️ Transactions\'!$H$5:$H$200,9999999)),{n}),\'🗄️ Transactions\'!$H$5:$H$200,0)),"—")'
    ).font = Font(size=10, name="Calibri")
    ws_dash.cell(r, 3,
        f'=IFERROR(INDEX(\'🗄️ Transactions\'!$C$5:$C$200,MATCH(SMALL(IF(\'🗄️ Transactions\'!$Y$5:$Y$200="Active",IF(\'🗄️ Transactions\'!$H$5:$H$200>=TODAY()-1,\'🗄️ Transactions\'!$H$5:$H$200,9999999)),{n}),\'🗄️ Transactions\'!$H$5:$H$200,0)),"—")'
    ).font = Font(size=10, name="Calibri")
    c_date = ws_dash.cell(r, 4,
        f'=IFERROR(SMALL(IF(\'🗄️ Transactions\'!$Y$5:$Y$200="Active",IF(\'🗄️ Transactions\'!$H$5:$H$200>=TODAY()-1,\'🗄️ Transactions\'!$H$5:$H$200,9999999)),{n}),"—")'
    )
    c_date.number_format = "mm/dd/yyyy"
    c_date.font = Font(size=10, name="Calibri")
    c_days = ws_dash.cell(r, 5,
        f'=IFERROR(D{r}-TODAY(),"—")'
    )
    c_days.number_format = "0"
    c_days.font = Font(size=10, name="Calibri")
    c_days.alignment = Alignment(horizontal="center")
    # Alt row
    bg = C_ROW_ALT if i % 2 == 0 else C_WHITE
    for col_idx in [2,3,4,5]:
        ws_dash.cell(r, col_idx).fill = fill(bg)
    ws_dash.row_dimensions[r].height = 18

# ─── EM Not Received (row 25+) ────────────────────────────────────────────────
ws_dash.row_dimensions[24].height = 8

ws_dash.merge_cells("B25:F25")
c = ws_dash["B25"]
c.value = "🚨  EARNEST MONEY NOT RECEIVED  (EM Holder = HH, but EM Amount > 0)"
c.fill = fill("843C0C")
c.font = Font(bold=True, size=12, color=C_WHITE, name="Calibri")
c.alignment = Alignment(horizontal="left", vertical="center", indent=2)
ws_dash.row_dimensions[25].height = 24

for col_idx, label in [(2,"Txn ID"),(3,"Address"),(4,"EM Amount"),(5,"Closing Date")]:
    c = ws_dash.cell(26, col_idx, label)
    c.fill = fill(C_HEADER_ROW)
    c.font = Font(bold=True, size=10, color=C_WHITE, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center")
ws_dash.row_dimensions[26].height = 20

for i in range(8):
    r = 27 + i
    n = i + 1
    # Transactions where EMHold=HH and EMAmt>0 and status=Active
    ws_dash.cell(r, 2,
        f'=IFERROR(INDEX(\'🗄️ Transactions\'!$A$5:$A$200,SMALL(IF((\'🗄️ Transactions\'!$Q$5:$Q$200="HH")*(\'🗄️ Transactions\'!$P$5:$P$200>0)*(\'🗄️ Transactions\'!$Y$5:$Y$200="Active"),ROW(\'🗄️ Transactions\'!$A$5:$A$200)-ROW(\'🗄️ Transactions\'!$A$5)+1),{n})),"—")'
    ).font = Font(size=10, name="Calibri")
    ws_dash.cell(r, 3,
        f'=IFERROR(INDEX(\'🗄️ Transactions\'!$C$5:$C$200,SMALL(IF((\'🗄️ Transactions\'!$Q$5:$Q$200="HH")*(\'🗄️ Transactions\'!$P$5:$P$200>0)*(\'🗄️ Transactions\'!$Y$5:$Y$200="Active"),ROW(\'🗄️ Transactions\'!$A$5:$A$200)-ROW(\'🗄️ Transactions\'!$A$5)+1),{n})),"—")'
    ).font = Font(size=10, name="Calibri")
    c_em = ws_dash.cell(r, 4,
        f'=IFERROR(INDEX(\'🗄️ Transactions\'!$P$5:$P$200,SMALL(IF((\'🗄️ Transactions\'!$Q$5:$Q$200="HH")*(\'🗄️ Transactions\'!$P$5:$P$200>0)*(\'🗄️ Transactions\'!$Y$5:$Y$200="Active"),ROW(\'🗄️ Transactions\'!$A$5:$A$200)-ROW(\'🗄️ Transactions\'!$A$5)+1),{n})),"—")'
    )
    c_em.number_format = '"$"#,##0'
    c_em.font = Font(size=10, name="Calibri")
    c_em.alignment = Alignment(horizontal="right")
    c_cl = ws_dash.cell(r, 5,
        f'=IFERROR(INDEX(\'🗄️ Transactions\'!$H$5:$H$200,SMALL(IF((\'🗄️ Transactions\'!$Q$5:$Q$200="HH")*(\'🗄️ Transactions\'!$P$5:$P$200>0)*(\'🗄️ Transactions\'!$Y$5:$Y$200="Active"),ROW(\'🗄️ Transactions\'!$A$5:$A$200)-ROW(\'🗄️ Transactions\'!$A$5)+1),{n})),"—")'
    )
    c_cl.number_format = "mm/dd/yyyy"
    c_cl.font = Font(size=10, name="Calibri")
    bg = C_ROW_ALT if i % 2 == 0 else C_WHITE
    for col_idx in [2,3,4,5]:
        ws_dash.cell(r, col_idx).fill = fill(bg)
    ws_dash.row_dimensions[r].height = 18

# ─── Revenue by Agent (row 36+) ───────────────────────────────────────────────
ws_dash.row_dimensions[35].height = 8
ws_dash.merge_cells("B36:F36")
c = ws_dash["B36"]
c.value = "💼  HH REVENUE SUMMARY  (Active Transactions)"
c.fill = fill(C_DARK_BLUE)
c.font = Font(bold=True, size=12, color=C_WHITE, name="Calibri")
c.alignment = Alignment(horizontal="left", vertical="center", indent=2)
ws_dash.row_dimensions[36].height = 24

for col_idx, label in [(2,"Metric"),(3,"Amount"),(4,"Count"),(5,"Avg/Transaction")]:
    c = ws_dash.cell(37, col_idx, label)
    c.fill = fill(C_HEADER_ROW)
    c.font = Font(bold=True, size=10, color=C_WHITE, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center")
ws_dash.row_dimensions[37].height = 20

summary_rows = [
    ("Total HH Revenue (Active)",
     "=SUMIF('🗄️ Transactions'!Y5:Y200,\"Active\",'🗄️ Transactions'!AI5:AI200)",
     "=COUNTIF('🗄️ Transactions'!Y5:Y200,\"Active\")",
     ),
    ("Total Commission Generated",
     "=SUMIF('🗄️ Transactions'!Y5:Y200,\"Active\",'🗄️ Transactions'!AE5:AE200)",
     "=COUNTIF('🗄️ Transactions'!Y5:Y200,\"Active\")",
     ),
    ("Total Sale Value (Active)",
     "=SUMIF('🗄️ Transactions'!Y5:Y200,\"Active\",'🗄️ Transactions'!F5:F200)",
     "=COUNTIF('🗄️ Transactions'!Y5:Y200,\"Active\")",
     ),
    ("HH Listing Revenue",
     "=SUMPRODUCT((\'🗄️ Transactions\'!Y5:Y200=\"Active\")*(\'🗄️ Transactions\'!AB5:AB200=\"Yes\"),\'🗄️ Transactions\'!AG5:AG200)",
     "=COUNTIFS('🗄️ Transactions'!Y5:Y200,\"Active\",'🗄️ Transactions'!AB5:AB200,\"Yes\")",
     ),
    ("HH Sale Revenue",
     "=SUMPRODUCT((\'🗄️ Transactions\'!Y5:Y200=\"Active\")*(\'🗄️ Transactions\'!AC5:AC200=\"Yes\"),\'🗄️ Transactions\'!AH5:AH200)",
     "=COUNTIFS('🗄️ Transactions'!Y5:Y200,\"Active\",'🗄️ Transactions'!AC5:AC200,\"Yes\")",
     ),
    ("Flat Fee Revenue",
     "=SUMPRODUCT((\'🗄️ Transactions\'!Y5:Y200=\"Active\")*(\'🗄️ Transactions\'!U5:U200=\"Yes\"),\'🗄️ Transactions\'!AF5:AF200)",
     "=COUNTIFS('🗄️ Transactions'!Y5:Y200,\"Active\",'🗄️ Transactions'!U5:U200,\"Yes\")",
     ),
]
for i, (label, amt_formula, count_formula) in enumerate(summary_rows):
    r = 38 + i
    ws_dash.cell(r, 2, label).font = Font(size=11, name="Calibri", bold=(i==0))
    c_amt = ws_dash.cell(r, 3, amt_formula)
    c_amt.number_format = '"$"#,##0.00'
    c_amt.font = Font(size=11, name="Calibri", bold=(i==0))
    c_amt.alignment = Alignment(horizontal="right")
    c_cnt = ws_dash.cell(r, 4, count_formula)
    c_cnt.font = Font(size=11, name="Calibri")
    c_cnt.alignment = Alignment(horizontal="center")
    c_avg = ws_dash.cell(r, 5, f'=IFERROR(C{r}/D{r},"—")')
    c_avg.number_format = '"$"#,##0.00'
    c_avg.font = Font(size=11, name="Calibri")
    c_avg.alignment = Alignment(horizontal="right")
    bg = C_GREEN_KPI if i == 0 else (C_ROW_ALT if i % 2 == 0 else C_WHITE)
    for col_idx in [2,3,4,5]:
        ws_dash.cell(r, col_idx).fill = fill(bg)
    ws_dash.row_dimensions[r].height = 20

# Protect dashboard (no edits allowed at all)
ws_dash.protection = SheetProtection(sheet=True, password="HHRealty2026",
                                      selectLockedCells=True, selectUnlockedCells=True)
print("Sheet 7 (Dashboard) built.")


# ─────────────────────────────────────────────────────────────────────────────
# FINAL: TAB ORDER, SHEET COLORS, SAVE
# ─────────────────────────────────────────────────────────────────────────────

# Set tab colors and order
TAB_COLORS = {
    "📋 Instructions":  "17375E",
    "🏠 Input Form":    "F4B942",
    "📊 Dashboard":     "1F3864",
    "🗄️ Transactions":  "2F5496",
    "✅ Checklist":     "375623",
    "👤 Agents":        "7F4000",
    "⚙️ Lookups":       "808080",
}
for sheet_name, color in TAB_COLORS.items():
    if sheet_name in wb.sheetnames:
        wb[sheet_name].sheet_properties.tabColor = color

# Set sheet order: Dashboard first, then Input, Transactions, Checklist, Agents, Instructions, Lookups
DESIRED_ORDER = [
    "📊 Dashboard",
    "🏠 Input Form",
    "🗄️ Transactions",
    "✅ Checklist",
    "👤 Agents",
    "📋 Instructions",
    "⚙️ Lookups",
]
# Reorder sheets
for i, name in enumerate(DESIRED_ORDER):
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        if idx != i:
            ws_move = wb[name]
            wb.move_sheet(ws_move, offset=i - idx)

# Make Dashboard the active sheet
wb.active = wb["📊 Dashboard"]

# ─── Save ───────────────────────────────────────────────────────────────────
output_path = "/home/user/Real-Estate-Tracker/HH Realty Tracker V14.xlsx"
wb.save(output_path)
print(f"\n✅ Workbook saved: {output_path}")
print(f"   Sheets: {wb.sheetnames}")

import os
size = os.path.getsize(output_path)
print(f"   File size: {size:,} bytes ({size/1024:.1f} KB)")

