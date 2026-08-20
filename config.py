import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "Input"
OUTPUT_DIR = BASE_DIR / "Output"
OUTPUT_FILE = OUTPUT_DIR / "Result.xlsx"
OUTLET_MASTER_FILE = INPUT_DIR / "Outlet_Master.xlsx"
LUBRICANT_LITER_MASTER_FILE = INPUT_DIR / "Lubricant_Product_Liters.xlsx"

FISCAL_MONTHS = [
    "July", "August", "September", "October", "November", "December",
    "January", "February", "March", "April", "May", "June",
]

PRODUCTS = ["PMG", "HSD", "R95", "Lubricants"]
AREAS = ["A00", "A01", "A02", "A03", "A04", "A05", "A06", "A07", "A08"]

# Objective workbook sheet names. The keys are dashboard product names.
OBJECTIVE_SHEET_ALIASES = {
    "HSD": ("HSD",),
    "PMG": ("PMG",),
    "R95": ("HOBC", "R95", "R-95"),
    "Lubricants": ("Lubes", "Lubricants", "LUBE", "HOC"),
}

# Sales Material values used in the sales files.
PRODUCT_MATERIAL_ALIASES = {
    "HSD": ("HSD",),
    "PMG": ("PMG",),
    "R95": ("R95", "HOBC", "R-95", "ALTRON X"),
}

LUBRICANT_MATERIAL_CODES = {
    "4103903", "4009451", "4020103", "4020951", "4024051", "4017551",
    "4020139", "4010417", "4017539", "4050879", "4023669", "4023251",
    "4023269", "4022451", "4022469", "4022417", "4023751", "4050817",
    "4023217", "4020117", "4023651", "4020151", "4014551", "4017503",
    "4051217", "4017517", "4016251", "4023969", "4014539", "4016239",
    "4023569", "4020251", "4014703", "4023951", "4041517", "4024039",
    "4009439", "4014503", "4023151", "4041569", "4103978", "4050017",
    "4050917", "4109403", "4023551", "4041417", "4050979", "4003851",
    "4014451", "4020939", "4103999", "4009478", "4020203", "4022403",
    "4044003", "4044117", "4023739", "4103904", "4014603", "4041303",
    "4014403", "4017403", "4040903", "4024551", "4024751", "4024451",
    "4024469", "4024151", "4024169", "4024651", "4112503", "4041503",
}
