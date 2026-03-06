# ================================================================
# ODOO CONNECTION CONFIGURATION
# ⚠ KEEP THIS FILE PRIVATE — never share or paste its contents
# ================================================================

URL      = "https://my.thorasys.com"
DB       = "production"
USERNAME = "guy.drapeau@thorasys.com"
API_KEY  = "2d57ea208e7431a13f797377a2b7d2b32f7aa2d0"

# ================================================================
# PRODUCT LINE DEFINITIONS
# ================================================================
# C2 device: tickets where product internal ref is one of these
C2_PRODUCT_REFS  = ['101711', '102237', '101769']
# Explicitly C-100 products (treated as C-100 even if logic changes)
C100_PRODUCT_REFS = ['101769']

# C-100 device: everything else (including no product assigned)
# No need to list — anything not in C2_PRODUCT_REFS is C-100

DEVICE_LINES = ['C2', 'C-100']

# ================================================================
# HELPDESK DIVISIONS
# ================================================================
DIVISIONS = ["Support - Americas", "Support - EMEA", "Support - APAC"]

# Stages considered closed (excluded from time-based/open stats)
CLOSED_STAGES = {'Solved', 'Cancelled'}

# ================================================================
# REPAIR DIVISIONS
# ================================================================
REPAIR_DIVISIONS = ["THORASYS-Americas", "THORASYS-EMEA", "THORASYS-APAC"]

# Short labels for display (same order as REPAIR_DIVISIONS)
REPAIR_DIVISION_LABELS = {
    "THORASYS-Americas": "Americas",
    "THORASYS-EMEA":     "EMEA",
    "THORASYS-APAC":     "APAC",
}

# ================================================================
# REPAIR STATE MAPPING
# ================================================================
REPAIR_STATE_MAP = {
    'draft':        'RMA Created',
    'confirmed':    'Device Received',
    'under_repair': 'Under Repair',
}
REPAIR_STATE_ORDER = ['RMA Created', 'Device Received', 'Under Repair']

# ================================================================
# REPAIR — EXCLUDED CUSTOMERS
# ================================================================
REPAIR_EXCLUDED_CUSTOMERS = ['eResearch Technology GmbH', 'Clario']

# ================================================================
# REPAIR — EXCLUDED TAGS
# ================================================================
REPAIR_EXCLUDED_TAGS = ['Refurbishment']

# ================================================================
# PRODUCTION PLAN (units per month — finished goods)
# ================================================================
MONTHLY_PRODUCTION_PLAN = {
    '101336': 150,   # tremoFlo C-100 Device ClearFlo compatible
    '101711': 40,    # tremoflo C2 Device - clearflo
    '101769': 20,    # tremoflo C2 Device - Bird - Purple
    '102237': 20,    # tF C2 Device - Yellow - VTG 28554
}
PLANNING_MONTHS = 3  # current month + next 2
