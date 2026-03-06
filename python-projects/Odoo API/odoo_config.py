# ================================================================
# ODOO CONNECTION CONFIGURATION
# ⚠ KEEP THIS FILE PRIVATE — never share or paste its contents
# ================================================================

URL      = "https://my.thorasys.com"
DB       = "production"
USERNAME = "guy.drapeau@thorasys.com"
API_KEY  = "7e68dd7d7d93a778571b25161b4e0d97784431d1"

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
    
   '101485': 5,   # [normal]  devices: 101490
               # tremoflo C-100 Airwave Oscillometry System & Accessories - Suregard Australia Edition
    '101524': 10,   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - US Edition
    '101761': 5,   # [normal]  devices: 101759
               # tf C-100 Kit Vitalograph - Clinical Trial Edition (28300)
    '101762': 5,   # [normal]  devices: 101760
               # tf C-100 Kit Vitalograph - Healthcare Edition (28500)
    '101769': 5,   # [normal]  C2 devices: 101769 AU
               
    '101711': 40,   # [normal]  C2 devices: 101711 US/CAN IND
    '102239': 5,   # [normal]  C2 devices: 102239 Blue VTG 28501
    '102240': 5,   # [normal]  C2 devices: 102239 Blue VTG 28554
               
    '101963': 10,   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - India
    '101969': 1,   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - Canada
    '102002': 1,   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - UK Edition
    '102003': 10,   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - Europe Edition
    '102005': 1,   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - Switzerland Edition
    
    '102036': 100,  # tremoflo C2 Airwave Oscillometry System & Accessories
    '102181': 5,   # [normal]  devices: 101760
               # tf C-100 Kit Vitalograph - Healthcare US Edition (28500)
    '102585': 125,   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - ERT US Edition
}
PLANNING_MONTHS = 3  # current month + next 2

# ================================================================
# KANBAN SUPPLIERS — parts from these suppliers are always available
# on-demand and should be excluded from buildable units limiting calc
# and from the PO plan (they are replenished via kanban/MRO)
# ================================================================
KANBAN_SUPPLIERS = [
    'McMaster Carr',
    'Digi-Key',
    'Mouser Electronics',
    'Lee Spring',
    'Caplugs',
    'Maverick Label',
    'THORASYS Thoracic Medical Systems Inc.',  # internal — not a real supplier
]

# All finished products whose BOMs are included in procurement/lead time planning
ALL_PLAN_PRODUCTS = [
    '101485',   # [normal]  devices: 101490
               # tremoflo C-100 Airwave Oscillometry System & Accessories - Suregard Australia Edition
    '101524',   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - US Edition
    '101761',   # [normal]  devices: 101759
               # tf C-100 Kit Vitalograph - Clinical Trial Edition (28300)
    '101762',   # [normal]  devices: 101760
               # tf C-100 Kit Vitalograph - Healthcare Edition (28500)
    '101769',   # [normal]  C2 devices: 101769 AU
               
    '101711',   # [normal]  C2 devices: 101711 US/CAN IND
    '102239',   # [normal]  C2 devices: 102239 Blue VTG 28501
    '102240',   # [normal]  C2 devices: 102239 Blue VTG 28554
               
    '101963',   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - India
    '101969',   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - Canada
    '102002',   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - UK Edition
    '102003',   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - Europe Edition
    '102005',   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - Switzerland Edition
    
    '102036',  # tremoflo C2 Airwave Oscillometry System & Accessories
    '102181',   # [normal]  devices: 101760
               # tf C-100 Kit Vitalograph - Healthcare US Edition (28500)
    '102585',   # [normal]  devices: 101336
               # tremoflo C-100 Airwave Oscillometry System & Accessories - ERT US Edition
]
