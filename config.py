# --- Centralized Configuration File ---

# Google Sheet settings
SHEET_NAME = '2025_HR Planner'
EMPLOYEES_TAB = 'Employees'
SHIFTS_TAB = 'Shifts'
REQUESTS_TAB = 'Absence_Requests'
OFFICIAL_SCHEDULE_TAB = 'Official_Schedule'
SANDBOX_SCHEDULE_TAB = 'Sandbox_Schedule'
OFFERS_TAB = 'Offers'

# Email settings
HR_EMAIL = 'hr.scheduler@example.com' # The address for sending/receiving offers

# Solver settings
NUM_PARALLEL_WORKERS = 4

# --- Column Name Mappings ---
# This helps if the column names in the Google Sheet change.

# Absence_Requests sheet
COL_REQUEST_NAME = 'Nom'
COL_REQUEST_START = 'Début'
COL_REQUEST_END = 'Fin'
COL_REQUEST_TOKENS = 'Tokens'

# Employees sheet
COL_EMPLOYEE_NAME = 'Employee_Name'
COL_EMPLOYEE_EMAIL = 'Email'
COL_EMPLOYEE_ROLE = 'Role'
COL_EMPLOYEE_TOKENS = 'Tokens_Official'

# Shifts sheet
COL_SHIFT_ID = 'Shift_ID'
COL_SHIFT_DURATION = 'Duration_Hours'
COL_SHIFT_ROLE = 'Role'
COL_SHIFT_DAYS = 'Applicable_Days'

# Official_Schedule sheet
COL_SCHEDULE_SHIFT = 'Shift'

# Offers sheet
COL_OFFER_ID = 'Offer_ID'
COL_OFFER_EMPLOYEE = 'Employee_Name'
COL_OFFER_STATUS = 'Status'
COL_OFFER_EXPIRY = 'Expiry_Time'
COL_OFFER_REQUESTER = 'Requester_Name'
