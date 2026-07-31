import os

# --- PENGATURAN DIREKTORI (FOLDER) ---
# BASE_DIR otomatis mencari folder utama (PROJECT_OTOMATISASI)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DB_DIR = os.path.join(BASE_DIR, "database")
TPL_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "Output_FI_PL_SI")

# --- PENGATURAN NAMA FILE DATABASE ---
MASTER_LIST_FILE = os.path.join(DB_DIR, "Master Order List Adi.xlsx")
DB_SHIPMENT_FILE = os.path.join(DB_DIR, "guideline_shipment.xlsx")

# --- KAMUS HS CODE ---
HS_KAMUS = {
    "211354": "64041990",
}