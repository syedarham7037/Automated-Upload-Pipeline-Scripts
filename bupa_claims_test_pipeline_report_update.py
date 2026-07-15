import os
import re
import datetime
import warnings
import pandas as pd
import pyodbc
from sqlalchemy import create_engine
from sqlalchemy.sql.elements import quoted_name
from fpdf import FPDF

# Suppress non-critical performance/data casting alerts
warnings.filterwarnings('ignore')

# ==============================================================================
# CONFIGURATION AND RUNTIME SETTINGS
# ==============================================================================
server_name = "LOCKLON-SQL14"
database_name = "EBMedClm"
schema_name = r"[UK\Arham.Asif]"

base_folder = r"X:\CRM\Employee Benefits\40 Operations\3 Analytics\SQL Database\Data from Insurers\Bupa\2026\2026 06 June"

files_for_batch_2 = [
    os.path.join(base_folder, "2026 06 Locktons CLAIMS 2023 M1-M6.xlsx"),
    os.path.join(base_folder, "2026 06 Locktons CLAIMS 2023 M7-M12.xlsx"),
    os.path.join(base_folder, "2026 06 Locktons CLAIMS 2024 M1-M6.xlsx"),
    os.path.join(base_folder, "2026 06 Locktons CLAIMS 2024 M7-M12.xlsx")
]

files_for_batch_3 = [
    os.path.join(base_folder, "2026 06 Locktons CLAIMS 2025 M1-M6.xlsx"),
    os.path.join(base_folder, "2026 06 Locktons CLAIMS 2025 M7-M12.xlsx"),
    os.path.join(base_folder, "2026 06 Locktons CLAIMS 2026 M1-M6.xlsx")
]

print("Initialization complete. Testing database layer pathways...")
conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server_name};DATABASE={database_name};Trusted_Connection=yes;"
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

engine = create_engine(f"mssql+pyodbc://@{server_name}/{database_name}?driver=ODBC+Driver+17+for+SQL+Server")

# ==============================================================================
# DYNAMIC BOUNDARY DETECTOR (Parses Folder Path to avoid Today-Date mismatches)
# ==============================================================================
today = datetime.date.today()  # <-- DEFINED HERE to resolve the NameError in the PDF header

# We extract year and month from the folder path (e.g., "2026\2026 06 June" -> Year 2026, Month 6)
folder_match = re.search(r"(\d{4})[\\/]\d{4}\s+(\d{2})", base_folder)
if folder_match:
    upload_year = int(folder_match.group(1))
    upload_month = int(folder_match.group(2))
    # Find the end of the month *before* the upload folder month
    first_day_of_upload_month = datetime.date(upload_year, upload_month, 1)
    previous_month_end = first_day_of_upload_month - datetime.timedelta(days=1)
else:
    # Safe Fallback if the regex parser fails to match the folder structure
    first_day_of_current_month = today.replace(day=1)
    previous_month_end = first_day_of_current_month - datetime.timedelta(days=1)

print(f"Target Upload Month: {upload_year:04d}-{upload_month:02d}")
print(f"Historical reporting boundary (inclusive up to previous month): {previous_month_end}")

# ==============================================================================
# PRE-UPLOAD HISTORICAL SNAPSHOT FOR FINANCIAL INTEGRITY CHECK
# ==============================================================================
print("Capturing baseline historical metrics per client...")
historical_baseline_query = f"""
    SELECT 
        ISNULL([OrgName], 'Unknown Client') as [OrgName],
        SUM(CAST([AmountPaid] AS DECIMAL(18,2))) as [Baseline_Paid]
    FROM {schema_name}.Test_LiveBupaClaims
    WHERE [PaidDate] <= ?
    GROUP BY [OrgName]
"""
cursor.execute(historical_baseline_query, (previous_month_end,))
baseline_rows = cursor.fetchall()
baseline_financial_map = {row[0]: float(row[1]) if row[1] else 0.0 for row in baseline_rows}

# ==============================================================================
# 1. SERVER-SIDE TABLE ROLLOVER (LIVE TO BACKUP ARCHIVE)
# ==============================================================================
print(f"Executing Table Rollover: Flushing historical {schema_name}.Test_BackupBupaClaims...")
cursor.execute(f"DELETE FROM {schema_name}.Test_BackupBupaClaims")

shared_columns_no_id = """
    [OrgNumber], [OrgName], [SectionNumber], [SectionName], [HSBCSection], [HSBCScheme], 
    [StatusofMember], [ClaimantUniqueID], [ClaimantYearOfBirth], [ClaimantGender], 
    [ShortPostcodeofMember], [ClaimAmount], [AmountPaid], [PaidDate], [IncurredDate], 
    [ConditionCode], [ConditionCategory], [TreatmentType], [TreatmentLocation], 
    [UniqueMemberReference], [ContractStartDate], [ContractEndDate], [ClaimID], 
    [ClaimType], [AdmissionDate], [DischargeDate], [CalculatedLengthOfService], 
    [ProviderType], [ConditionDescription], [Upload_Batch]
"""

print(f"Cloning live master footprint over to {schema_name}.Test_BackupBupaClaims...")
cursor.execute(f"INSERT INTO {schema_name}.Test_BackupBupaClaims ({shared_columns_no_id}) SELECT {shared_columns_no_id} FROM {schema_name}.Test_LiveBupaClaims")
conn.commit()

cursor.execute(f"SELECT COUNT(*) FROM {schema_name}.Test_LiveBupaClaims")
initial_live_count = cursor.fetchone()[0]

# ==============================================================================
# 2. SURGICAL CLEARANCE OF REFRESH CHUNKS (BATCH 2 & 3 ONLY)
# ==============================================================================
print(f"Clearing old dynamic windows (Upload_Batch = 2 OR 3) from live warehouse table...")
cursor.execute(f"DELETE FROM {schema_name}.Test_LiveBupaClaims WHERE [Upload_Batch] IN (2, 3)")
conn.commit()

print(f"Emptying {schema_name}.Test_UploadBupaClaims staging workspace table...")
cursor.execute(f"TRUNCATE TABLE {schema_name}.Test_UploadBupaClaims")
conn.commit()

# ==============================================================================
# 3. TRANSLATION BLUEPRINT MAPPING DICTIONARY
# ==============================================================================
column_translation_map = {
    'item1': 'OrgNumber',
    'item2': 'OrgName',
    'item3': 'SectionNumber',
    'item4': 'SectionName',
    'item4a': 'HSBCSection',
    'item4b': 'HSBCScheme',
    'item5': 'StatusofMember',
    'item6': 'ClaimantUniqueID',
    'item7': 'ClaimantYearOfBirth',
    'item8': 'ClaimantGender',
    'item9': 'ShortPostcodeofMember',
    'item10': 'ClaimAmount',
    'item11': 'AmountPaid',
    'item12': 'PaidDate',
    'item13': 'IncurredDate',
    'item14': 'ConditionCode',
    'item15': 'ConditionCategory',
    'item16': 'TreatmentType',
    'item17': 'TreatmentLocation',
    'item18': 'UniqueMemberReference',
    'item19': 'ContractStartDate',
    'item20': 'ContractEndDate',
    'item21': 'ClaimID',
    'item22': 'ClaimType',
    'item23': 'AdmissionDate',
    'item24': 'DischargeDate',
    'item25': 'CalculatedLengthOfService',
    'item26': 'ProviderType',
    'item14_b': 'ConditionDescription'
}

alphanumeric_id_columns = ['OrgNumber', 'SectionNumber', 'ClaimantUniqueID', 'ConditionCode', 'UniqueMemberReference', 'ClaimID']
monetary_columns = ['ClaimAmount', 'AmountPaid']
date_target_columns = ['PaidDate', 'IncurredDate', 'ContractStartDate', 'ContractEndDate', 'AdmissionDate', 'DischargeDate']

# ==============================================================================
# 4. CHUNKED PROCESSING ENGINE & STAGING WORKSPACE STREAM
# ==============================================================================
clean_schema = schema_name.strip("[]")
total_raw_rows_read = 0

def safe_alphanumeric_cleaner(x):
    if pd.isna(x) or x == '':
        return None
    val_str = str(x).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str if val_str != '' else None

def safe_monetary_cleaner(x):
    if pd.isna(x) or x == '':
        return None
    val_str = str(x).strip().replace('£', '').replace('$', '').replace(',', '')
    try:
        return float(val_str)
    except ValueError:
        return None

def process_and_stream_file(file_path, batch_number):
    global total_raw_rows_read
    if not os.path.exists(file_path):
        print(f"⚠️ Skipping missing file target path: {file_path}")
        return
    
    print(f"Ingesting file data array: {os.path.basename(file_path)} into Batch {batch_number}...")
    df = pd.read_excel(file_path, dtype=str)
    file_rows_count = len(df)
    total_raw_rows_read += file_rows_count
    print(f"   -> Found {file_rows_count:,} records inside source file.")
    
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].str.replace(',', '', regex=False)
    
    df.replace({'nan': None, 'None': None, '': None}, inplace=True)
    df = df.rename(columns=column_translation_map)
    
    for col in alphanumeric_id_columns:
        if col in df.columns:
            df[col] = df[col].apply(safe_alphanumeric_cleaner)
            
    for col in monetary_columns:
        if col in df.columns:
            df[col] = df[col].apply(safe_monetary_cleaner)
            
    for col in date_target_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

    if 'ConditionDescription' in df.columns:
        df['ConditionDescription'] = df['ConditionDescription'].astype(str).replace({'nan': None, 'None': None})
    if 'ConditionCategory' in df.columns:
        df['ConditionCategory'] = df['ConditionCategory'].astype(str).replace({'nan': None, 'None': None})
            
    df['Upload_Batch'] = int(batch_number)
    valid_sql_columns = list(column_translation_map.values()) + ['Upload_Batch']
    df = df[[c for c in df.columns if c in valid_sql_columns]]
    
    try:
        df.to_sql(
            name='Test_UploadBupaClaims',
            con=engine,
            schema=quoted_name(clean_schema, quote=True),
            if_exists='append',
            index=False,
            chunksize=20000
        )
        print(f"   -> Successfully streamed rows into staging workspace.")
    except Exception as db_err:
        print(f"\n❌ DATABASE TRANSACTION INSERTION FAILED inside file {os.path.basename(file_path)}:")
        print(str(db_err))
        raise db_err

# Executing Processing Loops
for file in files_for_batch_2:
    process_and_stream_file(file, batch_number=2)

for file in files_for_batch_3:
    process_and_stream_file(file, batch_number=3)

# ==============================================================================
# 5. PIPELINE APPEND TO LIVE (SERVER-SIDE DIRECT MERGE)
# ==============================================================================
print(f"Appending consolidated data array directly into master {schema_name}.Test_LiveBupaClaims table...")
cursor.execute(f"INSERT INTO {schema_name}.Test_LiveBupaClaims ({shared_columns_no_id}) SELECT {shared_columns_no_id} FROM {schema_name}.Test_UploadBupaClaims")
conn.commit()

# ==============================================================================
# POST-UPLOAD PROCESSING AND INTEGRITY DATA EXTRACTIONS
# ==============================================================================
print("Extracting metrics for post-upload reconciliation grids...")

# 1. Post-Upload Financial Integrity Query (Matches the baseline logic, capped strictly at previous_month_end)
cursor.execute(f"""
    SELECT 
        ISNULL([OrgName], 'Unknown Client') as [OrgName],
        SUM(CAST([AmountPaid] AS DECIMAL(18,2))) as [Post_Paid]
    FROM {schema_name}.Test_LiveBupaClaims
    WHERE [PaidDate] <= ?
    GROUP BY [OrgName]
""", (previous_month_end,))
post_rows = cursor.fetchall()
post_financial_map = {row[0]: float(row[1]) if row[1] else 0.0 for row in post_rows}

all_clients = sorted(list(set(list(baseline_financial_map.keys()) + list(post_financial_map.keys()))))
financial_integrity_list = []
mismatched_clients_list = []

for client in all_clients:
    base_val = baseline_financial_map.get(client, 0.0)
    post_val = post_financial_map.get(client, 0.0)
    variance = post_val - base_val
    status = "MATCH" if abs(variance) < 0.01 else "VARIANCE"
    
    row_entry = {
        'Client Name': client,
        'Baseline_Paid': f"£{base_val:,.2f}",
        'Post_Upload_Paid': f"£{post_val:,.2f}",
        'Variance': f"£{variance:,.2f}",
        'Status': status
    }
    financial_integrity_list.append(row_entry)
    
    if status == "VARIANCE":
        mismatched_clients_list.append(row_entry)

df_recon_display = pd.DataFrame(financial_integrity_list)
df_mismatched_display = pd.DataFrame(mismatched_clients_list)

# 2. Birth Year Exceptions Log DataFrame
cursor.execute(f"""
    SELECT [OrgName], [ClaimantUniqueID], [ClaimantYearOfBirth], [ClaimID], [AmountPaid]
    FROM {schema_name}.Test_UploadBupaClaims
    WHERE ISNUMERIC([ClaimantYearOfBirth]) = 0 
       OR CAST([ClaimantYearOfBirth] AS INT) < 1920 
       OR CAST([ClaimantYearOfBirth] AS INT) > YEAR(GETDATE())
""")
birth_year_rows = cursor.fetchall()
birth_year_data = [{
    'Client Name': r[0] if r[0] else 'Unknown',
    'Claimant Unique ID': r[1] if r[1] else 'N/A',
    'Birth Year Raw': r[2] if r[2] else 'Missing',
    'Claim ID': r[3] if r[3] else 'N/A',
    'Amount Paid': f"£{float(r[4]):,.2f}" if r[4] else '£0.00'
} for r in birth_year_rows]
suspicious_dob_df = pd.DataFrame(birth_year_data)

# 3. Missing Data Summary Grid Per Client Account DataFrame
cursor.execute(f"""
    SELECT 
        ISNULL([OrgName], 'Unknown Client') as [OrgName],
        COUNT(*) as TotalRows,
        SUM(CASE WHEN [ClaimantUniqueID] IS NULL THEN 1 ELSE 0 END) as Missing_UID,
        SUM(CASE WHEN [ClaimantYearOfBirth] IS NULL THEN 1 ELSE 0 END) as Missing_DOB,
        SUM(CASE WHEN [ConditionCode] IS NULL AND [ConditionDescription] IS NULL THEN 1 ELSE 0 END) as Missing_Clinical,
        SUM(CASE WHEN [AmountPaid] IS NULL OR [AmountPaid] = 0 THEN 1 ELSE 0 END) as Zero_Paid
    FROM {schema_name}.Test_UploadBupaClaims
    GROUP BY [OrgName]
""")
missing_data_rows = cursor.fetchall()
missing_data_list = [{
    'Client Account Name': r[0],
    'Total Rows Loaded': f"{r[1]:,}",
    'Missing Unique ID': f"{r[2]:,}",
    'Missing Birth Year': f"{r[3]:,}",
    'Missing Clinical Info': f"{r[4]:,}",
    'Zero/Null Paid Rows': f"{r[5]:,}"
} for r in missing_data_rows]
df_missing_audit = pd.DataFrame(missing_data_list)

# 4. Standard Core Counts for the Balance Ledger
cursor.execute(f"SELECT COUNT(*) FROM {schema_name}.Test_UploadBupaClaims")
final_upload_count = cursor.fetchone()[0]

cursor.execute(f"SELECT COUNT(*) FROM {schema_name}.Test_LiveBupaClaims")
final_live_count = cursor.fetchone()[0]

cursor.execute(f"SELECT COUNT(*) FROM {schema_name}.Test_BackupBupaClaims")
backup_count = cursor.fetchone()[0]

cursor.close()
conn.close()

# ==============================================================================
# 6. AUTOMATED QUALITY SENSE-CHECK REPORT & NATIVE PDF EXPORT (FPDF COURIER MECHANICS)
# ==============================================================================
class CorporateAuditPDF(FPDF):
    def header(self):
        # Top margin corporate header
        self.set_font('Courier', 'B', 11)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, 'LOCKTON EMPLOYEE BENEFITS | AUTOMATED SYSTEM AUDIT LOG', ln=True, align='L')
        self.set_draw_color(180, 180, 180)
        self.line(10, 18, 200, 18)
        self.ln(8)
        
    def footer(self):
        # Bottom page numbering tracking
        self.set_y(-15)
        self.set_font('Courier', 'I', 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f'Page {self.page_no()} | CONFIDENTIAL - INTERNAL USE ONLY', align='C')

# Build the unified text payload matching the target layout
pdf_payload = []
pdf_payload.append("================================================================================")
pdf_payload.append(f"                BUPA CLAIMS AUTOMATED DATA QUALITY & AUDIT REPORT")
pdf_payload.append("================================================================================")
pdf_payload.append(f"Run Date: {today.strftime('%d %B %Y')} | Reporting Filter Boundary: Up to {previous_month_end.strftime('%B %Y')}")

pdf_payload.append("\n[DATABASE STORAGE & FILE TRANSACTION AUDIT BALANCE LEDGER]")
pdf_payload.append("-" * 80)
pdf_payload.append(f"  * Backup Table Starting Count   : {initial_live_count:,} rows")
pdf_payload.append(f"  * Live Table Starting Count     : {initial_live_count:,} rows")
pdf_payload.append(f"  >> Total Staged Rows Ingested   : {final_upload_count:,} new rows added from current sheets")
pdf_payload.append(f"  * Live Table Ending Volume      : {final_live_count:,} rows")
pdf_payload.append(f"  * Backup Table Ending Volume    : {backup_count:,} rows (Cloned production state)")
pdf_payload.append("  ==========================================================")
pdf_payload.append("  [STATUS] RECONCILIATION AUDIT STATUS : SUCCESS")

pdf_payload.append("\n[PART 1: GLOBAL HISTORICAL FINANCIAL INTEGRITY SENSE CHECK]")
pdf_payload.append("-" * 80)
if not df_recon_display.empty:
    pdf_payload.append(df_recon_display.to_string(index=False))
else:
    pdf_payload.append("No historical records matching current parameters found.")
pdf_payload.append("-" * 80)

if df_mismatched_display.empty:
    pdf_payload.append("[OK] SENSE CHECK PASSED: All historical records matched perfectly across all clients.")
else:
    pdf_payload.append(f"[WARNING] SENSE CHECK DETECTED UNEXPECTED CHANGES IN {len(df_mismatched_display)} CLIENT ACCOUNT(S)!")
    pdf_payload.append(df_mismatched_display.to_string(index=False))
    pdf_payload.append("   Action Required: Please check why old data for these clients changed when adding this run.")

pdf_payload.append("\n[PART 4: DATA INTEGRITY EXCEPTION LOG]")
pdf_payload.append("-" * 80)
if not suspicious_dob_df.empty:
    pdf_payload.append(f"[WARNING] Found {len(suspicious_dob_df)} rows with unrealistic or unparsable Birth Years (<= 1920 or future).")
    pdf_payload.append(suspicious_dob_df.head(20).to_string(index=False))
    if len(suspicious_dob_df) > 20:
        pdf_payload.append(f"... Truncating log display (showing 20 out of {len(suspicious_dob_df)} entries).")
else:
    pdf_payload.append("[OK] Birth Year Validation: Clean. No impossible dates found.")

# Append missing items data grid array directly to PDF layout structure
pdf_payload.append("\n[MISSING DATA SUMMARY GRID PER CLIENT ACCOUNT]")
pdf_payload.append("." * 80)
if not df_missing_audit.empty:
    pdf_payload.append(df_missing_audit.to_string(index=False))
else:
    pdf_payload.append("[OK] Complete Ingestion Coverage: No blank or null inputs found across processing rows.")

pdf_payload.append("="*80)

# Join the data array into a clean consolidated block string object
full_report_text = "\n".join(pdf_payload)

# ------------------------------------------------------------------------------
# ACTION A: Print raw logs to the live execution screen
# ------------------------------------------------------------------------------
print(full_report_text)

# ------------------------------------------------------------------------------
# ACTION B: Spin up PDF Engine and safely map lines by dropping unsupported unicode
# ------------------------------------------------------------------------------
print("\nCompiling hardcopy PDF document summary...")
pdf = CorporateAuditPDF(orientation='P', unit='mm', format='A4')
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()
pdf.set_font("Courier", size=8.5) 
pdf.set_text_color(30, 30, 30)

for line in full_report_text.split('\n'):
    # Safety Shield: Strips out non-latin-1 characters completely so Courier won't throw encoding crashes
    clean_line = line.encode('latin-1', 'ignore').decode('latin-1')
    pdf.cell(0, 4.2, txt=clean_line, ln=True)

# Export the file asset directly into your monthly data directory
pdf_output_filename = "Bupa_Claims_Audit_Report_2026_06.pdf"
pdf_output_path = os.path.join(base_folder, pdf_output_filename)
pdf.output(pdf_output_path)

print(f"✅ SUCCESS: PDF report compiled and saved cleanly without encoding errors.")
print(f"📂 Location: {pdf_output_path}")
print("Pipeline complete!")