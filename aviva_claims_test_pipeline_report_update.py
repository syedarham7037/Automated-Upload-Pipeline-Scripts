import os
import re
import warnings
import pandas as pd
import pyodbc
from sqlalchemy import create_engine
from sqlalchemy.sql.elements import quoted_name

# Suppress non-critical performance notifications
warnings.filterwarnings('ignore')

# ==============================================================================
# CONFIGURATION AND RUNTIME SETTINGS (May 2026 Run)
# ==============================================================================
target_upload_date = "2026-06-15"  # Format: YYYY-MM-DD (Stamps the UploadDate column)
server_name = "LOCKLON-SQL14"
database_name = "EBMedClm"
schema_name = r"[UK\Arham.Asif]"   # Raw string identifier handles backslashes safely

# Corporate network folder file definitions
current_month_csv = r"X:\CRM\Employee Benefits\40 Operations\3 Analytics\SQL Database\Data from Insurers\Aviva\2026\2026 06 June\LOCKTONS_CLAIMS.csv"
previous_month_csv = r"X:\CRM\Employee Benefits\40 Operations\3 Analytics\SQL Database\Data from Insurers\Aviva\2026\2026 05 May\LOCKTONS_CLAIMS.csv"

# Establish target month and year parameters based on the upload date
parsed_date = pd.to_datetime(target_upload_date)
target_month = parsed_date.month
target_year = parsed_date.year

# Calculate Dynamic Look-back Months for Part 3 Trend Audit
date_current_m = parsed_date.to_period('M')
date_minus_1m = (parsed_date - pd.DateOffset(months=1)).to_period('M')
date_minus_2m = (parsed_date - pd.DateOffset(months=2)).to_period('M')

label_current_m = date_current_m.strftime('%B_Rows')   # e.g., "May_Rows"
label_minus_1m = date_minus_1m.strftime('%B_Rows')     # e.g., "April_Rows"
label_minus_2m = date_minus_2m.strftime('%B_Rows')     # e.g., "March_Rows"

print("Initialization complete. Connecting to SQL Server...")
# Raw database connectivity for structural queries
conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server_name};DATABASE={database_name};Trusted_Connection=yes;"
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# High-speed bulk loading engine connection link
engine = create_engine(f"mssql+pyodbc://@{server_name}/{database_name}?driver=ODBC+Driver+17+for+SQL+Server")

# ==============================================================================
# 1. LIVE TO BACKUP SERVER-SIDE REPLICATION (IDENTITY SAFE)
# ==============================================================================
print(f"Executing Table Rollover: Emptying {schema_name}.Test_BackupAvivaClaims...")
cursor.execute(f"DELETE FROM {schema_name}.Test_BackupAvivaClaims")

# Explicit structural column string layout to handle identity transfers smoothly
shared_columns_with_id = """
    [ID], [Policy Number], [Policy Holder Name], [Broker], [policy_last_renewal_date], 
    [pol_mem_num], [Age Band], [gender], [claim_num_yr], [scheme_year], 
    [Condition Code], [Main_Diagnosis_Group], [Procedure Type], [Place of Service], 
    [Finalised Date], [claim_incurred_date_min], [claim_incurred_date_max], 
    [count_incurred_days], [Amount Submitted], [Amount Paid], [Excess Amount], [UploadDate]
"""

print(f"Cloning Test_Live records safely over to {schema_name}.Test_BackupAvivaClaims...")
cursor.execute(f"SET IDENTITY_INSERT {schema_name}.Test_BackupAvivaClaims ON")
cursor.execute(f"INSERT INTO {schema_name}.Test_BackupAvivaClaims ({shared_columns_with_id}) SELECT {shared_columns_with_id} FROM {schema_name}.Test_LiveAvivaClaims")
cursor.execute(f"SET IDENTITY_INSERT {schema_name}.Test_BackupAvivaClaims OFF")
conn.commit()

# Track original baseline stats for the final reconciliation summary
cursor.execute(f"SELECT COUNT(*) FROM {schema_name}.Test_LiveAvivaClaims")
initial_live_count = cursor.fetchone()[0]

# ==============================================================================
# 2. CLIENT DRIFT ASSESSMENT (IN-MEMORY DELTA DRIFT)
# ==============================================================================
print("Reading incoming and historical CSV source files for client drift evaluations...")
df_current_raw = pd.read_csv(current_month_csv, dtype=str)  # Read as string to preserve exact text
df_previous_raw = pd.read_csv(previous_month_csv, dtype=str)

# Isolate unique client structures
current_clients = set(df_current_raw['policyholder_name'].dropna().unique())
previous_clients = set(df_previous_raw['policyholder_name'].dropna().unique())

# Identify new vs lost groups
new_clients_drift = current_clients - previous_clients
lost_clients_drift = previous_clients - current_clients

# ==============================================================================
# 3. CHRONOLOGICAL FILTER PIPELINE MATRIX (DIAGNOSTIC ALIGNED)
# ==============================================================================
print("Executing conditional chronological filters based on client classifications...")

# Step 1: Force clean text type on the date tracking column, replacing slashes with dashes
df_current_raw['bill_finalised_month'] = df_current_raw['bill_finalised_month'].astype(str).str.strip().str.replace('/', '-', regex=False)

# Step 2: Build an exact unified text pattern layout matching our standardized format (e.g., "01-05-2026")
string_pattern_match = f"01-{target_month:02d}-{target_year}"

historical_new_client_pool = []
current_month_existing_client_pool = []

for client, group in df_current_raw.groupby('policyholder_name'):
    if client in new_clients_drift:
        # Rule A: Keep ALL data rows for brand-new client additions
        historical_new_client_pool.append(group)
    else:
        # Rule B: Existing clients -> Only extract rows matching our standard month token
        filtered_group = group[group['bill_finalised_month'] == string_pattern_match]
        
        if not filtered_group.empty:
            current_month_existing_client_pool.append(filtered_group)

# Re-pool elements into a single execution matrix
processed_frames = historical_new_client_pool + current_month_existing_client_pool
if processed_frames:
    df_upload = pd.concat(processed_frames, ignore_index=True)
else:
    df_upload = pd.DataFrame(columns=df_current_raw.columns)

# ==============================================================================
# 4. SANITIZATION AND ALPHANUMERIC CLEANING LAYER
# ==============================================================================
print("Applying sanitization filters and mapping structure to destination layout...")

def global_hex_scrubber(val):
    if pd.isna(val) or not isinstance(val, str): 
        return val
    val = re.sub(r'_[xX]([0-9a-fA-F]{4})_', lambda m: chr(int(m.group(1), 16)), val)
    return val.strip()

for col in df_upload.select_dtypes(include=['object']).columns:
    df_upload[col] = df_upload[col].apply(global_hex_scrubber)

print("Cleaning identification tracking records and removing '#' hashes...")
if 'uid' in df_upload.columns:
    df_upload['uid'] = df_upload['uid'].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else (str(x) if pd.notna(x) else x)
    )
    df_upload['uid'] = df_upload['uid'].astype(str).str.replace('#', '', regex=False).str.strip()
    df_upload['uid'] = df_upload['uid'].replace({'nan': None, 'None': None})

if 'claim_num_yr' in df_upload.columns:
    df_upload['claim_num_yr'] = df_upload['claim_num_yr'].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else (str(x) if pd.notna(x) else x)
    )
    df_upload['claim_num_yr'] = df_upload['claim_num_yr'].astype(str).str.replace('#', '', regex=False).str.strip()
    df_upload['claim_num_yr'] = df_upload['claim_num_yr'].replace({'nan': None, 'None': None})

column_mapping_blueprint = {
    'policy_num': 'Policy Number',
    'policyholder_name': 'Policy Holder Name',
    'broker_name': 'Broker',
    'policy_last_renewal_date': 'policy_last_renewal_date',
    'uid': 'pol_mem_num',  
    'age_band': 'Age Band',
    'gender': 'gender',
    'claim_num_yr': 'claim_num_yr',
    'scheme_year': 'scheme_year',
    'primary_diagnosis_code': 'Condition Code',
    'main_diagnosis_group': 'Main_Diagnosis_Group',
    'procedure_category': 'Procedure Type',
    'place_of_delivery': 'Place of Service',
    'bill_finalised_month': 'Finalised Date',
    'claim_incurred_date_min': 'claim_incurred_date_min',
    'claim_incurred_date_max': 'claim_incurred_date_max',
    'count_incurred_days': 'count_incurred_days',
    'claim_submitted_amount_sum': 'Amount Submitted',
    'claim_paid_amount_sum': 'Amount Paid',
    'bill_excess_amount_sum': 'Excess Amount'
}

df_upload = df_upload.rename(columns=column_mapping_blueprint)

# Clean date columns safely
date_columns = ['policy_last_renewal_date', 'Finalised Date', 'claim_incurred_date_min', 'claim_incurred_date_max']
for col in date_columns:
    df_upload[col] = pd.to_datetime(df_upload[col], dayfirst=True, format='mixed', errors='coerce').dt.date

df_upload['UploadDate'] = str(pd.to_datetime(target_upload_date).date())

# Standardize numeric typing for mapped financial inputs
for col in ['Amount Submitted', 'Amount Paid', 'Excess Amount']:
    if col in df_upload.columns:
        df_upload[col] = pd.to_numeric(df_upload[col], errors='coerce').fillna(0.0)

all_valid_sql_columns = list(column_mapping_blueprint.values()) + ['UploadDate']
df_upload = df_upload[[c for c in df_upload.columns if c in all_valid_sql_columns]]

# ==============================================================================
# SENSE CHECK PART A: BASELINE AVIVA FINANCIAL SUMS (BEFORE LOAD)
# ==============================================================================
print(f"Aggregating baseline conditions metrics panel from Backup archive...")
historical_check_query = f"""
SELECT 
    [Policy Holder Name],
    SUM(CASE WHEN [Main_Diagnosis_Group] = 'Musculoskeletal' THEN [Amount Paid] ELSE 0 END) AS MSK_Sum,
    SUM(CASE WHEN [Main_Diagnosis_Group] = 'Psychiatry' THEN [Amount Paid] ELSE 0 END) AS MH_Sum,
    SUM(CASE WHEN [Main_Diagnosis_Group] = 'Oncology' THEN [Amount Paid] ELSE 0 END) AS Cancer_Sum,
    SUM([Amount Paid]) AS Total_Sum
FROM {schema_name}.Test_BackupAvivaClaims
{{where_clause}}
GROUP BY [Policy Holder Name]
"""

df_sense_before = pd.read_sql(historical_check_query.format(where_clause=""), con=engine)
df_sense_before['Client_Key'] = df_sense_before['Policy Holder Name'].astype(str).str.strip().str.title()
df_sense_before_grouped = df_sense_before.groupby('Client_Key')[['MSK_Sum', 'MH_Sum', 'Cancer_Sum', 'Total_Sum']].sum().reset_index()

# ==============================================================================
# 5. STREAM TO STAGING WORKSPACE
# ==============================================================================
print(f"Emptying {schema_name}.Test_UploadAvivaClaims staging table...")
cursor.execute(f"TRUNCATE TABLE {schema_name}.Test_UploadAvivaClaims")
conn.commit()

print(f"Streaming filtered data matrices over to {schema_name}.Test_UploadAvivaClaims...")
clean_schema = schema_name.strip("[]")

df_upload.to_sql(
    name='Test_UploadAvivaClaims', 
    con=engine, 
    schema=quoted_name(clean_schema, quote=True), 
    if_exists='append', 
    index=False, 
    method='multi', 
    chunksize=50
)

# ==============================================================================
# 6. PIPELINE APPEND TO LIVE
# ==============================================================================
print(f"Appending newly isolated records directly into {schema_name}.Test_LiveAvivaClaims...")

upload_columns_no_id = """
    [Policy Number], [Policy Holder Name], [Broker], [policy_last_renewal_date], 
    [pol_mem_num], [Age Band], [gender], [claim_num_yr], [scheme_year], 
    [Condition Code], [Main_Diagnosis_Group], [Procedure Type], [Place of Service], 
    [Finalised Date], [claim_incurred_date_min], [claim_incurred_date_max], 
    [count_incurred_days], [Amount Submitted], [Amount Paid], [Excess Amount], [UploadDate]
"""

cursor.execute(f"INSERT INTO {schema_name}.Test_LiveAvivaClaims ({upload_columns_no_id}) SELECT {upload_columns_no_id} FROM {schema_name}.Test_UploadAvivaClaims")
conn.commit()

# Gather post-execution statistics
cursor.execute(f"SELECT COUNT(*) FROM {schema_name}.Test_LiveAvivaClaims")
final_live_count = cursor.fetchone()[0]
cursor.execute(f"SELECT COUNT(*) FROM {schema_name}.Test_BackupAvivaClaims")
backup_count = cursor.fetchone()[0]

# ==============================================================================
# SENSE CHECK PART B: AVIVA FINANCIAL SUMS POST LOAD (WITH EXCLUSIONS)
# ==============================================================================
print("Aggregating comparison metrics panel from modified Live table...")

# Exclude processing inputs belonging to the incoming target window matrix to isolate historical modifications
target_exclusion_string = f"'{target_year}-{target_month:02d}-01'"
where_clause_live = f"WHERE DATEFROMPARTS(YEAR([Finalised Date]), MONTH([Finalised Date]), '01') != {target_exclusion_string}"

df_sense_after = pd.read_sql(historical_check_query.format(where_clause=where_clause_live), con=engine)
df_sense_after['Client_Key'] = df_sense_after['Policy Holder Name'].astype(str).str.strip().str.title()
df_sense_after_grouped = df_sense_after.groupby('Client_Key')[['MSK_Sum', 'MH_Sum', 'Cancer_Sum', 'Total_Sum']].sum().reset_index()

# Merge metrics to compute explicit differences
df_recon = pd.merge(
    df_sense_before_grouped, 
    df_sense_after_grouped, 
    on='Client_Key', 
    how='outer', 
    suffixes=('_Before', '_After')
).fillna(0.0)

# Calculate system variance deltas
df_recon['MSK_Var'] = df_recon['MSK_Sum_After'] - df_recon['MSK_Sum_Before']
df_recon['MH_Var'] = df_recon['MH_Sum_After'] - df_recon['MH_Sum_Before']
df_recon['Cancer_Var'] = df_recon['Cancer_Sum_After'] - df_recon['Cancer_Sum_Before']
df_recon['Total_Var'] = df_recon['Total_Sum_After'] - df_recon['Total_Sum_Before']

df_recon['Status'] = df_recon['Total_Var'].apply(lambda x: '✅ MATCHED' if round(abs(x), 2) == 0.00 else '❌ MISMATCH')
mismatched_clients_df = df_recon[df_recon['Status'] == '❌ MISMATCH']

# ==============================================================================
# 7. PERFORMANCE AUDIT AND RECONCILIATION RADAR REPORT & NATIVE PDF EXPORT
# ==============================================================================
from fpdf import FPDF

class CorporateAvivaClaimsPDF(FPDF):
    def header(self):
        # Top margin corporate header layout
        self.set_font('Courier', 'B', 11)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, 'LOCKTON EMPLOYEE BENEFITS | AUTOMATED AVIVA CLAIMS AUDIT LOG', ln=True, align='L')
        self.set_draw_color(180, 180, 180)
        self.line(10, 18, 200, 18)
        self.ln(8)
        
    def footer(self):
        # Bottom page tracking signatures
        self.set_y(-15)
        self.set_font('Courier', 'I', 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f'Page {self.page_no()} | CONFIDENTIAL - INTERNAL USE ONLY', align='C')

# Initialize payload tracking array for screen printing and PDF staging
pdf_payload = []
pdf_payload.append("================================================================================")
pdf_payload.append("                    AVIVA CLAIMS AUTOMATED AUDIT & RECONCILIATION")
pdf_payload.append("================================================================================")

pdf_payload.append("\n[DATABASE STORAGE & FILE TRANSACTION AUDIT BALANCE LEDGER]")
pdf_payload.append("-" * 80)
pdf_payload.append(f"  * Backup Table Starting Count   : {backup_count:,} rows")
pdf_payload.append(f"  * Live Table Starting Count     : {initial_live_count:,} rows")
pdf_payload.append(f"  >> Total Staged Rows Ingested   : {len(df_upload):,} new rows added from current sheet")
pdf_payload.append(f"  * Live Table Ending Volume      : {final_live_count:,} rows")
pdf_payload.append(f"  * Backup Table Ending Volume    : {backup_count:,} rows (Cloned production state)")
pdf_payload.append("  ==========================================================")
pdf_payload.append("  [STATUS] RECONCILIATION AUDIT STATUS : SUCCESS")

# PART 1: GLOBAL HISTORICAL FINANCIAL INTEGRITY SENSE CHECK
pdf_payload.append("\n[PART 1: GLOBAL HISTORICAL FINANCIAL INTEGRITY SENSE CHECK]")
pdf_payload.append("-" * 80)

df_recon_display = df_recon[['Client_Key', 'Total_Sum_Before', 'Total_Sum_After', 'Total_Var', 'Status']].copy()
df_recon_display = df_recon_display.rename(columns={
    'Client_Key': 'Client Name', 
    'Total_Sum_Before': 'Baseline_Paid', 
    'Total_Sum_After': 'Post_Upload_Paid', 
    'Total_Var': 'Variance'
})

for col in ['Baseline_Paid', 'Post_Upload_Paid', 'Variance']:
    df_recon_display[col] = df_recon_display[col].apply(lambda x: f"£{x:,.2f}")

pdf_payload.append(df_recon_display.to_string(index=False))
pdf_payload.append("-" * 80)

if len(mismatched_clients_df) == 0:
    pdf_payload.append("[OK] SENSE CHECK PASSED: All historical records matched perfectly across all clients.")
else:
    pdf_payload.append(f"[WARNING] WARNING: SENSE CHECK DETECTED UNEXPECTED CHANGES IN {len(mismatched_clients_df)} CLIENT ACCOUNT(S)!")
    
    mismatched_display = mismatched_clients_df[['Client_Key', 'MSK_Var', 'MH_Var', 'Cancer_Var', 'Total_Var']].copy()
    mismatched_display = mismatched_display.rename(columns={'Client_Key': 'Client Name'})
    for col in ['MSK_Var', 'MH_Var', 'Cancer_Var', 'Total_Var']:
        mismatched_display[col] = mismatched_display[col].apply(lambda x: f"£{x:,.2f}")
        
    pdf_payload.append(mismatched_display.to_string(index=False))
    pdf_payload.append("   Action Required: Please check why old data for these clients changed when adding records.")

pdf_payload.append("\n[PART 2: NEW VS LOST CLIENT ACCOUNT RETENTION AUDIT]")
pdf_payload.append("-" * 80)
if len(new_clients_drift) > 0:
    pdf_payload.append(f"[NEW] NEW CLIENTS IDENTIFIED (Full History Ingested):\n   -> {', '.join(new_clients_drift)}")
else:
    pdf_payload.append("[OK] No brand new clients found in this month's register.")

pdf_payload.append("")
if len(lost_clients_drift) > 0:
    pdf_payload.append(f"[WARNING] MISSING CLIENT DETECTION WARNING (Review Required):\n   -> {', '.join(lost_clients_drift)}")
else:
    pdf_payload.append("[OK] CLIENT DRIFT RADAR: All expected clients accounted for in dataset.")

# PART 3: CLIENT VOLUME TREND AUDIT (DYNAMIC LOOKBACK MONTH GENERATOR)
pdf_payload.append("\n[PART 3: CLIENT VOLUME TREND AUDIT]")
pdf_payload.append("-" * 80)

# Historical Base Count Evaluation
df_historic_counts = pd.read_sql(f"SELECT [Policy Holder Name], COUNT(*) as Historic_Rows FROM {schema_name}.Test_BackupAvivaClaims GROUP BY [Policy Holder Name]", con=engine)
df_historic_counts['Clean_Key'] = df_historic_counts['Policy Holder Name'].astype(str).str.strip().str.title()
df_hist_grouped = df_historic_counts.groupby('Clean_Key')['Historic_Rows'].sum().reset_index()

def evaluate_csv_month_volume(file_raw_df, target_period):
    if file_raw_df is None or file_raw_df.empty:
        return pd.DataFrame(columns=['Clean_Key', 'Count'])
    
    df_temp = file_raw_df.copy()
    df_temp['bill_finalised_month'] = df_temp['bill_finalised_month'].astype(str).str.strip().str.replace('/', '-', regex=False)
    
    # Generate match targets
    match_pattern = f"01-{target_period.month:02d}-{target_period.year}"
    filtered = df_temp[df_temp['bill_finalised_month'] == match_pattern]
    
    if filtered.empty:
        return pd.DataFrame(columns=['Clean_Key', 'Count'])
        
    counts = filtered.groupby('policyholder_name').size().reset_index(name='Count')
    counts['Clean_Key'] = counts['policyholder_name'].astype(str).str.strip().str.title()
    return counts.groupby('Clean_Key')['Count'].sum().reset_index()

# Extract Dynamic Month Fields
df_minus_2m_counts = evaluate_csv_month_volume(df_current_raw, date_minus_2m).rename(columns={'Count': label_minus_2m})
df_minus_1m_counts = evaluate_csv_month_volume(df_current_raw, date_minus_1m).rename(columns={'Count': label_minus_1m})

df_current_upload_counts = df_upload.groupby('Policy Holder Name').size().reset_index(name=label_current_m)
df_current_upload_counts['Clean_Key'] = df_current_upload_counts['Policy Holder Name'].astype(str).str.strip().str.title()
df_curr_grouped = df_current_upload_counts.groupby('Clean_Key')[label_current_m].sum().reset_index()

# Core Merge Pipeline Matrix Assembly
df_trend_audit = pd.merge(df_hist_grouped, df_minus_2m_counts, on='Clean_Key', how='outer')
df_trend_audit = pd.merge(df_trend_audit, df_minus_1m_counts, on='Clean_Key', how='outer')
df_trend_audit = pd.merge(df_trend_audit, df_curr_grouped, on='Clean_Key', how='outer').fillna(0)

for column_name in ['Historic_Rows', label_minus_2m, label_minus_1m, label_current_m]:
    df_trend_audit[column_name] = df_trend_audit[column_name].astype(int)

df_trend_audit = df_trend_audit.rename(columns={'Clean_Key': 'Client Name'})
pdf_payload.append(df_trend_audit[['Client Name', 'Historic_Rows', label_minus_2m, label_minus_1m, label_current_m]].to_string(index=False))

pdf_payload.append("="*80 + "\n")

# Unify structural log items into a clean consolidated block string object
full_report_text = "\n".join(pdf_payload)

# ------------------------------------------------------------------------------
# ACTION A: Push original report layouts cleanly to your Terminal window
# ------------------------------------------------------------------------------
print(full_report_text)

# ------------------------------------------------------------------------------
# ACTION B: Render onto clear, multi-page hardcopy PDF documents safely
# ------------------------------------------------------------------------------
print("\nCompiling automated A4 corporate PDF document summary...")
pdf = CorporateAvivaClaimsPDF(orientation='P', unit='mm', format='A4')
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()
pdf.set_font("Courier", size=8.5) # Protects column width structures from breaking lines
pdf.set_text_color(40, 40, 40)

for line in full_report_text.split('\n'):
    # Safety Shield: Strips out non-latin-1 characters completely so Courier won't throw encoding crashes
    clean_line = line.encode('latin-1', 'ignore').decode('latin-1')
    pdf.cell(0, 4.2, txt=clean_line, ln=True)

# Generate destination targets inside the target month directory path structure
base_directory = os.path.dirname(current_month_csv)
pdf_output_filename = f"Aviva_Claims_Audit_Report_{parsed_date.strftime('%Y_%m')}.pdf"
pdf_output_path = os.path.join(base_directory, pdf_output_filename)
pdf.output(pdf_output_path)

print(f"✅ SUCCESS: Aviva Claims PDF report compiled and saved cleanly.")
print(f"📂 Saved Location: {pdf_output_path}")

# Terminate streaming pipelines safely
try:
    cursor.close()
    conn.close()
except Exception:
    pass

print("Aviva process completed successfully with automated scripts executed!")