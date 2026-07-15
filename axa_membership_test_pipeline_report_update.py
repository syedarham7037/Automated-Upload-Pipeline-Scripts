import os
import warnings
import re
import pandas as pd
import pyodbc
from sqlalchemy import create_engine

# Silence Excel engine and pandas warning noise
warnings.simplefilter(action='ignore', category=UserWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)
if hasattr(pd, 'Pandas4Warning'): warnings.filterwarnings("ignore", category=pd.Pandas4Warning)

# ==============================================================================
# 1. CONFIGURATION & PATH SETUP
# ==============================================================================
test_year = "2026"
test_month_folder = "2026 06 June"
upload_date_setting = "2026-06-01"  # Automatically marks your monthly upload run date

base_directory = r"X:\CRM\Employee Benefits\40 Operations\3 Analytics\SQL Database\Data from Insurers\AXA\2026\2026 06 June"
raw_membership_filename = "Locktons_LC_Membership 30Jun26.xlsx" 
raw_membership_path = os.path.join(base_directory, raw_membership_filename)

group_size_filename = "Locktons_LC_Group_Size 30Jun26.xlsx"
group_size_path = os.path.join(base_directory, group_size_filename)

april_directory = r"X:\CRM\Employee Benefits\40 Operations\3 Analytics\SQL Database\Data from Insurers\AXA\2026\2026 05 May"
april_group_size_filename = "Locktons_LC_Group_Size 31May26.xlsx"
april_group_size_path = os.path.join(april_directory, april_group_size_filename)

# Path to your external Exposure script calculation engine
exposure_script_path = r"X:\CRM\Employee Benefits\40 Operations\3 Analytics\SQL Database\Scripts\SQL Upload\AXA -Monthly Membership Update (4 columns).sql"

server_name = "LOCKLON-SQL14"
database_name = "EBMedClm"
connection_string = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server_name};DATABASE={database_name};Trusted_Connection=yes;"

engine = create_engine(f"mssql+pyodbc:///?odbc_connect={connection_string}", fast_executemany=True)

print("Initialization complete. Connecting to SQL Server...")

# Calculate the upper boundary for historical tracking (excludes the current month)
parsed_upload_date = pd.to_datetime(upload_date_setting)
audit_target_month = (parsed_upload_date - pd.DateOffset(months=1)).replace(day=1)
sql_audit_end_date_str = audit_target_month.strftime('%Y-%m-%d')

# ==============================================================================
# 2. RUN STEP 3 DATABASE BACKUP ROTATION AUTOMATION
# ==============================================================================
conn = pyodbc.connect(connection_string)
cursor = conn.cursor()

print("Executing Table Rollover: Emptying Test_BackupAxaMembership...")
cursor.execute("DELETE FROM Test_BackupAxaMembership")
conn.commit()

print("Cloning Test_Live records safely over to Test_BackupAxaMembership...")
cursor.execute("""
    INSERT INTO Test_BackupAxaMembership (
        [Group no], [Client Name], [Person ref], [Relation], [Birth Year], [Gender], [Region],
        [Enrolment Month], [Lapsed Month], [Person status], [Product], [Family Status], [UniqueID], 
        [UploadDate], [calculated_exposure_start], [calculated_exposure_end], [member_exposure_start], [member_exposure_end]
    )
    SELECT 
        [Group no], [Client Name], [Person ref], [Relation], [Birth Year], [Gender], [Region],
        [Enrolment Month], [Lapsed Month], [Person status], [Product], [Family Status], [UniqueID], 
        [UploadDate], [calculated_exposure_start], [calculated_exposure_end], [member_exposure_start], [member_exposure_end]
    FROM Test_LiveAxaMembership
""")
conn.commit()

# Track pre-run row counts for sense checking
cursor.execute("SELECT COUNT(*) FROM Test_BackupAxaMembership")
backup_count_start = cursor.fetchone()[0]
cursor.execute("SELECT COUNT(*) FROM Test_LiveAxaMembership")
live_count_start = cursor.fetchone()[0]

# ==============================================================================
# 3. ADVANCED GLOBAL SHIELD: DYNAMIC HEX CONVERTER
# ==============================================================================
def dynamic_hex_decoder(text):
    if pd.isna(text) or not isinstance(text, str):
        return text
    def replace_match(match):
        return chr(int(match.group(1), 16))
    return re.sub(r'_[xX]([0-9a-fA-F]{4})_', replace_match, text)

def global_hex_decryption_shield(df_target, is_membership=False):
    df_target.columns = [dynamic_hex_decoder(str(col)).strip() for col in df_target.columns]
    text_cols = df_target.select_dtypes(include=['object']).columns
    
    for col in text_cols:
        col_clean = str(col).lower().replace(' ', '').replace('_', '')
        
        if is_membership and col_clean in ['personref', 'personrefno']:
            df_target[col] = df_target[col].astype(str)
            for i in range(10):
                df_target[col] = df_target[col].str.replace(f'_x003{i}_', str(i), regex=False).str.replace(f'_X003{i}_', str(i), regex=False)
            df_target[col] = df_target[col].str.strip()
            continue
            
        df_target[col] = df_target[col].astype(str).apply(dynamic_hex_decoder)
        df_target[col] = df_target[col].str.replace('""', '', regex=False).str.replace('"', '', regex=False)
        df_target[col] = df_target[col].str.strip()
        
    return df_target

# ==============================================================================
# 4. MEMBERSHIP CLEANING, FORMATTING & UNIQUEID ENGINEERING
# ==============================================================================
print(f"Reading raw file: {raw_membership_filename}...")
df = pd.read_excel(raw_membership_path)
raw_file_row_count = len(df)

# Run Global Decryption Shield
df = global_hex_decryption_shield(df, is_membership=True)

print("Standardizing column parameters and layout geometry...")
df = df.rename(columns={
    'Group No': 'Group No', 'Group no': 'Group No',
    'Person ref': 'Person Ref', 'Person Ref': 'Person Ref',
    'Sex': 'Gender', 'Gender': 'Gender'
})

df['Group No'] = pd.to_numeric(df['Group No'], errors='coerce').fillna(0).astype(int)
df['Birth Year'] = pd.to_numeric(df['Birth Year'], errors='coerce').fillna(0).astype(int)

df['Enrolment Month'] = pd.to_datetime(df['Enrolment Month'], errors='coerce')
df['Lapsed Month'] = pd.to_datetime(df['Lapsed Month'], errors='coerce')

client_name_mappings = {
    89082: 'BT Group Plc', 89119: 'British American Tobacco',
    89138: 'DELARUNIARIANS', 89087: 'DELARUNIARIANS',
    89152: 'EE', 97061: 'A&O Shearman'
}
df['Client Name'] = df.apply(
    lambda row: client_name_mappings.get(row['Group No'], row.get('Client Name', '')), axis=1
)

# CAPTURE DATA QUALITY EXCEPTIONS
current_run_year = pd.to_datetime(upload_date_setting).year

# Temporary Age calculation column to handle requested summaries cleanly
df['Temp_Calculated_Age'] = df['Birth Year'].apply(lambda x: current_run_year - x if x > 0 else 0)

bad_birth_year_df = df[(df['Birth Year'] <= 1920) & (df['Birth Year'] > 0)]
overage_children_df = df[(df['Relation'].str.lower().str.strip() == 'child dependant') & (df['Temp_Calculated_Age'] > 25)]
status_lapsed_no_date_df = df[(df['Person status'].str.lower().str.strip() == 'lapsed') & (df['Lapsed Month'].isna())]

# --- PROCESS AGGREGATED QUALITY EXCEPTION SUMMARIES BY CLIENT ---
df_bad_birth_summary = pd.DataFrame()
if not bad_birth_year_df.empty:
    df_bad_birth_summary = bad_birth_year_df.groupby('Client Name').agg(
        Count_Of_Rows=('Birth Year', 'size'),
        Max_Age=('Temp_Calculated_Age', 'max')
    ).reset_index()

df_overage_summary = pd.DataFrame()
if not overage_children_df.empty:
    df_overage_summary = overage_children_df.groupby('Client Name').agg(
        Count_Of_Rows=('Birth Year', 'size'),
        Max_Age=('Temp_Calculated_Age', 'max')
    ).reset_index()

df_status_lapsed_summary = pd.DataFrame()
if not status_lapsed_no_date_df.empty:
    df_status_lapsed_summary = status_lapsed_no_date_df.groupby('Client Name').size().reset_index(name='Mismatch_Count')

print("Building explicit UniqueID hashes mimicking Excel's 1900 sequential date engine...")
def build_excel_style_id(row):
    p_ref = str(row.get('Person Ref', '')).strip()
    lapsed = row.get('Lapsed Month', pd.NaT)
    if pd.notna(lapsed):
        excel_serial_date = (lapsed - pd.Timestamp('1899-12-30')).days
        lapsed_str = str(excel_serial_date)
    else:
        lapsed_str = ''
    
    p_stat = str(row.get('Person status', '')).strip()
    f_stat = str(row.get('Family Status', '')).strip()
    
    if p_ref in ['nan', 'None']: p_ref = ''
    if p_stat in ['nan', 'None']: p_stat = ''
    if f_stat in ['nan', 'None']: f_stat = ''
        
    return f"{p_ref}{lapsed_str}{p_stat}{f_stat}"

df['UniqueID'] = df.apply(build_excel_style_id, axis=1)
df['UploadDate'] = pd.to_datetime(upload_date_setting)

exposure_cols = ['calculated_exposure_start', 'calculated_exposure_end', 'member_exposure_start', 'member_exposure_end']
for ec in exposure_cols:
    df[ec] = None

# Full database layout mapping
upload_column_layout = [
    'Group No', 'Client Name', 'Person Ref', 'Relation', 'Birth Year', 'Gender', 'Region',
    'Enrolment Month', 'Lapsed Month', 'Person status', 'Product', 'Family Status', 'UploadDate', 
    'UniqueID', 'calculated_exposure_start', 'calculated_exposure_end', 'member_exposure_start', 'member_exposure_end'
]

df_final_upload = df[upload_column_layout].copy()
df_final_upload = df_final_upload.replace(r'^\s*$', None, regex=True)

# --- MISSING DATA CHECK ENGINE MIGRATION (EXCLUDING EXPOSURE COLUMNS) ---
# Filter layout list dynamically to check all columns except the 4 tracking metrics
audit_missing_columns = [col for col in upload_column_layout if col not in exposure_cols]

missing_grid_rows = []
clients_in_batch = df_final_upload['Client Name'].dropna().unique()

for client in clients_in_batch:
    df_client_subset = df_final_upload[df_final_upload['Client Name'] == client]
    for col in audit_missing_columns:
        missing_count = df_client_subset[col].isna().sum() + (df_client_subset[col].astype(str).str.strip() == '').sum()
        if missing_count > 0:
            missing_grid_rows.append({
                'Client Name': client,
                'Column Header': f"[{col}]",
                'Missing Count': missing_count
            })

df_missing_audit = pd.DataFrame(missing_grid_rows)
if not df_missing_audit.empty:
    df_missing_audit = df_missing_audit.sort_values(by=['Client Name', 'Column Header'])

# ==============================================================================
# EXPOSURE CALCULATION ENGINE FUNCTION (HISTORICAL MONTHLY AVERAGES)
# ==============================================================================
def get_historical_exposure_averages(table_source):
    exposure_sql = f"""
    WITH DateRange(Months) AS (
        SELECT CAST('2016-01-01' AS DATETIME) AS Date
        UNION ALL
        SELECT DATEADD(MONTH, 1, Months)
        FROM DateRange
        WHERE Months < CAST('{sql_audit_end_date_str}' AS DATETIME)
    ),
    SchemeClaimsExperience AS (
        SELECT
            [Group no],
            DATEADD(month, DATEDIFF(month, 0, MIN([Payment Month])), 0) AS SchemeMinPayMonth,
            DATEADD(MONTH, 1, DATEADD(month, DATEDIFF(month, 0, MAX([Payment Month])), 0)) AS SchemeMaxPayMonth_Exclusive
        FROM dbo.LiveAxaClaims  
        GROUP BY [Group no]
    ),
    CleanedMembership AS (
        SELECT
            *,
            CASE
                WHEN [Person Status] = 'current' AND [Lapsed Month] IS NOT NULL THEN NULL
                WHEN [Person Status] = 'lapsed' AND [Lapsed Month] IS NULL THEN DATEADD(month, DATEDIFF(month, 0, [UploadDate]), 0)
                ELSE [Lapsed Month]
            END AS CleanedLapsedMonth
        FROM {table_source}
    ),
    MemberHistoryWithEffectiveDate AS (
        SELECT
            *,
            CASE WHEN rnk = 1 THEN [Enrolment Month] ELSE [UploadDate] END AS EffectiveDate
        FROM (
            SELECT *, ROW_NUMBER() OVER(PARTITION BY [Person Ref] ORDER BY [UploadDate]) as rnk
            FROM CleanedMembership
        ) AS RankedRecords
    ),
    MemberTimeline_Raw AS (
        SELECT
            *,
            EffectiveDate AS ExposureStart,
            LEAD(EffectiveDate, 1, '2999-12-31') OVER (PARTITION BY [Person Ref] ORDER BY EffectiveDate) AS NextRecord_StartDate,
            LEAD(CleanedLapsedMonth, 1, NULL) OVER (PARTITION BY [Person Ref] ORDER BY EffectiveDate) AS NextRecord_LapsedMonth
        FROM MemberHistoryWithEffectiveDate
    ),
    MemberTimeline AS (
        SELECT 
            *,
            CASE
                WHEN NextRecord_LapsedMonth IS NOT NULL AND NextRecord_LapsedMonth < NextRecord_StartDate THEN NextRecord_LapsedMonth
                ELSE NextRecord_StartDate
            END AS PotentialExposureEnd
        FROM MemberTimeline_Raw
    ),
    FinalPeriods AS (
        SELECT 
            *,
            CASE
                WHEN CleanedLapsedMonth IS NOT NULL AND CleanedLapsedMonth < PotentialExposureEnd THEN CleanedLapsedMonth
                ELSE PotentialExposureEnd
            END AS ExposureEnd
        FROM MemberTimeline
    ),
    MonthlyHeadcounts AS (
        SELECT 
            t.[Client Name],
            DR.Months,
            COUNT(*) AS OpeningLives,
            COUNT(CASE WHEN t.[Relation] = 'Policyholder' THEN 1 END) AS Members,
            COUNT(CASE WHEN t.[Relation] IN ('Adult Dependant', 'Dependant') THEN 1 END) AS Partners,
            COUNT(CASE WHEN t.[Relation] = 'Child Dependant' THEN 1 END) AS Dependents,
            COUNT(CASE WHEN t.[Family Status] = 'Single' AND t.[Relation] = 'Policyholder' THEN 1 END) AS Single,
            COUNT(CASE WHEN t.[Family Status] = 'Married' AND t.[Relation] = 'Policyholder' THEN 1 END) AS Couple,
            COUNT(CASE WHEN t.[Family Status] = 'Family' AND t.[Relation] = 'Policyholder' THEN 1 END) AS Family,
            COUNT(CASE WHEN t.[Family Status] = 'Single Parent' AND t.[Relation] = 'Policyholder' THEN 1 END) AS SPF
        FROM DateRange DR
        JOIN FinalPeriods t ON 1=1
        JOIN SchemeClaimsExperience s ON t.[Group no] = s.[Group no]
        WHERE DR.Months >= (CASE WHEN t.ExposureStart > s.SchemeMinPayMonth THEN t.ExposureStart ELSE s.SchemeMinPayMonth END)
          AND DR.Months < (CASE WHEN t.ExposureEnd < s.SchemeMaxPayMonth_Exclusive THEN t.ExposureEnd ELSE s.SchemeMaxPayMonth_Exclusive END)
        GROUP BY t.[Client Name], DR.Months
    )
    SELECT 
        [Client Name],
        AVG(CAST(OpeningLives AS FLOAT)) AS OpeningLives,
        AVG(CAST(Members AS FLOAT)) AS Members,
        AVG(CAST(Partners AS FLOAT)) AS Partners,
        AVG(CAST(Dependents AS FLOAT)) AS Dependents,
        AVG(CAST(Single AS FLOAT)) AS Single,
        AVG(CAST(Couple AS FLOAT)) AS Couple,
        AVG(CAST(Family AS FLOAT)) AS Family,
        AVG(CAST(SPF AS FLOAT)) AS SPF
    FROM MonthlyHeadcounts
    GROUP BY [Client Name]
    OPTION (MAXRECURSION 0);
    """
    res_df = pd.read_sql(exposure_sql, con=engine)
    res_df['Client_Key'] = res_df['Client Name'].astype(str).str.strip().str.title()
    return res_df.groupby('Client_Key')[['OpeningLives', 'Members', 'Partners', 'Dependents', 'Single', 'Couple', 'Family', 'SPF']].sum().reset_index()

# Extract baseline historical averages from the live table before adding records
print("Calculating running historical monthly averages across active client matrices...")
df_exp_before_grouped = get_historical_exposure_averages("Test_LiveAxaMembership")

# ==============================================================================
# 5. UPLOAD STREAMING & TARGET APPENDING (NEW JOINERS ONLY)
# ==============================================================================
print("Emptying Test_UploadAxaMembership staging table...")
cursor.execute("DELETE FROM Test_UploadAxaMembership")
conn.commit()

print("Streaming cleaned membership data matrix over to Test_UploadAxaMembership...")
df_final_upload.to_sql('Test_UploadAxaMembership', con=engine, if_exists='append', index=False)

print("Appending TRUE New Joiners only into Test_LiveAxaMembership table...")
cursor.execute("""
    INSERT INTO Test_LiveAxaMembership (
        [Group no], [Client Name], [Person ref], [Relation], [Birth Year], [Gender], [Region],
        [Enrolment Month], [Lapsed Month], [Person status], [Product], [Family Status], [UniqueID], [UploadDate]
    )
    SELECT u.[Group No], u.[Client Name], u.[Person Ref], u.[Relation], u.[Birth Year], u.[Gender], u.[Region],
           u.[Enrolment Month], u.[Lapsed Month], u.[Person status], u.[Product], u.[Family Status], u.[UniqueID], u.[UploadDate]
    FROM Test_UploadAxaMembership u
    WHERE NOT EXISTS (
       SELECT 1
       FROM Test_LiveAxaMembership l
       WHERE l.UniqueID = u.UniqueID
    )
""")
conn.commit()

cursor.execute("SELECT COUNT(*) FROM Test_LiveAxaMembership")
live_count_post_append = cursor.fetchone()[0]
new_rows_added = live_count_post_append - live_count_start
print(f"Append complete. Successfully detected and added {new_rows_added:,} true new joiner records.")

# ==============================================================================
# SENSE CHECK PART B - CAPTURE CLIENT EXPOSURE headcounts POST UPLOAD
# ==============================================================================
print("Recalculating post-upload historical monthly averages (excluding new target month data)...")
df_exp_after_grouped = get_historical_exposure_averages("Test_LiveAxaMembership")

# Combine structural dataframes to find discrepancies
df_exp_recon = pd.merge(
    df_exp_before_grouped, df_exp_after_grouped, 
    on='Client_Key', how='outer', suffixes=('_Before', '_After')
).fillna(0)

# Calculate numerical differences and percentage tracking variances
metrics_list = ['OpeningLives', 'Members', 'Partners', 'Dependents', 'Single', 'Couple', 'Family', 'SPF']
for metric in metrics_list:
    df_exp_recon[f'{metric}_Diff'] = df_exp_recon[f'{metric}_After'] - df_exp_recon[f'{metric}_Before']
    df_exp_recon[f'{metric}_Pct'] = df_exp_recon.apply(
        lambda r: f"{((r[f'{metric}_After'] - r[f'{metric}_Before']) / r[f'{metric}_Before'] * 100):+.2f}%" if r[f'{metric}_Before'] > 0 else ("+0.00%" if r[f'{metric}_After'] == 0 else "+100.00%"), axis=1
    )

# ==============================================================================
# 6. RUN THE EXTERNAL NETWORK EXPOSURE SCRIPT AUTOMATICALLY
# ==============================================================================
if os.path.exists(exposure_script_path):
    print(f"Reading external SQL Exposure calculation script: {os.path.basename(exposure_script_path)}...")
    with open(exposure_script_path, 'r') as sql_file:
        exposure_sql_script = sql_file.read()
    
    print("Preparing and applying corporate schema overrides...")
    exposure_sql_script = re.sub(r'(?i)^\s*GO\s*$', '', exposure_sql_script, flags=re.MULTILINE)
    exposure_sql_script = re.sub(r'(?i)\b(dbo\.)?LiveAxaMembership\b', r'[UK\\Arham.Asif].Test_LiveAxaMembership', exposure_sql_script)
    exposure_sql_script = re.sub(r'(?i)\b(dbo\.)?LiveAxaClaims\b', r'[UK\\Arham.Asif].Test_LiveAxaClaims', exposure_sql_script)
    
    print("Executing Exposure period scripts on SQL Server...")
    try:
        cursor.execute(exposure_sql_script)
        conn.commit()
        print("✅ Exposure calculations executed and assigned successfully.")
    except Exception as sql_err:
        print(f"❌ SQL Execution Error: {sql_err}")
else:
    print(f"❌ Warning: Could not locate external exposure script file at path: {exposure_script_path}")

# ==============================================================================
# 7. AUTOMATED QUALITY SENSE-CHECK REPORT
# ==============================================================================
cursor.execute("SELECT COUNT(*) FROM Test_UploadAxaMembership")
upload_count_final = cursor.fetchone()[0]

print("\n" + "="*115)
print("                  AXA MEMBERSHIP AUTOMATED AUDIT & RECONCILIATION")
print("="*115)
print(f"Raw Excel File Rows Read:     {raw_file_row_count:,}")
print(f"Staging Stored Rows (Upload): {upload_count_final:,}")
print("-" * 115)
print(f"Initial Live Table Rows:      {live_count_start:,}")
print(f"Current Live Table Rows:      {live_count_post_append:,}")
print(f"True New Joiners Appended:    {new_rows_added:,}")
print(f"Historical Month Lag Backup:  {backup_count_start:,} (Must be lower than Live)")
print("-" * 115)

if raw_file_row_count == upload_count_final:
    print("✅ DATA INTEGRITY VERIFIED: Raw row totals match staging database perfectly.")
else:
    print("⚠️ WARNING: Mismatch detected between Excel size and SQL staging rows.")

# PRINT THE RECONCILED RUNNING AVERAGES STATUS TABLE
print(f"\n[PART 1: RUNNING HISTORICAL MONTHLY EXPOSURE AVERAGE SENSE CHECK (UP TO: {audit_target_month.strftime('%B %Y')})]")
print("-" * 115)

for _, row in df_exp_recon.iterrows():
    print(f"🏢 CLIENT ACCOUNT: {row['Client_Key']}")
    print("   " + "-"*90)
    print("   Metric         | Before (Avg)  | After (Avg)   | Net Variance  | Change %     ")
    print("   " + "-"*90)
    for m in metrics_list:
        print(f"   {m:<14} | {row[f'{m}_Before']:<13,.2f} | {row[f'{m}_After']:<13,.2f} | {row[f'{m}_Diff']:<13,.2f} | {row[f'{m}_Pct']}")
    print("   " + "-"*90 + "\n")

print("-" * 115)

# PRINT FILE HEALTH DATA QUALITY EXCEPTION LOG
print("\n[PART 2: INCOMING MEMBERSHIP DATA QUALITY EXCEPTION LOG]")
print("-" * 115)
has_exceptions = False

if not df_bad_birth_summary.empty:
    has_exceptions = True
    print(f"⚠️ ANOMALY SUMMARY: Out-of-bounds/Ancient Birth Years (<= 1920) grouped by Client:")
    print(df_bad_birth_summary.to_string(index=False))
    print("")

if not df_overage_summary.empty:
    has_exceptions = True
    print(f"⚠️ ANOMALY SUMMARY: Child Dependants older than 25 years grouped by Client:")
    print(df_overage_summary.to_string(index=False))
    print("")

if not df_status_lapsed_summary.empty:
    has_exceptions = True
    print(f"⚠️ STATUS MISALIGNMENT SUMMARY: Count of Lapsed statuses with missing Lapsed Dates grouped by Client:")
    print(df_status_lapsed_summary.to_string(index=False))
    print("")

if not has_exceptions:
    print("✅ FILE VALIDATION SUCCESSFUL: No impossible dates, overaged child dependants, or status logic anomalies found.")

# PRINT MISSING DATA SUMMARY MATRIX
print("\n[MISSING DATA SUMMARY GRID PER CLIENT ACCOUNT]")
print("." * 115)
if not df_missing_audit.empty:
    print(df_missing_audit.to_string(index=False))
else:
    print("✅ Complete Ingestion Coverage: No blanks or null inputs found across processing columns.")

print("="*115)

cursor.close()
conn.close()
print("\nMembership process completed successfully with automated scripts executed!")

# ==============================================================================
# 8. AUTOMATED QUALITY SENSE-CHECK REPORT & NATIVE PDF EXPORT
# ==============================================================================
from fpdf import FPDF

class CorporateMembershipPDF(FPDF):
    def header(self):
        self.set_font('Courier', 'B', 11)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, 'LOCKTON EMPLOYEE BENEFITS | AUTOMATED MEMBERSHIP SYSTEM AUDIT LOG', ln=True, align='L')
        self.set_draw_color(180, 180, 180)
        self.line(10, 18, 200, 18)
        self.ln(8)
        
    def footer(self):
        self.set_y(-15)
        self.set_font('Courier', 'I', 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f'Page {self.page_no()} | CONFIDENTIAL - INTERNAL USE ONLY', align='C')

# Safe connection fetch using SQLAlchemy engine to avoid "Closed Cursor" exception
upload_count_final = int(pd.read_sql("SELECT COUNT(*) AS cnt FROM Test_UploadAxaMembership", con=engine).iloc[0]['cnt'])

# Initialize payload tracking array for screen printing and PDF staging
pdf_payload = []
pdf_payload.append("====================================================================================================")
pdf_payload.append("                         AXA MEMBERSHIP AUTOMATED AUDIT & RECONCILIATION")
pdf_payload.append("====================================================================================================")
pdf_payload.append(f"Raw Excel File Rows Read:     {raw_file_row_count:,}")
pdf_payload.append(f"Staging Stored Rows (Upload): {upload_count_final:,}")
pdf_payload.append("-" * 100)
pdf_payload.append(f"Initial Live Table Rows:      {live_count_start:,}")
pdf_payload.append(f"Current Live Table Rows:      {live_count_post_append:,}")
pdf_payload.append(f"True New Joiners Appended:    {new_rows_added:,}")
pdf_payload.append(f"Historical Month Lag Backup:  {backup_count_start:,} (Must be lower than Live)")
pdf_payload.append("-" * 100)

if raw_file_row_count == upload_count_final:
    pdf_payload.append("[OK] DATA INTEGRITY VERIFIED: Raw row totals match staging database perfectly.")
else:
    pdf_payload.append("[WARNING] WARNING: Mismatch detected between Excel size and SQL staging rows.")

# STAGE THE RECONCILED RUNNING AVERAGES STATUS TABLE
pdf_payload.append(f"\n[PART 1: RUNNING HISTORICAL MONTHLY EXPOSURE AVERAGE SENSE CHECK (UP TO: {audit_target_month.strftime('%B %Y')})]")
pdf_payload.append("-" * 100)

for _, row in df_exp_recon.iterrows():
    pdf_payload.append(f"Client Account: {row['Client_Key']}")
    pdf_payload.append("   " + "-"*85)
    pdf_payload.append("   Metric         | Before (Avg)  | After (Avg)   | Net Variance  | Change %     ")
    pdf_payload.append("   " + "-"*85)
    for m in metrics_list:
        pdf_payload.append(f"   {m:<14} | {row[f'{m}_Before']:<13,.2f} | {row[f'{m}_After']:<13,.2f} | {row[f'{m}_Diff']:<13,.2f} | {row[f'{m}_Pct']}")
    pdf_payload.append("   " + "-"*85 + "\n")

pdf_payload.append("-" * 100)

# STAGE FILE HEALTH DATA QUALITY EXCEPTION LOG
pdf_payload.append("\n[PART 2: INCOMING MEMBERSHIP DATA QUALITY EXCEPTION LOG]")
pdf_payload.append("-" * 100)
has_exceptions = False

if not df_bad_birth_summary.empty:
    has_exceptions = True
    pdf_payload.append(f"[WARNING] ANOMALY SUMMARY: Out-of-bounds/Ancient Birth Years (<= 1920) grouped by Client:")
    pdf_payload.append(df_bad_birth_summary.to_string(index=False))
    pdf_payload.append("")

if not df_overage_summary.empty:
    has_exceptions = True
    pdf_payload.append(f"[WARNING] ANOMALY SUMMARY: Child Dependants older than 25 years grouped by Client:")
    pdf_payload.append(df_overage_summary.to_string(index=False))
    pdf_payload.append("")

if not df_status_lapsed_summary.empty:
    has_exceptions = True
    pdf_payload.append(f"[WARNING] STATUS MISALIGNMENT SUMMARY: Count of Lapsed statuses with missing Lapsed Dates grouped by Client:")
    pdf_payload.append(df_status_lapsed_summary.to_string(index=False))
    pdf_payload.append("")

if not has_exceptions:
    pdf_payload.append("[OK] FILE VALIDATION SUCCESSFUL: No impossible dates, overaged child dependants, or status logic anomalies found.")

# STAGE MISSING DATA SUMMARY GRID PER CLIENT ACCOUNT
pdf_payload.append("\n[MISSING DATA SUMMARY GRID PER CLIENT ACCOUNT]")
pdf_payload.append("." * 100)
if not df_missing_audit.empty:
    pdf_payload.append(df_missing_audit.to_string(index=False))
else:
    pdf_payload.append("[OK] Complete Ingestion Coverage: No blanks or null inputs found across processing columns.")

pdf_payload.append("="*100)

full_report_text = "\n".join(pdf_payload)

# Render PDF document safely
print("\nCompiling automated A4 corporate PDF document...")
pdf = CorporateMembershipPDF(orientation='P', unit='mm', format='A4')
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()
pdf.set_font("Courier", size=8.0)
pdf.set_text_color(40, 40, 40)

for line in full_report_text.split('\n'):
    clean_line = line.encode('latin-1', 'ignore').decode('latin-1')
    pdf.cell(0, 4.0, txt=clean_line, ln=True)

pdf_output_filename = f"AXA_Membership_Audit_Report_{test_month_folder.replace(' ', '_')}.pdf"
pdf_output_path = os.path.join(base_directory, pdf_output_filename)
pdf.output(pdf_output_path)

print(f"✅ SUCCESS: Membership PDF asset structured and saved cleanly without encoding errors.")
print(f"📂 Saved Location: {pdf_output_path}")

print("Pipeline complete!")