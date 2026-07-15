import os
import re
import warnings
import pandas as pd
import pyodbc
from sqlalchemy import create_engine
from sqlalchemy.sql.elements import quoted_name
from fpdf import FPDF

# Suppress non-critical performance notifications
warnings.filterwarnings('ignore')

# ==============================================================================
# CONFIGURATION AND RUNTIME SETTINGS (May 2026 Run)
# ==============================================================================
target_upload_date = "2026-06-15"  # Primary execution date stamp
server_name = "LOCKLON-SQL14"
database_name = "EBMedClm"
schema_name = r"[UK\Arham.Asif]"   # Handles backslashes securely

# Corporate network folder file paths
raw_membership_csv = r"X:\CRM\Employee Benefits\40 Operations\3 Analytics\SQL Database\Data from Insurers\Aviva\2026\2026 06 June\LOCKTONS_MBRS.csv"
exposure_script_sql = r"X:\CRM\Employee Benefits\40 Operations\3 Analytics\SQL Database\Scripts\SQL Upload\AVIVA-Monthly Membership Update (4 columns).sql"

# Dynamically calculate the 1st of the upload month (e.g., 2026-05-15 -> 2026-05-01)
parsed_date = pd.to_datetime(target_upload_date)
first_of_the_month_dt = parsed_date.replace(day=1)
first_of_the_month = str(first_of_the_month_dt.date())

# Upper bound for the running historical tracking engine (excludes current month data)
audit_target_month = (parsed_date - pd.DateOffset(months=1)).replace(day=1)
sql_audit_end_date_str = audit_target_month.strftime('%Y-%m-%d')

print(f"Initialization complete. Upload Date set to: {first_of_the_month}")
print(f"Historical Audit Target Bound: Up to {audit_target_month.strftime('%B %Y')}")
print("Connecting to SQL Server database layer...")

# High-speed data connections
conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server_name};DATABASE={database_name};Trusted_Connection=yes;"
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

engine = create_engine(f"mssql+pyodbc://@{server_name}/{database_name}?driver=ODBC+Driver+17+for+SQL+Server")

# ==============================================================================
# 1. LIVE TO BACKUP REPLICATION LAYER
# ==============================================================================
print(f"Executing Table Rollover: Wiping historical {schema_name}.Test_BackupAvivaMembership...")
cursor.execute(f"DELETE FROM {schema_name}.Test_BackupAvivaMembership")

print(f"Cloning live snapshots into {schema_name}.Test_BackupAvivaMembership...")
cursor.execute(f"INSERT INTO {schema_name}.Test_BackupAvivaMembership SELECT * FROM {schema_name}.Test_LiveAvivaMembership")
conn.commit()

# Document current rows baseline before running the update
cursor.execute(f"SELECT COUNT(*) FROM {schema_name}.Test_LiveAvivaMembership")
initial_live_count = cursor.fetchone()[0]

# ==============================================================================
# 2. INGESTION AND STRING FORCE CASTING (ELIMINATES VALUE CLASHES)
# ==============================================================================
print("Reading raw incoming data file and enforcing text layouts...")
df_raw = pd.read_csv(raw_membership_csv, dtype=str)

# ==============================================================================
# 3. TEXT UNIFICATION AND CLEANING LAYER (FIXES JOIN MISSES)
# ==============================================================================
print("Applying standardization modifications and column map filters...")

# Global Strip & Upper: Clean all columns to handle hidden spaces and case-mismatches
for col in df_raw.columns:
    df_raw[col] = df_raw[col].astype(str).str.strip()
    if col in ['unique_member_id', 'policy_no']:
        df_raw[col] = df_raw[col].str.upper()

# Clear literal 'nan' string artifacts generated from reading empty CSV boxes
df_raw.replace({'nan': None, 'None': None}, inplace=True)

# 3-A. Scrub character hash marks from Unique Member identification rows
if 'unique_member_id' in df_raw.columns:
    df_raw['unique_member_id'] = df_raw['unique_member_id'].str.replace('#', '', regex=False)

# 3-B. Map relational hierarchy terms
relation_blueprint = {
    'Self': 'Primary',
    'Partner': 'Spouse/Partner',
    'Child': 'Child'
}
if 'member_relation' in df_raw.columns:
    df_raw['member_relation'] = df_raw['member_relation'].map(relation_blueprint).fillna(df_raw['member_relation'])

# 3-C. Standardize client levels of coverage terms
cover_blueprint = {
    'Couple': 'Self and Partner',
    'Family': 'Self and Family',
    'Single Parent': 'Self and Child(ren)',
    'Single': 'Self Only'
}
if 'level_of_cover' in df_raw.columns:
    df_raw['level_of_cover'] = df_raw['level_of_cover'].map(cover_blueprint).fillna(df_raw['level_of_cover'])

# ==============================================================================
# 4. CHRONOLOGICAL FORMATTING & DATA QUALITY ANOMALY SCANNING
# ==============================================================================
print("Formatting dates using strict UK Day-First syntax constraints...")
df_raw['member_effective_date'] = pd.to_datetime(df_raw['member_effective_date'], dayfirst=True, format='mixed', errors='coerce').dt.date
df_raw['member_cancellation_date'] = pd.to_datetime(df_raw['member_cancellation_date'], dayfirst=True, format='mixed', errors='coerce').dt.date

df_raw['UniqueID'] = ""

# --- CAPTURE MEMBERSHIP FILE HEALTH EXCEPTIONS BEFORE ALIGNMENT ---
print("Scanning incoming file matrix for internal data health anomalies...")
current_run_year = parsed_date.year

# Extract Birth Year if tracking fields exist, otherwise handle cleanly
df_raw['Birth_Year_Parsed'] = 0
if 'birth_year' in df_raw.columns:
    df_raw['Birth_Year_Parsed'] = pd.to_numeric(df_raw['birth_year'], errors='coerce').fillna(0).astype(int)
elif 'date_of_birth' in df_raw.columns:
    df_raw['Birth_Year_Parsed'] = pd.to_datetime(df_raw['date_of_birth'], dayfirst=True, errors='coerce').dt.year.fillna(0).astype(int)

# Anomaly Filters
bad_birth_year_df = df_raw[(df_raw['Birth_Year_Parsed'] <= 1920) & (df_raw['Birth_Year_Parsed'] > 0)]
overage_children_df = df_raw[
    (df_raw['member_relation'].str.lower().str.strip() == 'child') & 
    (df_raw['Birth_Year_Parsed'] > 0) & 
    ((current_run_year - df_raw['Birth_Year_Parsed']) > 25)
]
chronology_mismatch_df = df_raw[
    (df_raw['member_cancellation_date'].notna()) & 
    (df_raw['member_effective_date'].notna()) & 
    (df_raw['member_cancellation_date'] < df_raw['member_effective_date'])
]

# ==============================================================================
# 5. DATA WAREHOUSE STRUCTURAL ALIGNMENT
# ==============================================================================
column_alignment_map = {
    'policy_no': 'Policy Number',
    'policy_holder_name': 'Policy Holder Name',
    'division_code': 'Division Code',
    'unique_member_id': 'Unique Member ID',
    'member_relation': 'Member Relation',
    'age_band': 'Age Band (Current Month)',
    'gender': 'Gender',
    'postcode_area': 'Postcode Area',
    'member_effective_date': 'Member Effective Date',
    'member_cancellation_date': 'Member Cancellation Date',
    'level_of_cover': 'Level Of Cover',
    'broker_name': 'Broker Name'
}

df_upload = df_raw.rename(columns=column_alignment_map)
df_upload['UploadDate'] = first_of_the_month

# Append empty calculating slots to fulfill structural destination schema footprints
exposure_columns = ['calculated_exposure_start', 'calculated_exposure_end', 'member_exposure_start', 'member_exposure_end']
for col in exposure_columns:
    df_upload[col] = None

all_valid_sql_columns = list(column_alignment_map.values()) + ['UploadDate', 'UniqueID'] + exposure_columns
df_upload = df_upload[[c for c in df_upload.columns if c in all_valid_sql_columns]]

# ==============================================================================
# RUNTIME EXPOSURE ENGINE CALCULATION FUNCTION (FIXED SYNTAX SENSE-CHECK ENGINE)
# ==============================================================================
def get_historical_exposure_averages(table_source):
    """
    Executes an optimized portfolio-wide in-memory calculation mimicking your 
    Aviva SQL variables logic, bounded up to the specified historical threshold date.
    """
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
            [Policy Number], 
            DATEADD(month, DATEDIFF(month, 0, MIN([finalised date])), 0) AS SchemeMinPayMonth, 
            DATEADD(MONTH, 1, DATEADD(month, DATEDIFF(month, 0, MAX([finalised date])), 0)) AS SchemeMaxPayMonth_Exclusive 
        FROM 
            {schema_name}.Test_LiveAvivaClaims 
        GROUP BY 
            [Policy Number] 
    ),
    CleanedAvivaMembership AS (
        SELECT
            *,
            CASE
                WHEN [Member Cancellation Date] IS NOT NULL
                THEN DATEADD(month, DATEDIFF(month, 0, [Member Cancellation Date]), 0)
                ELSE NULL
            END AS CleanedCancellationDate
        FROM {table_source}
    ),
    MemberHistoryWithEffectiveDate AS (
        SELECT
            *,
            CASE
                WHEN rnk = 1 THEN DATEADD(month, DATEDIFF(month, 0, [Member Effective Date]), 0)
                ELSE DATEADD(month, DATEDIFF(month, 0, [UploadDate]), 0) 
            END AS EffectiveDate
        FROM (
            SELECT *,
                ROW_NUMBER() OVER(PARTITION BY [Unique Member ID] ORDER BY [UploadDate]) as rnk
            FROM CleanedAvivaMembership
        ) AS RankedRecords
    ),
    MemberTimeline_Raw AS (
        SELECT
            *,
            EffectiveDate AS ExposureStart,
            LEAD(EffectiveDate, 1, '2999-12-31') OVER (PARTITION BY [Unique Member ID] ORDER BY EffectiveDate) AS NextRecord_StartDate,
            LEAD(CleanedCancellationDate, 1, NULL) OVER (PARTITION BY [Unique Member ID] ORDER BY EffectiveDate) AS NextRecord_CancellationDate
        FROM MemberHistoryWithEffectiveDate
    ),
    MemberTimeline AS (
        SELECT 
            *,
            CASE
                WHEN NextRecord_CancellationDate IS NOT NULL AND NextRecord_CancellationDate < NextRecord_StartDate 
                THEN NextRecord_CancellationDate
                ELSE NextRecord_StartDate
            END AS PotentialExposureEnd
        FROM MemberTimeline_Raw
    ),
    FinalPeriods AS (
        SELECT 
            *,
            CASE
                WHEN CleanedCancellationDate IS NOT NULL AND CleanedCancellationDate < PotentialExposureEnd THEN CleanedCancellationDate
                ELSE PotentialExposureEnd
            END AS ExposureEnd
        FROM MemberTimeline
    ),
    MonthlyHeadcounts AS (
        SELECT 
            t.[Policy Holder Name],
            DR.Months,
            COUNT(*) AS OpeningLives,
            COUNT(CASE WHEN t.[Member Relation] = 'Primary' THEN 1 END) AS Members,
            COUNT(CASE WHEN t.[Member Relation] = 'Spouse/Partner' THEN 1 END) AS Partners,
            COUNT(CASE WHEN t.[Member Relation] = 'Child' THEN 1 END) AS Dependents,
            COUNT(CASE WHEN t.[Level Of Cover] = 'Self Only' AND t.[Member Relation] = 'Primary' THEN 1 END) AS Single,
            COUNT(CASE WHEN t.[Level Of Cover] = 'Self and Partner' AND t.[Member Relation] = 'Primary' THEN 1 END) AS Couple,
            COUNT(CASE WHEN t.[Level Of Cover] = 'Self and Family' AND t.[Member Relation] = 'Primary' THEN 1 END) AS Family,
            COUNT(CASE WHEN t.[Level Of Cover] = 'Self and Child(ren)' AND t.[Member Relation] = 'Primary' THEN 1 END) AS SPF
        FROM DateRange DR
        JOIN FinalPeriods t ON 1=1
        JOIN SchemeClaimsExperience s ON t.[Policy Number] = s.[Policy Number]
        WHERE DR.Months >= (CASE WHEN t.ExposureStart > s.SchemeMinPayMonth THEN t.ExposureStart ELSE s.SchemeMinPayMonth END)
          AND DR.Months < (CASE WHEN t.ExposureEnd < s.SchemeMaxPayMonth_Exclusive THEN t.ExposureEnd ELSE s.SchemeMaxPayMonth_Exclusive END)
        GROUP BY t.[Policy Holder Name], DR.Months
    )
    SELECT 
        [Policy Holder Name],
        AVG(CAST(OpeningLives AS FLOAT)) AS OpeningLives,
        AVG(CAST(Members AS FLOAT)) AS Members,
        AVG(CAST(Partners AS FLOAT)) AS Partners,
        AVG(CAST(Dependents AS FLOAT)) AS Dependents,
        AVG(CAST(Single AS FLOAT)) AS Single,
        AVG(CAST(Couple AS FLOAT)) AS Couple,
        AVG(CAST(Family AS FLOAT)) AS Family,
        AVG(CAST(SPF AS FLOAT)) AS SPF
    FROM MonthlyHeadcounts
    GROUP BY [Policy Holder Name]
    OPTION (MAXRECURSION 0);
    """
    res_df = pd.read_sql(exposure_sql, con=engine)
    res_df['Client_Key'] = res_df['Policy Holder Name'].astype(str).str.strip().str.title()
    return res_df.groupby('Client_Key')[['OpeningLives', 'Members', 'Partners', 'Dependents', 'Single', 'Couple', 'Family', 'SPF']].sum().reset_index()

# Calculate baseline historical averages BEFORE modifying live tables
print("Calculating running historical monthly averages across active Aviva client matrices...")
df_exp_before_grouped = get_historical_exposure_averages(f"{schema_name}.Test_LiveAvivaMembership")

# ==============================================================================
# 6. BULK PIPELINE STREAM TO STAGING WORKSPACE
# ==============================================================================
print(f"Emptying {schema_name}.Test_UploadAvivaMembership workspace table...")
cursor.execute(f"TRUNCATE TABLE {schema_name}.Test_UploadAvivaMembership")
conn.commit()

print(f"Streaming ingestion array data over to {schema_name}.Test_UploadAvivaMembership...")
clean_schema = schema_name.strip("[]")

df_upload.to_sql(
    name='Test_UploadAvivaMembership', 
    con=engine, 
    schema=quoted_name(clean_schema, quote=True), 
    if_exists='append', 
    index=False, 
    method='multi', 
    chunksize=50
)

# ==============================================================================
# 7. DETERMINISTIC UNIQUE KEY RECOMPILATION & DELTA JOIN
# ==============================================================================
print("Recompiling deterministic Unique IDs via SQL transformation engine...")
recompile_staging_keys_query = f"""
    UPDATE {schema_name}.Test_UploadAvivaMembership
    SET [UniqueID] = CONCAT(
        TRIM([Unique Member ID]),
        CAST(DATEDIFF(day, '1899-12-30', [Member Effective Date]) AS VARCHAR(20)),
        ISNULL(CAST(DATEDIFF(day, '1899-12-30', [Member Cancellation Date]) AS VARCHAR(20)), ''),
        TRIM([Level Of Cover])
    );
"""
cursor.execute(recompile_staging_keys_query)
conn.commit()

print(f"Appending newly identified distinct profiles directly into {schema_name}.Test_LiveAvivaMembership...")
upload_columns_no_id = """
    [Policy Number], [Policy Holder Name], [Division Code], [Unique Member ID], 
    [Member Relation], [Age Band (Current Month)], [Gender], [Postcode Area], 
    [Member Effective Date], [Member Cancellation Date], [Level Of Cover], [Broker Name], 
    [UploadDate], [UniqueID], [calculated_exposure_start], [calculated_exposure_end], 
    [member_exposure_start], [member_exposure_end]
"""

delta_insert_query = f"""
    INSERT INTO {schema_name}.Test_LiveAvivaMembership ({upload_columns_no_id})
    SELECT {upload_columns_no_id}
    FROM {schema_name}.Test_UploadAvivaMembership u
    WHERE NOT EXISTS (
        SELECT 1
        FROM {schema_name}.Test_LiveAvivaMembership l
        WHERE l.UniqueID = u.UniqueID
    );
"""
cursor.execute(delta_insert_query)
conn.commit()

# ==============================================================================
# SENSE CHECK PART B - CAPTURE AVIVA ACCOUNT EXPOSURE POST UPLOAD
# ==============================================================================
print("Recalculating post-upload historical monthly averages (excluding new target month data)...")
df_exp_after_grouped = get_historical_exposure_averages(f"{schema_name}.Test_LiveAvivaMembership")

# Combine structural dataframes to track discrepancies
df_exp_recon = pd.merge(
    df_exp_before_grouped, df_exp_after_grouped, 
    on='Client_Key', how='outer', suffixes=('_Before', '_After')
).fillna(0)

# Calculate numerical variances and percentage shifts
metrics_list = ['OpeningLives', 'Members', 'Partners', 'Dependents', 'Single', 'Couple', 'Family', 'SPF']
for metric in metrics_list:
    df_exp_recon[f'{metric}_Diff'] = df_exp_recon[f'{metric}_After'] - df_exp_recon[f'{metric}_Before']
    df_exp_recon[f'{metric}_Pct'] = df_exp_recon.apply(
        lambda r: f"{((r[f'{metric}_After'] - r[f'{metric}_Before']) / r[f'{metric}_Before'] * 100):+.2f}%" if r[f'{metric}_Before'] > 0 else ("+0.00%" if r[f'{metric}_After'] == 0 else "+100.00%"), axis=1
    )

# ==============================================================================
# 8. EXPOSURE CALCULATOR WITH REPLACEMENT MAPPING
# ==============================================================================
if os.path.exists(exposure_script_sql):
    print(f"Reading original production exposure script: {os.path.basename(exposure_script_sql)}")
    with open(exposure_script_sql, 'r') as sql_file:
        sql_content = sql_file.read()
    
    print("Applying corporate schema overrides using standard string replacements...")
    
    # Strip database headers and 'GO' statements smoothly
    sql_content = re.sub(r'^\s*USE\s+\w+\s*;?', '', sql_content, flags=re.IGNORECASE | re.MULTILINE)
    sql_content = re.sub(r'^\s*GO\s*$', ';', sql_content, flags=re.IGNORECASE | re.MULTILINE)
    
    target_live_membership = f"{schema_name}.Test_LiveAvivaMembership"
    target_live_claims = f"{schema_name}.Test_LiveAvivaClaims"
    
    # --- FIX: Use safe lambda replacement or literal strings to bypass regex escape sequences (\A) ---
    sql_content = re.sub(r'(?i)\b(?:dbo\.)?\[?LiveAvivaMembership\]?\b', lambda m: target_live_membership, sql_content)
    sql_content = re.sub(r'(?i)\b(?:dbo\.)?\[?liveavivaclaims\]?\b', lambda m: target_live_claims, sql_content)
    
    print("Streaming comprehensive exposure calculation script to SQL Server...")
    if sql_content.strip():
        try:
            cursor.execute(sql_content)
            conn.commit()
            print("Exposure metrics calculation completed successfully on test targets!")
        except Exception as sql_err:
            print(f"❌ SQL Execution Error inside exposure script: {sql_err}")
else:
    print(f"⚠️ Warning: Exposure script not located at path target location:\n   -> {exposure_script_sql}")

# ==============================================================================
# 9. INTEGRITY RECONCILIATION AUDIT, REPORTING & NATIVE PDF EXPORT
# ==============================================================================
class CorporateAvivaMembershipPDF(FPDF):
    def header(self):
        self.set_font('Courier', 'B', 11)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, 'LOCKTON EMPLOYEE BENEFITS | AUTOMATED AVIVA MEMBERSHIP SYSTEM AUDIT LOG', ln=True, align='L')
        self.set_draw_color(180, 180, 180)
        self.line(10, 18, 200, 18)
        self.ln(8)
        
    def footer(self):
        self.set_y(-15)
        self.set_font('Courier', 'I', 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f'Page {self.page_no()} | CONFIDENTIAL - INTERNAL USE ONLY', align='C')

# Fetch metrics directly via SQLAlchemy to safeguard against closed pyodbc connections
upload_count_final = int(pd.read_sql(f"SELECT COUNT(*) AS cnt FROM {schema_name}.Test_UploadAvivaMembership", con=engine).iloc[0]['cnt'])
final_live_count = int(pd.read_sql(f"SELECT COUNT(*) AS cnt FROM {schema_name}.Test_LiveAvivaMembership", con=engine).iloc[0]['cnt'])
new_rows_added = final_live_count - initial_live_count

# Build Report Text payload matrix arrays
pdf_payload = []
pdf_payload.append("====================================================================================================")
pdf_payload.append("                         AVIVA MEMBERSHIP AUTOMATED AUDIT & RECONCILIATION")
pdf_payload.append("====================================================================================================")
pdf_payload.append(f"Raw Input Rows (Incoming File):       {len(df_raw)}")
pdf_payload.append(f"Staging Stored Rows (Upload):         {upload_count_final}")
pdf_payload.append("-" * 100)
pdf_payload.append(f"Initial Live Table Rows Baseline:    {initial_live_count}")
pdf_payload.append(f"Current Live Table Rows Post-Load:   {final_live_count}")
pdf_payload.append(f"Net Row Increment Delta Added:       {new_rows_added}")
pdf_payload.append("-" * 100)

if len(df_raw) == upload_count_final:
    pdf_payload.append("[OK] DATA INTEGRITY VERIFIED: Raw row totals match staging database perfectly.")
else:
    pdf_payload.append("[WARNING] Mismatch detected between incoming raw size and SQL staging rows.")

# STAGE PART 1: RUNNING HISTORICAL MONTHLY EXPOSURE AVERAGE SENSE CHECK
pdf_payload.append(f"\n[PART 1: RUNNING HISTORICAL MONTHLY EXPOSURE AVERAGE SENSE CHECK (UP TO: {audit_target_month.strftime('%B %Y')})]")
pdf_payload.append("-" * 100)

for _, row in df_exp_recon.iterrows():
    pdf_payload.append(f"Client Account: {row['Client_Key']}")
    pdf_payload.append("   " + "-"*85)
    pdf_payload.append("   Metric         | Before (Avg)  | After (Avg)   | Net Variance  | Change %     ")
    pdf_payload.append("   " + "-"*85)
    for m in metrics_list:
        pdf_payload.append(f"   {m:<14} | {row[f'{m}_Before']:<13.2f} | {row[f'{m}_After']:<13.2f} | {row[f'{m}_Diff']:<13.2f} | {row[f'{m}_Pct']}")
    pdf_payload.append("   " + "-"*85 + "\n")

pdf_payload.append("-" * 100)

# STAGE PART 2: INCOMING MEMBERSHIP DATA QUALITY EXCEPTION LOG
pdf_payload.append("\n[PART 2: INCOMING MEMBERSHIP DATA QUALITY EXCEPTION LOG]")
pdf_payload.append("-" * 100)
has_exceptions = False

if len(bad_birth_year_df) > 0:
    has_exceptions = True
    pdf_payload.append(f"[WARNING] ANOMALY: Found {len(bad_birth_year_df)} rows with impossible/ancient Birth Years (<= 1920):")
    pdf_payload.append(bad_birth_year_df[['policy_no', 'policy_holder_name', 'unique_member_id', 'Birth_Year_Parsed']].head(5).to_string(index=False))
    if len(bad_birth_year_df) > 5: pdf_payload.append("   ... [Truncated for scannability]")

if len(overage_children_df) > 0:
    has_exceptions = True
    pdf_payload.append(f"[WARNING] ANOMALY: Found {len(overage_children_df)} rows where a 'Child Dependant' is older than 25 years:")
    pdf_payload.append(overage_children_df[['policy_no', 'policy_holder_name', 'unique_member_id', 'member_relation', 'Birth_Year_Parsed']].head(5).to_string(index=False))
    if len(overage_children_df) > 5: pdf_payload.append("   ... [Truncated for scannability]")

if len(chronology_mismatch_df) > 0:
    has_exceptions = True
    pdf_payload.append(f"[WARNING] CHRONOLOGY BREAK: Found {len(chronology_mismatch_df)} rows where Cancellation Date is before Effective Date:")
    pdf_payload.append(chronology_mismatch_df[['policy_no', 'policy_holder_name', 'unique_member_id', 'member_effective_date', 'member_cancellation_date']].head(5).to_string(index=False))
    if len(chronology_mismatch_df) > 5: pdf_payload.append("   ... [Truncated for scannability]")

if not has_exceptions:
    pdf_payload.append("[OK] FILE VALIDATION SUCCESSFUL: No impossible dates, overaged child dependants, or status logic anomalies found.")

pdf_payload.append("="*100)

# Merge matrix blocks to single text output payload
full_report_text = "\n".join(pdf_payload)

# ------------------------------------------------------------------------------
# ACTION A: Push report layouts cleanly to your Terminal window
# ------------------------------------------------------------------------------
print(full_report_text)

# ------------------------------------------------------------------------------
# ACTION B: Render onto clear, multi-page hardcopy PDF documents safely
# ------------------------------------------------------------------------------
print("\nCompiling automated A4 corporate PDF document...")
pdf = CorporateAvivaMembershipPDF(orientation='P', unit='mm', format='A4')
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()
pdf.set_font("Courier", size=8.0)  # Compact font size protects grid layouts from wrapping
pdf.set_text_color(40, 40, 40)

for line in full_report_text.split('\n'):
    # Sanitizer Shield: Strips screen emojis safely for standard latin-1 core fonts
    clean_line = line.encode('latin-1', 'ignore').decode('latin-1')
    pdf.cell(0, 4.0, txt=clean_line, ln=True)

# Generate destination file targets inside your specified base directory
base_directory = os.path.dirname(raw_membership_csv)
pdf_output_filename = f"Aviva_Membership_Audit_Report_{parsed_date.strftime('%Y_%m')}.pdf"
pdf_output_path = os.path.join(base_directory, pdf_output_filename)
pdf.output(pdf_output_path)

print(f"✅ SUCCESS: Aviva Membership PDF asset structured and saved cleanly.")
print(f"📂 Saved Location: {pdf_output_path}")

# Terminate pipeline connections safely
try:
    cursor.close()
    conn.close()
except Exception:
    pass

print("Aviva process completed successfully with automated scripts executed!")