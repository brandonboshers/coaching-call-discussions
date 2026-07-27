"""
Coaching Call Discussions - Weekly Refresh to Vertica
Runs the full SQL pipeline and persists results to CAREFIRST_SANDBOX.

Target tables (drop/recreate each run):
  - CAREFIRST_SANDBOX.COACHING_CALL_TOPICS
  - CAREFIRST_SANDBOX.COACHING_CALL_GOALS

Usage:
    python3 weekly_refresh.py                     # all customers, all dates
    python3 weekly_refresh.py HP_SCCareFirst      # single customer
    python3 weekly_refresh.py HP_SCCareFirst 2025-01-01 2025-06-30  # customer + date range
"""
import sys
import os
import logging
from datetime import datetime

sys.path.append(os.path.expanduser('~/Documents/dev/automation'))
from db_connect import get_connection

# --- Logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(levelname)-8s  %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(os.path.expanduser('~/Library/Logs/coaching-call-discussions.log'))
    ]
)
log = logging.getLogger(__name__)

# --- Parse args ---
customer_id = sys.argv[1] if len(sys.argv) > 1 else None
start_date = sys.argv[2] if len(sys.argv) > 2 else None
end_date = sys.argv[3] if len(sys.argv) > 3 else None

TARGET_SCHEMA = 'CAREFIRST_SANDBOX'
TOPICS_TABLE = f'{TARGET_SCHEMA}.COACHING_CALL_TOPICS'
GOALS_TABLE = f'{TARGET_SCHEMA}.COACHING_CALL_GOALS'


def run_refresh():
    """Execute the full pipeline and persist to Vertica."""
    start_time = datetime.now()
    scope = customer_id or 'ALL CUSTOMERS'
    log.info(f"Starting coaching call discussions refresh | Scope: {scope}")
    if start_date:
        log.info(f"Date range: {start_date} to {end_date or 'present'}")

    # --- Read SQL ---
    sql_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'coaching_call_topics_goals.sql')
    with open(sql_path, 'r') as f:
        full_sql = f.read()

    # --- Inject filters if specified ---
    if customer_id or start_date:
        filters = []
        if customer_id:
            filters.append(f"AND MC.CUSTOMERID = '{customer_id}'")
        if start_date:
            filters.append(f"AND TRUNC(MC.ENCOUNTERDATETIME)::DATE >= '{start_date}'")
        if end_date:
            filters.append(f"AND TRUNC(MC.ENCOUNTERDATETIME)::DATE <= '{end_date}'")
        inject_point = "AND UPPER(MC.DIRECTION) = 'OUTBOUND'"
        full_sql = full_sql.replace(inject_point, inject_point + '\n      ' + '\n      '.join(filters))

    # --- Split into setup statements (DROP/CREATE) and output queries (SELECT) ---
    raw_stmts = [s.strip() for s in full_sql.split(';') if s.strip()]
    setup_statements = []
    for stmt in raw_stmts:
        first_keyword = None
        for line in stmt.split('\n'):
            stripped = line.strip()
            if stripped and not stripped.startswith('--'):
                first_keyword = stripped.split()[0].upper()
                break
        if first_keyword in ('DROP', 'CREATE'):
            setup_statements.append(stmt)

    log.info(f"Setup statements: {len(setup_statements)}")

    # --- Execute pipeline ---
    with get_connection(connection_timeout=None) as conn:
        cursor = conn.cursor()

        # Step 1: Run all temp table creation (the full pipeline)
        log.info("Running temp table pipeline...")
        for i, stmt in enumerate(setup_statements):
            # Log the first meaningful line
            for line in stmt.split('\n'):
                if line.strip() and not line.strip().startswith('--'):
                    log.info(f"  [{i+1}/{len(setup_statements)}] {line.strip()[:70]}")
                    break
            cursor.execute(stmt)

        log.info("Temp table pipeline complete.")

        # Step 2: Persist COACHING_TOPICS to Vertica
        log.info(f"Persisting topics to {TOPICS_TABLE}...")
        cursor.execute(f"DROP TABLE IF EXISTS {TOPICS_TABLE}")
        cursor.execute(f"""
            CREATE TABLE {TOPICS_TABLE} AS
            SELECT
                ACCOUNT,
                GUID,
                CUSTOMERID,
                CALL_DATE,
                CALL_TYPE,
                RAW_TOPIC,
                REPORT_TOPIC,
                TOPIC_SOURCE,
                PROGRAM_TYPE,
                CURRENT_TIMESTAMP AS REFRESH_TIMESTAMP
            FROM COACHING_TOPICS
        """)
        cursor.execute(f"SELECT COUNT(*) FROM {TOPICS_TABLE}")
        topics_count = cursor.fetchone()[0]
        log.info(f"  {TOPICS_TABLE}: {topics_count:,} rows")

        # Step 3: Persist COACHING_GOALS_NUMBERED to Vertica
        log.info(f"Persisting goals to {GOALS_TABLE}...")
        cursor.execute(f"DROP TABLE IF EXISTS {GOALS_TABLE}")
        cursor.execute(f"""
            CREATE TABLE {GOALS_TABLE} AS
            SELECT
                ACCOUNT,
                GOAL_TYPE,
                GOAL_DOMAIN,
                GOAL_STATUS,
                GOAL_DESCRIPTION,
                RAW_ACTION_NAME,
                MEMBERACTION_ID,
                ACTIONTYPE_ID,
                ACTIONSTATUS_ID,
                FOCUSAREA_ID,
                GOAL_SET_DATE,
                GOAL_CLOSE_DATE,
                CURRENTGUID,
                GUID,
                GOAL_NUMBER,
                CURRENT_TIMESTAMP AS REFRESH_TIMESTAMP
            FROM COACHING_GOALS_NUMBERED
        """)
        cursor.execute(f"SELECT COUNT(*) FROM {GOALS_TABLE}")
        goals_count = cursor.fetchone()[0]
        log.info(f"  {GOALS_TABLE}: {goals_count:,} rows")

        # Step 4: Also persist the tobacco flag for easy lookup
        tobacco_table = f'{TARGET_SCHEMA}.COACHING_CALL_TOBACCO'
        log.info(f"Persisting tobacco flag to {tobacco_table}...")
        cursor.execute(f"DROP TABLE IF EXISTS {tobacco_table}")
        cursor.execute(f"""
            CREATE TABLE {tobacco_table} AS
            SELECT
                ACCOUNT,
                GUID,
                TOBACCO_DISCUSSED,
                CURRENT_TIMESTAMP AS REFRESH_TIMESTAMP
            FROM COACHING_TOBACCO
        """)
        cursor.execute(f"SELECT COUNT(*) FROM {tobacco_table}")
        tobacco_count = cursor.fetchone()[0]
        log.info(f"  {tobacco_table}: {tobacco_count:,} rows")

        conn.commit()

    elapsed = datetime.now() - start_time
    log.info(f"Done! Elapsed: {elapsed}")
    log.info(f"Summary: Topics={topics_count:,} | Goals={goals_count:,} | Tobacco={tobacco_count:,}")


if __name__ == '__main__':
    try:
        run_refresh()
    except Exception as e:
        log.error(f"FAILED: {e}", exc_info=True)
        sys.exit(1)
