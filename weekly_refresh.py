"""
Coaching Call Discussions — Monthly Refresh

Executes weekly_refresh.sql against Vertica to rebuild the persistent
Carefirst_Sandbox tables (COACHING_CALL_TOPICS, COACHING_CALL_TOBACCO,
COACHING_CALL_GOALS).

Designed for:
  - Local execution via launchd (monthly, 1st of each month)
  - Airflow execution (import main() or call as script)
  - Manual execution (python3 weekly_refresh.py)

Usage:
    python3 weekly_refresh.py
"""
import logging
import os
import sys
import time

# --- Logging setup ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(levelname)-8s  %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
log = logging.getLogger(__name__)

# --- Path setup ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SQL_FILE = os.path.join(SCRIPT_DIR, 'weekly_refresh.sql')

# --- Connection ---
# When running locally, use shared db_connect helper.
# When running on Airflow, the connection is passed in or configured via env vars.
sys.path.append(os.path.expanduser('~/Documents/dev/automation'))
try:
    from db_connect import get_connection
except ImportError:
    # Airflow or other environments: build connection from env vars
    import vertica_python
    def get_connection():
        conn_info = {
            'host': os.environ['VERTICA_HOST'],
            'port': int(os.environ.get('VERTICA_PORT', 5433)),
            'user': os.environ['VERTICA_USER'],
            'password': os.environ['VERTICA_PASSWORD'],
            'database': os.environ.get('VERTICA_DATABASE', 'SharecareBI'),
            'connection_timeout': None,  # no socket timeout for long-running SQL
        }
        return vertica_python.connect(**conn_info)


def load_sql(path):
    """Read and split SQL file into executable statements."""
    with open(path, 'r') as f:
        raw = f.read()
    # Split on semicolons, skip empty/comment-only fragments
    stmts = []
    for s in raw.split(';'):
        stripped = s.strip()
        if not stripped:
            continue
        # Check if there's any non-comment content
        has_code = False
        for line in stripped.split('\n'):
            line = line.strip()
            if line and not line.startswith('--'):
                has_code = True
                break
        if has_code:
            stmts.append(stripped)
    return stmts


def get_first_keyword(stmt):
    """Extract the first SQL keyword from a statement (skipping comments)."""
    for line in stmt.split('\n'):
        line = line.strip()
        if line and not line.startswith('--'):
            return line.split()[0].upper()
    return ''


def main():
    """Execute the weekly refresh pipeline."""
    log.info("=" * 60)
    log.info("COACHING CALL DISCUSSIONS - WEEKLY REFRESH")
    log.info("=" * 60)

    # Load SQL
    if not os.path.exists(SQL_FILE):
        log.error(f"SQL file not found: {SQL_FILE}")
        sys.exit(1)

    stmts = load_sql(SQL_FILE)
    log.info(f"Loaded {len(stmts)} SQL statements from {os.path.basename(SQL_FILE)}")

    # Execute
    start = time.time()
    try:
        with get_connection() as conn:
            cursor = conn.cursor()
            for i, stmt in enumerate(stmts, 1):
                keyword = get_first_keyword(stmt)
                # Build a short description for logging
                first_line = ''
                for line in stmt.split('\n'):
                    line = line.strip()
                    if line and not line.startswith('--'):
                        first_line = line[:80]
                        break
                log.info(f"  [{i}/{len(stmts)}] {first_line}...")

                step_start = time.time()
                cursor.execute(stmt)
                elapsed = time.time() - step_start

                # Log row counts for INSERT/TRUNCATE
                if keyword == 'INSERT':
                    rowcount = cursor.rowcount if cursor.rowcount >= 0 else '?'
                    log.info(f"           -> {rowcount} rows inserted ({elapsed:.1f}s)")
                elif keyword == 'TRUNCATE':
                    log.info(f"           -> truncated ({elapsed:.1f}s)")
                elif elapsed > 5:
                    log.info(f"           -> done ({elapsed:.1f}s)")

            conn.commit()

    except Exception as e:
        log.error(f"FAILED: {e}", exc_info=True)
        sys.exit(1)

    total = time.time() - start
    log.info(f"COMPLETED in {total:.0f}s ({total/60:.1f} min)")
    log.info("Tables refreshed:")
    log.info("  - Carefirst_Sandbox.COACHING_CALL_TOPICS")
    log.info("  - Carefirst_Sandbox.COACHING_CALL_TOBACCO")
    log.info("  - Carefirst_Sandbox.COACHING_CALL_GOALS")


if __name__ == '__main__':
    main()
