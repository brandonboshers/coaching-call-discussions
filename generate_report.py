"""
Coaching Call Discussions — Monthly PowerPoint Report Generator

Queries the persistent Carefirst_Sandbox tables and generates a branded
PowerPoint report with dashboard metrics.

Usage:
    python3 generate_report.py                              # HP_SCCareFirst, current year
    python3 generate_report.py ER_SHBP                      # specific customer
    python3 generate_report.py HP_SCCareFirst 2025-01-01 2025-06-30  # customer + date range
"""
import sys
import os
import logging
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

sys.path.append(os.path.expanduser('~/Documents/dev/automation'))
from db_connect import get_connection
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# --- Config ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(SCRIPT_DIR, 'template.pptx')
DEFAULT_CUSTOMER = 'HP_SCCAREFIRST'

# Brand colors
COLOR_PRIMARY = RGBColor(0x00, 0x4D, 0x3D)    # Dark teal
COLOR_ACCENT = RGBColor(0x00, 0xBF, 0xA5)     # Bright teal
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_DARK = RGBColor(0x33, 0x33, 0x33)
COLOR_LIGHT_BG = RGBColor(0xF5, 0xF5, 0xF5)
COLOR_HEADER_BG = RGBColor(0x00, 0x4D, 0x3D)

# Fonts
FONT_HEADING = 'Roboto Serif 20pt'
FONT_BODY = 'Proxima Nova Rg'

logging.basicConfig(level=logging.INFO, format='%(asctime)s  %(levelname)-8s  %(message)s')
log = logging.getLogger(__name__)


# --- Parse args ---
customer_id = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_CUSTOMER
start_date = sys.argv[2] if len(sys.argv) > 2 else None
end_date = sys.argv[3] if len(sys.argv) > 3 else None

# Default: prior full month
if not start_date:
    today = date.today()
    first_of_this_month = today.replace(day=1)
    last_month_end = first_of_this_month - relativedelta(days=1)
    last_month_start = last_month_end.replace(day=1)
    start_date = last_month_start.strftime('%Y-%m-%d')
    end_date = last_month_end.strftime('%Y-%m-%d')
elif not end_date:
    end_date = date.today().strftime('%Y-%m-%d')


def build_filter(table_alias='', date_col='CALL_DATE'):
    """Build WHERE clause fragments for customer and date filtering."""
    parts = []
    prefix = f"{table_alias}." if table_alias else ""
    parts.append(f"UPPER({prefix}CUSTOMERID) = UPPER('{customer_id}')")
    if date_col:
        parts.append(f"{prefix}{date_col} >= '{start_date}'")
        parts.append(f"{prefix}{date_col} <= '{end_date}'")
    return ' AND '.join(parts)


# --- Query data ---
log.info(f"Customer: {customer_id}")
log.info(f"Period: {start_date} to {end_date}")

with get_connection() as conn:
    # Section 1: Coaching Engagement by Wellbeing Topic
    df_engagement = pd.read_sql(f"""
        SELECT
            T.REPORT_TOPIC AS WELLBEING_TOPIC,
            COUNT(DISTINCT T.CURRENTGUID) AS MEMBERS,
            ROUND(COUNT(DISTINCT T.CURRENTGUID) * 100.0
                / NULLIFZERO(TOTALS.TOTAL_MEMBERS), 1) AS PCT_OF_MEMBERS,
            COUNT(DISTINCT CASE WHEN G.GOAL_STATUS = 'Completed' THEN G.MEMBERACTION_ID END) AS COMPLETED_GOALS,
            COUNT(DISTINCT CASE WHEN G.GOAL_STATUS IN ('In Progress','Not Started') THEN G.MEMBERACTION_ID END) AS OPEN_IN_PROGRESS_GOALS
        FROM Carefirst_Sandbox.COACHING_CALL_TOPICS T
        CROSS JOIN (
            SELECT COUNT(DISTINCT CURRENTGUID) AS TOTAL_MEMBERS
            FROM Carefirst_Sandbox.COACHING_CALL_TOPICS
            WHERE {build_filter()}
        ) TOTALS
        LEFT JOIN Carefirst_Sandbox.COACHING_CALL_GOALS G ON T.CURRENTGUID = G.CURRENTGUID
        WHERE {build_filter('T')}
        GROUP BY 1, TOTALS.TOTAL_MEMBERS
        ORDER BY 2 DESC
    """, conn)

    # Section 2: Goal Status by Wellbeing Topic
    df_goal_status_topic = pd.read_sql(f"""
        SELECT
            T.REPORT_TOPIC AS WELLBEING_TOPIC,
            COUNT(DISTINCT CASE WHEN G.GOAL_STATUS = 'Completed' THEN G.MEMBERACTION_ID END) AS COMPLETED,
            COUNT(DISTINCT CASE WHEN G.GOAL_STATUS = 'In Progress' THEN G.MEMBERACTION_ID END) AS IN_PROGRESS,
            COUNT(DISTINCT CASE WHEN G.GOAL_STATUS = 'Not Started' THEN G.MEMBERACTION_ID END) AS NOT_STARTED
        FROM Carefirst_Sandbox.COACHING_CALL_TOPICS T
        JOIN Carefirst_Sandbox.COACHING_CALL_GOALS G ON T.CURRENTGUID = G.CURRENTGUID
        WHERE {build_filter('T')}
        GROUP BY 1
        ORDER BY 2 DESC
    """, conn)

    # Section 3: Goal Status Distribution
    df_goal_dist = pd.read_sql(f"""
        SELECT
            G.GOAL_STATUS,
            COUNT(*) AS COUNT,
            ROUND(COUNT(*) * 100.0 / NULLIFZERO(SUM(COUNT(*)) OVER()), 1) AS GOAL_PCT
        FROM Carefirst_Sandbox.COACHING_CALL_GOALS G
        JOIN Carefirst_Sandbox.COACHING_CALL_TOPICS T ON G.CURRENTGUID = T.CURRENTGUID
        WHERE G.GOAL_STATUS IN ('Completed','In Progress','Not Started','Withdrawn')
          AND {build_filter('T')}
        GROUP BY 1
        ORDER BY CASE G.GOAL_STATUS
            WHEN 'Completed' THEN 1 WHEN 'In Progress' THEN 2 WHEN 'Not Started' THEN 3 WHEN 'Withdrawn' THEN 4 END
    """, conn)

    # Section 4: Goal Progression by Domain
    df_goal_prog = pd.read_sql(f"""
        SELECT
            G.GOAL_DOMAIN,
            COUNT(*) AS TOTAL_GOALS,
            SUM(CASE WHEN G.GOAL_STATUS = 'Completed' THEN 1 ELSE 0 END) AS COMPLETED,
            ROUND(SUM(CASE WHEN G.GOAL_STATUS = 'Completed' THEN 1 ELSE 0 END) * 100.0
                / NULLIFZERO(COUNT(*)), 1) AS COMPLETION_RATE
        FROM Carefirst_Sandbox.COACHING_CALL_GOALS G
        JOIN Carefirst_Sandbox.COACHING_CALL_TOPICS T ON G.CURRENTGUID = T.CURRENTGUID
        WHERE {build_filter('T')}
        GROUP BY 1
        ORDER BY CASE G.GOAL_DOMAIN
            WHEN 'Gaps in Care' THEN 1 WHEN 'Exercise' THEN 2 WHEN 'Nutrition' THEN 3
            WHEN 'Weight Management' THEN 4 WHEN 'Tobacco Cessation' THEN 5
            WHEN 'Mental/Behavioral Health' THEN 6 WHEN 'Stress Management' THEN 7
            WHEN 'Condition Management' THEN 8 WHEN 'Financial' THEN 9
            WHEN 'Social' THEN 10 WHEN 'Spiritual' THEN 11 ELSE 12 END
    """, conn)

    # Section 5: Tobacco Coaching Focus
    df_tobacco = pd.read_sql(f"""
        SELECT 'Tobacco Participants' AS METRIC,
            COUNT(DISTINCT TB.CURRENTGUID)::VARCHAR AS VALUE
        FROM Carefirst_Sandbox.COACHING_CALL_TOBACCO TB
        JOIN Carefirst_Sandbox.COACHING_CALL_TOPICS T ON TB.CURRENTGUID = T.CURRENTGUID
        WHERE {build_filter('T')}

        UNION ALL
        SELECT 'Active Tobacco Participants',
            COUNT(DISTINCT G.CURRENTGUID)::VARCHAR
        FROM Carefirst_Sandbox.COACHING_CALL_TOBACCO TB
        JOIN Carefirst_Sandbox.COACHING_CALL_TOPICS T ON TB.CURRENTGUID = T.CURRENTGUID
        JOIN Carefirst_Sandbox.COACHING_CALL_GOALS G
            ON TB.CURRENTGUID = G.CURRENTGUID AND G.GOAL_DOMAIN = 'Tobacco Cessation'
            AND G.GOAL_STATUS IN ('In Progress','Not Started')
        WHERE {build_filter('T')}

        UNION ALL
        SELECT 'Goals Completed',
            COUNT(*)::VARCHAR
        FROM Carefirst_Sandbox.COACHING_CALL_GOALS G
        JOIN Carefirst_Sandbox.COACHING_CALL_TOPICS T ON G.CURRENTGUID = T.CURRENTGUID
        WHERE G.GOAL_DOMAIN = 'Tobacco Cessation' AND G.GOAL_STATUS = 'Completed'
          AND {build_filter('T')}

        UNION ALL
        SELECT 'Goals In Progress',
            COUNT(*)::VARCHAR
        FROM Carefirst_Sandbox.COACHING_CALL_GOALS G
        JOIN Carefirst_Sandbox.COACHING_CALL_TOPICS T ON G.CURRENTGUID = T.CURRENTGUID
        WHERE G.GOAL_DOMAIN = 'Tobacco Cessation' AND G.GOAL_STATUS = 'In Progress'
          AND {build_filter('T')}

        UNION ALL
        SELECT 'Completion Rate',
            ROUND(SUM(CASE WHEN G.GOAL_STATUS = 'Completed' THEN 1 ELSE 0 END) * 100.0
                / NULLIFZERO(COUNT(*)), 1)::VARCHAR || '%'
        FROM Carefirst_Sandbox.COACHING_CALL_GOALS G
        JOIN Carefirst_Sandbox.COACHING_CALL_TOPICS T ON G.CURRENTGUID = T.CURRENTGUID
        WHERE G.GOAL_DOMAIN = 'Tobacco Cessation'
          AND G.GOAL_STATUS IN ('Completed','In Progress','Not Started')
          AND {build_filter('T')}
    """, conn)

log.info(f"Data loaded: {len(df_engagement)} topics, {len(df_goal_prog)} domains")


# --- PowerPoint generation ---
def set_cell_format(cell, text, font_size=10, bold=False, color=COLOR_DARK, alignment=PP_ALIGN.LEFT, fill=None):
    """Format a table cell with consistent styling."""
    cell.text = str(text) if text is not None else ''
    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = alignment
        for run in paragraph.runs:
            run.font.name = FONT_BODY
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.color.rgb = color
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill


def add_branded_table(slide, df, left, top, width, height, title=None, pct_cols=None):
    """Add a formatted table to a slide with branded header row."""
    rows = len(df) + 1  # +1 for header
    cols = len(df.columns)
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    # Set column widths proportionally
    col_width = width // cols
    for i in range(cols):
        table.columns[i].width = col_width

    # Header row
    for c, col_name in enumerate(df.columns):
        cell = table.cell(0, c)
        display_name = col_name.replace('_', ' ').title()
        set_cell_format(cell, display_name, font_size=9, bold=True,
                       color=COLOR_WHITE, alignment=PP_ALIGN.CENTER, fill=COLOR_HEADER_BG)

    # Data rows
    for r, row in df.iterrows():
        for c, val in enumerate(row):
            cell = table.cell(r + 1, c)
            # Format percentages
            if pct_cols and df.columns[c] in pct_cols and val is not None:
                display_val = f"{val}%"
            else:
                display_val = val
            is_numeric = isinstance(val, (int, float))
            align = PP_ALIGN.RIGHT if is_numeric else PP_ALIGN.LEFT
            # Alternating row shading
            fill_color = COLOR_LIGHT_BG if r % 2 == 0 else None
            set_cell_format(cell, display_val, font_size=9, color=COLOR_DARK,
                          alignment=align, fill=fill_color)

    return table_shape


def add_section_title(slide, text, left, top):
    """Add a bold section title text box."""
    txBox = slide.shapes.add_textbox(left, top, Inches(8), Pt(30))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = FONT_BODY
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY
    return txBox


# Build presentation
prs = Presentation(TEMPLATE_PATH)

# Remove all existing slides (we only want the template's theme/master)
while len(prs.slides) > 0:
    rId = prs.slides._sldIdLst[0].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[0]

# --- SLIDE 1: Title ---
slide_layout = prs.slide_layouts[7]  # Blank
slide = prs.slides.add_slide(slide_layout)

# Title text
txBox = slide.shapes.add_textbox(Inches(0.75), Inches(2.0), Inches(8.5), Inches(1.5))
tf = txBox.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Coaching Call Discussions"
p.font.name = FONT_HEADING
p.font.size = Pt(36)
p.font.color.rgb = COLOR_PRIMARY

# Subtitle
p2 = tf.add_paragraph()
p2.text = "Monthly Report"
p2.font.name = FONT_BODY
p2.font.size = Pt(20)
p2.font.color.rgb = COLOR_ACCENT

# Customer + period
txBox2 = slide.shapes.add_textbox(Inches(0.75), Inches(4.0), Inches(6), Inches(1))
tf2 = txBox2.text_frame
p3 = tf2.paragraphs[0]
p3.text = f"{customer_id}  |  {start_date} to {end_date}"
p3.font.name = FONT_BODY
p3.font.size = Pt(14)
p3.font.color.rgb = COLOR_DARK


# --- SLIDE 2: Coaching Engagement by Wellbeing Topic ---
slide = prs.slides.add_slide(prs.slide_layouts[7])
add_section_title(slide, "Coaching Engagement by Wellbeing Topic", Inches(0.4), Inches(0.3))
add_branded_table(slide, df_engagement,
                  left=Inches(0.4), top=Inches(0.8),
                  width=Inches(9.2), height=Inches(5.5),
                  pct_cols=['PCT_OF_MEMBERS'])


# --- SLIDE 3: Goal Status by Wellbeing Topic ---
slide = prs.slides.add_slide(prs.slide_layouts[7])
add_section_title(slide, "Goal Status by Wellbeing Topic", Inches(0.4), Inches(0.3))
add_branded_table(slide, df_goal_status_topic,
                  left=Inches(0.4), top=Inches(0.8),
                  width=Inches(9.2), height=Inches(5.5))


# --- SLIDE 4: Goal Status Distribution + Goal Progression ---
slide = prs.slides.add_slide(prs.slide_layouts[7])

# Goal Status Distribution (left side)
add_section_title(slide, "Goal Status Distribution", Inches(0.4), Inches(0.3))
add_branded_table(slide, df_goal_dist,
                  left=Inches(0.4), top=Inches(0.8),
                  width=Inches(4.3), height=Inches(2.0),
                  pct_cols=['GOAL_PCT'])

# Goal Progression (below)
add_section_title(slide, "Goal Progression", Inches(0.4), Inches(3.2))
add_branded_table(slide, df_goal_prog,
                  left=Inches(0.4), top=Inches(3.7),
                  width=Inches(9.2), height=Inches(3.5),
                  pct_cols=['COMPLETION_RATE'])


# --- SLIDE 5: Tobacco Coaching Focus ---
slide = prs.slides.add_slide(prs.slide_layouts[7])
add_section_title(slide, "Tobacco Coaching Focus", Inches(0.4), Inches(0.3))
add_branded_table(slide, df_tobacco,
                  left=Inches(0.4), top=Inches(0.8),
                  width=Inches(5.0), height=Inches(2.5))


# --- Save ---
period_str = f"{start_date.replace('-','')}_{end_date.replace('-','')}"
output_filename = f"coaching_report_{customer_id}_{period_str}.pptx"
output_path = os.path.join(SCRIPT_DIR, output_filename)
prs.save(output_path)
log.info(f"Report saved: {output_path}")
print(f"\nDone! {output_path}")
