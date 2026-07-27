# Coaching Call Discussion Report

This report captures what happens during coaching calls: what topics are discussed, whether tobacco was addressed, and what goals are set with members.

---

## What This Report Produces

| Section | What It Answers | Output |
|---------|----------------|--------|
| Topics | What was discussed on each coaching call? | Count of calls and members per topic category |
| Tobacco | Was tobacco ever a focus of coaching? | Yes/No per member |
| Goals | What goals were set, and what's their status? | Goals by domain, type, status, and number |

---

## How Topics Are Assigned

The report starts from **every successful outbound coaching call**. For each call, we determine the discussion topic using a priority system. The first tier that produces a result wins:

```
Call comes in
  |
  v
[Tier 1] Did the coach fill out the topic form on this date?
  YES --> use that topic
  NO  --> continue
         |
         v
[Tier 1.5] Did the coach write a goal or topic detail on this date?
  YES (keywords match) --> use inferred topic
  NO  --> continue
         |
         v
[Tier 2] Is the call type itself a topic? (Tobacco, Dietary Referral, Specialty, Clinical)
  YES --> use call type as topic
  NO  --> continue
         |
         v
[Tier 3] Did this member have a topic recorded in the last 180 days?
  YES --> carry forward that topic
  NO  --> continue
         |
         v
[Tier 4] --> "General"
```

### Priority 1: Direct Recording

The coach selected a topic in the call workflow form on the same date as the call.

- Question 502533: "Highest priority topic" (Exercise, Healthy Eating, Weight, Stress, Tobacco, Other)
- Question 502758: "Select a pathway" (Physical activity, Nutrition/Weight management, Mental well-being, Restorative sleep, Positive social connections, Avoidance of risky substances)
- Question 502599: "Primary call focus" (DM only: Diabetes, Asthma, CAD, etc.)

### Priority 1.5: Inferred from Notes

No topic form was filled out, but the coach wrote a goal or topic detail on the same date. We classify the free text using keyword matching (see full mapping in the Topic Mapping Reference section below).

Sources: Question 502534 (topic details), 502833 (long-term goal), 502616 (DM goal)

### Priority 2: Call Type

The call type itself tells us the topic:

- Tobacco call --> Tobacco Cessation
- Dietary Referral call --> Nutrition
- Specialty call --> Chronic Disease State
- Clinical call --> Chronic Disease State

### Priority 3: Last Known Topic

The member has discussed a topic on a previous call (within the last 180 days). We carry forward their most recent topic. If their last recorded topic was more than 6 months ago, we don't use it — it may no longer be relevant.

### Priority 4: General

None of the above produced a topic. Typically:

- Engagement/outreach calls (not a discussion, just reconnecting)
- Calls before 2021 (topic recording didn't exist yet)
- Members whose coach never fills out the topic form

---

## Topic Mapping Reference

Each coaching call is assigned exactly one topic category. The mapping depends on which tier resolved the topic:

### Tier 1 Mapping (Direct form responses to report categories)

Two versions of the topic form exist. Both map to the same report categories:

| Coach's Selection (Question 502533) | Coach's Selection (Question 502758) | Report Category |
|-------------------------------------|-------------------------------------|-----------------|
| Exercise | Physical activity | **Exercise** |
| Healthy Eating | Nutrition/Weight management | **Nutrition** |
| Weight | *(not available)* | **Weight Management** |
| Stress | Mental well-being | **Stress Management** |
| *(not available)* | Restorative sleep | **Sleep Management** |
| Tobacco | *(not available)* | **Tobacco Cessation** |
| *(not available)* | Avoidance of risky substances | **Behavioral Health** |
| *(not available)* | Positive social connections | **Social Support** |
| Other - physical/social | *(not available)* | **Social Support** |
| Other | *(not available)* | **Other** |

For DM calls (question 502599), all responses (Diabetes, Asthma, CAD, COPD, HF, etc.) map to **Chronic Disease State**.

### Tier 1.5 Mapping (Keyword inference from free text)

When the topic form wasn't filled out but the coach wrote notes on that date, we classify by looking for keywords. The CASE evaluates top-down, so the first keyword match wins:

| Keywords Found | Report Category |
|----------------|-----------------|
| weight, lbs, pound, BMI | **Weight Management** |
| exercise, walk, run, gym, active, steps | **Exercise** |
| stress, anxiety, depression, mental, mindful | **Stress Management** |
| sleep | **Sleep Management** |
| eat, diet, nutrition, meal, food, calorie | **Nutrition** |
| cholesterol, A1C, blood pressure, diabetes, asthma, COPD | **Chronic Disease State** |
| tobacco, smoking, quit, nicotine | **Tobacco Cessation** |
| *(none of the above)* | *(skip — let Tier 2/3/4 handle)* |

Note: "quit" only triggers Tobacco Cessation because it's evaluated AFTER the exercise/nutrition keywords. A note saying "quit eating junk food" would match "eat" first and become Nutrition.

### Tier 2 Mapping (Call type inference)

| Call Type | Report Category |
|-----------|-----------------|
| Tobacco | **Tobacco Cessation** |
| Dietary Referral | **Nutrition** |
| Specialty | **Chronic Disease State** |
| Clinical | **Chronic Disease State** |

### Tier 3 Mapping (Prior topic)

Uses the same mapping as Tier 1 — the prior topic is a raw response from question 502533 or 502758.

### Tier 4

All remaining calls: **General**

---

## How Goals Work

Goals come from the coaching platform's formal goal system (`SCP.AH_MEMBER_ACTION`). Each goal has:

### Goal Type (how it was created)

| Type | What It Means |
|------|---------------|
| Coach-Created | The coach wrote a custom goal for the member (e.g., "Walk 3x/week for 30 minutes") |
| System-Recommended | The platform suggested a goal from an evidence-based menu and the coach activated it (e.g., "Get cholesterol screened annually") |

### Goal Domain (what area of wellbeing)

| Domain | Examples |
|--------|----------|
| Gaps in Care | Screening, vaccines, PCP, preventive care |
| Exercise | Exercise |
| Nutrition | Nutrition, healthy eating, diet |
| Weight Management | Weight, BMI, weight loss |
| Tobacco Cessation | Tobacco |
| Mental/Behavioral Health | Stress, depression, mental health, mindfulness, alcohol |
| Stress Management | Sleep, stress management (explicit goal name), CPAP |
| Condition Management | Diabetes, blood pressure, cholesterol, asthma, COPD, CAD, HF, medication adherence, appointments, self-management, work items |
| Financial | Finances, financial wellness |
| Social | Social wellness |
| Spiritual | Spiritual health |

### Goal Status

| Status | ActionStatus_ID | Meaning |
|--------|-----------------|---------|
| Not Started | 1 | Goal was set but member hasn't begun working on it |
| In Progress | 2 | Goal is active and being worked on |
| Completed | 3, 5 | Goal was achieved |
| Withdrawn | 4 | Member refused or goal was abandoned |

### Goal Number

Goals are numbered 1-6 per member in the order they were set. The report spec allows up to 3 concurrent goals and a maximum of 6 goals per enrollment.

---

## Tobacco Flag

A member is flagged "Yes" for tobacco if they ever had:
- "Tobacco" explicitly selected as their topic (question 502533), OR
- A call with call type "Tobacco"

This does NOT include members who discussed alcohol or marijuana (those fall under Mental/Behavioral Health but are not tobacco-specific).

---

## Important Notes

### Topic Coverage

Topic recording was introduced in mid-2021. Before that, all calls show as "General." Current coverage (2024+) is approximately 70% with all tiers combined. The call universe only includes billable/interaction-eligible call types (filtered via CALLTYPE_XREF_VW), which removes non-coaching calls like surveys and supervisor support.

### Call Filtering

Only calls that are billable (`PPPY_BILL_ELIG = 'Y'`) or interaction-eligible (`INTERACTION_ELIG = 'Y'`) per the CALLTYPE_XREF_VW lookup are included. This keeps the denominator aligned with standard coaching metrics and removes noise from administrative call types.

### Multi-Topic Calls

Coaches can select multiple topics per call (38% of calls have 2+ topics selected). Since the report requires one topic per call, we pick the member's **most frequently discussed topic** as the tiebreaker.

### Goal Type Limitation

The original report spec references Achievement/Habit/Learning as goal types. This classification only exists in the CareFirst Asset Health system (one client). For the broader platform, Coach-Created vs System-Recommended is the available distinction.

---

## Report Output Tables

### Output 1: Topics Summary

| Call Topic | Members | % Members | Calls | % of Calls |
|------------|---------|-----------|-------|------------|
| Exercise | ... | ... | ... | ... |
| Nutrition | ... | ... | ... | ... |
| General | ... | ... | ... | ... |

### Output 2: Tobacco

| Tobacco Discussed | Members | % Members | Calls | % of Calls |
|-------------------|---------|-----------|-------|------------|
| Yes | ... | ... | ... | ... |
| No | ... | ... | ... | ... |
| Grand Total | ... | ... | ... | ... |

### Output 3: Goal Type & Domain

| Goal Domain | Coach-Created | % | System-Recommended | % |
|-------------|---------------|---|--------------------|----|
| Gaps in Care | ... | ... | ... | ... |
| Exercise | ... | ... | ... | ... |
| Nutrition | ... | ... | ... | ... |
| Weight Management | ... | ... | ... | ... |
| Tobacco Cessation | ... | ... | ... | ... |
| Mental/Behavioral Health | ... | ... | ... | ... |
| Stress Management | ... | ... | ... | ... |
| Condition Management | ... | ... | ... | ... |

### Output 4: Goal Domain & Status

| Goal Domain | Not Started | % | In Progress | % | Completed | % | Withdrawn | % |
|-------------|-------------|---|-------------|---|-----------|---|-----------|---|
| Gaps in Care | ... | ... | ... | ... | ... | ... | ... | ... |
| Exercise | ... | ... | ... | ... | ... | ... | ... | ... |

### Output 5: Goal Number & Status

| Goal Number | Not Started | % | In Progress | % | Completed | % | Withdrawn | % |
|-------------|-------------|---|-------------|---|-----------|---|-----------|---|
| 1 | ... | ... | ... | ... | ... | ... | ... | ... |
| 2 | ... | ... | ... | ... | ... | ... | ... | ... |

### Output 6: Goal Type & Status

| Goal Type | Not Started | % | In Progress | % | Completed | % | Withdrawn | % |
|-----------|-------------|---|-------------|---|-----------|---|-----------|---|
| Coach-Created | ... | ... | ... | ... | ... | ... | ... | ... |
| System-Recommended | ... | ... | ... | ... | ... | ... | ... | ... |

---

## How to Run

### Automated (monthly refresh to Vertica)

The pipeline runs on the **1st of each month at 6 AM** via launchd (local) or Airflow (production). It rebuilds the data from scratch and writes to `Carefirst_Sandbox`:

| Table | Contents |
|-------|----------|
| `Carefirst_Sandbox.COACHING_CALL_TOPICS` | One row per member per call date — topic, tier source, call type |
| `Carefirst_Sandbox.COACHING_CALL_GOALS` | One row per goal — domain, type, status, goal number |
| `Carefirst_Sandbox.COACHING_CALL_TOBACCO` | One row per member flagged for tobacco discussion |

All tables include a `REFRESH_DATE` column showing when the data was last rebuilt.

```bash
# Manual trigger
python3 monthly_refresh.py

# Via launchd
launchctl start com.sharecare.ETL_Monthly_1st6am_CoachingDiscussions
```

Log: `~/Library/Logs/coaching-discussions.log`

### Airflow deployment

The same `monthly_refresh.py` can be called from an Airflow DAG:

```python
from coaching_call_discussions.monthly_refresh import main

# In a PythonOperator:
task = PythonOperator(task_id='coaching_discussions_refresh', python_callable=main)
```

Or as a BashOperator pointing to the script. Connection is handled via `VERTICA_*` env vars (see `.env.example`) when the shared `db_connect.py` helper is not available.

### Excel export (ad-hoc formatted report)
```bash
# All customers, all dates
python3 run_report.py

# Single customer
python3 run_report.py HP_SCCareFirst

# Customer + date range
python3 run_report.py HP_SCCareFirst 2025-04-01 2025-06-30
```

Output: `coaching_call_discussions.xlsx` (or `coaching_call_discussions_HP_SCCareFirst.xlsx`)

### DbVisualizer / Manual SQL
Run the entire `coaching_call_topics_goals.sql` file in a Vertica session. Add filters to the `CALLS_ONE_PER_DAY` subquery:
```sql
    WHERE UPPER(MC.CALL_STATUS) = 'SUCCESSFUL'
      AND UPPER(MC.DIRECTION) = 'OUTBOUND'
      AND MC.CUSTOMERID = 'ER_SHBP'                        -- customer filter
      AND TRUNC(MC.ENCOUNTERDATETIME)::DATE >= '2024-01-01' -- date filter
```

---

## Scheduling

| Label | Schedule | Log |
|-------|----------|-----|
| `com.sharecare.ETL_Monthly_1st6am_CoachingDiscussions` | 1st of month, 6:00 AM | `~/Library/Logs/coaching-discussions.log` |

### Install/reload
```bash
cp com.sharecare.ETL_Monthly_1st6am_CoachingDiscussions.plist ~/Library/LaunchAgents/
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.sharecare.ETL_Monthly_1st6am_CoachingDiscussions.plist
```

### Disable
```bash
launchctl bootout gui/$(id -u) ~/Library/LaunchAgents/com.sharecare.ETL_Monthly_1st6am_CoachingDiscussions.plist
```

---

## What's Not Included

| Item | Why |
|------|-----|
| Inbound calls | Report spec covers outbound coaching touches only (coach-initiated) |
| Unsuccessful calls | Only completed/connected calls count as a coaching discussion |
| Non-billable call types | UNKNOWN, Survey - Follow Up, Support [Supervisor] are administrative, not coaching |
| Goal Type: Achievement/Habit/Learning | Only exists in CareFirst Asset Health (one client). Not in the platform. |
| Historical goal status changes | SCP stores current state only. Use SC_KPI.goal_history for "when did status change" |
| Calls before 2017 | MEMBER_CALL_DATA starts in 2017 |

---

## Technical Details

### Source Tables

| Table | Schema | Role |
|-------|--------|------|
| MEMBER_CALL_DATA | BI_REPORTING | Starting point: all successful outbound coaching calls |
| COACH_NOTES_WORKFLOW | ENT_WH | Topic responses and goal text (LM programs) |
| COACH_NOTES_WORKFLOW_DM | ENT_WH | Topic responses and goal text (DM programs) |
| AH_MEMBER_ACTION | SCP | Formal coaching goals (current state) |
| COACHING_ENROLLMENT_MODEL | BI_REPORTING | Links coaching account IDs to member GUIDs |

### Files in This Folder

| File | Purpose |
|------|---------|
| `coaching_call_topics_goals.sql` | Interactive SQL query with output SELECTs (for DbVisualizer / ad-hoc) |
| `monthly_refresh.sql` | Pipeline SQL — same logic but ends with TRUNCATE+INSERT to Vertica |
| `monthly_refresh.py` | Automated pipeline — executes SQL, logs progress, persists to Carefirst_Sandbox |
| `monthly_report_queries.sql` | Dashboard queries — reads from persistent tables to produce report output |
| `run_report.py` | Ad-hoc — runs interactive SQL and exports formatted Excel |
| `ddl_create_tables.sql` | One-time DDL to create the 3 persistent tables |
| `run_coaching_refresh.sh` | Shell wrapper for launchd |
| `com.sharecare.ETL_Monthly_1st6am_CoachingDiscussions.plist` | launchd schedule (1st of month, 6 AM) |
| `.env.example` | Template showing credential env vars (for Airflow / non-local) |
| `requirements.txt` | Python dependencies |
| `README.md` | This documentation |
| `.gitignore` | Excludes .xlsx, .log, .env, __pycache__ |

---

## Known Issues

| Issue | Impact | Status |
|-------|--------|--------|
| Multi-topic calls deduped to one | 38% of calls have 2+ topics. Report shows one per call. | Accepted: report spec requires one per call. Most-frequent-topic tiebreak used. |
| Goal Type (Achievement/Habit/Learning) | Not available platform-wide | Accepted: Coach-Created/System-Recommended used as alternative |
| Not Started goals dominate volume | 5.6M "Not Started" vs 67K "In Progress" | Monitor: may need to exclude or separate in output |
| ActionStatus_ID 5 mapped to Completed | 211K rows with close dates, same behavior as status 3 | Accepted: both represent achieved goals |

---

## Next Steps

### Immediate (before first client delivery)

1. **Run the full SQL for one customer** (e.g., HP_SCCareFirst or ER_SHBP) and validate output counts match expectations. Add `WHERE CUSTOMERID = 'X'` to the calls step.
2. **Review the Tier breakdown output** (Output 1B) to confirm the fallback tiers are contributing as expected and the General bucket is acceptable.
3. **Validate goal counts** against a known client's coaching report to confirm alignment.

### Short-term improvements

4. **Add question 503721 (Progress toward goal)** to supplement goal status. Members at 100% progress in the workflow could be marked Completed even if their SCP goal wasn't formally closed.
5. **Add DM prior topic fallback** (Tier 3 currently only uses LM topics for lookback; DM members with no DM topic on a call date could use their most recent DM topic).
6. **Consider excluding "Not Started" goals from summary outputs** — 5.6M goals are status 1 (Not Started) vs 67K In Progress. May overwhelm the output. Could show as a separate line or exclude from domain/status crosstabs.

### Future enhancements

7. **Multi-topic output** -- optional output showing ALL topics per call (not deduped) for stakeholders who want the full picture of what was discussed.
8. **Trend analysis** -- topic distribution by month/quarter to show how coaching focus evolves over time.
9. **Goal-to-call linkage** -- tie goals to specific calls by date proximity to show "goals set during this call" rather than just "all goals for this member."
10. **Coaching Call Types investigation** -- explore SCP.AH_MEMBER_ENCOUNTERS for a richer call-type taxonomy.

