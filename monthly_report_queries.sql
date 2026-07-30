-- =========================================================================
-- MONTHLY COACHING REPORT: Dashboard Queries
-- =========================================================================
-- Reads from the persistent Carefirst_Sandbox tables (populated by
-- monthly_refresh.sql). Run these after the monthly refresh completes.
--
-- Filter by customer: add WHERE CUSTOMERID = 'HP_SCCareFirst' to each query.
-- Filter by date range: add WHERE CALL_DATE BETWEEN '2025-01-01' AND '2025-06-30'.
-- =========================================================================


-- =========================================================================
-- SECTION 1: Coaching Engagement by Wellbeing Topic
-- Members who had calls in each topic, with their goal counts
-- =========================================================================
SELECT
    T.REPORT_TOPIC                              AS WELLBEING_TOPIC,
    COUNT(DISTINCT T.CURRENTGUID)               AS MEMBERS,
    ROUND(COUNT(DISTINCT T.CURRENTGUID) * 100.0
        / NULLIFZERO(TOTALS.TOTAL_MEMBERS), 1)  AS PCT_OF_MEMBERS,
    COUNT(DISTINCT CASE WHEN G.GOAL_STATUS = 'Completed' THEN G.MEMBERACTION_ID END) AS COMPLETED_GOALS,
    COUNT(DISTINCT CASE WHEN G.GOAL_STATUS = 'In Progress' THEN G.MEMBERACTION_ID END) AS OPEN_IN_PROGRESS_GOALS
FROM Carefirst_Sandbox.COACHING_CALL_TOPICS T
CROSS JOIN (
    SELECT COUNT(DISTINCT CURRENTGUID) AS TOTAL_MEMBERS
    FROM Carefirst_Sandbox.COACHING_CALL_TOPICS
) TOTALS
LEFT JOIN Carefirst_Sandbox.COACHING_CALL_GOALS G
    ON T.CURRENTGUID = G.CURRENTGUID
GROUP BY 1, TOTALS.TOTAL_MEMBERS
ORDER BY 2 DESC;


-- =========================================================================
-- SECTION 2: Goal Status by Wellbeing Topic
-- Goal counts per topic broken out by status
-- =========================================================================
SELECT
    T.REPORT_TOPIC                              AS WELLBEING_TOPIC,
    COUNT(DISTINCT CASE WHEN G.GOAL_STATUS = 'Completed' THEN G.MEMBERACTION_ID END) AS COMPLETED,
    COUNT(DISTINCT CASE WHEN G.GOAL_STATUS = 'In Progress' THEN G.MEMBERACTION_ID END) AS IN_PROGRESS,
FROM Carefirst_Sandbox.COACHING_CALL_TOPICS T
JOIN Carefirst_Sandbox.COACHING_CALL_GOALS G
    ON T.CURRENTGUID = G.CURRENTGUID
GROUP BY 1
ORDER BY 2 DESC;


-- =========================================================================
-- SECTION 3: Goal Status Distribution (overall)
-- =========================================================================
SELECT
    GOAL_STATUS,
    COUNT(*)                                    AS COUNT,
    ROUND(COUNT(*) * 100.0
        / NULLIFZERO(SUM(COUNT(*)) OVER()), 1) AS GOAL_PCT
FROM Carefirst_Sandbox.COACHING_CALL_GOALS
WHERE GOAL_STATUS IN ('Completed','In Progress','Withdrawn')
GROUP BY 1
ORDER BY CASE GOAL_STATUS
    WHEN 'Completed' THEN 1
    WHEN 'In Progress' THEN 2
    WHEN 'Withdrawn' THEN 4
END;


-- =========================================================================
-- SECTION 4: Goal Progression by Domain
-- Total goals, completed, and completion rate per goal domain
-- =========================================================================
SELECT
    GOAL_DOMAIN,
    COUNT(*)                                    AS TOTAL_GOALS,
    SUM(CASE WHEN GOAL_STATUS = 'Completed' THEN 1 ELSE 0 END) AS COMPLETED,
    ROUND(SUM(CASE WHEN GOAL_STATUS = 'Completed' THEN 1 ELSE 0 END) * 100.0
        / NULLIFZERO(COUNT(*)), 1)             AS COMPLETION_RATE
FROM Carefirst_Sandbox.COACHING_CALL_GOALS
GROUP BY 1
ORDER BY CASE GOAL_DOMAIN
    WHEN 'Gaps in Care' THEN 1
    WHEN 'Exercise' THEN 2
    WHEN 'Nutrition' THEN 3
    WHEN 'Weight Management' THEN 4
    WHEN 'Tobacco Cessation' THEN 5
    WHEN 'Mental/Behavioral Health' THEN 6
    WHEN 'Stress Management' THEN 7
    WHEN 'Condition Management' THEN 8
    WHEN 'Financial' THEN 9
    WHEN 'Social' THEN 10
    WHEN 'Spiritual' THEN 11
    ELSE 12
END;


-- =========================================================================
-- SECTION 5: Tobacco Coaching Focus
-- =========================================================================
SELECT
    'Tobacco Participants' AS METRIC,
    COUNT(DISTINCT TB.CURRENTGUID)::VARCHAR AS VALUE
FROM Carefirst_Sandbox.COACHING_CALL_TOBACCO TB

UNION ALL

SELECT
    'Active Tobacco Participants',
    COUNT(DISTINCT G.CURRENTGUID)::VARCHAR
FROM Carefirst_Sandbox.COACHING_CALL_TOBACCO TB
JOIN Carefirst_Sandbox.COACHING_CALL_GOALS G
    ON TB.CURRENTGUID = G.CURRENTGUID
    AND G.GOAL_DOMAIN = 'Tobacco Cessation'
    AND G.GOAL_STATUS = 'In Progress'

UNION ALL

SELECT
    'Goals Completed',
    COUNT(*)::VARCHAR
FROM Carefirst_Sandbox.COACHING_CALL_GOALS G
WHERE G.GOAL_DOMAIN = 'Tobacco Cessation'
  AND G.GOAL_STATUS = 'Completed'

UNION ALL

SELECT
    'Goals In Progress',
    COUNT(*)::VARCHAR
FROM Carefirst_Sandbox.COACHING_CALL_GOALS G
WHERE G.GOAL_DOMAIN = 'Tobacco Cessation'
  AND G.GOAL_STATUS = 'In Progress'

UNION ALL

SELECT
    'Completion Rate',
    ROUND(
        SUM(CASE WHEN GOAL_STATUS = 'Completed' THEN 1 ELSE 0 END) * 100.0
        / NULLIFZERO(COUNT(*)), 1
    )::VARCHAR || '%'
FROM Carefirst_Sandbox.COACHING_CALL_GOALS
WHERE GOAL_DOMAIN = 'Tobacco Cessation'
  AND GOAL_STATUS IN ('Completed','In Progress');
