-- =========================================================================
-- COACHING DISCUSSION REPORT: TOPICS, TOBACCO STATUS, AND GOALS (v5)
-- =========================================================================
-- APPROACH: CALL-FIRST with TIERED TOPIC FALLBACK
--   Tier 1:   Exact date match (502533/502758 LM + 502599 DM)
--   Tier 1.5: Goal/detail text inference (502534, 502833, 502616)
--   Tier 2:   Call type inference (Tobacco->Tobacco Cessation, Specialty/Clinical->Chronic Disease State, etc.)
--   Tier 3:   Most recent prior topic (180-day lookback)
--   Tier 4:   "General" (no topic determinable)
--
-- v5 CHANGES:
--   - Tier 1.5 scoped to only account/dates without Tier 1 match (efficiency)
--   - Tier 3 scoped to only calls not resolved by Tiers 1-2 (efficiency)
--   - Removed redundant correlated subqueries in Output 2
--   - Single-pass keyword classification via reusable pattern
-- =========================================================================


-- =========================================================================
-- STEP 1A: CALLS - one per account per day, billable types only
-- =========================================================================
DROP TABLE IF EXISTS CALLS_ONE_PER_DAY;
CREATE LOCAL TEMP TABLE CALLS_ONE_PER_DAY ON COMMIT PRESERVE ROWS AS
SELECT ACCOUNT, GUID, CUSTOMERID, CALL_TYPE, CALL_DATE
FROM (
    SELECT
        MC.ACCOUNT,
        NVL(MC.DERIVED_GUID, MC.GUID)           AS GUID,
        MC.CUSTOMERID,
        TRIM(MC.DESCRIPTION)                     AS CALL_TYPE,
        TRUNC(MC.ENCOUNTERDATETIME)::DATE        AS CALL_DATE,
        ROW_NUMBER() OVER (
            PARTITION BY MC.ACCOUNT, TRUNC(MC.ENCOUNTERDATETIME)::DATE
            ORDER BY MC.ENCOUNTERDATETIME DESC
        ) AS RN
    FROM BI_REPORTING.MEMBER_CALL_DATA MC
    JOIN (
        SELECT DISTINCT TRIM(CALL_DESCRIPTIONS) AS CALL_DESC
        FROM ENT_WH.CALLTYPE_XREF_VW
        WHERE PPPY_BILL_ELIG = 'Y' OR INTERACTION_ELIG = 'Y'
    ) CT ON UPPER(TRIM(MC.DESCRIPTION)) = UPPER(CT.CALL_DESC)
    WHERE UPPER(MC.CALL_STATUS) = 'SUCCESSFUL'
      AND UPPER(MC.DIRECTION) = 'OUTBOUND'
) X WHERE RN = 1;


-- =========================================================================
-- STEP 1B: LM TOPICS - one per account per day (frequency tiebreak)
-- =========================================================================
DROP TABLE IF EXISTS LM_TOPICS_DEDUPED;
CREATE LOCAL TEMP TABLE LM_TOPICS_DEDUPED ON COMMIT PRESERVE ROWS AS
SELECT ACCOUNT, TOPIC_DATE, RESPONSE_TEXT
FROM (
    SELECT T.ACCOUNT, TRUNC(T.RESPONSE_DATE)::DATE AS TOPIC_DATE, T.RESPONSE_TEXT,
        ROW_NUMBER() OVER (
            PARTITION BY T.ACCOUNT, TRUNC(T.RESPONSE_DATE)::DATE
            ORDER BY FREQ.CNT DESC, T.RESPONSE_TEXT ASC
        ) AS RN
    FROM ENT_WH.COACH_NOTES_WORKFLOW T
    JOIN (
        SELECT ACCOUNT, RESPONSE_TEXT, COUNT(*) AS CNT
        FROM ENT_WH.COACH_NOTES_WORKFLOW
        WHERE QUESTION_ID IN ('502533','502758')
        GROUP BY 1, 2
    ) FREQ ON T.ACCOUNT = FREQ.ACCOUNT AND T.RESPONSE_TEXT = FREQ.RESPONSE_TEXT
    WHERE T.QUESTION_ID IN ('502533','502758')
) X WHERE RN = 1;

-- =========================================================================
-- STEP 1C: DM TOPICS - one per account per day
-- =========================================================================
DROP TABLE IF EXISTS DM_TOPICS_DEDUPED;
CREATE LOCAL TEMP TABLE DM_TOPICS_DEDUPED ON COMMIT PRESERVE ROWS AS
SELECT ACCOUNT, TOPIC_DATE, RESPONSE_TEXT
FROM (
    SELECT T.ACCOUNT, TRUNC(T.RESPONSE_DATE)::DATE AS TOPIC_DATE, T.RESPONSE_TEXT,
        ROW_NUMBER() OVER (
            PARTITION BY T.ACCOUNT, TRUNC(T.RESPONSE_DATE)::DATE
            ORDER BY FREQ.CNT DESC, T.RESPONSE_TEXT ASC
        ) AS RN
    FROM ENT_WH.COACH_NOTES_WORKFLOW_DM T
    JOIN (
        SELECT ACCOUNT, RESPONSE_TEXT, COUNT(*) AS CNT
        FROM ENT_WH.COACH_NOTES_WORKFLOW_DM
        WHERE QUESTION_ID = '502599'
        GROUP BY 1, 2
    ) FREQ ON T.ACCOUNT = FREQ.ACCOUNT AND T.RESPONSE_TEXT = FREQ.RESPONSE_TEXT
    WHERE T.QUESTION_ID = '502599'
) X WHERE RN = 1;


-- =========================================================================
-- STEP 1D: Identify calls NOT resolved by Tier 1
-- These are the only calls that need fallback tiers (1.5, 2, 3)
-- =========================================================================
DROP TABLE IF EXISTS CALLS_NO_TIER1;
CREATE LOCAL TEMP TABLE CALLS_NO_TIER1 ON COMMIT PRESERVE ROWS AS
SELECT C.*
FROM CALLS_ONE_PER_DAY C
LEFT JOIN LM_TOPICS_DEDUPED LM ON C.ACCOUNT = LM.ACCOUNT AND C.CALL_DATE = LM.TOPIC_DATE
LEFT JOIN DM_TOPICS_DEDUPED DM ON C.ACCOUNT = DM.ACCOUNT AND C.CALL_DATE = DM.TOPIC_DATE
WHERE LM.ACCOUNT IS NULL AND DM.ACCOUNT IS NULL;


-- =========================================================================
-- STEP 1E: TIER 1.5 - Text inference (only for calls without Tier 1)
-- Keyword-classify 502534 (topic detail), 502833 (goal), 502616 (DM goal)
-- =========================================================================
DROP TABLE IF EXISTS GOAL_TEXT_TOPICS;
CREATE LOCAL TEMP TABLE GOAL_TEXT_TOPICS ON COMMIT PRESERVE ROWS AS
SELECT ACCOUNT, TOPIC_DATE, INFERRED_TOPIC
FROM (
    SELECT ACCOUNT, TOPIC_DATE, INFERRED_TOPIC,
        ROW_NUMBER() OVER (PARTITION BY ACCOUNT, TOPIC_DATE ORDER BY SRC_PRI ASC) AS RN
    FROM (
        SELECT
            T.ACCOUNT,
            TRUNC(T.RESPONSE_DATE)::DATE AS TOPIC_DATE,
            -- Keyword classification. Order matters: more specific matches first.
            -- "walking to lose weight" -> Weight Management (not Exercise)
            -- "quit smoking" -> Tobacco Cessation (not caught by Exercise "quit")
            CASE
                WHEN UPPER(T.RESPONSE_TEXT) LIKE '%WEIGHT%' OR UPPER(T.RESPONSE_TEXT) LIKE '%LBS%'
                    OR UPPER(T.RESPONSE_TEXT) LIKE '%POUND%' OR UPPER(T.RESPONSE_TEXT) LIKE '%BMI%'
                    THEN 'Weight Management'
                WHEN UPPER(T.RESPONSE_TEXT) LIKE '%EXERCIS%' OR UPPER(T.RESPONSE_TEXT) LIKE '%WALK%'
                    OR UPPER(T.RESPONSE_TEXT) LIKE '%RUN %' OR UPPER(T.RESPONSE_TEXT) LIKE '%GYM%'
                    OR UPPER(T.RESPONSE_TEXT) LIKE '%ACTIVE%' OR UPPER(T.RESPONSE_TEXT) LIKE '%STEPS%'
                    THEN 'Exercise'
                WHEN UPPER(T.RESPONSE_TEXT) LIKE '%STRESS%' OR UPPER(T.RESPONSE_TEXT) LIKE '%ANXI%'
                    OR UPPER(T.RESPONSE_TEXT) LIKE '%DEPRESS%' OR UPPER(T.RESPONSE_TEXT) LIKE '%MENTAL%'
                    OR UPPER(T.RESPONSE_TEXT) LIKE '%MINDFUL%'
                    THEN 'Stress Management'
                WHEN UPPER(T.RESPONSE_TEXT) LIKE '%SLEEP%'
                    THEN 'Sleep Management'
                WHEN UPPER(T.RESPONSE_TEXT) LIKE '%EAT%' OR UPPER(T.RESPONSE_TEXT) LIKE '%DIET%'
                    OR UPPER(T.RESPONSE_TEXT) LIKE '%NUTRI%' OR UPPER(T.RESPONSE_TEXT) LIKE '%MEAL%'
                    OR UPPER(T.RESPONSE_TEXT) LIKE '%FOOD%' OR UPPER(T.RESPONSE_TEXT) LIKE '%CALORI%'
                    THEN 'Nutrition'
                WHEN UPPER(T.RESPONSE_TEXT) LIKE '%CHOLEST%' OR UPPER(T.RESPONSE_TEXT) LIKE '%A1C%'
                    OR UPPER(T.RESPONSE_TEXT) LIKE '%BLOOD PRESSURE%' OR UPPER(T.RESPONSE_TEXT) LIKE '%DIABET%'
                    OR UPPER(T.RESPONSE_TEXT) LIKE '%ASTHMA%' OR UPPER(T.RESPONSE_TEXT) LIKE '%COPD%'
                    THEN 'Chronic Disease State'
                WHEN UPPER(T.RESPONSE_TEXT) LIKE '%TOBACCO%' OR UPPER(T.RESPONSE_TEXT) LIKE '%SMOK%'
                    OR UPPER(T.RESPONSE_TEXT) LIKE '%QUIT%' OR UPPER(T.RESPONSE_TEXT) LIKE '%NICOTINE%'
                    THEN 'Tobacco Cessation'
                ELSE NULL
            END AS INFERRED_TOPIC,
            CASE WHEN T.QUESTION_ID = '502534' THEN 1 ELSE 2 END AS SRC_PRI
        FROM ENT_WH.COACH_NOTES_WORKFLOW T
        -- Only process accounts/dates that need Tier 1.5
        JOIN CALLS_NO_TIER1 C ON T.ACCOUNT = C.ACCOUNT AND TRUNC(T.RESPONSE_DATE)::DATE = C.CALL_DATE
        WHERE T.QUESTION_ID IN ('502534','502833')
          AND T.RESPONSE_TEXT IS NOT NULL AND LENGTH(TRIM(T.RESPONSE_TEXT)) > 2
          AND UPPER(TRIM(T.RESPONSE_TEXT)) NOT IN ('N/A','-','NA','NONE','.','N/A.')
          -- Skip calls already handled by Tier 2
          AND UPPER(TRIM(C.CALL_TYPE)) NOT IN ('TOBACCO','DIETARY REFERRAL','SPECIALTY','CLINICAL')

        UNION ALL

        SELECT
            T.ACCOUNT,
            TRUNC(T.RESPONSE_DATE)::DATE AS TOPIC_DATE,
            'Chronic Disease State' AS INFERRED_TOPIC,
            3 AS SRC_PRI
        FROM ENT_WH.COACH_NOTES_WORKFLOW_DM T
        JOIN CALLS_NO_TIER1 C ON T.ACCOUNT = C.ACCOUNT AND TRUNC(T.RESPONSE_DATE)::DATE = C.CALL_DATE
        WHERE T.QUESTION_ID = '502616'
          AND T.RESPONSE_TEXT IS NOT NULL AND LENGTH(TRIM(T.RESPONSE_TEXT)) > 2
          AND UPPER(TRIM(T.RESPONSE_TEXT)) NOT IN ('N/A','-','NA','NONE','.','N/A.')
          AND UPPER(TRIM(C.CALL_TYPE)) NOT IN ('TOBACCO','DIETARY REFERRAL','SPECIALTY','CLINICAL')
    ) RAW
    WHERE INFERRED_TOPIC IS NOT NULL
) RANKED WHERE RN = 1;


-- =========================================================================
-- STEP 1F: TIER 3 - Most recent prior topic (180-day lookback)
-- Only for calls not resolved by Tiers 1, 1.5, or 2.
-- Excludes Tobacco/Dietary Referral (handled by Tier 2 in final assembly).
-- =========================================================================
DROP TABLE IF EXISTS PRIOR_TOPICS;
CREATE LOCAL TEMP TABLE PRIOR_TOPICS ON COMMIT PRESERVE ROWS AS
SELECT ACCOUNT, CALL_DATE, PRIOR_RESPONSE_TEXT
FROM (
    SELECT C.ACCOUNT, C.CALL_DATE, T.RESPONSE_TEXT AS PRIOR_RESPONSE_TEXT,
        ROW_NUMBER() OVER (PARTITION BY C.ACCOUNT, C.CALL_DATE ORDER BY T.RESPONSE_DATE DESC) AS RN
    FROM CALLS_NO_TIER1 C
    JOIN ENT_WH.COACH_NOTES_WORKFLOW T
        ON C.ACCOUNT = T.ACCOUNT
        AND TRUNC(T.RESPONSE_DATE)::DATE < C.CALL_DATE
        AND TRUNC(T.RESPONSE_DATE)::DATE >= C.CALL_DATE - 180
    WHERE T.QUESTION_ID IN ('502533','502758')
      AND UPPER(TRIM(C.CALL_TYPE)) NOT IN ('TOBACCO','DIETARY REFERRAL','SPECIALTY','CLINICAL')
      AND NOT EXISTS (
          SELECT 1 FROM GOAL_TEXT_TOPICS GT
          WHERE GT.ACCOUNT = C.ACCOUNT AND GT.TOPIC_DATE = C.CALL_DATE
      )
) X WHERE RN = 1;


-- =========================================================================
-- STEP 1G: FINAL COACHING_TOPICS - assemble all tiers
-- =========================================================================
DROP TABLE IF EXISTS COACHING_TOPICS;
CREATE LOCAL TEMP TABLE COACHING_TOPICS ON COMMIT PRESERVE ROWS AS
SELECT
    C.ACCOUNT, C.GUID, C.CUSTOMERID, C.CALL_DATE, C.CALL_TYPE,
    -- RAW_TOPIC: the direct response text from Tier 1 only (LM or DM form).
    -- NULL for calls resolved by Tiers 1.5-4. Used by Tobacco flag.
    COALESCE(LM.RESPONSE_TEXT, DM.RESPONSE_TEXT) AS RAW_TOPIC,
    -- REPORT_TOPIC: final assigned topic. CASE evaluates top-down = tier priority.
    -- First non-NULL tier match wins. No fan-out risk because each tier temp table
    -- is pre-deduped to one row per account/date.
    CASE
        WHEN UPPER(LM.RESPONSE_TEXT) = 'EXERCISE'                  THEN 'Exercise'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'HEALTHY EATING'            THEN 'Nutrition'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'WEIGHT'                    THEN 'Weight Management'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'STRESS'                    THEN 'Stress Management'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'TOBACCO'                   THEN 'Tobacco Cessation'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'OTHER'                     THEN 'Other'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'OTHER - PHYSICAL/SOCIAL'   THEN 'Social Support'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'PHYSICAL ACTIVITY'             THEN 'Exercise'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'NUTRITION/WEIGHT MANAGEMENT'   THEN 'Nutrition'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'MENTAL WELL-BEING'             THEN 'Stress Management'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'RESTORATIVE SLEEP'             THEN 'Sleep Management'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'POSITIVE SOCIAL CONNECTIONS'   THEN 'Social Support'
        WHEN UPPER(LM.RESPONSE_TEXT) = 'AVOIDANCE OF RISKY SUBSTANCES' THEN 'Behavioral Health'
        WHEN DM.RESPONSE_TEXT IS NOT NULL       THEN 'Chronic Disease State'
        WHEN GT.INFERRED_TOPIC IS NOT NULL      THEN GT.INFERRED_TOPIC
        -- TIER 2: CALL TYPE INFERENCE
        WHEN UPPER(TRIM(C.CALL_TYPE)) = 'TOBACCO'          THEN 'Tobacco Cessation'
        WHEN UPPER(TRIM(C.CALL_TYPE)) = 'DIETARY REFERRAL' THEN 'Nutrition'
        WHEN UPPER(TRIM(C.CALL_TYPE)) IN ('SPECIALTY','CLINICAL') THEN 'Chronic Disease State'
        WHEN PT.PRIOR_RESPONSE_TEXT IS NOT NULL THEN
            CASE
                WHEN UPPER(PT.PRIOR_RESPONSE_TEXT) IN ('EXERCISE','PHYSICAL ACTIVITY') THEN 'Exercise'
                WHEN UPPER(PT.PRIOR_RESPONSE_TEXT) IN ('HEALTHY EATING','NUTRITION/WEIGHT MANAGEMENT') THEN 'Nutrition'
                WHEN UPPER(PT.PRIOR_RESPONSE_TEXT) = 'WEIGHT' THEN 'Weight Management'
                WHEN UPPER(PT.PRIOR_RESPONSE_TEXT) IN ('STRESS','MENTAL WELL-BEING') THEN 'Stress Management'
                WHEN UPPER(PT.PRIOR_RESPONSE_TEXT) = 'RESTORATIVE SLEEP' THEN 'Sleep Management'
                WHEN UPPER(PT.PRIOR_RESPONSE_TEXT) = 'TOBACCO' THEN 'Tobacco Cessation'
                WHEN UPPER(PT.PRIOR_RESPONSE_TEXT) = 'AVOIDANCE OF RISKY SUBSTANCES' THEN 'Behavioral Health'
                WHEN UPPER(PT.PRIOR_RESPONSE_TEXT) IN ('POSITIVE SOCIAL CONNECTIONS','OTHER - PHYSICAL/SOCIAL') THEN 'Social Support'
                WHEN UPPER(PT.PRIOR_RESPONSE_TEXT) = 'OTHER' THEN 'Other'
                ELSE 'General'
            END
        ELSE 'General'
    END AS REPORT_TOPIC,
    CASE
        WHEN LM.RESPONSE_TEXT IS NOT NULL THEN 'Tier 1: Date Match (LM)'
        WHEN DM.RESPONSE_TEXT IS NOT NULL THEN 'Tier 1: Date Match (DM)'
        WHEN GT.INFERRED_TOPIC IS NOT NULL THEN 'Tier 1.5: Text Inference'
        WHEN UPPER(TRIM(C.CALL_TYPE)) IN ('TOBACCO','DIETARY REFERRAL','SPECIALTY','CLINICAL') THEN 'Tier 2: Call Type'
        WHEN PT.PRIOR_RESPONSE_TEXT IS NOT NULL THEN 'Tier 3: Prior Topic'
        ELSE 'Tier 4: None'
    END AS TOPIC_SOURCE,
    CASE WHEN LM.RESPONSE_TEXT IS NOT NULL THEN 'LM'
         WHEN DM.RESPONSE_TEXT IS NOT NULL THEN 'DM' ELSE NULL END AS PROGRAM_TYPE
FROM CALLS_ONE_PER_DAY C
LEFT JOIN LM_TOPICS_DEDUPED LM ON C.ACCOUNT = LM.ACCOUNT AND C.CALL_DATE = LM.TOPIC_DATE
LEFT JOIN DM_TOPICS_DEDUPED DM ON C.ACCOUNT = DM.ACCOUNT AND C.CALL_DATE = DM.TOPIC_DATE AND LM.ACCOUNT IS NULL
LEFT JOIN GOAL_TEXT_TOPICS GT ON C.ACCOUNT = GT.ACCOUNT AND C.CALL_DATE = GT.TOPIC_DATE AND LM.ACCOUNT IS NULL AND DM.ACCOUNT IS NULL
LEFT JOIN PRIOR_TOPICS PT ON C.ACCOUNT = PT.ACCOUNT AND C.CALL_DATE = PT.CALL_DATE AND LM.ACCOUNT IS NULL AND DM.ACCOUNT IS NULL AND GT.ACCOUNT IS NULL;


-- =========================================================================
-- SECTION 2: TOBACCO FLAG (scoped to actual tobacco only)
-- =========================================================================
DROP TABLE IF EXISTS COACHING_TOBACCO;
CREATE LOCAL TEMP TABLE COACHING_TOBACCO ON COMMIT PRESERVE ROWS AS
SELECT DISTINCT ACCOUNT, GUID, 'Yes' AS TOBACCO_DISCUSSED
FROM COACHING_TOPICS
WHERE UPPER(RAW_TOPIC) = 'TOBACCO'
   OR UPPER(CALL_TYPE) = 'TOBACCO';

-- =========================================================================
-- SECTION 3: GOALS (SCP.AH_MEMBER_ACTION, current enrollment only)
-- =========================================================================
-- GOAL DOMAINS (updated):
--   Gaps in Care       = screening, vaccines, PCP, preventive care
--   Exercise           = exercise
--   Nutrition          = nutrition, healthy eating, diet
--   Weight Management  = weight, BMI
--   Tobacco Cessation  = tobacco
--   Mental/Behavioral Health = stress, depression, mental health, mindfulness, alcohol
--   Stress Management  = sleep, stress management (standalone)
--   Condition Management = diabetes, blood pressure, cholesterol, medication adherence,
--                          appointments, self-management, work items, chronic conditions
--   Financial          = finances, financial wellness
--   Social             = social wellness
--   Spiritual          = spiritual health
--
-- GOAL STATUSES:
--   1 = Not Started, 2 = In Progress, 3 = Completed, 4 = Withdrawn, 5 = Completed
-- =========================================================================
DROP TABLE IF EXISTS COACHING_GOALS;
CREATE LOCAL TEMP TABLE COACHING_GOALS ON COMMIT PRESERVE ROWS AS
SELECT
    MA.ACCOUNT,
    CASE WHEN MA.ActionType_ID = 2 THEN 'Coach-Created'
         WHEN MA.ActionType_ID = 3 THEN 'System-Recommended' END AS GOAL_TYPE,
    CASE
        -- Exercise
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('EXERCISE','[EXERCISE]','*EXERCISE','[[*EXERCISE]]','*CIS EXERCISE','*CIS EXERICSE','EXERICSE','EXERISE','EXERCSE') THEN 'Exercise'
        -- Nutrition
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('HEALTHY EATING','[HEALTHY EATING]','[*HEALTHY EATING]','NUTRTION','HEATLHY EATING','HEALTH EATING','[NUTRITION]','DIET') THEN 'Nutrition'
        -- Weight Management
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('WEIGHT','[WEIGHT]','WEIGHT LOSS') THEN 'Weight Management'
        -- Tobacco Cessation
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('TOBACCO','[TOBACCO]') THEN 'Tobacco Cessation'
        -- Mental/Behavioral Health (stress, depression, mental health, mindfulness, alcohol)
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('STRESS','[*STRESS]','STRESS/EMOTIONAL WELL-BEING','DEPRESSION','MENTAL HEALTH','MINDFULNESS','ALCOHOL') THEN 'Mental/Behavioral Health'
        -- Stress Management (sleep, stress management as explicit goal name)
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('STRESS MANAGEMENT','SLEEP','CPAP') THEN 'Stress Management'
        -- Gaps in Care (screening, preventive, vaccines, PCP)
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('SCREENING','[PREVENTATIVE CARE]','[*PREVENTATIVE CARE]','[*CARE GAPS]','SCREENING SOC','VACCINE','[VACCINE]') THEN 'Gaps in Care'
        -- Condition Management (chronic conditions, medication, appointments, self-management)
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('DIABETES','[DIABETES]','DIABETES/CAD','CAD/DIABETES','DIABETES SOC','ASTHMA/COPD','ASTHMA','COPD','COPD/HF','[RESPIRATORY (ASTHMA/COPD)]','[ASTHMA/COPD]','[ASTHMA]','[COPD]','CAD/HF','CAD','HF','[CARDIAC (CAD/HF)]','[CAD]','[HF]','AFIB','FIBRO','OA','IBS','IBD','LBP','ARSD') THEN 'Condition Management'
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('CHOLESTEROL','CHOL','TRIG','BP','BLOOD PRESSURE') THEN 'Condition Management'
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('UTILIZATION RISK','HOSPITAL','[HOSPITALIZATION]','CRITICAL RISK EVENT FOLLOW-UP','APPOINTMENT ADHERENCE','APPOINTMENT','MAKE AND KEEP APPOINTMENTS','MEDICATION ADHERENCE','MEDICATION','MED ADHERENCE','SELF MANAGEMENT','SELF CARE','WORK ITEM','TIME MANAGEMENT','SCHEDULE ADHERENCE') THEN 'Condition Management'
        -- Financial
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('FINANCES','FINANCIAL WELLNESS','SAVINGS') THEN 'Financial'
        -- Social
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) = 'SOCIAL WELLNESS' THEN 'Social'
        -- Spiritual
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) = 'SPIRITUAL HEALTH' THEN 'Spiritual'
        -- Remaining physical items that don't fit cleanly elsewhere
        WHEN UPPER(SPLIT_PART(MA.Action_Name,':',1)) IN ('WATER','WATER INTAKE','PAIN MANAGEMENT') THEN 'Condition Management'
        -- Keyword fallback for full-sentence System-Recommended goals (no colon prefix)
        WHEN MA.Action_Name NOT LIKE '%:%' THEN
            CASE
                WHEN UPPER(MA.Action_Name) LIKE '%SCREENING%' OR UPPER(MA.Action_Name) LIKE '%VACCINE%'
                    OR UPPER(MA.Action_Name) LIKE '%FLU SHOT%' OR UPPER(MA.Action_Name) LIKE '%ANNUAL WELLNESS VISIT%'
                    OR UPPER(MA.Action_Name) LIKE '%CANCER SCREENING%' OR UPPER(MA.Action_Name) LIKE '%DENTAL EXAM%'
                    OR UPPER(MA.Action_Name) LIKE '%PRIMARY CARE PROVIDER%'
                    THEN 'Gaps in Care'
                WHEN UPPER(MA.Action_Name) LIKE '%EXERCISE%' OR UPPER(MA.Action_Name) LIKE '%AEROBIC%'
                    OR UPPER(MA.Action_Name) LIKE '%STRENGTHENING%' OR UPPER(MA.Action_Name) LIKE '%MINUTES A WEEK%'
                    OR UPPER(MA.Action_Name) LIKE '%PHYSICAL ACTIVITY%'
                    THEN 'Exercise'
                WHEN UPPER(MA.Action_Name) LIKE '%SUGAR%' OR UPPER(MA.Action_Name) LIKE '%FRUIT%'
                    OR UPPER(MA.Action_Name) LIKE '%VEGETABLE%' OR UPPER(MA.Action_Name) LIKE '%EATING%'
                    OR UPPER(MA.Action_Name) LIKE '%FOOD%' OR UPPER(MA.Action_Name) LIKE '%NUTRITION%'
                    THEN 'Nutrition'
                WHEN UPPER(MA.Action_Name) LIKE '%WEIGHT%' OR UPPER(MA.Action_Name) LIKE '%BMI%'
                    THEN 'Weight Management'
                WHEN UPPER(MA.Action_Name) LIKE '%TOBACCO%' OR UPPER(MA.Action_Name) LIKE '%SMOK%'
                    THEN 'Tobacco Cessation'
                WHEN UPPER(MA.Action_Name) LIKE '%DEPRESS%' OR UPPER(MA.Action_Name) LIKE '%STRESS%'
                    OR UPPER(MA.Action_Name) LIKE '%ANXIETY%' OR UPPER(MA.Action_Name) LIKE '%MENTAL%'
                    OR UPPER(MA.Action_Name) LIKE '%ALCOHOL%'
                    THEN 'Mental/Behavioral Health'
                WHEN UPPER(MA.Action_Name) LIKE '%SLEEP%'
                    THEN 'Stress Management'
                WHEN UPPER(MA.Action_Name) LIKE '%BLOOD PRESSURE%' OR UPPER(MA.Action_Name) LIKE '%CHOLESTEROL%'
                    OR UPPER(MA.Action_Name) LIKE '%A1C%' OR UPPER(MA.Action_Name) LIKE '%BLOOD SUGAR%'
                    OR UPPER(MA.Action_Name) LIKE '%DIABETES%' OR UPPER(MA.Action_Name) LIKE '%INHALER%'
                    OR UPPER(MA.Action_Name) LIKE '%STATIN%' OR UPPER(MA.Action_Name) LIKE '%BETA-BLOCKER%'
                    OR UPPER(MA.Action_Name) LIKE '%MEDICATION%' OR UPPER(MA.Action_Name) LIKE '%SPIROMETRY%'
                    OR UPPER(MA.Action_Name) LIKE '%RETINAL%' OR UPPER(MA.Action_Name) LIKE '%FOOT EXAM%'
                    OR UPPER(MA.Action_Name) LIKE '%MICROALBUMIN%' OR UPPER(MA.Action_Name) LIKE '%CAD%'
                    OR UPPER(MA.Action_Name) LIKE '%HEART FAILURE%'
                    THEN 'Condition Management'
                ELSE 'Other'
            END
        ELSE 'Other'
    END AS GOAL_DOMAIN,
    CASE
        WHEN MA.ActionStatus_ID = 1 THEN 'Not Started'
        WHEN MA.ActionStatus_ID = 2 THEN 'In Progress'
        WHEN MA.ActionStatus_ID IN (3, 5) THEN 'Completed'
        WHEN MA.ActionStatus_ID = 4 THEN 'Withdrawn'
    END AS GOAL_STATUS,
    CASE WHEN MA.Action_Name LIKE '%:%' THEN TRIM(SUBSTR(MA.Action_Name, POSITION(':' IN MA.Action_Name)+1)) ELSE MA.Action_Name END AS GOAL_DESCRIPTION,
    MA.Action_Name AS RAW_ACTION_NAME,
    MA.MemberAction_ID, MA.ActionType_ID, MA.ActionStatus_ID, MA.FocusArea_ID,
    MA.Action_Date AS GOAL_SET_DATE, MA.Close_Date AS GOAL_CLOSE_DATE
FROM SCP.AH_MEMBER_ACTION MA
WHERE MA.ActionType_ID IN (2,3) AND MA.ActionStatus_ID IN (1, 2, 3, 4, 5)
  AND MA.Action_Name NOT LIKE 'Survey:%' AND MA.Action_Name NOT LIKE 'RED:%' AND MA.Action_Name NOT LIKE 'RED/%';

-- =========================================================================
-- SECTION 4: BRIDGE GOALS TO GUID (current enrollment)
-- =========================================================================
DROP TABLE IF EXISTS COACHING_GOALS_NUMBERED;
CREATE LOCAL TEMP TABLE COACHING_GOALS_NUMBERED ON COMMIT PRESERVE ROWS AS
SELECT G.*, CE.CURRENTGUID, CE.GUID,
    ROW_NUMBER() OVER (PARTITION BY G.ACCOUNT ORDER BY G.GOAL_SET_DATE ASC, G.RAW_ACTION_NAME ASC) AS GOAL_NUMBER
FROM COACHING_GOALS G
JOIN (
    SELECT ACCOUNT, CURRENTGUID, GUID,
        ROW_NUMBER() OVER (PARTITION BY ACCOUNT ORDER BY BIENDDATE DESC) AS RN
    FROM BI_REPORTING.COACHING_ENROLLMENT_MODEL
    WHERE BIENDDATE >= CURRENT_DATE
) CE ON G.ACCOUNT = CE.ACCOUNT AND CE.RN = 1;


-- =========================================================================
-- OUTPUTS
-- =========================================================================
-- Column definitions:
--   MEMBERS     = distinct members (GUIDs) with at least one call in that category
--   PCT_MEMBERS = members / total distinct members across ALL calls (not sum of groups)
--   CALLS       = number of call-days in that category
--   PCT_OF_CALLS= calls / total calls (sums to 100% across all categories)
-- =========================================================================

-- OUTPUT 1: TOPICS SUMMARY
-- Note: PCT_MEMBERS uses total distinct members as denominator (not sum of per-group).
-- A member can have calls in multiple topic categories, so per-group counts don't sum to total.
SELECT REPORT_TOPIC AS CALL_TOPIC,
    COUNT(DISTINCT CT.GUID) AS MEMBERS,
    ROUND(COUNT(DISTINCT CT.GUID)*100.0 / NULLIFZERO(TOTALS.TOTAL_MEMBERS), 2) AS PCT_MEMBERS,
    COUNT(*) AS CALLS,
    ROUND(COUNT(*)*100.0/NULLIFZERO(TOTALS.TOTAL_CALLS),2) AS PCT_OF_CALLS
FROM COACHING_TOPICS CT
CROSS JOIN (
    SELECT COUNT(DISTINCT GUID) AS TOTAL_MEMBERS, COUNT(*) AS TOTAL_CALLS
    FROM COACHING_TOPICS
) TOTALS
GROUP BY 1, TOTALS.TOTAL_MEMBERS, TOTALS.TOTAL_CALLS
ORDER BY 4 DESC;

-- OUTPUT 1B: TIER BREAKDOWN
SELECT TOPIC_SOURCE, COUNT(*) AS CALLS,
    ROUND(COUNT(*)*100.0/NULLIFZERO(SUM(COUNT(*)) OVER()),2) AS PCT
FROM COACHING_TOPICS GROUP BY 1 ORDER BY 2 DESC;

-- OUTPUT 2: TOBACCO (simplified - no correlated subqueries)
SELECT
    T.TOBACCO_DISCUSSED,
    T.MEMBERS,
    ROUND(T.MEMBERS * 100.0 / NULLIFZERO(TOTALS.TOTAL_MEMBERS), 2) AS PCT_MEMBERS,
    T.CALLS,
    ROUND(T.CALLS * 100.0 / NULLIFZERO(TOTALS.TOTAL_CALLS), 2) AS PCT_OF_CALLS
FROM (
    SELECT 'Yes' AS TOBACCO_DISCUSSED,
        COUNT(DISTINCT GUID) AS MEMBERS,
        COUNT(*) AS CALLS
    FROM COACHING_TOPICS
    WHERE UPPER(RAW_TOPIC) = 'TOBACCO' OR UPPER(CALL_TYPE) = 'TOBACCO'
    UNION ALL
    SELECT 'No',
        COUNT(DISTINCT GUID),
        COUNT(*)
    FROM COACHING_TOPICS
    WHERE GUID NOT IN (SELECT GUID FROM COACHING_TOBACCO WHERE GUID IS NOT NULL)
    UNION ALL
    SELECT 'Grand Total',
        COUNT(DISTINCT GUID),
        COUNT(*)
    FROM COACHING_TOPICS
) T
CROSS JOIN (
    SELECT COUNT(DISTINCT GUID) AS TOTAL_MEMBERS, COUNT(*) AS TOTAL_CALLS
    FROM COACHING_TOPICS
) TOTALS
ORDER BY CASE T.TOBACCO_DISCUSSED WHEN 'Yes' THEN 1 WHEN 'No' THEN 2 ELSE 3 END;

-- OUTPUT 3: GOAL TYPE & DOMAIN
SELECT GOAL_DOMAIN,
    SUM(CASE WHEN GOAL_TYPE='Coach-Created' THEN 1 ELSE 0 END) AS COACH_CREATED,
    ROUND(SUM(CASE WHEN GOAL_TYPE='Coach-Created' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS CC_PCT,
    SUM(CASE WHEN GOAL_TYPE='System-Recommended' THEN 1 ELSE 0 END) AS SYS_RECOMMENDED,
    ROUND(SUM(CASE WHEN GOAL_TYPE='System-Recommended' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS SR_PCT
FROM COACHING_GOALS_NUMBERED GROUP BY 1
ORDER BY CASE GOAL_DOMAIN WHEN 'Gaps in Care' THEN 1 WHEN 'Exercise' THEN 2 WHEN 'Nutrition' THEN 3 WHEN 'Weight Management' THEN 4 WHEN 'Tobacco Cessation' THEN 5 WHEN 'Mental/Behavioral Health' THEN 6 WHEN 'Stress Management' THEN 7 WHEN 'Condition Management' THEN 8 WHEN 'Financial' THEN 9 WHEN 'Social' THEN 10 WHEN 'Spiritual' THEN 11 ELSE 12 END;

-- OUTPUT 4: GOAL DOMAIN & STATUS
SELECT GOAL_DOMAIN,
    SUM(CASE WHEN GOAL_STATUS='Not Started' THEN 1 ELSE 0 END) AS NOT_STARTED,
    ROUND(SUM(CASE WHEN GOAL_STATUS='Not Started' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS NS_PCT,
    SUM(CASE WHEN GOAL_STATUS='In Progress' THEN 1 ELSE 0 END) AS IN_PROGRESS,
    ROUND(SUM(CASE WHEN GOAL_STATUS='In Progress' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS IP_PCT,
    SUM(CASE WHEN GOAL_STATUS='Completed' THEN 1 ELSE 0 END) AS COMPLETED,
    ROUND(SUM(CASE WHEN GOAL_STATUS='Completed' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS COMP_PCT,
    SUM(CASE WHEN GOAL_STATUS='Withdrawn' THEN 1 ELSE 0 END) AS WITHDRAWN,
    ROUND(SUM(CASE WHEN GOAL_STATUS='Withdrawn' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS WD_PCT
FROM COACHING_GOALS_NUMBERED GROUP BY 1
ORDER BY CASE GOAL_DOMAIN WHEN 'Gaps in Care' THEN 1 WHEN 'Exercise' THEN 2 WHEN 'Nutrition' THEN 3 WHEN 'Weight Management' THEN 4 WHEN 'Tobacco Cessation' THEN 5 WHEN 'Mental/Behavioral Health' THEN 6 WHEN 'Stress Management' THEN 7 WHEN 'Condition Management' THEN 8 WHEN 'Financial' THEN 9 WHEN 'Social' THEN 10 WHEN 'Spiritual' THEN 11 ELSE 12 END;

-- OUTPUT 5: GOAL NUMBER & STATUS
SELECT GOAL_NUMBER,
    SUM(CASE WHEN GOAL_STATUS='Not Started' THEN 1 ELSE 0 END) AS NOT_STARTED,
    ROUND(SUM(CASE WHEN GOAL_STATUS='Not Started' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS NS_PCT,
    SUM(CASE WHEN GOAL_STATUS='In Progress' THEN 1 ELSE 0 END) AS IN_PROGRESS,
    ROUND(SUM(CASE WHEN GOAL_STATUS='In Progress' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS IP_PCT,
    SUM(CASE WHEN GOAL_STATUS='Completed' THEN 1 ELSE 0 END) AS COMPLETED,
    ROUND(SUM(CASE WHEN GOAL_STATUS='Completed' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS COMP_PCT,
    SUM(CASE WHEN GOAL_STATUS='Withdrawn' THEN 1 ELSE 0 END) AS WITHDRAWN,
    ROUND(SUM(CASE WHEN GOAL_STATUS='Withdrawn' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS WD_PCT
FROM COACHING_GOALS_NUMBERED WHERE GOAL_NUMBER <= 6 GROUP BY 1 ORDER BY 1;

-- OUTPUT 6: GOAL TYPE & STATUS
SELECT GOAL_TYPE,
    SUM(CASE WHEN GOAL_STATUS='Not Started' THEN 1 ELSE 0 END) AS NOT_STARTED,
    ROUND(SUM(CASE WHEN GOAL_STATUS='Not Started' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS NS_PCT,
    SUM(CASE WHEN GOAL_STATUS='In Progress' THEN 1 ELSE 0 END) AS IN_PROGRESS,
    ROUND(SUM(CASE WHEN GOAL_STATUS='In Progress' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS IP_PCT,
    SUM(CASE WHEN GOAL_STATUS='Completed' THEN 1 ELSE 0 END) AS COMPLETED,
    ROUND(SUM(CASE WHEN GOAL_STATUS='Completed' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS COMP_PCT,
    SUM(CASE WHEN GOAL_STATUS='Withdrawn' THEN 1 ELSE 0 END) AS WITHDRAWN,
    ROUND(SUM(CASE WHEN GOAL_STATUS='Withdrawn' THEN 1 ELSE 0 END)*100.0/NULLIFZERO(COUNT(*)),2) AS WD_PCT
FROM COACHING_GOALS_NUMBERED GROUP BY 1 ORDER BY 1;

-- VALIDATION: TIER COUNTS
SELECT 'ALL CALLS' AS SECTION, COUNT(*) AS ROWS, COUNT(DISTINCT GUID) AS MEMBERS FROM COACHING_TOPICS
UNION ALL SELECT 'Tier 1: Date Match', COUNT(*), COUNT(DISTINCT GUID) FROM COACHING_TOPICS WHERE TOPIC_SOURCE LIKE 'Tier 1:%'
UNION ALL SELECT 'Tier 1.5: Text', COUNT(*), COUNT(DISTINCT GUID) FROM COACHING_TOPICS WHERE TOPIC_SOURCE = 'Tier 1.5: Text Inference'
UNION ALL SELECT 'Tier 2: Call Type', COUNT(*), COUNT(DISTINCT GUID) FROM COACHING_TOPICS WHERE TOPIC_SOURCE = 'Tier 2: Call Type'
UNION ALL SELECT 'Tier 3: Prior', COUNT(*), COUNT(DISTINCT GUID) FROM COACHING_TOPICS WHERE TOPIC_SOURCE = 'Tier 3: Prior Topic'
UNION ALL SELECT 'Tier 4: General', COUNT(*), COUNT(DISTINCT GUID) FROM COACHING_TOPICS WHERE TOPIC_SOURCE = 'Tier 4: None'
UNION ALL SELECT 'TOBACCO', COUNT(*), COUNT(DISTINCT GUID) FROM COACHING_TOBACCO
UNION ALL SELECT 'GOALS', COUNT(*), COUNT(DISTINCT CURRENTGUID) FROM COACHING_GOALS_NUMBERED;
