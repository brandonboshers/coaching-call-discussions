-- =========================================================================
-- DDL: Create persistent tables in Carefirst_Sandbox
-- Run once to create the tables. monthly_refresh.sql handles ongoing loads.
-- =========================================================================

-- COACHING_CALL_TOPICS: One row per account per call day with assigned topic
CREATE TABLE IF NOT EXISTS Carefirst_Sandbox.COACHING_CALL_TOPICS (
    REPORT_MONTH    VARCHAR(6),
    ACCOUNT         VARCHAR(50),
    GUID            VARCHAR(50),
    CUSTOMERID      VARCHAR(50),
    CALL_DATE       DATE,
    CALL_TYPE       VARCHAR(200),
    RAW_TOPIC       VARCHAR(500),
    REPORT_TOPIC    VARCHAR(100),
    TOPIC_SOURCE    VARCHAR(50),
    PROGRAM_TYPE    VARCHAR(10),
    REFRESH_DATE    TIMESTAMP DEFAULT SYSDATE
);

-- COACHING_CALL_TOBACCO: Members flagged for tobacco discussion
CREATE TABLE IF NOT EXISTS Carefirst_Sandbox.COACHING_CALL_TOBACCO (
    REPORT_MONTH    VARCHAR(6),
    ACCOUNT         VARCHAR(50),
    GUID            VARCHAR(50),
    TOBACCO_DISCUSSED VARCHAR(10),
    REFRESH_DATE    TIMESTAMP DEFAULT SYSDATE
);

-- COACHING_CALL_GOALS: One row per goal per member with domain, type, status
CREATE TABLE IF NOT EXISTS Carefirst_Sandbox.COACHING_CALL_GOALS (
    REPORT_MONTH    VARCHAR(6),
    ACCOUNT         VARCHAR(50),
    GOAL_TYPE       VARCHAR(50),
    GOAL_DOMAIN     VARCHAR(50),
    GOAL_STATUS     VARCHAR(20),
    GOAL_DESCRIPTION VARCHAR(2000),
    RAW_ACTION_NAME VARCHAR(2000),
    MEMBERACTION_ID INT,
    ACTIONTYPE_ID   INT,
    ACTIONSTATUS_ID INT,
    FOCUSAREA_ID    INT,
    GOAL_SET_DATE   DATE,
    GOAL_CLOSE_DATE DATE,
    CURRENTGUID     VARCHAR(50),
    GUID            VARCHAR(50),
    GOAL_NUMBER     INT,
    REFRESH_DATE    TIMESTAMP DEFAULT SYSDATE
);
