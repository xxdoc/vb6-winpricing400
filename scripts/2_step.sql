
CREATE TABLE STEP
(
   STEP_ID               ID_TYPE NOT NULL,
   SOC_FEATURE_ID        ID_TYPE NOT NULL,
   FROM_QUANTITY         MONEY_TYPE NOT NULL,
   TO_QUANTITY           MONEY_TYPE NOT NULL,
   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
