
CREATE TABLE UC_RATE
(
   UC_RATE_ID            ID_TYPE NOT NULL,
   SOC_FEATURE_ID        ID_TYPE NOT NULL,
   RATE_AMOUNT           MONEY_TYPE NOT NULL,
   STEP_ID               ID_TYPE,
   TIER_ID               ID_TYPE,
   STEP_VOLUME_ID        ID_TYPE,
   TIER_VOLUME_ID        ID_TYPE,
   STPTIER_VOL_ID        ID_TYPE,
   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);