
CREATE TABLE SOC
(
   SOC_ID                ID_TYPE NOT NULL,
   SOC_CODE              CODE_TYPE NOT NULL,
   SOC_DESC              DESC_TYPE,
   SOC_STATUS            ID_TYPE,
   SOC_LEVEL             FLAG_TYPE,
   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
