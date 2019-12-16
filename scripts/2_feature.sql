
CREATE TABLE FEATURE
(
   FEATURE_ID            ID_TYPE NOT NULL,
   FEATURE_CODE          CODE_TYPE NOT NULL,
   FEATURE_TYPE          ID_TYPE NOT NULL,
   FEATURE_DESC          DESC_TYPE,
   FEATURE_STATUS        FLAG_TYPE,
   FEATURE_LVEL          ID_TYPE,
   FEATURE_UNIT          ID_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
