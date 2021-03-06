
CREATE TABLE SUPPLIER
(
   SUPPLIER_ID           ID_TYPE NOT NULL,
   SUPPLIER_CODE         CODE_TYPE NOT NULL,
   SUPPLIER_GRADE        ID_TYPE,
   SUPPLIER_TYPE         ID_TYPE,
   CREDIT                LEN_TYPE,
   TAX_ID                TAXID_TYPE,
   BIRTH_DATE            DATE_TYPE,
   EMAIL                 EMAIL_TYPE,
   WEBSITE               WEB_TYPE,
   PASSWORD1             PASSWORD_TYPE,
   SUPPLIER_STATUS       ID_TYPE,
   BUSINESS_DESC         DESC_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
