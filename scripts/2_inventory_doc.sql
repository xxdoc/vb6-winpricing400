CREATE TABLE INVENTORY_DOC
(
   INVENTORY_DOC_ID      ID_TYPE NOT NULL,
   DOCUMENT_NO           CODE_TYPE NOT NULL,
   DOCUMENT_DATE         DATE_TYPE NOT NULL,
   DOCUMENT_DESC         DESC_TYPE,
   BILL_NO               CODE_TYPE,
   DO_NO                 CODE_TYPE,
   TRUCK_NO              CODE_TYPE,
   SUPPLIER_ID           ID_TYPE,
   DELIVERY_ID           ID_TYPE,
   DELIVERY_FEE          MONEY_TYPE,
   SENDER_NAME           NAME_TYPE,
   RECEIVE_NAME          NAME_TYPE,
   DOCUMENT_TYPE         ID_TYPE NOT NULL,
   EMP_ID                ID_TYPE,
   COMMIT_FLAG           FLAG_TYPE,
   SALE_FLAG             FLAG_TYPE,
   REASON_ID             ID_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);