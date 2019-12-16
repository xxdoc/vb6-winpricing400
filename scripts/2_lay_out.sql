
CREATE TABLE LAY_OUT
(
   LAY_OUT_ID             ID_TYPE NOT NULL,
   LAY_OUT_NO             CODE_TYPE NOT NULL,
   LAY_OUT_NAME           CODE_TYPE NOT NULL,
   LOCATION_ID            ID_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);