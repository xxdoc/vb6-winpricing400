CREATE TABLE PART_LOCATION
(
   PART_LOCATION_ID      ID_TYPE NOT NULL,
   PART_ITEM_ID          ID_TYPE,
   LOCATION_ID           ID_TYPE,
   AVG_PRICE             MONEY_TYPE,
   LAST_PRICE            MONEY_TYPE,
   MINIMUM_ALLOW         MONEY_TYPE,
   CURRENT_AMOUNT        MONEY_TYPE,
   TEMP_AVG_PRICE        MONEY_TYPE,
   TEMP_LAST_PRICE       MONEY_TYPE,
   TEMP_CURRENT_AMOUNT   MONEY_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
