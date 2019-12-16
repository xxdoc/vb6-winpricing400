CREATE TABLE DO_ITEM
(
   DO_ITEM_ID            ID_TYPE NOT NULL,
   DO_ID                 ID_TYPE NOT NULL,
   PART_ITEM_ID          ID_TYPE NOT NULL,
   LOCATION_ID           ID_TYPE NOT NULL,
   ITEM_AMOUNT           MONEY_TYPE,
   TOTAL_WEIGHT          MONEY_TYPE,
   AVG_WEIGHT            MONEY_TYPE,
   TOTAL_PRICE           MONEY_TYPE,
   AVG_PRICE             MONEY_TYPE,
   GUI_ID                ID_TYPE,
   PIG_STATUS            ID_TYPE,
   LINK_ID               ID_TYPE,
   DEPOSIT_AMOUNT        MONEY_TYPE,
   DISCOUNT_AMOUNT       MONEY_TYPE,
   FEATURE_ID            ID_TYPE,
   SOC_ID                ID_TYPE,
   ITEM_DESC             DESC_TYPE,
   DISPLAY_ID            ID_TYPE,
   PO_ID                 ID_TYPE,
   PO_NO                 CODE_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
