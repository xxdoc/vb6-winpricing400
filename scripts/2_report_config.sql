CREATE TABLE REPORT_CONFIG
(
   REPORT_CONFIG_ID      ID_TYPE NOT NULL,
   REPORT_KEY            CODE_TYPE,
   PAPER_SIZE            MONEY_TYPE,
   PAPER_WIDTH           MONEY_TYPE,
   PAPER_HEIGHT          MONEY_TYPE,
   ORIENTATION           MONEY_TYPE,
   MARGIN_BOTTOM         MONEY_TYPE,
   MARGIN_FOOTER         MONEY_TYPE,
   MARGIN_HEADER         MONEY_TYPE,
   MARGIN_LEFT           MONEY_TYPE,
   MARGIN_RIGHT          MONEY_TYPE,
   MARGIN_TOP            MONEY_TYPE,
   FONT_NAME             NAME_TYPE,
   FONT_SIZE             MONEY_TYPE,
   HEAD_OFFSET           MONEY_TYPE,
   DUMMY_OFFSET          MONEY_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
