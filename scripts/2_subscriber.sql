
CREATE TABLE SUBSCRIBER
(
   SUBSCRIBER_ID         ID_TYPE NOT NULL,
   SUBSCRIBER_NO         CODE_TYPE NOT NULL,
   ACCOUNT_ID            ID_TYPE NOT NULL,
   SUBSCRIBER_STATUS     FLAG_TYPE,
   DUMMY_FLAG            FLAG_TYPE NOT NULL,
   SUBSCRIBER_DESC       DESC_TYPE,
   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
