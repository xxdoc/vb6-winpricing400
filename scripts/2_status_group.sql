CREATE TABLE STATUS_GROUP
(
   STATUS_GROUP_ID       ID_TYPE NOT NULL,
   STATUS_GROUP_NO       CODE_TYPE,
   STATUS_GROUP_NAME     NAME_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
