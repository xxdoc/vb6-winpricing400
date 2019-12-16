
CREATE TABLE COUNTRY
(
   COUNTRY_ID         ID_TYPE NOT NULL,
   COUNTRY_NO         CODE_TYPE NOT NULL,
   COUNTRY_NAME       COUNTRY_TYPE NOT NULL,
   CONTINENT_ID       ID_TYPE,

   CREATE_DATE        DATE_TYPE NOT NULL,
   CREATE_BY          ID_TYPE NOT NULL,
   MODIFY_DATE        DATE_TYPE NOT NULL,
   MODIFY_BY          ID_TYPE NOT NULL
);
