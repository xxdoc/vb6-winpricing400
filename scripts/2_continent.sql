
CREATE TABLE CONTINENT
(
   CONTINENT_ID       ID_TYPE NOT NULL,
   CONTINENT_NAME     CODE_TYPE NOT NULL UNIQUE,
   CREATE_DATE        DATE_TYPE NOT NULL,
   CREATE_BY          ID_TYPE NOT NULL,
   MODIFY_DATE        DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
