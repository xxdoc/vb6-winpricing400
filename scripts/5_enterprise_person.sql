
ALTER TABLE ENTERPRISE_PERSON ADD CONSTRAINT ENTERPRISE_PERSON_ID_PK PRIMARY KEY (ENTERPRISE_PERSON_ID);
ALTER TABLE ENTERPRISE_PERSON ADD CONSTRAINT ENTERPRISE_PERSON_ENP_ID_FK2 FOREIGN KEY (ENTERPRISE_ID) REFERENCES ENTERPRISE;
ALTER TABLE ENTERPRISE_PERSON ADD CONSTRAINT ENTERPRISE_PERSON_NM_ID_FK3 FOREIGN KEY (NAME_ID) REFERENCES NAME;