
ALTER TABLE EMPLOYEE_NAME ADD CONSTRAINT EMPLOYEE_NAME_ID_PK PRIMARY KEY (EMPLOYEE_NAME_ID);
ALTER TABLE EMPLOYEE_NAME ADD CONSTRAINT EMPLOYEE_EMP_ID_FK2 FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE;
ALTER TABLE EMPLOYEE_NAME ADD CONSTRAINT EMPLOYEE_NAME_ID_FK3 FOREIGN KEY (NAME_ID) REFERENCES NAME;