
ALTER TABLE CUSTOMER ADD CONSTRAINT CUSTOMER_ID_PK PRIMARY KEY (CUSTOMER_ID);
ALTER TABLE CUSTOMER ADD CONSTRAINT CUSTOMER_GRADE_ID_FK FOREIGN KEY (CUSTOMER_GRADE) REFERENCES CUSTOMER_GRADE;
ALTER TABLE CUSTOMER ADD CONSTRAINT CUSTOMER_TYPE_ID_FK FOREIGN KEY (CUSTOMER_TYPE) REFERENCES CUSTOMER_TYPE;
ALTER TABLE CUSTOMER ADD CONSTRAINT CUSTOMER_BUSINESS_ID_FK FOREIGN KEY (BUSINESS_TYPE) REFERENCES BUSINESS_TYPE;
