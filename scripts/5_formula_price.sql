
ALTER TABLE FORMULA_PRICE ADD CONSTRAINT FORMULA_PRICE_ID_PK PRIMARY KEY (FORMULA_PRICE_ID);
ALTER TABLE FORMULA_PRICE ADD CONSTRAINT FORMULA_PRICE_FORMULA_ID_FK FOREIGN KEY (FORMULA_ID) REFERENCES FORMULA;
