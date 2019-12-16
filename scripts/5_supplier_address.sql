
ALTER TABLE SUPPLIER_ADDRESS ADD CONSTRAINT SUPPLIER_ADDRESS_ID_PK PRIMARY KEY (SUPPLIER_ADDRESS_ID);
ALTER TABLE SUPPLIER_ADDRESS ADD CONSTRAINT SUPPLIER_ADDR_SUPPLIER_ID_FK FOREIGN KEY (SUPPLIER_ID) REFERENCES SUPPLIER;
ALTER TABLE SUPPLIER_ADDRESS ADD CONSTRAINT SUPPLIER_ADDRESS_ID_FK FOREIGN KEY (ADDRESS_ID) REFERENCES ADDRESS;
ALTER TABLE SUPPLIER_ADDRESS ADD CONSTRAINT SUPPLIER_ADDRESS_TYPE_FK FOREIGN KEY (ADDRESS_TYPE) REFERENCES ADDRESS_TYPE;
