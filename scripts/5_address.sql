
ALTER TABLE ADDRESS ADD CONSTRAINT ADDRESS_ID_PK PRIMARY KEY (ADDRESS_ID);
ALTER TABLE ADDRESS ADD CONSTRAINT ADDRESS_COUNTRY_ID_FK FOREIGN KEY (COUNTRY_ID) REFERENCES COUNTRY;
