
ALTER TABLE STPTIER_VOL ADD CONSTRAINT STPTIER_VOLUME_ID_PK PRIMARY KEY (STPTIER_VOL_ID);
ALTER TABLE STPTIER_VOL ADD CONSTRAINT STPTIER_VOL_SOC_FEATURE_ID_FK FOREIGN KEY (SOC_FEATURE_ID) REFERENCES SOC_FEATURE;
