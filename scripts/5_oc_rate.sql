
ALTER TABLE OC_RATE ADD CONSTRAINT OC_RATE_ID_PK PRIMARY KEY (OC_RATE_ID);
ALTER TABLE OC_RATE ADD CONSTRAINT OC_SOC_FEATURE_ID_FK FOREIGN KEY (SOC_FEATURE_ID) REFERENCES SOC_FEATURE;
