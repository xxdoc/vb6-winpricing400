
ALTER TABLE STEP_VOLUME ADD CONSTRAINT STEP_VOLUME_ID_PK PRIMARY KEY (STEP_VOLUME_ID);
ALTER TABLE STEP_VOLUME ADD CONSTRAINT STEP_VOLUME_SOC_FEATURE_ID_FK FOREIGN KEY (SOC_FEATURE_ID) REFERENCES SOC_FEATURE;
