
ALTER TABLE GROUP_RIGHT ADD CONSTRAINT GROUP_RIGHT_PK PRIMARY KEY (GROUP_RIGHT_ID);
ALTER TABLE GROUP_RIGHT ADD CONSTRAINT GROUP_RIGHT_UK UNIQUE (GROUP_ID, RIGHT_ID);
ALTER TABLE GROUP_RIGHT ADD CONSTRAINT RIGHT_ID_FK FOREIGN KEY (RIGHT_ID) REFERENCES RIGHT_ITEM;
ALTER TABLE GROUP_RIGHT ADD CONSTRAINT GROUP_ID_FK FOREIGN KEY (GROUP_ID) REFERENCES USER_GROUP;