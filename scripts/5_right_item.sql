
ALTER TABLE RIGHT_ITEM ADD CONSTRAINT RIGHT_ITEM_PK PRIMARY KEY (RIGHT_ID);
ALTER TABLE RIGHT_ITEM ADD CONSTRAINT RIGHT_ITEM_NAME_UK UNIQUE (RIGHT_ITEM_NAME);
