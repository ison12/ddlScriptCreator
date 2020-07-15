
-- -----------------------------------------------------------------------------
-- PK：PK_TEST
-- -----------------------------------------------------------------------------
ALTER TABLE "test" ADD CONSTRAINT pk_test PRIMARY KEY (
      "id"
);

-- -----------------------------------------------------------------------------
-- UK：UK_TEST_1
-- -----------------------------------------------------------------------------
ALTER TABLE "test" ADD CONSTRAINT uk_test_1 UNIQUE (
      "string1"
    , "string2"
);


-- -----------------------------------------------------------------------------
-- Index：IDX_TEST_01
-- -----------------------------------------------------------------------------
CREATE INDEX idx_test_01 ON "test" (
      `string1`
    , `string2`
);

