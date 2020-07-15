
-- -----------------------------------------------------------------------------
-- PK：PK_TEST_SUB
-- -----------------------------------------------------------------------------
ALTER TABLE "test_sub" ADD CONSTRAINT pk_test_sub PRIMARY KEY (
      "idsub"
);

-- -----------------------------------------------------------------------------
-- UK：UK_TEST_SUB_1
-- -----------------------------------------------------------------------------
ALTER TABLE "test_sub" ADD CONSTRAINT uk_test_sub_1 UNIQUE (
      "stringsub1"
);
-- -----------------------------------------------------------------------------
-- UK：UK_TEST_SUB_2
-- -----------------------------------------------------------------------------
ALTER TABLE "test_sub" ADD CONSTRAINT uk_test_sub_2 UNIQUE (
      "stringsub2"
);

-- -----------------------------------------------------------------------------
-- FK：FK_TEST_SUB_STRINGSUB1
-- -----------------------------------------------------------------------------
ALTER TABLE "test_sub" ADD CONSTRAINT FOREIGN KEY fk_test_sub_stringsub1 (
      "stringsub1"
) REFERENCES test (
      "string1"
);

-- -----------------------------------------------------------------------------
-- Index：IDX_TEST_SUB_01
-- -----------------------------------------------------------------------------
CREATE INDEX idx_test_sub_01 ON "test_sub" (
      `stringsub1`
    , `stringsub2`
);
