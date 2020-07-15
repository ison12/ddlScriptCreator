-- -----------------------------------------------------------------------------
-- テーブル：test
-- 作成者　：自動生成
-- -----------------------------------------------------------------------------
DROP TABLE IF EXISTS "test";
CREATE TABLE "test" (
      "id"                      bigserial          NOT NULL                    
    , "string1"                 varchar(100)       NOT NULL                    
    , "string2"                 varchar(100)       NOT NULL                    
);

COMMENT ON TABLE "test" IS 'テスト';
COMMENT ON COLUMN "test"."id" IS 'ID';
COMMENT ON COLUMN "test"."string1" IS '文字列1';
COMMENT ON COLUMN "test"."string2" IS '文字列2';
-- -----------------------------------------------------------------------------
-- テーブル：test_sub
-- 作成者　：自動生成
-- -----------------------------------------------------------------------------
DROP TABLE IF EXISTS "test_sub";
CREATE TABLE "test_sub" (
      "idsub"                   bigserial          NOT NULL                    
    , "stringsub1"              varchar(100)       NOT NULL                    
    , "stringsub2"              varchar(100)       NOT NULL                    
);

COMMENT ON TABLE "test_sub" IS 'テストサブ';
COMMENT ON COLUMN "test_sub"."idsub" IS 'ID';
COMMENT ON COLUMN "test_sub"."stringsub1" IS '文字列1';
COMMENT ON COLUMN "test_sub"."stringsub2" IS '文字列2';

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

