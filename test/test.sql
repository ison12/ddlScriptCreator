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
