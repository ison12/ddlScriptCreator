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
