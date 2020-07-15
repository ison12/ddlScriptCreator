VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImplTableDefineCreatorPostgres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITableDefineCreator

Private Function ITableDefineCreator_writeForTable(t As Table, ByVal filePath As String, isAppend As Boolean) As Variant

    Dim i       As Long
    Dim notNull As String
    Dim defVal  As String
    
    Dim c As Column
    Dim var As Variant
    
    Dim writer As New FileWriter
    writer.init filePath, "UTF-8", vbCr & vbLf, isAppend
    
    ' ----------------------------------------------------
    ' カラム定義
    ' ----------------------------------------------------
    writer.writeText "-- -----------------------------------------------------------------------------", True
    writer.writeText "-- テーブル：" & t.tableName, True
    writer.writeText "-- 作成者　：自動生成", True
    writer.writeText "-- -----------------------------------------------------------------------------", True
    
    writer.writeText "DROP TABLE IF EXISTS """ & t.tableName & """;", True
    writer.writeText "CREATE TABLE """ & t.tableName & """ (", True
    
    i = 0
    For Each c In t.columns.col
    
        If c.isNotNull = True Then
            notNull = "NOT NULL"
        Else
            notNull = Space(8)
        End If
        
        If c.defaultValue <> "" Then
            defVal = padRight(" DEFAULT '" & c.defaultValue & "'", 20, " ")
        Else
            defVal = Space(20)
        End If
        
        If i <> 0 Then
            writer.writeText "    , " & padRight("""" & c.columnName & """", 25, " ") & " " & padRight(c.dataType, 18, " ") & " " & notNull & defVal, True
        Else
            writer.writeText "      " & padRight("""" & c.columnName & """", 25, " ") & " " & padRight(c.dataType, 18, " ") & " " & notNull & defVal, True
        End If
        
        i = i + 1
    Next
    
    writer.writeText ");", True
    writer.writeText "", True
    
    writer.writeText "COMMENT ON TABLE """ & t.tableName & """ IS '" & t.tableComment & "';", True
    
    i = 0
    For Each c In t.columns.col
    
        writer.writeText "COMMENT ON COLUMN """ & t.tableName & """.""" & c.columnName & """ IS '" & c.columnComment & "';", True
        i = i + 1
    Next
    
    writer.makeUtf8WithoutBom
    
End Function

Private Function ITableDefineCreator_writeForConstraints(t As Table, ByVal filePath As String, isAppend As Boolean) As Variant

    Dim i       As Long
    Dim notNull As String
    Dim defVal  As String
    
    Dim c As Column
    Dim var As Variant
    
    Dim writer As New FileWriter
    writer.init filePath, "UTF-8", vbCr & vbLf, isAppend
    
    ' ----------------------------------------------------
    ' PK制約
    ' ----------------------------------------------------
    If t.constraintsPk.columns.count > 0 Then
        
        writer.writeText "", True
        writer.writeText "-- -----------------------------------------------------------------------------", True
        writer.writeText "-- PK：" & t.constraintsPk.constraintName, True
        writer.writeText "-- -----------------------------------------------------------------------------", True
    
        writer.writeText "ALTER TABLE """ & t.tableName & """ ADD CONSTRAINT " & LCase(t.constraintsPk.constraintName) & " PRIMARY KEY (", True
    End If
    
    i = 0
    For Each c In t.constraintsPk.columns.col
    
        If i <> 0 Then
            writer.writeText "    , " & """" & c.columnName & """", True
        Else
            writer.writeText "      " & """" & c.columnName & """", True
        End If
        
        i = i + 1
    Next
    
    If t.constraintsPk.columns.count > 0 Then
        writer.writeText ");", True
    End If
    
    writer.writeText "", True
        
    ' ----------------------------------------------------
    ' UK制約
    ' ----------------------------------------------------
    Dim constUk As ConstraintUnique
    For Each constUk In t.constraintsUks.col

        If constUk.columns.count > 0 Then
            writer.writeText "-- -----------------------------------------------------------------------------", True
            writer.writeText "-- UK：" & constUk.constraintName, True
            writer.writeText "-- -----------------------------------------------------------------------------", True
            writer.writeText "ALTER TABLE """ & t.tableName & """ ADD CONSTRAINT " & LCase(constUk.constraintName) & " UNIQUE (", True
        End If

        i = 0
        For Each c In constUk.columns.col

            If i <> 0 Then
                writer.writeText "    , " & """" & c.columnName & """", True
            Else
                writer.writeText "      " & """" & c.columnName & """", True
            End If

            i = i + 1
        Next

        If constUk.columns.count > 0 Then
            writer.writeText ");", True
        End If

    Next
    
    writer.writeText "", True
    
    ' ----------------------------------------------------
    ' FK制約
    ' ----------------------------------------------------
    Dim constFk As ConstraintForeign
    For Each constFk In t.constraintsFks.col
    
        If constFk.columns.count > 0 Then
            writer.writeText "-- -----------------------------------------------------------------------------", True
            writer.writeText "-- FK：" & constFk.constraintName, True
            writer.writeText "-- -----------------------------------------------------------------------------", True
            writer.writeText "ALTER TABLE """ & t.tableName & """ ADD CONSTRAINT FOREIGN KEY " & LCase(constFk.constraintName) & " (", True
        End If
        
        i = 0
        For Each c In constFk.columns.col
        
            If i <> 0 Then
                writer.writeText "    , " & """" & c.columnName & """", True
            Else
                writer.writeText "      " & """" & c.columnName & """", True
            End If
            
            i = i + 1
        Next
        
        If constFk.columns.count > 0 Then
            writer.writeText ")", False
        End If
        
        If constFk.refColumns.count > 0 Then
            writer.writeText " REFERENCES " & constFk.refTable & " (", True
        End If
        
        i = 0
        For Each var In constFk.refColumns.col
        
            If i <> 0 Then
                writer.writeText "    , " & """" & var & """", True
            Else
                writer.writeText "      " & """" & var & """", True
            End If
            
            i = i + 1
        Next
        
        If constFk.refColumns.count > 0 Then
        
            If constFk.refOption <> "" Then
                writer.writeText ") " & constFk.refOption & ";", True
            Else
                writer.writeText ");", True
            End If
        End If
        
    Next
    
    writer.writeText "", True
    
    ' ----------------------------------------------------
    ' インデックス制約
    ' ----------------------------------------------------
    Dim idx As index
    For Each idx In t.indexes.col
    
        If idx.columns.count > 0 Then
            writer.writeText "-- -----------------------------------------------------------------------------", True
            writer.writeText "-- Index：" & idx.indexName, True
            writer.writeText "-- -----------------------------------------------------------------------------", True
            writer.writeText "CREATE INDEX " & LCase(idx.indexName) & " ON """ & t.tableName & """ (", True
        End If
        
        i = 0
        For Each c In idx.columns.col

            If i <> 0 Then
                writer.writeText "    , " & "`" & c.columnName & "`", True
            Else
                writer.writeText "      " & "`" & c.columnName & "`", True
            End If

            i = i + 1
        Next
        
        If idx.columns.count > 0 Then
            writer.writeText ");", True
        End If

    Next
    
    writer.writeText "", True
    
    writer.makeUtf8WithoutBom
    
End Function


