VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImplTableDefineParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITableDefineParser

Private Function ITableDefineParser_parse(sheet As Worksheet) As Table

    Const START_CURRENT_ROW As Long = 11    '-- 設定開始行

    Dim i  As Long
    
    Dim t  As New Table
    Dim cl As New ValCollection
    
    With sheet
    
        Dim currentValue As String
        Dim currentRow   As Long
        
        ' -------------------------------------------
        ' テーブル情報読み込み
        ' -------------------------------------------
        t.tableName = .Cells(5, 3).value
        t.tableComment = .Cells(4, 3).value
        
        If t.tableName = "" Then
            err.Raise 999, Description:=vbNewLine & sheet.name & " : テーブル名が未入力のため処理できません。"
        End If
        
        ' -------------------------------------------
        ' カラム情報読み込み
        ' -------------------------------------------
        Dim c As Column
        
        currentRow = START_CURRENT_ROW
        Do While True
        
            currentValue = .Cells(currentRow, 2)
            If currentValue = "" Then
            
                Exit Do
            End If
            
            Set c = New Column
            c.columnName = .Cells(currentRow, 3)
            c.columnComment = .Cells(currentRow, 2)
            c.dataType = .Cells(currentRow, 4)
            
            If .Cells(currentRow, 5) <> "" Then
                c.isNotNull = True
            Else
                c.isNotNull = False
            End If
            
            c.defaultValue = .Cells(currentRow, 6)
            
            ' リストにカラムを追加
            cl.setItem c, c.columnName
        
            currentRow = currentRow + 1
        Loop
        
        Set t.columns = cl
        
        ' -------------------------------------------
        ' PK制約読み込み
        ' -------------------------------------------
        Dim pk As New ConstraintPrimary
        Set pk.Table = t
        Set pk.columns = New ValCollection
        
        currentRow = START_CURRENT_ROW
        Do While True
        
            currentValue = .Cells(currentRow, 3)
            If currentValue = "" Then
            
                Exit Do
            End If
            
            pk.constraintName = "PK_" & UCase(t.tableName)
            
            If .Cells(currentRow, 7) <> "" Then
               pk.columns.setItem (cl.getItem(currentValue)), currentValue
            End If
            
            currentRow = currentRow + 1
        Loop
        
        Set t.constraintsPk = pk
        
        ' -------------------------------------------
        ' UK制約読み込み
        ' -------------------------------------------
        Dim ukList As New ValCollection
        
        Dim uk As ConstraintUnique
        Dim ukName As String
        
        currentRow = START_CURRENT_ROW
        Do While True
        
            currentValue = .Cells(currentRow, 3)
            If currentValue = "" Then
        
                Exit Do
            End If
        
            If .Cells(currentRow, 8) <> "" Then
        
                Dim ukElements() As String
                ukElements = Split(.Cells(currentRow, 8), vbLf)
        
                For i = LBound(ukElements) To UBound(ukElements)
        
                    Set uk = ukList.getItem(ukElements(i))
                    If uk Is Nothing Then
        
                        Set uk = New ConstraintUnique
                        Set uk.Table = t
                        Set uk.columns = New ValCollection
        
                        If ukElements(i) = "〇" Or ukElements(i) = "○" Then
                            ukName = ""
                        Else
                            ukName = "_" & UCase(ukElements(i))
                        End If
        
                        uk.constraintName = "UK_" & UCase(t.tableName) & ukName
        
                        ukList.setItem uk, ukElements(i)
                    End If
        
                    uk.columns.setItem cl.getItem(currentValue)
                Next
        
            End If
        
            currentRow = currentRow + 1
        Loop
        
        Set t.constraintsUks = ukList
        
        ' -------------------------------------------
        ' 次の情報までスキップする
        ' -------------------------------------------
        currentRow = currentRow + 1
        Do While True
        
            currentValue = .Cells(currentRow, 2)
            If currentValue <> "" Then
            
                Exit Do
            End If
            
            currentRow = currentRow + 1
        Loop
        '----------------------------------------------------------
        
        ' -------------------------------------------
        ' fk制約読み込み
        ' -------------------------------------------
        Dim fkList As New ValCollection
        
        Dim fk As ConstraintForeign
        Dim fkNameLast  As String
        
        Dim fkRef()     As String
        Dim fkRefTable  As String
        Dim fkRefColumn As String
        
        currentRow = currentRow + 1
        '----------------------------------------------------------
        Do While True
        
            currentValue = .Cells(currentRow, 2)
            If currentValue = "" Then
            
                Exit Do
            End If
            
            ' FK名称が切り替わった場合
            If fkNameLast <> currentValue Then
                Set fk = New ConstraintForeign
                Set fk.columns = New ValCollection
                Set fk.refColumns = New ValCollection
                
                fk.constraintName = currentValue
                fkList.setItem fk, currentValue
            End If
            
            
            ' 参照カラムからテーブル名とカラム名を分割
            fkRef = Split(.Cells(currentRow, 4), ".")
            If UBound(fkRef) - LBound(fkRef) + 1 < 2 Then
                err.Raise 1000, Description:=vbNewLine & sheet.name & " : 外部キーの参照カラムが [table].[column] 形式になっていません。"
            End If
            
            fkRefTable = fkRef(0)
            fkRefColumn = fkRef(1)
            
            ' テーブル情報を設定
            Set fk.Table = t
            ' カラム情報が見つからない場合
            If cl.getItem(.Cells(currentRow, 3)) Is Nothing Then
                err.Raise 1001, Description:=vbNewLine & sheet.name & " : 外部キーのカラムがテーブル内に存在しません。"
            End If
            
            fk.columns.setItem cl.getItem(.Cells(currentRow, 3))
            ' 参照テーブル名を設定
            fk.refTable = fkRefTable
            ' 参照カラム名を設定
            fk.refColumns.setItem fkRefColumn, fkRefColumn
            ' 参照オプションを設定
            fk.refOption = .Cells(currentRow, 5)
            
            fkNameLast = currentValue
            
            currentRow = currentRow + 1
        Loop
        
        Set t.constraintsFks = fkList
        
        ' -------------------------------------------
        ' 次の情報までスキップする
        ' -------------------------------------------
        currentRow = currentRow + 1
        Do While True
        
            currentValue = .Cells(currentRow, 2)
            If currentValue <> "" Then
            
                Exit Do
            End If
            
            currentRow = currentRow + 1
        Loop

        
        ' -------------------------------------------
        ' INDEX制約読み込み
        ' -------------------------------------------
        Dim indexList As New ValCollection
        
        Dim idx          As index
        
        currentRow = currentRow + 1
        Do While True
        
            currentValue = .Cells(currentRow, 2)
            If currentValue = "" Then
            
                Exit Do
            End If
            
            Set idx = New index
            Set idx.columns = New ValCollection
            idx.indexName = currentValue
            indexList.setItem idx, currentValue
            
            ' テーブル情報を設定
            Set idx.Table = t
            
            ' インデックスのカラムを改行区切りで配列に変換（カラムは複数指定の可能性があるため）
            Dim indexElements() As String
            indexElements = Split(.Cells(currentRow, 4), vbLf)
    
            For i = LBound(indexElements) To UBound(indexElements)
    
                ' カラム情報が見つからない場合
                If cl.getItem(indexElements(i)) Is Nothing Then
                    err.Raise 1002, Description:=vbNewLine & sheet.name & " : インデックスのカラムがテーブル内に存在しません。"
                End If
                
                ' カラム情報を設定
                idx.columns.setItem cl.getItem(indexElements(i))
            Next
            
            currentRow = currentRow + 1
        Loop
        
        Set t.indexes = indexList
        
    End With
    
    Set ITableDefineParser_parse = t
    
End Function
