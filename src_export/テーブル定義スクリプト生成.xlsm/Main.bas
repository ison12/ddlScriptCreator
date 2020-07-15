Attribute VB_Name = "Main"
Option Explicit

' *********************************************************
' メインモジュール
'
' 作成者　：Ison
' 履歴　　：2020/07/14　新規作成
'
' 特記事項：
'
' *********************************************************

' MySQL
Public Const DBMS_TYPE_MYSQL      As String = "MySQL"
' PostgreSQL
Public Const DBMS_TYPE_POSTGRESQL As String = "PostgreSQL"
' SQLServer
Public Const DBMS_TYPE_SQLSERVER  As String = "SQLServer"
' DDLクエリ（全テーブル）
Const ALL_DDL_FILE_NAME   As String = "all_ddl.sql"

' ------------------------------------------------------------------
' DBテーブル定義スクリプトを生成するメインメソッド
' ------------------------------------------------------------------
Public Sub createDDLScript( _
    Optional ByVal dbmsType As String = DBMS_TYPE_MYSQL, _
    Optional ByVal inFolder As String = "", _
    Optional ByVal outFolder As String = "", _
    Optional ByVal silent As Boolean = True)

    On Error GoTo err
    
    Dim sheet As Worksheet
    
    Dim count As Long
    Dim tableName As String
    
    '---------------------------------------------
    ' DBテーブル定義格納フォルダ選択
    '---------------------------------------------
    If inFolder = "" Then
        ' ファイルパスが指定されない場合は、対話モード
        inFolder = selectInFolder
        If inFolder = Empty Then
            Exit Sub
        End If
    End If
   
    '---------------------------------------------
    ' DDLスクリプトを格納するフォルダ選択
    '---------------------------------------------
    If outFolder = "" Then
        ' ファイルパスが指定されない場合は、対話モード
        outFolder = selectOutFolder
        If outFolder = Empty Then
            Exit Sub
        End If
    End If

    '----------------------------------------------------------------
    ' テーブル定義フォルダ配下のExcelファイルを対象
    '----------------------------------------------------------------
    Dim tableDefineBookList     As ValCollection
    Dim tableDefineBookFileName As Variant
    Set tableDefineBookList = findTableDefineBookList(inFolder)
    
    Application.ScreenUpdating = False
    
    '----------------------------------------------------------------
    ' テーブル定義Excelファイルを順次処理
    '----------------------------------------------------------------
    Dim book            As Workbook
    Dim targetBook      As Workbook
    Dim t               As Table
    
    Dim needCloseBooks  As New ValCollection
    Dim alreadyExiests  As Boolean
    
    Dim parsedTableList As New ValCollection
    
    Dim outFilePath      As String
    Dim outFileConstPath As String
    Dim outFilePathAll   As String
    outFilePathAll = VBUtil.concatFilePath(outFolder, ALL_DDL_FILE_NAME)
    
    '-- 画面更新の非表示
    Application.ScreenUpdating = False
    
    For Each tableDefineBookFileName In tableDefineBookList.col
    
        ' --------------------------------------------
        '-- テーブル定義書を開く
        Set targetBook = openBookIfNotOpen(inFolder & "\" & tableDefineBookFileName, alreadyExiests)
        If alreadyExiests = False Then
            needCloseBooks.setItem targetBook
        End If
        ' --------------------------------------------
        
        '-- 対象シート選択
        For Each sheet In targetBook.Sheets
        
            '-- テーブル定義シートであるかを判定する
            If isTableDefineSheet(sheet) Then
                
                '------------------------------
                '-- テーブル定義シートからテーブル定義を解析
                '------------------------------
                Set t = parseTableFromSheet(dbmsType, sheet)
                parsedTableList.setItem t
    
                tableName = t.tableName
                outFilePath = VBUtil.concatFilePath(outFolder, tableName & ".sql")
                outFileConstPath = VBUtil.concatFilePath(outFolder, "const_" & tableName & ".sql")
                
                Application.StatusBar = "DBテーブル定義生成中・・・" & tableName
                DoEvents
                
                '------------------------------
                '-- DDL作成（テーブル毎）
                '------------------------------
                VBUtil.deleteFile outFilePath
                VBUtil.deleteFile outFileConstPath
                createTableDDL dbmsType, outFilePath, t
                createTableConstDDL dbmsType, outFileConstPath, t
                
                '------------------------------
                '-- DDL作成（全定義含むファイル）
                '------------------------------
                createTableDDL dbmsType, outFilePathAll, t
        
                count = count + 1
        
            End If
            
        Next
    
    Next

    '-- DDL作成（全定義含むファイル）
    ' 全部入りのファイルは最初に削除する
    VBUtil.deleteFile outFilePathAll
    ' テーブル定義を出力
    For Each t In parsedTableList.col
        createTableDDL dbmsType, outFilePathAll, t
    Next
    ' テーブル制約定義を出力（制約はテーブル間の制約もあるので最後に末尾に出力する）
    For Each t In parsedTableList.col
        createTableConstDDL dbmsType, outFilePathAll, t
    Next

    '-- テーブル定義書 CLOSE
    For Each book In needCloseBooks.col
        book.Close savechanges:=False
    Next
  
    If Not silent Then
        MsgBox "DBテーブル定義生成完了・・・" & count
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    Exit Sub
    
err:
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    If Not silent Then
        MsgBox "エラーが発生しました。" & err.Number & " " & err.Description
    End If
    
    Exit Sub

End Sub

' ------------------------------------------------------------------
' テーブル定義を解析する
' ------------------------------------------------------------------
Private Function parseTableFromSheet( _
    ByVal dbmsType As String, _
    ByRef sheet As Worksheet) As Table

    Dim tableParser  As ITableDefineParser
    Set tableParser = New ImplTableDefineParser
    
    Set parseTableFromSheet = tableParser.parse(sheet)
    
End Function

' ------------------------------------------------------------------
' DDLを生成する
' ------------------------------------------------------------------
Private Function createTableDDL( _
    ByVal dbmsType As String, _
    ByVal filePath As String, _
    ByRef t As Table)

    Dim tableCreator As ITableDefineCreator
    Set tableCreator = createTableCreator(dbmsType)
    
    tableCreator.writeForTable t, filePath, True
    
End Function

' ------------------------------------------------------------------
' DDL（制約関連）を生成する
' ------------------------------------------------------------------
Private Function createTableConstDDL( _
    ByVal dbmsType As String, _
    ByVal filePath As String, _
    ByRef t As Table)

    Dim tableCreator As ITableDefineCreator
    Set tableCreator = createTableCreator(dbmsType)
    
    tableCreator.writeForConstraints t, filePath, True
    
End Function

' --------------------------------------------------------------------
' テーブル定義シートであるかを判定する
' --------------------------------------------------------------------
Private Function isTableDefineSheet(sheet As Worksheet)

    On Error GoTo err

    ' テーブル定義であることを表すShapeオブジェクト
    Dim s As Shape
    Set s = sheet.Shapes("table_define")

    If Not s Is Nothing And sheet.Range("C4") <> "" Then
        ' テーブル定義シートとして有効
        isTableDefineSheet = True
    Else
        ' テーブル定義シートとして無効
        isTableDefineSheet = False
    End If
    
    Exit Function
    
err:

    ' テーブル定義シートとして無効
    isTableDefineSheet = False
    
End Function

' --------------------------------------------------------------------
' 入力フォルダを選択する
' --------------------------------------------------------------------
Private Function selectInFolder() As String

    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .title = "DBのテーブル定義が格納されているフォルダを指定してください"
        .initialFileName = VBUtil.convertOneDriveUrlToLocalFilePath(ThisWorkbook.Path) & "\"
    
        If .Show Then
        
            selectInFolder = .SelectedItems(1)
        Else
        
            selectInFolder = Empty
        End If
        
    End With

End Function

' --------------------------------------------------------------------
' 出力フォルダを選択する
' --------------------------------------------------------------------
Private Function selectOutFolder() As String
        
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .title = "DBのDDLスクリプトを格納するフォルダを指定してください"
        .initialFileName = VBUtil.convertOneDriveUrlToLocalFilePath(ThisWorkbook.Path) & "\"
        
    
        If .Show Then
        
            selectOutFolder = .SelectedItems(1)
        Else
        
            selectOutFolder = Empty
        End If
        
    End With

End Function

' --------------------------------------------------------------------
' テーブル定義ブックのファイルパスリストを取得する
' --------------------------------------------------------------------
Private Function findTableDefineBookList(ByVal folder As String) As ValCollection
    
    Dim tableDefineBookList As New ValCollection
    Dim filePath As String
    
    filePath = dir(folder & "\*.xlsx")
    Do While filePath <> ""
    
        tableDefineBookList.setItem filePath
        filePath = dir()
    Loop
    
    Set findTableDefineBookList = tableDefineBookList

End Function

' ------------------------------------------------------------------
' ブックを開く
' ------------------------------------------------------------------
Private Function openBookIfNotOpen(ByVal tableDefineBookFilePath As String, ByRef alreadyExiests As Boolean) As Workbook
    
    Dim book       As Workbook
    Dim targetBook As Workbook
    
    Dim fileName   As String
    fileName = VBUtil.extractFileName(tableDefineBookFilePath)
    
    alreadyExiests = False

    For Each book In Workbooks
        If book.name = fileName Then
            Set openBookIfNotOpen = book
            alreadyExiests = True
            Exit For
        End If
    Next
    
    If targetBook Is Nothing Then
        Set openBookIfNotOpen = Workbooks.Open(tableDefineBookFilePath)
    End If

End Function

' ------------------------------------------------------------------
' テーブル定義生成オブジェクトの生成
' ------------------------------------------------------------------
Private Function createTableCreator(ByVal dbmsType As String) As ITableDefineCreator

    Dim tableCreator As ITableDefineCreator
    
    Select Case dbmsType
    
        Case DBMS_TYPE_SQLSERVER
            Set tableCreator = New ImplTableDefineCreatorSqlserver
        
        Case DBMS_TYPE_POSTGRESQL
            Set tableCreator = New ImplTableDefineCreatorPostgres
        
        Case DBMS_TYPE_MYSQL
            Set tableCreator = New ImplTableDefineCreatorMySQL
    
    End Select
    
    Set createTableCreator = tableCreator

End Function

' --------------------------------------------------------------------
' テーブル定義一覧を出力する
' --------------------------------------------------------------------
Public Sub renderSheetListFromActiveBook()

    Dim book As Workbook
    Set book = ActiveWorkbook
    
    Dim sheet As Worksheet
    Set sheet = ActiveSheet
    
    Dim j As Long
    Dim i As Long
    For i = 1 To book.Sheets.count
    
        If isTableDefineSheet(Sheets(i)) Then
    
            Dim name As String
            Dim namePhi As String
            
            name = Sheets(i).Range("C4").value
            namePhi = Sheets(i).Range("C5").value
    
            sheet.Cells(j + Selection.Row, Selection.Column) = name
            sheet.Cells(j + Selection.Row, Selection.Column + 10) = namePhi
               
            sheet.Hyperlinks.Add Anchor:=sheet.Cells(j + Selection.Row, Selection.Column), Address:="", SubAddress:=Sheets(i).name & "!$A$1", TextToDisplay:=name
               
            j = j + 1
        End If
    
    Next i
    
End Sub
