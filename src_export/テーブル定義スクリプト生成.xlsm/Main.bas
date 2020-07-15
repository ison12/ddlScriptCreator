Attribute VB_Name = "Main"
Option Explicit

' *********************************************************
' ���C�����W���[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/07/14�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

' MySQL
Public Const DBMS_TYPE_MYSQL      As String = "MySQL"
' PostgreSQL
Public Const DBMS_TYPE_POSTGRESQL As String = "PostgreSQL"
' SQLServer
Public Const DBMS_TYPE_SQLSERVER  As String = "SQLServer"
' DDL�N�G���i�S�e�[�u���j
Const ALL_DDL_FILE_NAME   As String = "all_ddl.sql"

' ------------------------------------------------------------------
' DB�e�[�u����`�X�N���v�g�𐶐����郁�C�����\�b�h
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
    ' DB�e�[�u����`�i�[�t�H���_�I��
    '---------------------------------------------
    If inFolder = "" Then
        ' �t�@�C���p�X���w�肳��Ȃ��ꍇ�́A�Θb���[�h
        inFolder = selectInFolder
        If inFolder = Empty Then
            Exit Sub
        End If
    End If
   
    '---------------------------------------------
    ' DDL�X�N���v�g���i�[����t�H���_�I��
    '---------------------------------------------
    If outFolder = "" Then
        ' �t�@�C���p�X���w�肳��Ȃ��ꍇ�́A�Θb���[�h
        outFolder = selectOutFolder
        If outFolder = Empty Then
            Exit Sub
        End If
    End If

    '----------------------------------------------------------------
    ' �e�[�u����`�t�H���_�z����Excel�t�@�C����Ώ�
    '----------------------------------------------------------------
    Dim tableDefineBookList     As ValCollection
    Dim tableDefineBookFileName As Variant
    Set tableDefineBookList = findTableDefineBookList(inFolder)
    
    Application.ScreenUpdating = False
    
    '----------------------------------------------------------------
    ' �e�[�u����`Excel�t�@�C������������
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
    
    '-- ��ʍX�V�̔�\��
    Application.ScreenUpdating = False
    
    For Each tableDefineBookFileName In tableDefineBookList.col
    
        ' --------------------------------------------
        '-- �e�[�u����`�����J��
        Set targetBook = openBookIfNotOpen(inFolder & "\" & tableDefineBookFileName, alreadyExiests)
        If alreadyExiests = False Then
            needCloseBooks.setItem targetBook
        End If
        ' --------------------------------------------
        
        '-- �ΏۃV�[�g�I��
        For Each sheet In targetBook.Sheets
        
            '-- �e�[�u����`�V�[�g�ł��邩�𔻒肷��
            If isTableDefineSheet(sheet) Then
                
                '------------------------------
                '-- �e�[�u����`�V�[�g����e�[�u����`�����
                '------------------------------
                Set t = parseTableFromSheet(dbmsType, sheet)
                parsedTableList.setItem t
    
                tableName = t.tableName
                outFilePath = VBUtil.concatFilePath(outFolder, tableName & ".sql")
                outFileConstPath = VBUtil.concatFilePath(outFolder, "const_" & tableName & ".sql")
                
                Application.StatusBar = "DB�e�[�u����`�������E�E�E" & tableName
                DoEvents
                
                '------------------------------
                '-- DDL�쐬�i�e�[�u�����j
                '------------------------------
                VBUtil.deleteFile outFilePath
                VBUtil.deleteFile outFileConstPath
                createTableDDL dbmsType, outFilePath, t
                createTableConstDDL dbmsType, outFileConstPath, t
                
                '------------------------------
                '-- DDL�쐬�i�S��`�܂ރt�@�C���j
                '------------------------------
                createTableDDL dbmsType, outFilePathAll, t
        
                count = count + 1
        
            End If
            
        Next
    
    Next

    '-- DDL�쐬�i�S��`�܂ރt�@�C���j
    ' �S������̃t�@�C���͍ŏ��ɍ폜����
    VBUtil.deleteFile outFilePathAll
    ' �e�[�u����`���o��
    For Each t In parsedTableList.col
        createTableDDL dbmsType, outFilePathAll, t
    Next
    ' �e�[�u�������`���o�́i����̓e�[�u���Ԃ̐��������̂ōŌ�ɖ����ɏo�͂���j
    For Each t In parsedTableList.col
        createTableConstDDL dbmsType, outFilePathAll, t
    Next

    '-- �e�[�u����`�� CLOSE
    For Each book In needCloseBooks.col
        book.Close savechanges:=False
    Next
  
    If Not silent Then
        MsgBox "DB�e�[�u����`���������E�E�E" & count
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    Exit Sub
    
err:
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    If Not silent Then
        MsgBox "�G���[���������܂����B" & err.Number & " " & err.Description
    End If
    
    Exit Sub

End Sub

' ------------------------------------------------------------------
' �e�[�u����`����͂���
' ------------------------------------------------------------------
Private Function parseTableFromSheet( _
    ByVal dbmsType As String, _
    ByRef sheet As Worksheet) As Table

    Dim tableParser  As ITableDefineParser
    Set tableParser = New ImplTableDefineParser
    
    Set parseTableFromSheet = tableParser.parse(sheet)
    
End Function

' ------------------------------------------------------------------
' DDL�𐶐�����
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
' DDL�i����֘A�j�𐶐�����
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
' �e�[�u����`�V�[�g�ł��邩�𔻒肷��
' --------------------------------------------------------------------
Private Function isTableDefineSheet(sheet As Worksheet)

    On Error GoTo err

    ' �e�[�u����`�ł��邱�Ƃ�\��Shape�I�u�W�F�N�g
    Dim s As Shape
    Set s = sheet.Shapes("table_define")

    If Not s Is Nothing And sheet.Range("C4") <> "" Then
        ' �e�[�u����`�V�[�g�Ƃ��ėL��
        isTableDefineSheet = True
    Else
        ' �e�[�u����`�V�[�g�Ƃ��Ė���
        isTableDefineSheet = False
    End If
    
    Exit Function
    
err:

    ' �e�[�u����`�V�[�g�Ƃ��Ė���
    isTableDefineSheet = False
    
End Function

' --------------------------------------------------------------------
' ���̓t�H���_��I������
' --------------------------------------------------------------------
Private Function selectInFolder() As String

    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .title = "DB�̃e�[�u����`���i�[����Ă���t�H���_���w�肵�Ă�������"
        .initialFileName = VBUtil.convertOneDriveUrlToLocalFilePath(ThisWorkbook.Path) & "\"
    
        If .Show Then
        
            selectInFolder = .SelectedItems(1)
        Else
        
            selectInFolder = Empty
        End If
        
    End With

End Function

' --------------------------------------------------------------------
' �o�̓t�H���_��I������
' --------------------------------------------------------------------
Private Function selectOutFolder() As String
        
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .title = "DB��DDL�X�N���v�g���i�[����t�H���_���w�肵�Ă�������"
        .initialFileName = VBUtil.convertOneDriveUrlToLocalFilePath(ThisWorkbook.Path) & "\"
        
    
        If .Show Then
        
            selectOutFolder = .SelectedItems(1)
        Else
        
            selectOutFolder = Empty
        End If
        
    End With

End Function

' --------------------------------------------------------------------
' �e�[�u����`�u�b�N�̃t�@�C���p�X���X�g���擾����
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
' �u�b�N���J��
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
' �e�[�u����`�����I�u�W�F�N�g�̐���
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
' �e�[�u����`�ꗗ���o�͂���
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
