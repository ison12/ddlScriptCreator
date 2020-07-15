Attribute VB_Name = "VBUtil"
Option Explicit

' *********************************************************
' VB�֘A�̋��ʊ֐����W���[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/08/10�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

' �G���[�����i�[����\����
Public Type errInfo

    Source       As Variant
    Number       As Variant
    Description  As Variant
    LastDllError As Variant
    HelpFile     As Variant
    HelpContext  As Variant
    
End Type

' =========================================================
' ��Err�I�u�W�F�N�g�̏����\���̂ɑޔ�
'
' �T�v�@�@�@�FErr�I�u�W�F�N�g�̏����\���̂ɐݒ肵�ĕԂ��B
' �����@�@�@�F
' �߂�l�@�@�F�G���[���
'
' ���L�����@�F�G���[�n���h���ŕʂ̊֐����Ăяo����Err�I�u�W�F�N�g�̏�񂪏����Ă��܂����Ƃ�����
' �@�@�@�@�@�@���̏�ԂŁAErr.Raise����Ɛ�����������ʂ̃��W���[���ɂœ`�d�ł��Ȃ��B
' �@�@�@�@�@�@����������`�d����ꍇ�ɂ́A�{�֐��𗘗p���āA��x�G���[����ޔ����Ă���Err.Raise���Ă��Ɨǂ��B
'
' �@�@�@�@�@�@�g�p��F
' �@�@�@�@�@�@�@Dim errT As errInfo
' �@�@�@�@�@�@�@errT = VBUtil.swapErr

' �@�@�@�@�@�@�@�E�E�E�G���[���̌�n�������Ȃ�
'
' �@�@�@�@�@�@�@Err.Raise errT.Number, errT.Source�E�E�E
'
' =========================================================
Public Function swapErr() As errInfo

    swapErr.Source = err.Source
    swapErr.Number = err.Number
    swapErr.Description = err.Description
    swapErr.LastDllError = err.LastDllError
    swapErr.HelpFile = err.HelpFile
    swapErr.HelpContext = err.HelpContext

End Function

' =========================================================
' ���ۑ��_�C�A���O�\��
'
' �T�v�@�@�@�F�ۑ��_�C�A���O��\������
' �����@�@�@�Ftitle           �_�C�A���O�̃^�C�g��
' �@�@�@�@�@�@filter          �t�B���^
' �@�@�@�@�@�@initialFileName �����t�@�C����
' �߂�l�@�@�F�ۑ��t�@�C���p�X
'
' =========================================================
Public Function openFileSaveDialog(ByVal title As String, ByVal filter As String, ByVal initialFileName As String) As String

    ' �A�v���P�[�V����
    Dim xlsApp   As Application
    
    ' �t�@�C���p�X
    Dim filePath As Variant

    ' Application�I�u�W�F�N�g�擾
    Set xlsApp = Application
    
    ' �_�C�A���O�őI�����ꂽ�t�@�C�������i�[
    filePath = xlsApp.GetSaveAsFilename(initialFileName:=initialFileName _
                                      , fileFilter:=filter _
                                      , title:=title)
                                      
    ' �L�����Z�����ꂽ���𔻒肷��
    If filePath = False Then
    
        ' �L�����Z�����ꂽ�ꍇ �󕶎����Ԃ�
        openFileSaveDialog = ""
        
    Else
        ' �ۑ���I�����ꂽ�ꍇ �t�@�C������Ԃ�
        openFileSaveDialog = filePath
    End If

End Function

' =========================================================
' ���J���_�C�A���O�\��
'
' �T�v�@�@�@�F�J���_�C�A���O��\������
' �����@�@�@�Ftitle           �_�C�A���O�̃^�C�g��
' �@�@�@�@�@�@filter          �t�B���^
' �@�@�@�@�@�@multiSelect     �����I��
' �߂�l�@�@�F�I�������t�@�C���̃t�@�C���p�X
'
' =========================================================
Public Function openFileDialog(ByVal title As String, ByVal filter As String, Optional ByVal multiSelect As Boolean = False) As Variant

    ' �A�v���P�[�V����
    Dim xlsApp   As Application
    
    ' �t�@�C���p�X
    Dim filePath As Variant

    ' Application�I�u�W�F�N�g�擾
    Set xlsApp = Application
    
    ' �_�C�A���O�őI�����ꂽ�t�@�C�������i�[
    filePath = xlsApp.GetOpenFilename(fileFilter:=filter _
                                    , title:=title _
                                    , multiSelect:=multiSelect)

    ' �����I���̏ꍇ�A�߂�l�Ƃ��Ĕz�񂪕Ԃ����̂Ŕz�񂩂ǂ����𔻒肷��
    If IsArray(filePath) Then
    
        ' �ۑ���I�����ꂽ�ꍇ �t�@�C������Ԃ�
        openFileDialog = filePath
    
    ' �I�����L�����Z�����ꂽ�ꍇ
    ElseIf filePath = False Then
    
        ' �L�����Z�����ꂽ�ꍇ ���Ԃ�
        openFileDialog = Empty
        
    Else
        ' �ۑ���I�����ꂽ�ꍇ �t�@�C������Ԃ�
        openFileDialog = filePath
    
    End If

End Function

' =========================================================
' ���t�@�C���̊g���q�`�F�b�N
'
' �T�v�@�@�@�F�t�@�C���̊g���q���`�F�b�N����
' �����@�@�@�Ffile      �t�@�C����
' �@�@�@�@�@�@extension �g���q
' �߂�l�@�@�F�t�@�C���̊g���q���w�肳�ꂽ����extension�̏ꍇTrue��Ԃ�
'
' =========================================================
Public Function checkFileExtension(ByRef file As String _
                                 , ByRef extension As String) As Boolean

    ' �t�@�C�������璊�o�����g���q
    Dim fileExtension As String
    
    ' �C���f�b�N�X
    Dim index As Long
    
    ' �t�@�C�����Ɗg���q�̋�؂蕶���ł���h�b�g(.)����������
    index = InStrRev(file, ".")
    
    ' �h�b�g(.)��������Ȃ��ꍇ
    If index <= 0 Then
    
        Exit Function
    End If
    
    ' �t�@�C��������g���q�𒊏o����
    fileExtension = Mid$(file, index + 1, Len(file))

    If fileExtension = extension Then
    
        checkFileExtension = True
    Else
    
        checkFileExtension = False
    End If

End Function

' =========================================================
' ���t�@�C���p�X����t�@�C�������o
'
' �T�v�@�@�@�F�t�@�C���p�X����t�@�C�����𒊏o����
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�F�t�@�C����
'
' =========================================================
Public Function extractFileName(ByRef filePath As String) As String
    
    ' �t�@�C���p�X��؂蕶��
    Const FILE_SEPARATE As String = "\"

    ' �t�@�C���p�X�̉E�������͂��߂ɏo��������؂蕶���̕����ʒu
    Dim index As Long
    
    ' ��؂蕶���̈ʒu���擾����
    index = InStrRev(filePath, FILE_SEPARATE)

    ' ��؂蕶���𔭌������ꍇ
    If index > 0 Then
    
        extractFileName = Mid$(filePath, index + 1)
    
    ' ��؂蕶���𔭌��ł��Ȃ������ꍇ
    Else
        extractFileName = filePath
    
    End If

End Function

' =========================================================
' ���C���t�H���b�Z�[�W�{�b�N�X��\��
'
' �T�v�@�@�@�F�C���t�H���b�Z�[�W�{�b�N�X��\������
' �����@�@�@�FbasePrompt ��{���b�Z�[�W
'             title      ���b�Z�[�W�{�b�N�X�̃^�C�g��
' �@�@�@�@�@�@err        �G���[�I�u�W�F�N�g
'
' =========================================================
Public Sub showMessageBoxForInformation(ByRef basePrompt As String _
                                      , ByRef title As String _
                                      , Optional ByRef err As ErrObject = Nothing)

    MsgBox basePrompt _
         , vbInformation _
         , title
         
End Sub

' =========================================================
' ���G���[���b�Z�[�W�{�b�N�X��\��
'
' �T�v�@�@�@�F�G���[���b�Z�[�W�{�b�N�X��\������
' �����@�@�@�FbasePrompt ��{���b�Z�[�W
'             title      ���b�Z�[�W�{�b�N�X�̃^�C�g��
' �@�@�@�@�@�@err        �G���[�I�u�W�F�N�g
'
' =========================================================
Public Sub showMessageBoxForError(ByRef basePrompt As String _
                                , ByRef title As String _
                                , ByRef err As ErrObject)

    MsgBox basePrompt & vbNewLine & vbNewLine & _
           err.Description & vbNewLine & _
           "Error no [" & err.Number & "]" _
         , vbCritical _
         , title
         
End Sub

' =========================================================
' ���x�����b�Z�[�W�{�b�N�X��\��
'
' �T�v�@�@�@�F�x�����b�Z�[�W�{�b�N�X��\������
' �����@�@�@�FbasePrompt ��{���b�Z�[�W
'             title      ���b�Z�[�W�{�b�N�X�̃^�C�g��
' �@�@�@�@�@�@err        �G���[�I�u�W�F�N�g
'
' =========================================================
Public Sub showMessageBoxForWarning(ByVal basePrompt As String _
                                  , ByVal title As String _
                                  , ByRef err As ErrObject)

    If err Is Nothing Then
    
        MsgBox basePrompt _
             , vbExclamation _
             , title
    
    ElseIf err.Number = 0 Then
    
        MsgBox basePrompt _
             , vbExclamation _
             , title
    Else
    
        If basePrompt <> "" Then
        
            basePrompt = basePrompt & vbNewLine & vbNewLine
        End If
        
        MsgBox basePrompt & _
               err.Description & vbNewLine & _
               "Error no [" & err.Number & "]" _
             , vbExclamation _
             , title
    
    End If
         
End Sub

' =========================================================
' ���z��T�C�Y�擾
'
' �T�v�@�@�@�F�z��̃T�C�Y���擾����
' �����@�@�@�Fvar       �z��
' �@�@�@�@�@�@dimension ����
'
' =========================================================
Public Function arraySize(ByRef var As Variant, Optional ByVal dimension As Long = 1) As Long

    If IsArray(var) = True Then
    
        arraySize = UBound(var, dimension) - LBound(var, dimension) + 1
        
    Else
        arraySize = 0
    
    End If
    

End Function

' =========================================================
' ��2�����z��̔C�ӂ̍s��1�����z��Ƃ��ĕԂ�
'
' �T�v�@�@�@�F
' �����@�@�@�Fval �z��
'             i   �z��̃C���f�b�N�X
'
' =========================================================
Public Function convert2to1Array(ByRef val As Variant, ByVal i As Long) As Variant

    ' �߂�l
    Dim ret() As Variant

    Dim j As Long
    
    ReDim ret(LBound(val, 2) To UBound(val, 2))
    
    For j = LBound(ret) To UBound(ret)
    
        ret(j) = val(i, j)
    
    Next
    
    convert2to1Array = ret

End Function

' =========================================================
' ��2�����z����f�o�b�O�E�B���h�E�ɏo�͂���
'
' �T�v�@�@�@�F
' �����@�@�@�Fval �z��
'
' =========================================================
Public Function debugPrintArray(ByRef val As Variant)

    ' �z��̃C���f�b�N�X
    Dim i As Long
    Dim j As Long
    
    ' �f�o�b�O�E�B���h�E�ɏo�͂��镶����
    Dim str As String
    
    str = "Output Array" & vbNewLine
    
    ' -------------------------------------------------
    ' �z��Ƃ��ď���������Ă���ꍇ�ɏo�͂����{����
    ' -------------------------------------------------
    If VarType(val) = (vbArray + vbVariant) Then
    
        ' ���[�v����
        For i = LBound(val, 1) To UBound(val, 1)
        
            str = str & "+   [" & i & "] - {"
        
            For j = LBound(val, 2) To UBound(val, 2)
            
                str = str & val(i, j) & ", "
            Next
            
            str = str & "}" & vbNewLine
            
        Next
        
    Else
        str = str & "   ... Empty"
        
    End If
    ' -------------------------------------------------
    
    Debug.Print str
    
End Function

' =========================================================
' ��2�����z��̗v�f����ւ�
'
' �T�v�@�@�@�F2�����z��̗v�f��(x,y)����(y,x)�ɐݒ肵�Ȃ����B
' �����@�@�@�Fv 2�����z��
'
' �߂�l�@�@�F2�����z��
'
' =========================================================
Public Function transposeDim(ByRef v As Variant) As Variant
    
    Dim x As Long
    Dim y As Long
    
    Dim Xlower As Long
    Dim Xupper As Long
    
    Dim Ylower As Long
    Dim Yupper As Long
    
    Dim tempArray As Variant
    
    Xlower = LBound(v, 2)
    Xupper = UBound(v, 2)
    Ylower = LBound(v, 1)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xlower To Xupper, Ylower To Yupper)
    
    For x = Xlower To Xupper
        For y = Ylower To Yupper
        
            tempArray(x, y) = v(y, x)
        
        Next y
    Next x
    
    transposeDim = tempArray

End Function

' =========================================================
' �������`�F�b�N
'
' �T�v�@�@�@�F
' �����@�@�@�Fvalue �`�F�b�N������
' �߂�l�@�@�FTrue ����
'
' =========================================================
Public Function validInteger(ByVal value As String) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = False

    ' �`�F�b�N�Ώۂ����l�Ŋ��A�����_���܂܂Ȃ��ꍇ�AOK�Ƃ���
    If _
            IsNumeric(value) = True _
        And InStr(value, ".") = 0 Then
    
        ret = True
    
    End If

    ' �߂�l��Ԃ�
    validInteger = ret

End Function

' =========================================================
' �������`�F�b�N�i�����͊܂܂Ȃ��j
'
' �T�v�@�@�@�F
' �����@�@�@�Fvalue �`�F�b�N������
' �߂�l�@�@�FTrue ����
'
' =========================================================
Public Function validUnsignedInteger(ByVal value As String) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = False

    ' �`�F�b�N�Ώۂ����l�Ŋ��A�}�C�i�X�L�����܂܂������_���܂܂Ȃ��ꍇ�AOK�Ƃ���
    If _
            IsNumeric(value) = True _
        And InStr(value, ".") = 0 _
        And InStr(value, "-") = 0 _
    Then
    
        ret = True
    
    End If

    ' �߂�l��Ԃ�
    validUnsignedInteger = ret

End Function

' =========================================================
' ��16�i���`�F�b�N
'
' �T�v�@�@�@�F
' �����@�@�@�Fvalue �`�F�b�N������
' �߂�l�@�@�FTrue 16�i��
'
' =========================================================
Public Function validHex(ByVal value As String) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = True

    ' �C���f�b�N�X
    Dim i    As Long
    ' �����̃T�C�Y
    Dim size As Long
    
    ' �������1������
    Dim one    As String
    ' 1��������ASCII�R�[�h
    Dim oneAsc As Long
    
    ' �����̃T�C�Y���擾����
    size = Len(value)
    
    ' �����񂩂�1���������o�����[�v�����s����
    For i = 1 To size
    
        ' 1�������o��
        one = Mid$(value, i, 1)
        ' ���o����������ASCII�R�[�h�𒲂ׂ�
        oneAsc = Asc(one)
        
        ' �����񂪈ȉ��͈͓̔��ł��邩���m�F����
        ' 0-9 a-f A-F
        If _
             (65 <= oneAsc And oneAsc <= 70) _
          Or (97 <= oneAsc And oneAsc <= 102) _
          Or (48 <= oneAsc And oneAsc <= 57) Then
        
            ' ����
            
        Else
        
            ' �G���[��
            ret = False
            Exit For
        
        End If
        
    Next

    ' �߂�l��Ԃ�
    validHex = ret

End Function

' =========================================================
' �����l�ł��邩���`�F�b�N����
'
' �T�v�@�@�@�F
' �����@�@�@�Fvalue �`�F�b�N������
' �߂�l�@�@�FTrue ����
'
' =========================================================
Public Function validNumeric(ByVal value As String) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = False

    ' �`�F�b�N�Ώۂ����l�̏ꍇ�AOK�Ƃ���
    If _
            IsNumeric(value) = True Then
    
        ret = True
    
    End If

    ' �߂�l��Ԃ�
    validNumeric = ret

End Function

' =========================================================
' �����l�ł��邩���`�F�b�N����i�����͊܂܂Ȃ��j
'
' �T�v�@�@�@�F
' �����@�@�@�Fvalue �`�F�b�N������
' �߂�l�@�@�FTrue ����
'
' =========================================================
Public Function validUnsignedNumeric(ByVal value As String) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = False

    ' �`�F�b�N�Ώۂ����l�Ŋ��}�C�i�X�L�����܂܂Ȃ��ꍇ�AOK�Ƃ���
    If _
            IsNumeric(value) = True _
        And InStr(value, "-") = 0 _
    Then
    
        ret = True
    
    End If

    ' �߂�l��Ԃ�
    validUnsignedNumeric = ret

End Function

' =========================================================
' ���R�[�h�l�`�F�b�N
'
' �T�v�@�@�@�F�����ŗ^����ꂽ�R�[�h���X�g�Ɉ�v������̂����邩���`�F�b�N����B
' �����@�@�@�Fvalue    �`�F�b�N������
' �@�@�@�@�@�@codeList �R�[�h���X�g
' �߂�l�@�@�FTrue �R�[�h���X�g�Ɉ�v����l������
'
' =========================================================
Public Function validCode(ByVal value As String, ParamArray codeList() As Variant) As Boolean

    ' �`�F�b�N�Ώۂ���̏ꍇ�AOK�Ƃ���
    Dim i As Long
    
    ' value��enums�̉��ꂩ�̒l�ƈ�v���Ă��邩�ǂ������m�F����
    For i = LBound(codeList) To UBound(codeList)
    
        ' ��v���Ă���ꍇ
        If value = CStr(codeList(i)) Then
        
            ' True��Ԃ�
            validCode = True
            
            Exit Function
        End If
    
    Next
    
    ' ��v������̂��Ȃ������̂ŁAFalse��Ԃ�
    validCode = False

End Function

' =========================================================
' ��RGB���]
'
' �T�v�@�@�@�FRGB�𔽓]������B
' �����@�@�@�Fr ��
' �@�@�@�@�@�@g ��
' �@�@�@�@�@�@b ��
' �߂�l�@�@�F���]�F
'
' =========================================================
Public Function reverseRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long

    reverseRGB = (Not RGB(r, g, b)) And &HFFFFFF

End Function

' =========================================================
' ��NULL���󕶎���ϊ�
'
' �T�v�@�@�@�FNull���󕶎���ɕϊ�����B
' �����@�@�@�Fvalue VARIANT�f�[�^
' �߂�l�@�@�F�󕶎���
' ���L�����@�FNull �l�́A�f�[�^ �A�C�e�� �ɗL���ȃf�[�^��
' �@�@�@�@�@�@�i�[����Ă��Ȃ����Ƃ������̂Ɏg�p�����o���A���g�^ (Variant) �̓��������`���ł��B
'
' =========================================================
Public Function convertNullToEmptyStr(ByRef value As Variant) As String

    ' NULL�̏ꍇ
    If IsNull(value) = True Then
    
        ' �󕶎���ɕϊ�
        convertNullToEmptyStr = ""
        
    ' �z��̏ꍇ
    ElseIf IsArray(value) Then
    
        ' �󕶎���ɕϊ�
        convertNullToEmptyStr = ""
        
    ' ���̑�
    Else
    
        ' ������ɕϊ����Ċi�[����
        convertNullToEmptyStr = CStr(value)
    End If
    
End Function

' =========================================================
' ���t�@�C�������݂��邩���`�F�b�N����
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�FTrue �t�@�C�������݂���ꍇ
'
' =========================================================
Public Function isExistFile(ByVal filePath As String) As Boolean

    ' �t�@�C����
    Dim fileName As String

    ' �w��̃t�@�C���p�X�����݂��邩�ǂ������`�F�b�N����
    fileName = dir(filePath, vbNormal)
    
    ' �t�@�C�������擾�ł������ǂ������`�F�b�N����
    If fileName = "" Then
    
        ' �t�@�C�������݂��Ȃ��ꍇ
        isExistFile = False
    Else
        
        ' �t�@�C�������݂���ꍇ
        isExistFile = True
    End If

End Function

' =========================================================
' ���t�@�C�������݂��邩���`�F�b�N����
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�FTrue �t�@�C�������݂���ꍇ
'
' =========================================================
Public Function isExistDirectory(ByVal filePath As String) As Boolean

    ' �f�B���N�g���p�X
    Dim dirPath As String
    ' �t�@�C����
    Dim fileName As String

    ' �t�@�C���p�X����f�B���N�g���𒊏o����
    dirPath = VBUtil.extractDirPathFromFilePath(filePath)
    
    ' �w��̃f�B���N�g�������݂��邩���m�F����
    fileName = dir(dirPath, vbDirectory)
    
    ' �t�@�C�������擾�ł������ǂ������`�F�b�N����
    If fileName = "" Then
    
        ' �t�@�C�������݂��Ȃ��ꍇ
        isExistDirectory = False
    Else
        
        ' �t�@�C�������݂���ꍇ
        isExistDirectory = True
    End If

End Function

' =========================================================
' ���t�@�C���p�X����f�B���N�g���p�X�𒊏o����
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�F�f�B���N�g���p�X
'
' =========================================================
Public Function extractDirPathFromFilePath(filePath As String) As String

    ' �߂�l
    Dim ret As String
    
    ' �f�B���N�g���ʒu
    Dim dirPoint As Long

    ' ������̉E�[����"\"���������A���[����̈ʒu���擾����
    dirPoint = InStrRev(filePath, "\")
    
    ' "\"��������Ȃ��ꍇ
    If dirPoint <> 0 Then
    
        ' �f�B���N�g���p�X�̎擾
        ret = Left$(filePath, dirPoint - 1)
        
        extractDirPathFromFilePath = ret
    
    Else
        extractDirPathFromFilePath = ""
    
    End If
    
End Function

' =========================================================
' ���t�@�C���p�X����t�@�C�����𒊏o����
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�F�f�B���N�g���p�X
'
' =========================================================
Public Function extractFileNameFromFilePath(filePath As String) As String

    ' �߂�l
    Dim ret As String
    
    ' �f�B���N�g���ʒu
    Dim dirPoint As Long

    ' ������̉E�[����"\"���������A���[����̈ʒu���擾����
    dirPoint = InStrRev(filePath, "\")
    
    ' "\"�����������ꍇ
    If dirPoint <> 0 Then
    
        ' �f�B���N�g���p�X�̎擾
        ret = right$(filePath, Len(filePath) - dirPoint)
        
        extractFileNameFromFilePath = ret
    
    Else
    
        extractFileNameFromFilePath = filePath
    End If
    
End Function

' =========================================================
' ���f�B���N�g���p�X�ƃt�@�C���p�X��A������
'
' �T�v�@�@�@�F
' �����@�@�@�Fdir      �f�B���N�g���p�X
' �@�@�@�@�@�@filePath �t�@�C���p�X
' �߂�l�@�@�F�A����̕�����
'
' =========================================================
Public Function concatFilePath(ByVal dir As String, ByVal fileName As String) As String

    ' ������̍Ō���� "\" ���t���Ă��邩���m�F����
    If InStrRev(dir, "\") = Len(dir) Then
    
        concatFilePath = dir & fileName
    Else
    
        concatFilePath = dir & "\" & fileName
    End If
    
End Function

Public Function convertKeyCodeToKeyAscii(ByVal keyCode As Long) As String

    If vbKey0 = keyCode Then
        convertKeyCodeToKeyAscii = "0"
    ElseIf vbKey1 = keyCode Then convertKeyCodeToKeyAscii = "1"
    ElseIf vbKey2 = keyCode Then convertKeyCodeToKeyAscii = "2"
    ElseIf vbKey3 = keyCode Then convertKeyCodeToKeyAscii = "3"
    ElseIf vbKey4 = keyCode Then convertKeyCodeToKeyAscii = "4"
    ElseIf vbKey5 = keyCode Then convertKeyCodeToKeyAscii = "5"
    ElseIf vbKey6 = keyCode Then convertKeyCodeToKeyAscii = "6"
    ElseIf vbKey7 = keyCode Then convertKeyCodeToKeyAscii = "7"
    ElseIf vbKey8 = keyCode Then convertKeyCodeToKeyAscii = "8"
    ElseIf vbKey9 = keyCode Then convertKeyCodeToKeyAscii = "9"
    ElseIf vbKeyA = keyCode Then convertKeyCodeToKeyAscii = "A"
    ElseIf vbKeyB = keyCode Then convertKeyCodeToKeyAscii = "B"
    ElseIf vbKeyC = keyCode Then convertKeyCodeToKeyAscii = "C"
    ElseIf vbKeyD = keyCode Then convertKeyCodeToKeyAscii = "D"
    ElseIf vbKeyE = keyCode Then convertKeyCodeToKeyAscii = "E"
    ElseIf vbKeyF = keyCode Then convertKeyCodeToKeyAscii = "F"
    ElseIf vbKeyG = keyCode Then convertKeyCodeToKeyAscii = "G"
    ElseIf vbKeyH = keyCode Then convertKeyCodeToKeyAscii = "H"
    ElseIf vbKeyI = keyCode Then convertKeyCodeToKeyAscii = "I"
    ElseIf vbKeyJ = keyCode Then convertKeyCodeToKeyAscii = "J"
    ElseIf vbKeyK = keyCode Then convertKeyCodeToKeyAscii = "K"
    ElseIf vbKeyL = keyCode Then convertKeyCodeToKeyAscii = "L"
    ElseIf vbKeyM = keyCode Then convertKeyCodeToKeyAscii = "M"
    ElseIf vbKeyN = keyCode Then convertKeyCodeToKeyAscii = "N"
    ElseIf vbKeyO = keyCode Then convertKeyCodeToKeyAscii = "O"
    ElseIf vbKeyP = keyCode Then convertKeyCodeToKeyAscii = "P"
    ElseIf vbKeyQ = keyCode Then convertKeyCodeToKeyAscii = "Q"
    ElseIf vbKeyR = keyCode Then convertKeyCodeToKeyAscii = "R"
    ElseIf vbKeyS = keyCode Then convertKeyCodeToKeyAscii = "S"
    ElseIf vbKeyT = keyCode Then convertKeyCodeToKeyAscii = "T"
    ElseIf vbKeyU = keyCode Then convertKeyCodeToKeyAscii = "U"
    ElseIf vbKeyV = keyCode Then convertKeyCodeToKeyAscii = "V"
    ElseIf vbKeyW = keyCode Then convertKeyCodeToKeyAscii = "W"
    ElseIf vbKeyX = keyCode Then convertKeyCodeToKeyAscii = "X"
    ElseIf vbKeyY = keyCode Then convertKeyCodeToKeyAscii = "Y"
    ElseIf vbKeyZ = keyCode Then convertKeyCodeToKeyAscii = "Z"
    End If

End Function

' =========================================================
' ���|�C���g����s�N�Z���ɒP�ʂ�ϊ�����
'
' �T�v�@�@�@�F
' �����@�@�@�Fd     DPI
' �@�@�@�@�@�@pixel �s�N�Z��
' �߂�l�@�@�F�|�C���g
'
' =========================================================
Public Function convertPixelToPoint(ByVal d As Long, ByVal pixel As Long) As Single

    convertPixelToPoint = CSng(pixel) / d * 72

End Function

' =========================================================
' ���s�N�Z������|�C���g�ɒP�ʂ�ϊ�����
'
' �T�v�@�@�@�F
' �����@�@�@�Fd     DPI
' �@�@�@�@�@�@pixel �s�N�Z��
' �߂�l�@�@�F�|�C���g
'
' =========================================================
Public Function convertPointToPixel(ByVal d As Long, ByVal point As Single) As Long

    convertPointToPixel = point * d / 72
    
End Function

' =========================================================
' �����S���W���v�Z����
'
' �T�v�@�@�@�F�v�Z��̍��W���Adx�Edy�Ɋi�[�����
' �����@�@�@�Fsx ��ƂȂ��` ���WX
' �@�@�@�@�@�@sy ��ƂȂ��` ���WY
' �@�@�@�@�@�@sw ��ƂȂ��` ��
' �@�@�@�@�@�@sh ��ƂȂ��` ����
' �@�@�@�@�@�@dx ��r�����` ���WX
' �@�@�@�@�@�@dy ��r�����` ���WY
' �@�@�@�@�@�@dw ��r�����` ��
' �@�@�@�@�@�@dh ��r�����` ����
'
' =========================================================
Public Sub calcCenterPoint( _
                           ByVal sx As Single _
                         , ByVal sy As Single _
                         , ByVal sw As Single _
                         , ByVal sh As Single _
                         , ByRef dx As Single _
                         , ByRef dy As Single _
                         , ByVal dw As Single _
                         , ByVal dh As Single)

    ' ���S���v�Z����
    Dim newX As Single
    Dim newY As Single
    
    newX = sw / 2 - dw / 2 + sx
    newY = sh / 2 - dh / 2 + sy

    ' ���S��ݒ肷��
    dx = newX
    dy = newY

End Sub

' =========================================================
' ����`A��B���r��A��B���Ɏ��܂��Ă��邩���m�F����
'
' �T�v�@�@�@�F
' �����@�@�@�Fsx ��ƂȂ��` ���WX
' �@�@�@�@�@�@sy ��ƂȂ��` ���WY
' �@�@�@�@�@�@sw ��ƂȂ��` ��
' �@�@�@�@�@�@sh ��ƂȂ��` ����
' �@�@�@�@�@�@dx ��r�����` ���WX
' �@�@�@�@�@�@dy ��r�����` ���WY
' �@�@�@�@�@�@dw ��r�����` ��
' �@�@�@�@�@�@dh ��r�����` ����
' �߂�l�@�@�FTrue ��`A���Ɏ��܂��Ă���ꍇ
'
' =========================================================
Public Function isInnerScreen( _
                           ByVal sx As Single _
                         , ByVal sy As Single _
                         , ByVal sw As Single _
                         , ByVal sh As Single _
                         , ByRef dx As Single _
                         , ByRef dy As Single _
                         , ByRef dw As Single _
                         , ByRef dh As Single) As Boolean

    isInnerScreen = True

    ' �g���͂ݏo���Ă��Ȃ������m�F����
    If sx > dx Then
    
        isInnerScreen = False
        
    ElseIf sy > dy Then
    
        isInnerScreen = False
        
    ElseIf (sx + sw) < (dx + dw) Then
    
        isInnerScreen = False
        
    ElseIf (sy + sh) < (dy + dh) Then
    
        isInnerScreen = False
        
    End If

End Function

' =========================================================
' ���p�f�B���O�֐�
'
' �T�v�@�@�@�F������̍����ɓ���̕�����C�ӂ̌����ɂȂ�悤�ɋl�߂�
' �����@�@�@�Fvalue  �l
' �@�@�@�@�@�@length ����
' �@�@�@�@�@�@char   ����
' �߂�l�@�@�F�p�f�B���O����
'
' =========================================================
Public Function padLeft(ByVal value As String _
                      , ByVal length As Long _
                      , Optional ByVal char As String = "0") As String

    ' �p�f�B���O���錅��
    Dim padLen As Long
    padLen = length - Len(value)
    
    If padLen < 1 Then
    
        padLeft = value
        Exit Function
    End If

    padLeft = String(length - Len(value), char) & value

End Function

' =========================================================
' ���p�f�B���O�֐�
'
' �T�v�@�@�@�F������̉E���ɓ���̕�����C�ӂ̌����ɂȂ�悤�ɋl�߂�
' �����@�@�@�Fvalue  �l
' �@�@�@�@�@�@length ����
' �@�@�@�@�@�@char   ����
' �߂�l�@�@�F�p�f�B���O����
'
' =========================================================
Public Function padRight(ByVal value As String _
                      , ByVal length As Long _
                      , Optional ByVal char As String = "0") As String

    ' �p�f�B���O���錅��
    Dim padLen As Long
    padLen = length - Len(value)
    
    If padLen < 1 Then
    
        padRight = value
        Exit Function
    End If

    padRight = value & String(length - Len(value), char)

End Function

' =========================================================
' ���f�B���N�g�����쐬����
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�FTrue �f�B���N�g���쐬����True��ԋp
'
' =========================================================
Public Function createDir(ByVal filePath As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim rc As Long
    
    rc = WinAPI_Shell.SHCreateDirectoryEx(0&, filePath, 0&)
    
    If rc = 0 Then
        createDir = True
    Else
        createDir = False
    End If
        
End Function

' =========================================================
' ���t�@�C�����폜����
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�FTrue �t�@�C���폜����True��ԋp
'
' =========================================================
Public Function deleteFile(ByVal filePath As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.fileExists(filePath) = True Then
        fso.deleteFile filePath, True
        deleteFile = True
    End If

    deleteFile = False
        
End Function

' =========================================================
' ��One�h���C�u��URL���t�@�C���p�X�ɕϊ�����B
'
' �T�v�@�@�@�F
' �����@�@�@�FfullPath$ �t���p�X
' �߂�l�@�@�F�t�@�C���p�X
'
' =========================================================
Public Function convertOneDriveUrlToLocalFilePath(ByVal fullPath$)
    'Finds local path for a OneDrive file URL, using environment variables of OneDrive
    'Reference https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive
    'Authors: Philip Swannell 2019-01-14, MatChrupczalski 2019-05-19, Horoman 2020-03-29, P.G.Schild 2020-04-02

    Dim ii&
    Dim iPos&
    Dim oneDrivePath$
    Dim endFilePath$

    If Left(fullPath, 8) = "https://" Then 'Possibly a OneDrive URL
        If InStr(1, fullPath, "my.sharepoint.com") <> 0 Then 'Commercial OneDrive
            'For commercial OneDrive, path looks like "https://companyName-my.sharepoint.com/personal/userName_domain_com/Documents" & file.FullName)
            'Find "/Documents" in string and replace everything before the end with OneDrive local path
            iPos = InStr(1, fullPath, "/Documents") + Len("/Documents") 'find "/Documents" position in file URL
            endFilePath = Mid(fullPath, iPos) 'Get the ending file path without pointer in OneDrive. Include leading "/"
        Else 'Personal OneDrive
            'For personal OneDrive, path looks like "https://d.docs.live.net/d7bbaa#######1/" & file.FullName
            'We can get local file path by replacing "https.." up to the 4th slash, with the OneDrive local path obtained from registry
            iPos = 8 'Last slash in https://
            For ii = 1 To 2
                iPos = InStr(iPos + 1, fullPath, "/") 'find 4th slash
            Next ii
            endFilePath = Mid(fullPath, iPos) 'Get the ending file path without OneDrive root. Include leading "/"
        End If
        endFilePath = replace(endFilePath, "/", Application.PathSeparator) 'Replace forward slashes with back slashes (URL type to Windows type)
        For ii = 1 To 3 'Loop to see if the tentative LocalWorkbookName is the name of a file that actually exists, if so return the name
            oneDrivePath = Environ(Choose(ii, "OneDriveCommercial", "OneDriveConsumer", "OneDrive")) 'Check possible local paths. "OneDrive" should be the last one
            If 0 < Len(oneDrivePath) Then
                convertOneDriveUrlToLocalFilePath = oneDrivePath & endFilePath
                Exit Function 'Success (i.e. found the correct Environ parameter)
            End If
        Next ii
        'Possibly raise an error here when attempt to convert to a local file name fails - e.g. for "shared with me" files
        convertOneDriveUrlToLocalFilePath = vbNullString
    Else
        convertOneDriveUrlToLocalFilePath = fullPath
    End If
End Function

