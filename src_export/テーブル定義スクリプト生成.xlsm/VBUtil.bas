Attribute VB_Name = "VBUtil"
Option Explicit

' *********************************************************
' VB関連の共通関数モジュール
'
' 作成者　：Ison
' 履歴　　：2008/08/10　新規作成
'
' 特記事項：
'
' *********************************************************

' エラー情報を格納する構造体
Public Type errInfo

    Source       As Variant
    Number       As Variant
    Description  As Variant
    LastDllError As Variant
    HelpFile     As Variant
    HelpContext  As Variant
    
End Type

' =========================================================
' ▽Errオブジェクトの情報を構造体に退避
'
' 概要　　　：Errオブジェクトの情報を構造体に設定して返す。
' 引数　　　：
' 戻り値　　：エラー情報
'
' 特記事項　：エラーハンドラで別の関数を呼び出すとErrオブジェクトの情報が消えてしまうことがあり
' 　　　　　　この状態で、Err.Raiseすると正しい情報を上位のモジュールにで伝播できない。
' 　　　　　　正しい情報を伝播する場合には、本関数を利用して、一度エラー情報を退避してからErr.Raiseしてやると良い。
'
' 　　　　　　使用例：
' 　　　　　　　Dim errT As errInfo
' 　　　　　　　errT = VBUtil.swapErr

' 　　　　　　　・・・エラー時の後始末処理など
'
' 　　　　　　　Err.Raise errT.Number, errT.Source・・・
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
' ▽保存ダイアログ表示
'
' 概要　　　：保存ダイアログを表示する
' 引数　　　：title           ダイアログのタイトル
' 　　　　　　filter          フィルタ
' 　　　　　　initialFileName 初期ファイル名
' 戻り値　　：保存ファイルパス
'
' =========================================================
Public Function openFileSaveDialog(ByVal title As String, ByVal filter As String, ByVal initialFileName As String) As String

    ' アプリケーション
    Dim xlsApp   As Application
    
    ' ファイルパス
    Dim filePath As Variant

    ' Applicationオブジェクト取得
    Set xlsApp = Application
    
    ' ダイアログで選択されたファイル名を格納
    filePath = xlsApp.GetSaveAsFilename(initialFileName:=initialFileName _
                                      , fileFilter:=filter _
                                      , title:=title)
                                      
    ' キャンセルされたかを判定する
    If filePath = False Then
    
        ' キャンセルされた場合 空文字列を返す
        openFileSaveDialog = ""
        
    Else
        ' 保存を選択された場合 ファイル名を返す
        openFileSaveDialog = filePath
    End If

End Function

' =========================================================
' ▽開くダイアログ表示
'
' 概要　　　：開くダイアログを表示する
' 引数　　　：title           ダイアログのタイトル
' 　　　　　　filter          フィルタ
' 　　　　　　multiSelect     複数選択
' 戻り値　　：選択したファイルのファイルパス
'
' =========================================================
Public Function openFileDialog(ByVal title As String, ByVal filter As String, Optional ByVal multiSelect As Boolean = False) As Variant

    ' アプリケーション
    Dim xlsApp   As Application
    
    ' ファイルパス
    Dim filePath As Variant

    ' Applicationオブジェクト取得
    Set xlsApp = Application
    
    ' ダイアログで選択されたファイル名を格納
    filePath = xlsApp.GetOpenFilename(fileFilter:=filter _
                                    , title:=title _
                                    , multiSelect:=multiSelect)

    ' 複数選択の場合、戻り値として配列が返されるので配列かどうかを判定する
    If IsArray(filePath) Then
    
        ' 保存を選択された場合 ファイル名を返す
        openFileDialog = filePath
    
    ' 選択がキャンセルされた場合
    ElseIf filePath = False Then
    
        ' キャンセルされた場合 空を返す
        openFileDialog = Empty
        
    Else
        ' 保存を選択された場合 ファイル名を返す
        openFileDialog = filePath
    
    End If

End Function

' =========================================================
' ▽ファイルの拡張子チェック
'
' 概要　　　：ファイルの拡張子をチェックする
' 引数　　　：file      ファイル名
' 　　　　　　extension 拡張子
' 戻り値　　：ファイルの拡張子が指定された引数extensionの場合Trueを返す
'
' =========================================================
Public Function checkFileExtension(ByRef file As String _
                                 , ByRef extension As String) As Boolean

    ' ファイル名から抽出した拡張子
    Dim fileExtension As String
    
    ' インデックス
    Dim index As Long
    
    ' ファイル名と拡張子の区切り文字であるドット(.)を検索する
    index = InStrRev(file, ".")
    
    ' ドット(.)が見つからない場合
    If index <= 0 Then
    
        Exit Function
    End If
    
    ' ファイル名から拡張子を抽出する
    fileExtension = Mid$(file, index + 1, Len(file))

    If fileExtension = extension Then
    
        checkFileExtension = True
    Else
    
        checkFileExtension = False
    End If

End Function

' =========================================================
' ▽ファイルパスからファイル名抽出
'
' 概要　　　：ファイルパスからファイル名を抽出する
' 引数　　　：filePath ファイルパス
' 戻り値　　：ファイル名
'
' =========================================================
Public Function extractFileName(ByRef filePath As String) As String
    
    ' ファイルパス区切り文字
    Const FILE_SEPARATE As String = "\"

    ' ファイルパスの右後方からはじめに出現した区切り文字の文字位置
    Dim index As Long
    
    ' 区切り文字の位置を取得する
    index = InStrRev(filePath, FILE_SEPARATE)

    ' 区切り文字を発見した場合
    If index > 0 Then
    
        extractFileName = Mid$(filePath, index + 1)
    
    ' 区切り文字を発見できなかった場合
    Else
        extractFileName = filePath
    
    End If

End Function

' =========================================================
' ▽インフォメッセージボックスを表示
'
' 概要　　　：インフォメッセージボックスを表示する
' 引数　　　：basePrompt 基本メッセージ
'             title      メッセージボックスのタイトル
' 　　　　　　err        エラーオブジェクト
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
' ▽エラーメッセージボックスを表示
'
' 概要　　　：エラーメッセージボックスを表示する
' 引数　　　：basePrompt 基本メッセージ
'             title      メッセージボックスのタイトル
' 　　　　　　err        エラーオブジェクト
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
' ▽警告メッセージボックスを表示
'
' 概要　　　：警告メッセージボックスを表示する
' 引数　　　：basePrompt 基本メッセージ
'             title      メッセージボックスのタイトル
' 　　　　　　err        エラーオブジェクト
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
' ▽配列サイズ取得
'
' 概要　　　：配列のサイズを取得する
' 引数　　　：var       配列
' 　　　　　　dimension 次元
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
' ▽2次元配列の任意の行を1次元配列として返す
'
' 概要　　　：
' 引数　　　：val 配列
'             i   配列のインデックス
'
' =========================================================
Public Function convert2to1Array(ByRef val As Variant, ByVal i As Long) As Variant

    ' 戻り値
    Dim ret() As Variant

    Dim j As Long
    
    ReDim ret(LBound(val, 2) To UBound(val, 2))
    
    For j = LBound(ret) To UBound(ret)
    
        ret(j) = val(i, j)
    
    Next
    
    convert2to1Array = ret

End Function

' =========================================================
' ▽2次元配列をデバッグウィンドウに出力する
'
' 概要　　　：
' 引数　　　：val 配列
'
' =========================================================
Public Function debugPrintArray(ByRef val As Variant)

    ' 配列のインデックス
    Dim i As Long
    Dim j As Long
    
    ' デバッグウィンドウに出力する文字列
    Dim str As String
    
    str = "Output Array" & vbNewLine
    
    ' -------------------------------------------------
    ' 配列として初期化されている場合に出力を実施する
    ' -------------------------------------------------
    If VarType(val) = (vbArray + vbVariant) Then
    
        ' ループ処理
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
' ▽2次元配列の要素入れ替え
'
' 概要　　　：2次元配列の要素を(x,y)から(y,x)に設定しなおす。
' 引数　　　：v 2次元配列
'
' 戻り値　　：2次元配列
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
' ▽整数チェック
'
' 概要　　　：
' 引数　　　：value チェック文字列
' 戻り値　　：True 整数
'
' =========================================================
Public Function validInteger(ByVal value As String) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = False

    ' チェック対象が数値で且つ、小数点を含まない場合、OKとする
    If _
            IsNumeric(value) = True _
        And InStr(value, ".") = 0 Then
    
        ret = True
    
    End If

    ' 戻り値を返す
    validInteger = ret

End Function

' =========================================================
' ▽整数チェック（負数は含まない）
'
' 概要　　　：
' 引数　　　：value チェック文字列
' 戻り値　　：True 整数
'
' =========================================================
Public Function validUnsignedInteger(ByVal value As String) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = False

    ' チェック対象が数値で且つ、マイナス記号を含まず小数点を含まない場合、OKとする
    If _
            IsNumeric(value) = True _
        And InStr(value, ".") = 0 _
        And InStr(value, "-") = 0 _
    Then
    
        ret = True
    
    End If

    ' 戻り値を返す
    validUnsignedInteger = ret

End Function

' =========================================================
' ▽16進数チェック
'
' 概要　　　：
' 引数　　　：value チェック文字列
' 戻り値　　：True 16進数
'
' =========================================================
Public Function validHex(ByVal value As String) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = True

    ' インデックス
    Dim i    As Long
    ' 文字のサイズ
    Dim size As Long
    
    ' 文字列の1文字分
    Dim one    As String
    ' 1文字分のASCIIコード
    Dim oneAsc As Long
    
    ' 文字のサイズを取得する
    size = Len(value)
    
    ' 文字列から1文字ずつ取り出しループを実行する
    For i = 1 To size
    
        ' 1文字取り出す
        one = Mid$(value, i, 1)
        ' 取り出した文字のASCIIコードを調べる
        oneAsc = Asc(one)
        
        ' 文字列が以下の範囲内であるかを確認する
        ' 0-9 a-f A-F
        If _
             (65 <= oneAsc And oneAsc <= 70) _
          Or (97 <= oneAsc And oneAsc <= 102) _
          Or (48 <= oneAsc And oneAsc <= 57) Then
        
            ' 正常
            
        Else
        
            ' エラー時
            ret = False
            Exit For
        
        End If
        
    Next

    ' 戻り値を返す
    validHex = ret

End Function

' =========================================================
' ▽数値であるかをチェックする
'
' 概要　　　：
' 引数　　　：value チェック文字列
' 戻り値　　：True 整数
'
' =========================================================
Public Function validNumeric(ByVal value As String) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = False

    ' チェック対象が数値の場合、OKとする
    If _
            IsNumeric(value) = True Then
    
        ret = True
    
    End If

    ' 戻り値を返す
    validNumeric = ret

End Function

' =========================================================
' ▽数値であるかをチェックする（負数は含まない）
'
' 概要　　　：
' 引数　　　：value チェック文字列
' 戻り値　　：True 整数
'
' =========================================================
Public Function validUnsignedNumeric(ByVal value As String) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = False

    ' チェック対象が数値で且つマイナス記号を含まない場合、OKとする
    If _
            IsNumeric(value) = True _
        And InStr(value, "-") = 0 _
    Then
    
        ret = True
    
    End If

    ' 戻り値を返す
    validUnsignedNumeric = ret

End Function

' =========================================================
' ▽コード値チェック
'
' 概要　　　：引数で与えられたコードリストに一致するものがあるかをチェックする。
' 引数　　　：value    チェック文字列
' 　　　　　　codeList コードリスト
' 戻り値　　：True コードリストに一致する値がある
'
' =========================================================
Public Function validCode(ByVal value As String, ParamArray codeList() As Variant) As Boolean

    ' チェック対象が空の場合、OKとする
    Dim i As Long
    
    ' valueがenumsの何れかの値と一致しているかどうかを確認する
    For i = LBound(codeList) To UBound(codeList)
    
        ' 一致している場合
        If value = CStr(codeList(i)) Then
        
            ' Trueを返す
            validCode = True
            
            Exit Function
        End If
    
    Next
    
    ' 一致するものがなかったので、Falseを返す
    validCode = False

End Function

' =========================================================
' ▽RGB反転
'
' 概要　　　：RGBを反転させる。
' 引数　　　：r 赤
' 　　　　　　g 緑
' 　　　　　　b 青
' 戻り値　　：反転色
'
' =========================================================
Public Function reverseRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long

    reverseRGB = (Not RGB(r, g, b)) And &HFFFFFF

End Function

' =========================================================
' ▽NULL→空文字列変換
'
' 概要　　　：Nullを空文字列に変換する。
' 引数　　　：value VARIANTデータ
' 戻り値　　：空文字列
' 特記事項　：Null 値は、データ アイテム に有効なデータが
' 　　　　　　格納されていないことを示すのに使用されるバリアント型 (Variant) の内部処理形式です。
'
' =========================================================
Public Function convertNullToEmptyStr(ByRef value As Variant) As String

    ' NULLの場合
    If IsNull(value) = True Then
    
        ' 空文字列に変換
        convertNullToEmptyStr = ""
        
    ' 配列の場合
    ElseIf IsArray(value) Then
    
        ' 空文字列に変換
        convertNullToEmptyStr = ""
        
    ' その他
    Else
    
        ' 文字列に変換して格納する
        convertNullToEmptyStr = CStr(value)
    End If
    
End Function

' =========================================================
' ▽ファイルが存在するかをチェックする
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
' 戻り値　　：True ファイルが存在する場合
'
' =========================================================
Public Function isExistFile(ByVal filePath As String) As Boolean

    ' ファイル名
    Dim fileName As String

    ' 指定のファイルパスが存在するかどうかをチェックする
    fileName = dir(filePath, vbNormal)
    
    ' ファイル名が取得できたかどうかをチェックする
    If fileName = "" Then
    
        ' ファイルが存在しない場合
        isExistFile = False
    Else
        
        ' ファイルが存在する場合
        isExistFile = True
    End If

End Function

' =========================================================
' ▽ファイルが存在するかをチェックする
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
' 戻り値　　：True ファイルが存在する場合
'
' =========================================================
Public Function isExistDirectory(ByVal filePath As String) As Boolean

    ' ディレクトリパス
    Dim dirPath As String
    ' ファイル名
    Dim fileName As String

    ' ファイルパスからディレクトリを抽出する
    dirPath = VBUtil.extractDirPathFromFilePath(filePath)
    
    ' 指定のディレクトリが存在するかを確認する
    fileName = dir(dirPath, vbDirectory)
    
    ' ファイル名が取得できたかどうかをチェックする
    If fileName = "" Then
    
        ' ファイルが存在しない場合
        isExistDirectory = False
    Else
        
        ' ファイルが存在する場合
        isExistDirectory = True
    End If

End Function

' =========================================================
' ▽ファイルパスからディレクトリパスを抽出する
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
' 戻り値　　：ディレクトリパス
'
' =========================================================
Public Function extractDirPathFromFilePath(filePath As String) As String

    ' 戻り値
    Dim ret As String
    
    ' ディレクトリ位置
    Dim dirPoint As Long

    ' 文字列の右端から"\"を検索し、左端からの位置を取得する
    dirPoint = InStrRev(filePath, "\")
    
    ' "\"が見つからない場合
    If dirPoint <> 0 Then
    
        ' ディレクトリパスの取得
        ret = Left$(filePath, dirPoint - 1)
        
        extractDirPathFromFilePath = ret
    
    Else
        extractDirPathFromFilePath = ""
    
    End If
    
End Function

' =========================================================
' ▽ファイルパスからファイル名を抽出する
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
' 戻り値　　：ディレクトリパス
'
' =========================================================
Public Function extractFileNameFromFilePath(filePath As String) As String

    ' 戻り値
    Dim ret As String
    
    ' ディレクトリ位置
    Dim dirPoint As Long

    ' 文字列の右端から"\"を検索し、左端からの位置を取得する
    dirPoint = InStrRev(filePath, "\")
    
    ' "\"が見つかった場合
    If dirPoint <> 0 Then
    
        ' ディレクトリパスの取得
        ret = right$(filePath, Len(filePath) - dirPoint)
        
        extractFileNameFromFilePath = ret
    
    Else
    
        extractFileNameFromFilePath = filePath
    End If
    
End Function

' =========================================================
' ▽ディレクトリパスとファイルパスを連結する
'
' 概要　　　：
' 引数　　　：dir      ディレクトリパス
' 　　　　　　filePath ファイルパス
' 戻り値　　：連結後の文字列
'
' =========================================================
Public Function concatFilePath(ByVal dir As String, ByVal fileName As String) As String

    ' 文字列の最後尾に "\" が付いているかを確認する
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
' ▽ポイントからピクセルに単位を変換する
'
' 概要　　　：
' 引数　　　：d     DPI
' 　　　　　　pixel ピクセル
' 戻り値　　：ポイント
'
' =========================================================
Public Function convertPixelToPoint(ByVal d As Long, ByVal pixel As Long) As Single

    convertPixelToPoint = CSng(pixel) / d * 72

End Function

' =========================================================
' ▽ピクセルからポイントに単位を変換する
'
' 概要　　　：
' 引数　　　：d     DPI
' 　　　　　　pixel ピクセル
' 戻り値　　：ポイント
'
' =========================================================
Public Function convertPointToPixel(ByVal d As Long, ByVal point As Single) As Long

    convertPointToPixel = point * d / 72
    
End Function

' =========================================================
' ▽中心座標を計算する
'
' 概要　　　：計算後の座標が、dx・dyに格納される
' 引数　　　：sx 基準となる矩形 座標X
' 　　　　　　sy 基準となる矩形 座標Y
' 　　　　　　sw 基準となる矩形 幅
' 　　　　　　sh 基準となる矩形 高さ
' 　　　　　　dx 比較する矩形 座標X
' 　　　　　　dy 比較する矩形 座標Y
' 　　　　　　dw 比較する矩形 幅
' 　　　　　　dh 比較する矩形 高さ
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

    ' 中心を計算する
    Dim newX As Single
    Dim newY As Single
    
    newX = sw / 2 - dw / 2 + sx
    newY = sh / 2 - dh / 2 + sy

    ' 中心を設定する
    dx = newX
    dy = newY

End Sub

' =========================================================
' ▽矩形AとBを比較しAがB内に収まっているかを確認する
'
' 概要　　　：
' 引数　　　：sx 基準となる矩形 座標X
' 　　　　　　sy 基準となる矩形 座標Y
' 　　　　　　sw 基準となる矩形 幅
' 　　　　　　sh 基準となる矩形 高さ
' 　　　　　　dx 比較する矩形 座標X
' 　　　　　　dy 比較する矩形 座標Y
' 　　　　　　dw 比較する矩形 幅
' 　　　　　　dh 比較する矩形 高さ
' 戻り値　　：True 矩形A内に収まっている場合
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

    ' 枠をはみ出していないかを確認する
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
' ▽パディング関数
'
' 概要　　　：文字列の左側に特定の文字を任意の桁数になるように詰める
' 引数　　　：value  値
' 　　　　　　length 桁数
' 　　　　　　char   文字
' 戻り値　　：パディング結果
'
' =========================================================
Public Function padLeft(ByVal value As String _
                      , ByVal length As Long _
                      , Optional ByVal char As String = "0") As String

    ' パディングする桁数
    Dim padLen As Long
    padLen = length - Len(value)
    
    If padLen < 1 Then
    
        padLeft = value
        Exit Function
    End If

    padLeft = String(length - Len(value), char) & value

End Function

' =========================================================
' ▽パディング関数
'
' 概要　　　：文字列の右側に特定の文字を任意の桁数になるように詰める
' 引数　　　：value  値
' 　　　　　　length 桁数
' 　　　　　　char   文字
' 戻り値　　：パディング結果
'
' =========================================================
Public Function padRight(ByVal value As String _
                      , ByVal length As Long _
                      , Optional ByVal char As String = "0") As String

    ' パディングする桁数
    Dim padLen As Long
    padLen = length - Len(value)
    
    If padLen < 1 Then
    
        padRight = value
        Exit Function
    End If

    padRight = value & String(length - Len(value), char)

End Function

' =========================================================
' ▽ディレクトリを作成する
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
' 戻り値　　：True ディレクトリ作成時はTrueを返却
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
' ▽ファイルを削除する
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
' 戻り値　　：True ファイル削除時はTrueを返却
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
' ▽OneドライブのURLをファイルパスに変換する。
'
' 概要　　　：
' 引数　　　：fullPath$ フルパス
' 戻り値　　：ファイルパス
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

