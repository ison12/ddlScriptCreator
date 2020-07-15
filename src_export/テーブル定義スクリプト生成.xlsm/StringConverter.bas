Attribute VB_Name = "StringConverter"
Option Explicit

' *********************************************************
' 文字変換モジュール
'
' 作成者　：Ison
' 履歴　　：2020/07/14　新規作成
'
' 特記事項：
'
' *********************************************************

Public Function ConvertSnakeToCamel(ByVal val As String _
                                  , Optional isFirstUpper As Boolean = True) As String

    Dim i As Long
    
    ' アンダースコアを区切り文字として文字を分割する
    Dim tmp() As String
    tmp = Split(val, "_")

    ' 区切った文字列を連結する
    For i = LBound(tmp) To UBound(tmp)
        
        ConvertSnakeToCamel = ConvertSnakeToCamel & UCase$(Mid$(tmp(i), 1, 1)) & LCase(Mid$(tmp(i), 2))
    Next
    
    If Not isFirstUpper Then
    
        ' 先頭の文字を小文字にする
        If Len(ConvertSnakeToCamel) > 0 Then
        
            ConvertSnakeToCamel = LCase$(Mid$(ConvertSnakeToCamel, 1, 1)) & Mid$(ConvertSnakeToCamel, 2)
        End If
    End If

End Function

Public Function ConvertCamelToSnake(ByVal val As String _
                                  , Optional isUpper As Boolean = True) As String

    Dim i      As Long
    Dim length As Long
    
    Dim char As String
    
    length = Len(val)

    ' 区切った文字列を連結する
    For i = 1 To length
    
        ' 1文字抜き出す
        char = Mid$(val, i, 1)
        
        ' 大文字かを判定する
        If i <> 1 And Asc("A") <= Asc(char) And Asc(char) <= Asc("Z") Then
        
            ' 大文字の場合、単語の区切りとしてアンダースコアを付与する
            ConvertCamelToSnake = ConvertCamelToSnake & "_" & char
        Else
        
            ' 小文字の場合、単語の区切りとしてアンダースコアを付与する
            ConvertCamelToSnake = ConvertCamelToSnake & char
        End If
        
    Next
    
    If isUpper Then
        ConvertCamelToSnake = UCase$(ConvertCamelToSnake)
    Else
        ConvertCamelToSnake = LCase$(ConvertCamelToSnake)
    End If

End Function




