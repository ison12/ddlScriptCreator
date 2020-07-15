Attribute VB_Name = "StringConverter"
Option Explicit

' *********************************************************
' �����ϊ����W���[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/07/14�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

Public Function ConvertSnakeToCamel(ByVal val As String _
                                  , Optional isFirstUpper As Boolean = True) As String

    Dim i As Long
    
    ' �A���_�[�X�R�A����؂蕶���Ƃ��ĕ����𕪊�����
    Dim tmp() As String
    tmp = Split(val, "_")

    ' ��؂����������A������
    For i = LBound(tmp) To UBound(tmp)
        
        ConvertSnakeToCamel = ConvertSnakeToCamel & UCase$(Mid$(tmp(i), 1, 1)) & LCase(Mid$(tmp(i), 2))
    Next
    
    If Not isFirstUpper Then
    
        ' �擪�̕������������ɂ���
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

    ' ��؂����������A������
    For i = 1 To length
    
        ' 1���������o��
        char = Mid$(val, i, 1)
        
        ' �啶�����𔻒肷��
        If i <> 1 And Asc("A") <= Asc(char) And Asc(char) <= Asc("Z") Then
        
            ' �啶���̏ꍇ�A�P��̋�؂�Ƃ��ăA���_�[�X�R�A��t�^����
            ConvertCamelToSnake = ConvertCamelToSnake & "_" & char
        Else
        
            ' �������̏ꍇ�A�P��̋�؂�Ƃ��ăA���_�[�X�R�A��t�^����
            ConvertCamelToSnake = ConvertCamelToSnake & char
        End If
        
    Next
    
    If isUpper Then
        ConvertCamelToSnake = UCase$(ConvertCamelToSnake)
    Else
        ConvertCamelToSnake = LCase$(ConvertCamelToSnake)
    End If

End Function




