Attribute VB_Name = "strings"
Option Explicit

'������̕���
Sub splitStr()
    Dim ret() As String
    Dim i As Integer
    ret = Split("�x�m�R,3776", ",")
    
    For i = 0 To UBound(ret)
        Cells(1, i + 1).Value = ret(i)
    Next i
End Sub

'������̌���
Sub concat()
    Cells(2, 1).Value = Cells(1, 1).Value & "/" & Cells(1, 2).Value
    
    'worksheet�֐�concat/concatenate���g����
    Cells(2, 2) = WorksheetFunction.concat(Cells(1, 1).Value, ",", Cells(1, 2).Value)
End Sub

'�z��̕����������
Sub concat2()
    Dim arr() As Variant
    Dim ret As String
    arr = Array("�x�m", "�R�[��", "�I�E������")
    ret = Join(arr, "")
End Sub

Sub str2dec()
    Dim ret As Integer
    ret = Val("1234")
    Cells(3, 1) = ret
End Sub

Sub dec2str()
    Dim str As String
    str = CStr(1234)
    Cells(4, 1) = str
End Sub

Sub str2hex()
    Dim ret As Long
    ret = Val("&H" & "E0")
    Cells(5, 1) = ret
End Sub

Sub hex2str()
    Dim str As String
    str = CStr(1234)
    Cells(6, 1) = str
End Sub

'int32�^
'worksheet�֐�HEX2DEC��uint32?�܂őΉ��\
Function hex32dec(hexString As String) As Long
    Dim result As Long
    Dim i As Integer
    
    '�擪��"0x"������ꍇ�͍폜
    If Left(hexString, 2) = "0x" Or Left(hexString, 2) = "0X" Then
        hexString = Right(hexString, Len(hexString) - 2)
    End If
    
    '32�r�b�g��16�i���������10�i���ɕϊ�
    For i = Len(hexString) To 1 Step -1
        result = result + (16 ^ (Len(hexString) - i) * HexToDec(Mid(hexString, i, 1)))
    Next i
    
    hex32dec = result
    End Function

'int16�^
Function HexToDec(hexChar As String) As Long
    HexToDec = Val("&H" & hexChar)
End Function

'worksheet�֐�DEC3HEX��uint32?�܂őΉ��\
Function dec2hex(decNumber As Long) As String
    dec2hex = Hex(decNumber)
End Function
