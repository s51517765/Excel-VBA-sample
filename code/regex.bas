Attribute VB_Name = "regex"
Option Explicit
'�u�c�[���v-�u�Q�Ɛݒ�v�� �uMicrosoft VBScript Regular Expressions 5.5�v��L���ɂ���

Sub getIntegerNum()
    Dim re As New RegExp
    Dim mc As MatchCollection
    Dim m As Match
    Dim i As Long
    re.Pattern = "(\d+)"
    re.Global = True
    Set mc = re.Execute("123����������4567����������890.1����������1e20�����Ă�0x2D12")
    For Each m In mc
        For i = 0 To m.SubMatches.Count - 1
            Debug.Print m.SubMatches(i)
            '���� >> 123,4567,890,1,1,20,0,2,12
        Next
    Next
End Sub

Sub getFlaotNum()
    Dim re As New RegExp
    Dim mc As MatchCollection
    Dim m As Match
    Dim i As Long
    re.Pattern = "(\d+\.\d+)"
    re.Global = True
    Set mc = re.Execute("123����������4567����������890.1����������1e20�����Ă�0x2D12")
    For Each m In mc
        For i = 0 To m.SubMatches.Count - 1
            Debug.Print m.SubMatches(i)
            '���� >> 890.1
        Next
    Next
End Sub

Sub getHexNum()
    Dim re As New RegExp
    Dim mc As MatchCollection
    Dim m As Match
    Dim i As Long
    re.Pattern = "([0-9 a-f A-F]+)"
    re.Global = True
    Set mc = re.Execute("123����������4567����������890.1����������1e20�����Ă�0x2D12")
    For Each m In mc
        For i = 0 To m.SubMatches.Count - 1
            Debug.Print m.SubMatches(i)
            '���� >> 123,4567,890,1,1e20,0,2D12
        Next
    Next
End Sub

Sub getHexNum2()
    Dim re As New RegExp
    Dim mc As MatchCollection
    Dim m As Match
    Dim i As Long
    re.Pattern = "(0x[0-9 a-f A-F]+)"
    re.Global = True
    Set mc = re.Execute("123����������4567����������890.1����������1e20�����Ă�0x2D12")
    For Each m In mc
        For i = 0 To m.SubMatches.Count - 1
            Debug.Print m.SubMatches(i)
            '���� >> 0x2D12
        Next
    Next
End Sub

Sub replaceChar()
    Dim re As New RegExp
    re.Pattern = "([A-Z]+)"
    re.Global = True
    re.IgnoreCase = True
    Debug.Print re.replace("ABC1234DEF567G", "")
    '���� >> 1234567
End Sub

Sub replaceNum()
    Dim re As New RegExp
    re.Pattern = "([0-9])"
    re.Global = True
    re.IgnoreCase = True
    Debug.Print re.replace("ABC1234DEF567G", "*")
    '���� >> ABC****DEF***G
End Sub
