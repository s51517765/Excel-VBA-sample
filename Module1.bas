Attribute VB_Name = "Module1"
Option Explicit

'�t�@�C����I������_�C�A���O���g���ăt�@�C�������擾
Function OpenFileWithDialog() As String
    Dim filePath As Variant
    Dim fileContent As String
    Dim fileNum As Integer

    ' �t�@�C����I�����邽�߂̃_�C�A���O��\��
    filePath = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "�t�@�C����I�����Ă�������")

    ' ���[�U�[���t�@�C����I�����Ȃ������ꍇ�͏������I��
    If filePath = False Then
        MsgBox "�t�@�C�����I������܂���ł����B"
        Exit Function
    End If

    ' �ǂݍ��񂾃t�@�C����
    'MsgBox filePath
    OpenFileWithDialog = filePath
End Function

'CRLF�̃t�@�C����1�s�Âǂݍ���
Sub ReadTextFileByLine()
    Dim filePath, textLine, tmp As String
    Dim fileNumber As Integer
    Dim i As Long: i = 0
    
    Application.ScreenUpdating = False
    
    ' �e�L�X�g�t�@�C���̃p�X
    'filePath = "C:\Users\Public\outputCRLF.txt"     '���ڃt�@�C�������w��
    filePath = OpenFileWithDialog                           '�t�@�C���I�[�v���_�C�A���O���g��
    
    ' �t�@�C�����J��
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    
    ' �t�@�C������1�s���ǂݍ���
    Do Until EOF(fileNumber)
        Line Input #fileNumber, tmp                         'CRLF�ł���K�v������
        If i Mod 1000 = 0 Then
            ThisWorkbook.ActiveSheet.Cells(i / 1000 + 1, 1).Value = tmp
            Debug.Print (i)
        End If
        i = i + 1
    Loop
    
    ' �t�@�C�������
    Close fileNumber
End Sub

'���s�R�[�hLF�̃t�@�C����ǂݍ��ށi1�x�ɓǂݍ��܂��j
Sub Sample2()
    Dim buf As String
    Dim tmp As Variant, tmp2 As Variant
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    Open "C:\Users\Public\outputLF.txt" For Input As #1
    Line Input #1, buf   ' �����őS���ǂݍ��ނ̂Ŏ����t���[�Y����
    Close #1
        
    tmp = Split(buf, vbLf)
    For i = 0 To UBound(tmp)    '---(1)
        If i Mod 1000 = 0 Then
            ThisWorkbook.ActiveSheet.Cells(i / 1000 + 1, 1).Value = tmp(i)
        '   ThisWorkbook.Worksheets("Sheet1").Cells(i / 1000 + 1, 1).Value = tmp(i)
            Debug.Print (i)
        End If
    Next i
End Sub

'�����_���ȕ�����t�@�C���쐬
Sub GenerateRandomStringsToFile()
    Dim outputText As String
    Dim i As Long
    Dim numStrings As Long
    Dim stringLength As Integer
    Dim randomString As String
    Dim filePath As String
    Dim fileNum As Integer

    ' �o�͐�t�@�C���̃p�X���w�肵�܂�
    filePath = "C:\Users\Public\output.txt" ' �K�؂ȃp�X�ɕύX���Ă�������

    ' �������镶����̐��ƒ������w�肵�܂�
    numStrings = 10000000  ' �������镶����̐�
    stringLength = 8 ' �e������̒���
    
    ' �����_���ȕ�����𐶐����āAoutputText �ɒǉ����܂�
    For i = 1 To numStrings
        randomString = GenerateRandomString(stringLength)
        outputText = outputText & randomString & vbCrLf
    Next i
    
    ' �t�@�C���ɏo�͂��܂�
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, outputText
    Close #fileNum

    MsgBox "�t�@�C���ɏo�͂��܂����F" & filePath
End Sub

Function GenerateRandomString(ByVal length As Integer) As String
    Dim i As Integer
    Dim charset As String
    Dim result As String
    
    charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789" ' �g�p���镶���Z�b�g
    
    ' �����_���ȕ�����𐶐�
    For i = 1 To length
        result = result & Mid(charset, Int((Len(charset) * Rnd) + 1), 1)
    Next i
    
    GenerateRandomString = result
End Function

