Attribute VB_Name = "fileIO"
Option Explicit

'�t�@�C����I������_�C�A���O���g���ăt�@�C�������擾
Function OpenFileWithDialog() As String
    Dim FilePath As Variant
    Dim fileContent As String
    Dim fileNum As Integer

    '�t�@�C����I�����邽�߂̃_�C�A���O��\��
    FilePath = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "�t�@�C����I�����Ă�������")

    '���[�U�[���t�@�C����I�����Ȃ������ꍇ�͏������I��
    If FilePath = False Then
        MsgBox "�t�@�C�����I������܂���ł����B"
        Exit Function
    End If

    '�ǂݍ��񂾃t�@�C����
    'MsgBox filePath
    OpenFileWithDialog = FilePath
End Function

'CRLF�̃t�@�C����1�s�Âǂݍ���
Sub ReadTextFileByLine()
    Dim FilePath, textLine, tmp As String
    Dim fileNumber As Integer
    Dim i As Long: i = 0
    
    Application.ScreenUpdating = False
    
    '�e�L�X�g�t�@�C���̃p�X
    'filePath = "C:\Users\Public\outputCRLF.txt"     '���ڃt�@�C�������w��
    FilePath = OpenFileWithDialog                           '�t�@�C���I�[�v���_�C�A���O���g��
    
    '�t�@�C�����J��
    fileNumber = FreeFile
    Open FilePath For Input As fileNumber
    
    '�t�@�C������1�s���ǂݍ���
    Do Until EOF(fileNumber)
        Line Input #fileNumber, tmp                         'CRLF�ł���K�v������
        If i Mod 1000 = 0 Then
            ThisWorkbook.ActiveSheet.Cells(i / 1000 + 1, 1).Value = tmp
            Debug.Print (i)
        End If
        i = i + 1
    Loop
    
    '�t�@�C�������
    Close fileNumber
End Sub

'���s�R�[�hLF�̃t�@�C����ǂݍ��ށi1�x�ɓǂݍ��܂��j
Sub Sample2()
    Dim buf As String
    Dim tmp As Variant, tmp2 As Variant
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    Open "C:\Users\Public\outputLF.txt" For Input As #1
    Line Input #1, buf                          '�����őS���ǂݍ��ނ̂Ŏ����t���[�Y����
    Close #1
        
    tmp = Split(buf, vbLf)
    For i = 0 To UBound(tmp)
        If i Mod 1000 = 0 Then
            ThisWorkbook.ActiveSheet.Cells(i / 1000 + 1, 1).Value = tmp(i)
        '  ThisWorkbook.Worksheets("Sheet1").Cells(i / 1000 + 1, 1).Value = tmp(i)
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
    Dim FilePath As String
    Dim fileNum As Integer

    '�o�͐�t�@�C���̃p�X���w�肵�܂�
    FilePath = "C:\Users\Public\output.txt"

    '�������镶����̐��ƒ������w�肵�܂�
    numStrings = 10000                          '�������镶����̐�
    stringLength = 8 '�e������̒���
    
    '�����_���ȕ�����𐶐����āAoutputText �ɒǉ����܂�
    For i = 1 To numStrings
        randomString = GenerateRandomString(stringLength)
        outputText = outputText & randomString & vbCrLf
    Next i
    
    '�t�@�C���ɏo�͂��܂�
    fileNum = FreeFile
    Open FilePath For Output As #fileNum
    Print #fileNum, outputText
    Close #fileNum

    MsgBox "�t�@�C���ɏo�͂��܂����F" & FilePath
End Sub

Function GenerateRandomString(ByVal length As Integer) As String
    Dim i As Integer
    Dim charset As String
    Dim result As String
    
    charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789" '�g�p���镶���Z�b�g
    
    '�����_���ȕ�����𐶐�
    For i = 1 To length
        result = result & Mid(charset, Int((Len(charset) * Rnd) + 1), 1)
    Next i
    
    GenerateRandomString = result
End Function

Sub AddNewBook()
    Dim NewBook As Workbook
    Set NewBook = Workbooks.Add
        With NewBook
            .Title = "Title"                '�v���p�e�B�ɐݒ�
            .Subject = "Subject"            '�v���p�e�B�ɐݒ�
            .SaveAs Filename:="New.xlsx"    '�w�肵�Ȃ��ƃf�X�N�g�b�v�ɍ쐬
            
            '������Book�ւ̏������ݏ���
            Cells(1, 1).Value = "New"
            
        End With
End Sub

Sub ExistingBookOpen()
    Dim FilePath, Filename As String
    FilePath = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?")

    Workbooks.Open FilePath
    
    '�w�肵���t�@�C�������݂����Ƃ����̃t�@�C������Ԃ�
    Filename = Dir(FilePath)
    
    Dim aaa As String
    Workbooks(Filename).Worksheets("Sheet1").Range("A1").Value = "A1"
    
    '�㏑���ۑ����ĕ���
    Workbooks(Filename).Save
    Workbooks(Filename).Close

End Sub
