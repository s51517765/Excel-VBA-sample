Attribute VB_Name = "fileIO"
Option Explicit

'ファイルを選択するダイアログを使ってファイル名を取得
Function OpenFileWithDialog() As String
    Dim FilePath As Variant
    Dim fileContent As String
    Dim fileNum As Integer

    'ファイルを選択するためのダイアログを表示
    FilePath = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "ファイルを選択してください")

    'ユーザーがファイルを選択しなかった場合は処理を終了
    If FilePath = False Then
        MsgBox "ファイルが選択されませんでした。"
        Exit Function
    End If

    '読み込んだファイル名
    'MsgBox filePath
    OpenFileWithDialog = FilePath
End Function

'CRLFのファイルを1行づつ読み込む
Sub ReadTextFileByLine()
    Dim FilePath, textLine, tmp As String
    Dim fileNumber As Integer
    Dim i As Long: i = 0
    
    Application.ScreenUpdating = False
    
    'テキストファイルのパス
    'filePath = "C:\Users\Public\outputCRLF.txt"     '直接ファイル名を指定
    FilePath = OpenFileWithDialog                           'ファイルオープンダイアログを使う
    
    'ファイルを開く
    fileNumber = FreeFile
    Open FilePath For Input As fileNumber
    
    'ファイルから1行ずつ読み込む
    Do Until EOF(fileNumber)
        Line Input #fileNumber, tmp                         'CRLFである必要がある
        If i Mod 1000 = 0 Then
            ThisWorkbook.ActiveSheet.Cells(i / 1000 + 1, 1).Value = tmp
            Debug.Print (i)
        End If
        i = i + 1
    Loop
    
    'ファイルを閉じる
    Close fileNumber
End Sub

'改行コードLFのファイルを読み込む（1度に読み込まれる）
Sub Sample2()
    Dim buf As String
    Dim tmp As Variant, tmp2 As Variant
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    Open "C:\Users\Public\outputLF.txt" For Input As #1
    Line Input #1, buf                          'ここで全部読み込むので実質フリーズする
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

'ランダムな文字列ファイル作成
Sub GenerateRandomStringsToFile()
    Dim outputText As String
    Dim i As Long
    Dim numStrings As Long
    Dim stringLength As Integer
    Dim randomString As String
    Dim FilePath As String
    Dim fileNum As Integer

    '出力先ファイルのパスを指定します
    FilePath = "C:\Users\Public\output.txt"

    '生成する文字列の数と長さを指定します
    numStrings = 10000                          '生成する文字列の数
    stringLength = 8 '各文字列の長さ
    
    'ランダムな文字列を生成して、outputText に追加します
    For i = 1 To numStrings
        randomString = GenerateRandomString(stringLength)
        outputText = outputText & randomString & vbCrLf
    Next i
    
    'ファイルに出力します
    fileNum = FreeFile
    Open FilePath For Output As #fileNum
    Print #fileNum, outputText
    Close #fileNum

    MsgBox "ファイルに出力しました：" & FilePath
End Sub

Function GenerateRandomString(ByVal length As Integer) As String
    Dim i As Integer
    Dim charset As String
    Dim result As String
    
    charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789" '使用する文字セット
    
    'ランダムな文字列を生成
    For i = 1 To length
        result = result & Mid(charset, Int((Len(charset) * Rnd) + 1), 1)
    Next i
    
    GenerateRandomString = result
End Function

Sub AddNewBook()
    Dim NewBook As Workbook
    Set NewBook = Workbooks.Add
        With NewBook
            .Title = "Title"                'プロパティに設定
            .Subject = "Subject"            'プロパティに設定
            .SaveAs Filename:="New.xlsx"    '指定しないとデスクトップに作成
            
            'ここでBookへの書き込み処理
            Cells(1, 1).Value = "New"
            
        End With
End Sub

Sub ExistingBookOpen()
    Dim FilePath, Filename As String
    FilePath = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")

    Workbooks.Open FilePath
    
    '指定したファイルが存在したときそのファイル名を返す
    Filename = Dir(FilePath)
    
    Dim aaa As String
    Workbooks(Filename).Worksheets("Sheet1").Range("A1").Value = "A1"
    
    '上書き保存して閉じる
    Workbooks(Filename).Save
    Workbooks(Filename).Close

End Sub
