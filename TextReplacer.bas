Option Explicit

'マクロ実行情報
Dim debugFlag As Boolean
Dim startDate As Date
Dim endDate As Date
Dim macroWorkbook As Workbook
Dim macroExecutionSheet As Worksheet
Dim macroParameterSheet As Worksheet
Dim templateFileFolder As Variant
Dim resultFileFolder As Variant
Dim overwriteFlag As Boolean

'置換用のコレクション
Dim keyCollection As Collection
Dim valueCollection As Collection

'主処理
'（実行ボタンから呼ばれる）
Sub main()
    startDate = Now
    Set macroWorkbook = ThisWorkbook
    Set macroExecutionSheet = macroWorkbook.Sheets("実行")
    Set macroParameterSheet = macroWorkbook.Sheets("パラメータ")
    
    'デバッグ切り替え
    debugFlag = (macroExecutionSheet.Range("C5").Value = "する")
    If debugFlag Then
        On Error GoTo 0
    Else
        On Error GoTo ErrorHandler
    End If
    
    templateFileFolder = ThisWorkbook.Path & "\" & macroExecutionSheet.Range("C2").Text & "\"
    resultFileFolder = ThisWorkbook.Path & "\" & macroExecutionSheet.Range("C3").Text & "\"
    overwriteFlag = (macroExecutionSheet.Range("C4").Value <> "する")
    
    Dim paramCell As Range
    Dim keyCells As Range
    Set keyCells = macroParameterSheet.Range("G1:Z1")
    For Each paramCell In macroParameterSheet.Range("A2:A" & Rows.Count)
        'リストの最後なら抜ける
        If paramCell.Value = "" Then
            Exit For
        End If
        '作成="!"の行のみ処理
        If paramCell(, 2).Value = "!" Then
            initKeyValues keyCells, Range(paramCell(, 7), paramCell(, 26))
            makeDescription paramCell
        End If
    Next paramCell
    '終了
    endDate = Now
    MsgBox "完了しました" & vbCrLf & "開始：" & startDate & vbCrLf & "終了：" & endDate
    On Error GoTo 0
    Exit Sub

'エラーハンドラ
ErrorHandler:
    MsgBox Err.Description
    On Error GoTo 0
End Sub

'置換用のコレクションを初期化
Sub initKeyValues(keyCells As Range, valueCells As Range)
    Set keyCollection = New Collection
    Set valueCollection = New Collection
    Dim i As Long
    Dim valueCell As Range
    i = 1
    For Each valueCell In valueCells
        keyCollection.Add keyCells(, i).Text
        valueCollection.Add valueCell.Text
        i = i + 1
    Next valueCell
End Sub

'雛形ファイルから結果ファイルを作る
Sub makeDescription(paramCell As Range)
    Dim templateFileName As Variant
    Dim templateFileNum As Integer
    Dim resultFileNum As Integer
    Dim resultSubFolder As Variant
    Dim resultFileName As Variant

    '雛形ファイルのフルパスを生成
    templateFileName = templateFileFolder & paramCell(, 6).Text
    '結果ファイルのフルパスを生成
    resultSubFolder = paramCell(, 3).Text
    If resultSubFolder <> "" Then
        resultSubFolder = resultSubFolder & "\"
    End If
    resultFileName = resultFileFolder & resultSubFolder & paramCell(, 4).Text

    '結果ファイルのサブフォルダ作成
    If Dir(resultFileFolder & resultSubFolder, vbDirectory) = "" Then
        MkDir resultFileFolder & resultSubFolder
    End If

    '結果ファイルの上書きチェック
    If Not overwriteFlag And Dir(resultFileName) <> "" Then
        Err.Raise Number:=58, Description:=resultFileName & " が既に存在するのでマクロを終了します."
    End If

    '雛形ファイルをオープン
    templateFileNum = FreeFile
    Open templateFileName For Input Access Read As #templateFileNum
    '結果ファイルをオープン
    resultFileNum = FreeFile
    Open resultFileName For Output Access Write As #resultFileNum
    While Not EOF(templateFileNum)
        Dim inputBuffer As Variant
        Dim outputBuffer As Variant
        Line Input #templateFileNum, inputBuffer
        'key-value置換
        replaceKeywords inputBuffer, outputBuffer
        Print #resultFileNum, outputBuffer
    Wend
    '全てのファイルをクローズ
    Close
End Sub

Sub replaceKeywords(inputString As Variant, resultString As Variant)
    resultString = inputString
    Dim i As Integer
    For i = 1 To keyCollection.Count
        resultString = Replace(resultString, "%" & keyCollection(i) & "%", valueCollection(i))
    Next i
End Sub
