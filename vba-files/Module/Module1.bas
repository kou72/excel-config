Attribute VB_Name = "Module1"

Sub ExtractConfigInfo()
    ' ダイアログを表示し、選択したパスを取得
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim selectedPath As String
    With fd
        .Title = "Select Path"
        .AllowMultiSelect = False
        
        If .Show = True Then
            selectedPath = .SelectedItems(1)
        End If
    End With

    ' 選択したパスを "path" という名前のセルに書き込む
    ' Dim rng As Range
    ' Set rng = ThisWorkbook.Names("path").RefersToRange
    ' rng.Value = selectedPath

    ' Configファイル名（実際のファイル名に置き換えてください）
    Dim fileName As String
    fileName = selectedPath
    
    ' 検索対象のキーワードリスト
    Dim searchWords As Variant
    searchWords = Array("interface FastEthernet0/1", "interface FastEthernet0/2", "interface FastEthernet0/3") ' 必要なら更に追加
    
    ' ファイルを読み込みモードで開く
    Dim fileNo As Integer
    fileNo = FreeFile
    Open fileName For Input As fileNo
    
    ' フラグを初期化
    Dim hierarchyLevel As Integer
    Dim foundLine As Boolean
    Dim word As Variant
    Dim textLine As String
    
    ' 検索キーワードでループ
    For Each word In searchWords
        hierarchyLevel = 0
        foundLine = False
        
        ' ファイルを最初から読み込む
        Seek fileNo, 1
        
        ' ファイルを1行ずつ読み込む
        Do Until EOF(fileNo)
            Line Input #fileNo, textLine
            
            ' 行が目的の文字列を含むかチェック
            If InStr(textLine, word) > 0 Then
                hierarchyLevel = Len(textLine) - Len(LTrim(textLine))
                foundLine = True
            ElseIf foundLine Then
                ' 目的の文字列が見つかったら
                ' 現在の行が子要素かどうかチェック
                If Len(textLine) - Len(LTrim(textLine)) > hierarchyLevel Then
                    ' これは子要素なので、出力
                    Debug.Print textLine
                    MsgBox textLine
                Else
                    ' 子要素の範囲を超えたので、ループを抜ける
                    Exit Do
                End If
            End If
        Loop
    Next word
    
    ' ファイルを閉じる
    Close fileNo
End Sub
