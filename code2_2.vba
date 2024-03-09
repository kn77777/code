Sub SelectFile()
    
    '変数宣言
    Dim FileDialog As FileDialog
    
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)

    'ファイル選択ウィンドウ
    With FileDialog
        .Title = "ファイルを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "All Files", "*.*"

        If .Show = -1 Then
            Dim SelectedFilePath As String
            SelectedFilePath = .SelectedItems(1)

            '結果入力
            ThisWorkbook.Sheets("ファイルコピー").Range("C2").Value = SelectedFilePath
        End If
    End With
End Sub

