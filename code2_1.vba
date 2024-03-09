Sub SelectFolder()

    '変数宣言
    Dim Folder As String
    Dim Worksheet As Worksheet

    'フォルダ選択ウィンドウ
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "フォルダを選択してください"
        If .Show = -1 Then
            Folder = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With


    '結果入力
    ThisWorkbook.Sheets("ファイルコピー").Range("C3").Value = Folder

End Sub

