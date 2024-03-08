Option Explicit

Sub CopyFilesAndWriteKeywords()
    Dim SourceFilePath As String
    Dim DestinationFolderPath As String
    Dim NewFolderName As String
    Dim FinalFolderPath As String
    Dim FileExtension As String
    Dim FileSystem As Object
    Dim i As Integer
    Dim FileName As String
    Dim Keyword1 As String
    Dim Keyword2 As String
    Dim position1 As String
    Dim position2 As String
    Dim ExcelApp As Object
    Dim Workbook As Object
    
    ' オブジェクト生成
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set ExcelApp = CreateObject("Excel.Application")
    
    ' シートから値を読み込む
    With ThisWorkbook.Sheets(1)
        SourceFilePath = .Range("C2").Value
        DestinationFolderPath = .Range("C3").Value
        NewFolderName = .Range("C4").Value
        Keyword1 = .Range("C12").Value
        Keyword2 = .Range("C13").Value
        position1 = .Range("G12").Value
        position2 = .Range("G13").Value
        
    End With
    
    ' 最終的なフォルダパスの生成
    FinalFolderPath = DestinationFolderPath & "\" & NewFolderName
    FileExtension = FileSystem.GetExtensionName(SourceFilePath)
    
    ' フォルダが存在しなければ作成
    If Not FileSystem.FolderExists(FinalFolderPath) Then
        FileSystem.CreateFolder (FinalFolderPath)
    Else
        MsgBox "フォルダー名がすでに存在します。"
        Exit Sub
    
    End If
    
    ' C6:C10 のファイル名に対して処理
    For i = 6 To 10
        FileName = ThisWorkbook.Sheets(1).Range("C" & i).Value
        If FileName <> "" Then
            ' ファイルをコピー
            FileSystem.CopyFile SourceFilePath, FinalFolderPath & "\" & FileName & "." & FileExtension
            
            ' コピーしたファイルにキーワードを書き込む
            Set Workbook = ExcelApp.Workbooks.Open(FinalFolderPath & "\" & FileName)
            With Workbook
                .Sheets(1).Range(position1).Value = Keyword1
                .Sheets(1).Range(position2).Value = Keyword2
                .Save
                .Close
            End With
        End If
    Next i
    
    ' オブジェクトの解放
    Set Workbook = Nothing
    Set ExcelApp = Nothing
    Set FileSystem = Nothing
    
    ' 完了メッセージ
    MsgBox "ファイルが作成されました。", vbInformation
End Sub

