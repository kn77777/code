' C2 ~ C4 対象ファイル,コピー先フォルダ,ファルダー名
' C6 ~ C10 ファイル名1.2.3.4.5
' C12 ~ C13 キーワード1.2
' G12 ~ G13 セル1.2
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
    Dim RegExp As Object
    
    ' オブジェクト生成
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set ExcelApp = CreateObject("Excel.Application")
    Set RegExp = CreateObject("VBScript.RegExp")

    ' シートから値を読み込む
    With ThisWorkbook.Sheets("ファイルコピー")
        SourceFilePath = .Range("C2").Value
        DestinationFolderPath = .Range("C3").Value
        NewFolderName = .Range("C4").Value
        Keyword1 = .Range("C12").Value
        Keyword2 = .Range("C13").Value
        position1 = .Range("G12").Value
        position2 = .Range("G13").Value
    End With
    
    For i = 6 To 10
        FileName = ThisWorkbook.Sheets("ファイルコピー").Range("C" & i).Value
        If FileName <> "" Then
            If Left(FileName, 1) = " " Or Left(FileName, 1) = "　" Then
                MsgBox "ファイル名" & i - 5 & "は無効のファイル名です", vbCritical, "ファイル名エラー"
                Exit Sub
            End If
        End If
    Next i
    
    RegExp.Pattern = "^[A-Za-z]+\d+$"
    If Not RegExp.Test(position1) Or Not RegExp.Test(position2) Then
        MsgBox "セルの値が無効です", vbCritical, "形式エラー"
        Exit Sub
    End If
    ' 最終的なフォルダパスの生成
    FinalFolderPath = DestinationFolderPath & "\" & NewFolderName
    FileExtension = FileSystem.GetExtensionName(SourceFilePath)
    
    ' フォルダが存在しなければ作成
    If Not FileSystem.FolderExists(FinalFolderPath) Then
        FileSystem.CreateFolder (FinalFolderPath)
    Else
        MsgBox "フォルダー名がすでに存在します", vbCritical, "作成エラー"
        Exit Sub
    
    End If
    
    ' C6:C10 のファイル名に対して処理
    For i = 6 To 10
        FileName = ThisWorkbook.Sheets("ファイルコピー").Range("C" & i).Value
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
    Set RegExp = Nothing
    
    ' 完了メッセージ
    MsgBox "ファイルが作成されました。", vbInformation
End Sub
