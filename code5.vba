'CommandButton1,2,3
'Label1,2

Dim SelectedRangeA As Range
Dim SelectedRangeB As Range
Dim Cell As Range
Dim Dict As Object
Dim DuplicateFound As Boolean


Private Sub CommandButton1_Click()
    Set SelectedRangeA = Selection
    For Each Cell In SelectedRangeA
        valuesString = valuesString & Cell.Value & vbCrLf
    Next Cell
    Label1.Caption = "値数:" & SelectedRangeA.Count & vbCrLf & valuesString
End Sub

Private Sub CommandButton2_Click()
    Set SelectedRangeB = Selection
    For Each Cell In SelectedRangeB
        valuesString = valuesString & Cell.Value & vbCrLf
    Next Cell
    Label2.Caption = "値数:" & SelectedRangeB.Count & vbCrLf & valuesString
End Sub

Private Sub CommandButton3_Click()
On Error GoTo ErrHandler
    Set Dict = CreateObject("Scripting.Dictionary")
    DuplicateFound = False

    For Each Cell In SelectedRangeA
        If Not Dict.Exists(Cell.Value) Then
            Dict.Add Cell.Value, 1
            If Trim(Cell.Value) = "" Then
                MsgBox "Aの値を指定してください。", vbExclamation
                Exit Sub
            End If
        Else
            DuplicateFound = True
            Exit For
        End If
    Next Cell
    
    For Each Cell In SelectedRangeB
        If Not Dict.Exists(Cell.Value) Then
            Dict.Add Cell.Value, 1
            If Trim(Cell.Value) = "" Then
                MsgBox "Bの値を指定してください。", vbExclamation
                Exit Sub
            End If
        Else
            DuplicateFound = True
            Exit For
        End If
    Next Cell
    
    If DuplicateFound Then
        MsgBox "重複している値がありました。", vbExclamation, "重複検出"
    Else
        MsgBox "重複はありません。", vbInformation, "重複検出"
    End If
    Exit Sub
ErrHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub
