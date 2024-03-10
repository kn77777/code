'number　ラベル
'CommandButton1,2,3
'TextBox1

'Private Sub Workbook_Open()
'   UserForm1.Show
'End Sub



Dim SelectedRange As Range
Dim ElementCount As Integer
Dim SetCount As Integer


Private Sub CommandButton1_Click()
    GetSelectedRangeValues
End Sub


Sub GetSelectedRangeValues()
    ' ユーザーによって選択された範囲を取得
    Set SelectedRange = Selection
    ElementCount = SelectedRange.Count
    SetCount = 1
    TextBox1.Value = SelectedRange(SetCount)
End Sub

Private Sub CommandButton2_Click()
    If SetCount = 1 Or SetCount = 0 Then
        Exit Sub
    End If
    SetCount = SetCount - 1
    TextBox1.Value = SelectedRange(SetCount)
End Sub

Private Sub CommandButton3_Click()
    If SetCount = ElementCount Then
        Exit Sub
    End If
    SetCount = SetCount + 1
    TextBox1.Value = SelectedRange(SetCount)
End Sub

Private Sub TextBox1_Change()
    If SelectedRange Is Nothing Then
        Exit Sub
    End If
    SelectedRange(SetCount).Cells = TextBox1.Value
    number.Caption = SetCount
End Sub

