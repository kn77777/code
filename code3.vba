'number　ラベル
'CommandButton1,2,3
'TextBox1

'Private Sub Workbook_Open()
'   UserForm1.Show
'End Sub



Dim SelectedRange As Range
Dim ElementCount As Integer
Dim SetCount As Integer

Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Const HWND_TOPMOST As LongPtr = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_SHOWWINDOW As Long = &H40


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

Private Sub number_Click()

End Sub

Private Sub TextBox1_Change()
    If SelectedRange Is Nothing Then
        Exit Sub
    End If
    SelectedRange(SetCount).Cells = TextBox1.Value
    number.Caption = SetCount
End Sub

Public Sub KeepFormOnTop(ByVal FormName As String)
    Dim hWnd As LongPtr
    hWnd = FindWindowA(vbNullString, FormName)
    If hWnd > 0 Then
        SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    End If
End Sub

Private Sub UserForm_Activate()
    KeepFormOnTop Me.Caption
End Sub


