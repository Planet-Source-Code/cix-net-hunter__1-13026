Attribute VB_Name = "RemoveDubbel"
Public CancelOrder As Boolean
Function RemoveDupes(lst As ListBox)
    Dim iPos As Integer
    iPos = 0
    '-- if listbox empty then exit..
    If lst.ListCount < 1 Then Exit Function
    If StopAll = True Then Exit Function
    ProcessLineMaxValue = lst.ListCount
    Do While iPos < lst.ListCount
    DoEvents
    If StopAll = True Then Exit Function
    Dim Ik As Long
    Ik = iPos
    Hunter.Statusline.Caption = "Filtering list.."
    ProcessLineMaxValue = lst.ListCount
    ProcessLine_ValueChange Hunter.Picture1, Ik
        lst.Text = lst.List(iPos)
        '-- check if text already exists..


        If lst.ListIndex <> iPos Then
            '-- if so, remove it and keep iPos..
            lst.RemoveItem iPos
        Else
            '-- if not, increase iPos..
            iPos = iPos + 1
        End If
        
    Loop
    '-- used to unselect the last selected l
    '     ine..
    lst.Text = "~~~^^~~~"
End Function

