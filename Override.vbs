Option Explicit

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    'On Error Resume Next
    'CommonCheckSynWithMZ
    CommonCheckSynWithAllAvailableDrive
End Sub

Private Sub Workbook_Open()
    Dim tempAuthor As String
    tempAuthor = ActiveWorkbook.BuiltinDocumentProperties("Author").Value
    
    ' TODO:
    ' 1.Remove current owner in the owner list if have
    ' 2.Close the file directly without saving
    
    Dim Arr
    
    Dim idx, i As Integer
    
    If InStr(tempAuthor, theUser) > 0 Then
        Arr = Split(tempAuthor, ";")
        For i = 0 To UBound(Arr)
            If InStr(Arr(i), theUser) Then
                idx = i
                Exit For
            End If
            
        Next
    
        If idx = UBound(Arr) Then
            ActiveWorkbook.BuiltinDocumentProperties("Author").Value = Replace(tempAuthor, ";" & Arr(idx), "")
        Else
            ActiveWorkbook.BuiltinDocumentProperties("Author").Value = Replace(tempAuthor, Arr(idx) & ";", "")
        End If
    
    End If
    
    'On Error Resume Next
    If CommonCheckSynWithM_OPEN Then
    'If CommonCheckSynWithAllAvailableDrive_OPEN Then
        ActiveWorkbook.Saved = True
        ActiveWorkbook.Close
    Else
        ActiveWorkbook.BuiltinDocumentProperties("Author").Value = tempAuthor
        ActiveWindow.Caption = ActiveWorkbook.FullName
    End If
    

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Application.CommandBars("Staff & Workstation").Visible = False
    'MsgBox ActiveWorkbook.BuiltinDocumentProperties("Author").Value
    
    'On Error Resume Next
    'CommonCheckSynWithM
    
    ActiveWorkbook.Saved = True
    'ActiveWorkbook.Close savechanges:=False
End Sub
