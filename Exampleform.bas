' example  use on Microsoft Access Form Object
Public Window as Object

Private Sub Form_Open(Cancel As Integer)
    Set window = New ScaleWindow
    window.MeasureForm Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    window.ScaleForm Me, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRORS
    Set window = Nothing
    Exit Sub
ERRORS:
    Call errorh.errorhandler(Err.Number, Err.Description, Me.name, "Form_unload")
    Err.Clear
    Resume Next
End Sub
