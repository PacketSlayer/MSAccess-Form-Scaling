'Microsoft Access Class Modile

Option Compare Database 'integer to long 12-31
Option Explicit

'//---BEGIN Private Class Properties
Private CurrentControlSize As Object
Private OrigScaleFormWindow As Object
Private CurrentControl As ScaleFormControls
Private Const SCPROPORTIONALYES = 1
Private Const SCPROPORTIONALNO = 2
Private Const SCPROPORTIONALONLYX = 3
Private Const ADJ = 630
'//---END Public Class Properties

Public Sub MeasureForm(frm As Form)
    '//Measures and records the original form size
    '//and control sizes prior to resizing
    On Error Resume Next
    Dim db As Object
    Dim CurrentWindow As New ScaleFormControls
    Dim ctrl As Control
    Dim fh As Integer
    Dim ff As Integer
    Set db = CurrentDb
    Set OrigScaleFormWindow = New Collection
    If db.Properties("AllowBypassKey") = False Then
        '//Grab measurements for the detail section of the
        '//form if it exists (and it should)
        Set CurrentControlSize = New Collection
        CurrentWindow.X = frm.Width
        fh = frm.Section(1).Height
        ff = frm.Section(2).Height
        CurrentWindow.Y = frm.Form.Detail.Height + fh + ff
        CurrentWindow.N = frm.name
        OrigScaleFormWindow.Remove frm.name
        OrigScaleFormWindow.Add CurrentWindow, frm.name
        Set CurrentWindow = Nothing
        Set CurrentControl = New ScaleFormControls
        CurrentControl.h = frm.Detail.Height
        CurrentControl.N = "Detail"
        CurrentControlSize.Add CurrentControl, frm.name & "-Detail"
        Set CurrentControl = Nothing
        '//Grab measurements for the header section of the
        '//form if it exists
        Set CurrentControl = New ScaleFormControls
        CurrentControl.h = frm.FormHeader.Height
        CurrentControl.N = "Header"
        CurrentControlSize.Add CurrentControl, frm.name & "-Header"
        Set CurrentControl = Nothing
        '//Grab measurements for the footer section of the form
        '//if it exists (and it should)
        Set CurrentControl = New ScaleFormControls
        CurrentControl.h = frm.FormFooter.Height
        CurrentControl.N = "Footer"
        CurrentControlSize.Add CurrentControl, frm.name & "-Footer"
        Set CurrentControl = Nothing
        '//Grab measurements for the controls of the form
        '//if it exists (and it should)
        For Each ctrl In frm.Controls
            With ctrl
                Set CurrentControl = New ScaleFormControls
                CurrentControl.h = .Height
                CurrentControl.w = .Width
                CurrentControl.Y = .top
                CurrentControl.X = .Left
                CurrentControl.fn = .FontSize
                CurrentControl.N = .name
                CurrentControl.d = frm.DatasheetFontHeight
                CurrentControlSize.Add CurrentControl, frm.name & "-" & .name
                Set CurrentControl = Nothing
           End With
        Next
        Set ctrl = Nothing
    End If
    db.Close
    Set db = Nothing
    Exit Sub
errorhandle:
    Err.Clear
End Sub

Public Sub ScaleForm(frm As Form, Optional scaleproportionally, Optional paint)
    '//--- Starts the resizing
    On Error GoTo errorhandle
    If IsMissing(paint) Then paint = True
    Debug.Print paint
    Dim db As Object
    Dim ctrl As Control
    Dim ScaleXaxis As Single
    Dim ScaleYaxis As Single
    Dim dsf As Single
    Set db = CurrentDb
    '// select scaling type
    If db.Properties("AllowBypassKey") = False Then
        frm.Painting = paint
        Select Case scaleproportionally
            Case 0
                ScaleXaxis = 1
                ScaleYaxis = ScaleXaxis
            Case 1
                ScaleXaxis = (frm.WindowWidth - ADJ) / FormLookUp(frm.name, "x")
                ScaleYaxis = ScaleXaxis
            Case 2
                ScaleXaxis = (frm.WindowWidth - ADJ) / FormLookUp(frm.name, "x")
                ScaleYaxis = (frm.WindowHeight - ADJ) / FormLookUp(frm.name, "y")
            Case 3
                ScaleXaxis = (frm.WindowWidth - ADJ) / FormLookUp(frm.name, "x")
                ScaleYaxis = 1
        End Select
        '//Actually perform the resize of the form window
        frm.Width = frm.WindowWidth
        For Each ctrl In frm.Controls
            With ctrl
                If scaleproportionally <> 3 Then
                    .Height = ScaleYaxis * ControlLookUp(frm.name & "-" & ctrl.name, "h")
                    .top = ScaleYaxis * ControlLookUp(frm.name & "-" & ctrl.name, "y")
                    frm.Detail.Height = ScaleYaxis * ControlLookUp(frm.name & "-Detail", "h")
                    frm.FormHeader.Height = ScaleYaxis * ControlLookUp(frm.name & "-Header", "h")
                    frm.FormFooter.Height = ScaleYaxis * ControlLookUp(frm.name & "-Footer", "h")
                End If
                If .ControlType = acListBox Or .ControlType = acComboBox Then .ColumnWidths = AdjustColumns(.ColumnWidths, .Width)
                .Width = ScaleXaxis * ControlLookUp(frm.name & "-" & ctrl.name, "w")
                .Left = ScaleXaxis * ControlLookUp(frm.name & "-" & ctrl.name, "x")
                .FontSize = ScaleYaxis * ControlLookUp(frm.name & "-" & ctrl.name, "f") - (Int(.FontSize / 10))
                dsf = ScaleYaxis * ControlLookUp(frm.name & "-" & ctrl.name, "d") - (Int(frm.DatasheetFontHeight / 10))
                frm.DatasheetFontHeight = dsf
            End With
        Next
        Set ctrl = Nothing
    End If
    db.Close
    frm.Painting = paint
    Set db = Nothing
    Exit Sub
errorhandle:
    Err.Clear
    Resume Next
End Sub

Private Function AdjustColumns(columns As String, scaleX As Single)
    '// this adjusts comlumn widths on controls such as list and combo box that can have
    'multiple columns......
    Dim wd() As String
    Dim ub As Integer
    Dim Total As Double
    Dim final, l As Double
    wd = Split(columns, ";")
    ub = UBound(wd)
    Total = 0
    For l = 0 To ub
        Total = Val(wd(l)) + Total
    Next
    For l = 0 To ub
        wd(l) = str(Val(wd(l) / Total * scaleX)) & ";"
        final = final & wd(l)
    Next
    AdjustColumns = Left(final, Len(final) - 1)
End Function

Private Function ControlLookUp(N, value)
    '// looks up control from collection to return the original size
    On Error Resume Next
    Dim tmpCurrentControl As Object
    Set tmpCurrentControl = CurrentControlSize.Item(N)
    Select Case value
        Case "h"
            ControlLookUp = tmpCurrentControl.h
        Case "w"
            ControlLookUp = tmpCurrentControl.w
        Case "x"
            ControlLookUp = tmpCurrentControl.X
        Case "y"
            ControlLookUp = tmpCurrentControl.Y
        Case "f"
            ControlLookUp = tmpCurrentControl.fn
        Case "d"
            ControlLookUp = tmpCurrentControl.d
    End Select
    Set tmpCurrentControl = Nothing
    Exit Function
End Function

Private Function FormLookUp(N, value)
    On Error Resume Next
    '//Looks up form original  x and y size
    Dim tmpCurrentControl As Object
    Set tmpCurrentControl = OrigScaleFormWindow.Item(N)
    Select Case value
        Case "x"
            FormLookUp = tmpCurrentControl.X
        Case "y"
            FormLookUp = tmpCurrentControl.Y
    End Select
    Set tmpCurrentControl = Nothing
    Exit Function
End Function

Private Sub Class_Terminate()
    '// destroys opbects and frees up memory
    Set CurrentControlSize = Nothing
    Set OrigScaleFormWindow = Nothing
    Set CurrentControl = Nothing
End Sub

