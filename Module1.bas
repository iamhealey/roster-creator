Attribute VB_Name = "Module1"
Sub HighlightStrings(cFnd As String)
'Updateby Extendoffice 20160704
Application.ScreenUpdating = False
Dim Rng As Range
Dim xTmp As String
Dim x As Long
Dim m As Long
Dim y As Long
y = Len(cFnd)
For Each Rng In Selection
  With Rng
    m = UBound(Split(Rng.Value, cFnd))
    If m > 0 Then
      xTmp = ""
      For x = 0 To m - 1
        xTmp = xTmp & Split(Rng.Value, cFnd)(x)
        .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.FontStyle = "Bold"
        xTmp = xTmp & cFnd
      Next
    End If
  End With
Next Rng
Application.ScreenUpdating = True
End Sub

Sub userformstart()

    BookingWindow.Show
    
    Exit Sub
    
errormessage:

MsgBox "Uh oh! looks like something is broken. Please check your input!"
    
End Sub
