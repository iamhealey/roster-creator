VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BookingWindow 
   Caption         =   "Booking Window"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7035
   OleObjectBlob   =   "BookingWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BookingWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    For Each CheckBox In BookingWindow.Extras.Controls
        CheckBox.Value = True
    Next CheckBox

End Sub
Private Sub UserForm_Initialize()
'Change employee list items
    Employee.Clear
    With Employee
    .AddItem "Andre"
    .AddItem "Brian"
    .AddItem "Danny"
    .AddItem "Johnny"
    .AddItem "Rumesh"
    .AddItem "Shelton"
    End With
'Change booth list items
    Booth.Clear
    With Booth
    .AddItem "OpenBooth One (outdoors)"
    .AddItem "OpenBooth Two (outdoors)"
    .AddItem "OpenBooth Three"
    .AddItem "OpenBooth Four (outdoors)"
    .AddItem "OpenBooth Five"
    .AddItem "GIFBOOTH One"
    .AddItem "GIFBOOTH Two"
    End With
'Change backdrop list items
    Backdrop.Clear
    With Backdrop
    .AddItem "No Backdrop"
    .AddItem "White"
    .AddItem "Champagne"
    .AddItem "Rose Gold"
    .AddItem "Flower Wall"
    End With
'Change Event Type list items
    With EventType
    .AddItem "Wedding"
    .AddItem "Engagement"
    .AddItem "Birthday"
    .AddItem "Corporate"
    .AddItem "Other"
    End With
'Change Payment Type list items
    With PaymentType
    .AddItem "Collect"
    .AddItem "Pending"
    End With
'Change Stairs list items
    With Stairs
    .AddItem "Yes"
    .AddItem "No"
    .AddItem "Unsure"
    End With
End Sub
Sub CreateBooking_Click()
    On Error GoTo errormessage

    BookingWindow.Hide
    
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Sheets(1)
    Dim EventDate As Date
    EventDate = DateValue(BookingWindow.EventDate)
    
    Dim Employee As String
    Employee = BookingWindow.Employee
'Background colours for employees
    Dim Rval As Integer, Gval As Integer, Bval As Integer
    Select Case Employee
        Case Is = "Andre"
            Rval = 255
            Gval = 0
            Bval = 0
        Case Is = "Brian"
            Rval = 217
            Gval = 217
            Bval = 217
        Case Is = "Danny"
            Rval = 255
            Gval = 252
            Bval = 0
        Case Is = "Johnny"
            Rval = 0
            Gval = 176
            Bval = 240
        Case Is = "Rumesh"
            Rval = 255
            Gval = 64
            Bval = 255
        Case Is = "Shelton"
            Rval = 0
            Gval = 250
            Bval = 0
'No choice handler
        Case Is = "SUPPORT STAFF"
            MsgBox "No employee chosen."
            GoTo earlyexit
    End Select

'Booth selection - maps to column on spreadsheet
    Dim Booth As String
    Booth = BookingWindow.Booth

'No selection handler
    If Booth = "MACHINE" Then
        MsgBox "No booth chosen."
        GoTo earlyexit
    End If
    
    Dim col As Integer
    col = WorksheetFunction.Match(Booth, sht.Rows(1), 0)

'Extras selection - creates text in cell.
    Dim Extras As String
    For Each CheckBox In BookingWindow.Extras.Controls
        If CheckBox.Value Then
            Extras = Extras & CheckBox.Caption & vbNewLine
            ExtraCount = ExtraCount + 1
        End If
    Next CheckBox

'Cell content creation
    Dim output As String
    output = _
        "Client Name: " & BookingWindow.ClientName & vbNewLine & vbNewLine & _
        "Setup Time: " & BookingWindow.SetupTime & vbNewLine & _
        "Event Time: " & BookingWindow.EventStart & " - " & BookingWindow.EventEnd & vbNewLine & vbNewLine & _
        "Location: " & BookingWindow.Address & vbNewLine & vbNewLine & _
        "Event Type: " & BookingWindow.EventType & vbNewLine & _
        "Backdrop: " & BookingWindow.Backdrop & vbNewLine & _
        "Extras: " & vbNewLine & Extras & vbNewLine & _
        "Payment Type: " & BookingWindow.PaymentType & " " & BookingWindow.PaymentAmount & vbNewLine & _
        "Stairs: " & BookingWindow.Stairs & vbNewLine & vbNewLine & _
        "Contact: " & BookingWindow.PhoneNumber & vbNewLine & vbNewLine & _
        "Notes: " & BookingWindow.Notes

'Find correct row to insert data in
    Dim i As Integer
'Declaration causes error
   'Dim caldate As Date
    Dim datefound As Boolean
    
    i = 2
    caldate = sht.Range("A2")
    datefound = False
    While datefound = False
        caldate = sht.Cells(i, 1)
'Exact matching date
        If caldate = EventDate Then
            datefound = True
        End If
        
'New date is required
        If caldate > EventDate Then
            datefound = True
            sht.Rows(i).EntireRow.Insert
            sht.Rows(i).RowHeight = 300
            sht.Rows(i).EntireRow.Cells.Interior.ColorIndex = 0
        End If
        
'Bottom row - create new
        If caldate = "" Then
            datefound = True
            sht.Rows(i).RowHeight = 300
            sht.Rows(i).EntireRow.Cells.Interior.ColorIndex = 0
        End If
        
        If datefound = False Then i = i + 1
    Wend

'Create borders around selection
        With sht.Range(Cells(i, 1), Cells(i, 9))
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
'Format row
        With sht.Range(Cells(i, 3), Cells(1, 9)).Font
            '.Bold = False
            .Size = 12
        End With
        
'Check for existing appointment
    If sht.Cells(i, col) <> "" Then
        If MsgBox("An entry already exists for this combination of Attendant, Event Date and Machine. Would you like to create an additional entry?", vbYesNo) = vbYes Then
            sht.Rows(i).EntireRow.Insert
            sht.Rows(i).RowHeight = 300
            sht.Rows(i).EntireRow.Cells.Interior.ColorIndex = 0
            sht.Range(Cells(i, 1), Cells(i + 1, 2)).Interior.ColorIndex = 36
        Else
            GoTo earlyexit
        End If
    End If
    
    sht.Cells(i, 1) = EventDate
    sht.Cells(i, 2) = UCase(WorksheetFunction.Text(Weekday(EventDate), "dddd"))
    sht.Cells(i, col) = output
    sht.Cells(i, col).Interior.Color = RGB(Rval, Gval, Bval)
    sht.Cells(i, col).Characters(Start:=1, Length:=13 + Len(BookingWindow.ClientName)).Font.FontStyle = "Bold"
    sht.Cells(i, col).Select
    
    HighlightStrings ("Setup Time")
    HighlightStrings ("Event Time")
    HighlightStrings ("Location")
    HighlightStrings ("Event Type")
    HighlightStrings ("Backdrop")
    HighlightStrings ("Extras")
    HighlightStrings ("Contact")
    HighlightStrings ("Notes")
    HighlightStrings ("Stairs")
    HighlightStrings ("Payment Type")
    
earlyexit:
    Unload Me
    Exit Sub

errormessage:

MsgBox "Uh oh! looks like something is broken. Please check your input!"
        
End Sub

Private Sub Cancel_Click()
    Unload Me
    BookingWindow.Hide
End Sub




