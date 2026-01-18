VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim col As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Unprotect first (in case protected previously)
    ws.Unprotect Password:="Pass8371!"
    
    ' -------------------------
    ' Headers and formatting (main OT section)
    ' -------------------------
    With ws.Range("C5:J5")
        .Value = Array("User:", "Current Date", "Type", "Date of OT", "OT start time", "OT end time", "OT hr earn/use", "Reason")
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 102, 204)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    With ws.Range("C8:O8")
        .Value = Array("User:", "Current Date", "Type", "Date of OT", "OT start time", "OT end time", "OT hr earn/use", "Reason", _
                       "Approved Button", "Reject Button", "Approver/Rejector", "Status", "Approve Date & Time")
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 102, 204)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ws.Rows("7:7").Interior.Color = RGB(0, 0, 0)
    
    ' -------------------------
    ' Instruction row
    ' -------------------------
    With ws.Range("A2")
        .Value = "Only can fill in the yellow column"
        .Interior.Color = RGB(255, 255, 0)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
    ws.Columns("A:A").AutoFit
    
    ' -------------------------
    ' Row 6 input
    ' -------------------------
    ws.Range("C6").Value = Environ("Username")
    ws.Range("D6").Value = Format(Now, "yyyy-mm-dd hh:nn")
    ws.Range("E6:J6").ClearContents
    ws.Range("C6:J6").Font.Size = 10
    ws.Range("J6").WrapText = True
    ws.Rows("6:6").RowHeight = 15
    
    ' Drop-down for Type (E6)
    With ws.Range("E6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="earn,use"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' Highlight input cells
    With ws.Range("E6:J6")
        .Interior.Color = RGB(255, 255, 0)
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("I6").Interior.Pattern = xlNone
    
    ws.Range("C6:J6").HorizontalAlignment = xlRight
    ws.Range("C9:N9").HorizontalAlignment = xlRight
    ws.Range("C9:J100").HorizontalAlignment = xlRight
    ws.Range("M9:N100").HorizontalAlignment = xlRight
    
    ' -------------------------
    ' Column formatting
    ' -------------------------
    ws.Range("G6:H6").NumberFormat = "hh:mm:ss AM/PM"
    ws.Columns("C:O").AutoFit
    ws.Columns("J").ColumnWidth = ws.Columns("J").ColumnWidth + 15
    ws.Columns("L").ColumnWidth = ws.Columns("L").ColumnWidth + 3.5
    ws.Columns("N").ColumnWidth = ws.Columns("N").ColumnWidth + 1.5
    ws.Columns("N").AutoFit
    
    ' -------------------------
    ' Submit button
    ' -------------------------
    On Error Resume Next
    ws.Buttons("SubmitBtn").Delete
    On Error GoTo 0
    
    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range("K6").Left + 1, ws.Range("K6").Top, ws.Range("K6").Width, ws.Range("K6").Height)
    With btn
        .Caption = "Submit"
        .Name = "SubmitBtn"
        .OnAction = "SubmitRow6"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    ' -------------------------
    ' Formulas in I6 & O8
    ' -------------------------
    ws.Range("I6").Formula = _
        "=IF(AND(ISNUMBER(G6),ISNUMBER(H6),OR(E6=""earn"",E6=""use""))," & _
        "IF(E6=""earn"", MOD(H6-G6,1)*24, -MOD(H6-G6,1)*24), """")"
    ws.Range("O8").Formula = "=IF(COUNTA(N9:N100)=0,""Action Date & Time"",IF(COUNTIF(N9:N100,""Approved"")>0,""Approval Date & Time"",""Rejection Date & Time""))"
    
    ' -------------------------
    ' SUMMARY SECTION
    ' -------------------------
    With ws.Range("P1:AE1")
        .Merge
        .Value = "Summary"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 102, 204)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Row labels
    ws.Range("P2").Value = "Username"
    ws.Range("P3").Value = "User ID"
    ws.Range("P4").Value = "Total OT"
    ws.Range("P5").Value = "Used OT"
    ws.Range("P6").Value = "Remain OT hr"
    
    With ws.Range("P1:P6")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
    
    With ws.Range("Q2:AE6")
        .Borders.LineStyle = xlContinuous
    End With
    
    ws.Columns("P:AE").AutoFit
    
    lastCol = ws.Range("Q3").End(xlToRight).Column
    If lastCol < 31 Then lastCol = 31
    
    For col = 17 To lastCol ' Q = col 17
        ' Total OT (earn + approved)
        ws.Cells(4, col).Formula = _
            "=SUMIFS($I$9:$I$1000,$C$9:$C$1000," & ws.Cells(3, col).Address(False, False) & _
            ",$E$9:$E$1000,""earn"",$N$9:$N$1000,""Approved"")"
        
        ' Used OT (use + approved)
        ws.Cells(5, col).Formula = _
            "=SUMIFS($I$9:$I$1000,$C$9:$C$1000," & ws.Cells(3, col).Address(False, False) & _
            ",$E$9:$E$1000,""use"",$N$9:$N$1000,""Approved"")"
        
        ' Remaining OT (Earn - Use)
        ws.Cells(6, col).Formula = "=" & ws.Cells(4, col).Address(False, False) & "+" & ws.Cells(5, col).Address(False, False)
    Next col

    ' -------------------------
    ' Countdown header at P8
    ' -------------------------
    With ws.Range("P8")
        .Value = "Count Down"
        .Font.Bold = True
        .Font.Color = ws.Range("C8").Font.Color
        .Interior.Color = ws.Range("C8").Interior.Color
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' -------------------------
    ' Countdown formula in P9:P1000 (whole days)
    ' -------------------------
    With ws.Range("P9:P1000")
        .Formula = "=IF(N9=""Approved"", MAX(0, 93-ROUNDDOWN(TODAY()-O9,0)), """")"
        .Calculate
    End With
    
    ' -------------------------
    ' Disable expired rows automatically
    ' -------------------------
    Dim i As Long
    For i = 9 To 1000
        If ws.Cells(i, "P").Value = 0 And ws.Cells(i, "N").Value = "Approved" Then
            ' Clear OT hours so it doesn't count in summary
            ws.Cells(i, "I").Value = 0
            
            ' Lock the row
            ws.Rows(i).Locked = True
            
            ' Gray out row
            ws.Range("C" & i & ":J" & i).Interior.Color = RGB(220, 220, 220)
            
            ' Remove buttons if exist
            On Error Resume Next
            ws.Shapes("ApproveBtn_" & i).Delete
            ws.Shapes("RejectBtn_" & i).Delete
            On Error GoTo 0
        End If
    Next i

    
    ' -------------------------
    ' Lock sheet but leave inputs editable
    ' -------------------------
    ws.Cells.Locked = True
    ws.Range("E6,F6,G6,H6,J6").Locked = False
    ws.Range("I9:I1000").Locked = True
    ws.Rows("3:3").Locked = True
    
    ' Final protection with password
    ws.Protect Password:="Pass8371!", UserInterfaceOnly:=True
End Sub




