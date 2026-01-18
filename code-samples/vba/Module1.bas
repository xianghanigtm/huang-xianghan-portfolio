Attribute VB_Name = "Module1"
Option Explicit

' -------------------------------
' Check if a cell is green (#70AD47)
' -------------------------------
Public Function IsGreenHeader(rng As Range) As Boolean
    Application.Volatile
    If rng.Interior.Color = RGB(112, 173, 71) Then
        IsGreenHeader = True
    Else
        IsGreenHeader = False
    End If
End Function

' -------------------------------
' Auto fill user and current date
' -------------------------------
Sub AutoFillDetails()
    Range("C6").Value = Environ("Username") ' LAN ID
    Range("D6").Value = Format(Now, "yyyy-mm-dd hh:nn")
End Sub

' -------------------------------
' Initialize dynamic formula for I6
' -------------------------------
Sub InitI6Formula()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ws.Range("I6").Formula = _
        "=IF(AND(ISNUMBER(G6),ISNUMBER(H6),E6=""earn""), " & _
        "IF(AND(IsGreenHeader($C$3),WEEKDAY(F6,2)>5), " & _
        "IF((H6-G6)*24<4,4,8), (H6-G6)*24), " & _
        "IF(AND(ISNUMBER(G6),ISNUMBER(H6)),(H6-G6)*24,""""))"
End Sub

' -------------------------------
' Submit request from row 6 (C6:J6)
' -------------------------------
Public Sub SubmitRow6()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Check if all fields in row 6 (C6:J6) are filled
    If Application.WorksheetFunction.CountBlank(ws.Range("C6:J6")) > 0 Then
        MsgBox "Please fill in all fields before submitting!", vbExclamation
        Exit Sub
    End If

    Dim nextRow As Long
    Dim logStartRow As Long: logStartRow = 9
    Dim startTime As Variant, endTime As Variant, otHours As Double
    Dim otType As String
    Dim userID As String, userCol As Long
    Dim headerCell As Range
    
    ' Find next empty row in column B
    nextRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1
    If nextRow < logStartRow Then nextRow = logStartRow
    
    ' Unprotect to allow changes
    ws.Unprotect Password:="Pass8371!"
    
    ' Numbering
    ws.Cells(nextRow, "B").Value = (nextRow - logStartRow + 1) & ")"
    
    ' Copy details from row 6
    ws.Range("C" & nextRow & ":J" & nextRow).Value = ws.Range("C6:J6").Value
    
    ' Format Date of OT (column F) to show d/m/yyyy and weekday
    If IsDate(ws.Cells(nextRow, "F").Value) Then
        ws.Cells(nextRow, "F").NumberFormat = "d/m/yyyy ddd"
    End If
    
    ' Ensure G and H keep the time format
    ws.Cells(nextRow, "G").NumberFormat = ws.Cells(6, "G").NumberFormat
    ws.Cells(nextRow, "H").NumberFormat = ws.Cells(6, "H").NumberFormat
    
    ' -------------------------
    ' Compute OT Hours for submitted row
    ' -------------------------
    startTime = ws.Cells(nextRow, "G").Value
    endTime = ws.Cells(nextRow, "H").Value
    otType = LCase(Trim(ws.Cells(nextRow, "E").Value)) ' earn or use
    userID = ws.Cells(nextRow, "C").Value
    
    ' Find the column in row 3 where the User ID matches
    On Error Resume Next
    Set headerCell = ws.Rows(3).Find(What:=userID, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If IsNumeric(startTime) And IsNumeric(endTime) Then
        If Not headerCell Is Nothing Then
            ' Only round if column E = "earn", date is weekend, header is green
            If otType = "earn" And Weekday(ws.Cells(nextRow, "F").Value, vbMonday) > 5 And _
               headerCell.Interior.Color = RGB(112, 173, 71) Then
               
                otHours = (endTime - startTime) * 24
                If otHours < 4 Then
                    ws.Cells(nextRow, "I").Value = 4
                Else
                    ws.Cells(nextRow, "I").Value = 8
                End If
            Else
                ' "use" OT or non-weekend / non-green header
                If otType = "use" Then
                    ws.Cells(nextRow, "I").Value = -((endTime - startTime) * 24)
                Else
                    ws.Cells(nextRow, "I").Value = (endTime - startTime) * 24
                End If
            End If
        Else
            ' If User ID not found in row 3
            If otType = "use" Then
                ws.Cells(nextRow, "I").Value = -((endTime - startTime) * 24)
            Else
                ws.Cells(nextRow, "I").Value = (endTime - startTime) * 24
            End If
        End If
    Else
        ws.Cells(nextRow, "I").ClearContents
    End If
    
    ' Clear input fields in row 6 except I6
    ws.Range("E6:J6").ClearContents
    Call InitI6Formula ' Reapply formula in I6 for next entry
    
    ' Lock the submitted row except for Approve/Reject buttons
    ws.Rows(nextRow).Locked = True
    ws.Range("K" & nextRow & ":L" & nextRow).Locked = False
    
    ' Re-protect sheet
    ws.Protect Password:="Pass8371!", UserInterfaceOnly:=True
End Sub
' -------------------------------
' Add Approve & Reject buttons
' -------------------------------
Sub AddApproveRejectButtons(ByVal rowNum As Long)
    Dim ws As Worksheet
    Dim btnApprove As Shape, btnReject As Shape
    Dim cellLeft As Double, cellTop As Double, cellWidth As Double, cellHeight As Double
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Avoid duplicates
    On Error Resume Next
    ws.Shapes("ApproveBtn_" & rowNum).Delete
    ws.Shapes("RejectBtn_" & rowNum).Delete
    On Error GoTo 0
    
    ' Cell position in column K (Approve) and L (Reject)
    cellLeft = ws.Cells(rowNum, "K").Left
    cellTop = ws.Cells(rowNum, "K").Top
    cellWidth = ws.Cells(rowNum, "K").Width
    cellHeight = ws.Cells(rowNum, "K").Height
    
    ' ----- Approve button -----
    Set btnApprove = ws.Shapes.AddShape(msoShapeRoundedRectangle, cellLeft, cellTop, cellWidth, cellHeight)
    With btnApprove
        .TextFrame.Characters.Text = "Approve"
        .Fill.ForeColor.RGB = RGB(0, 176, 80)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Name = "ApproveBtn_" & rowNum
        .OnAction = "ApproveAction"
    End With
    
    ' ----- Reject button -----
    cellLeft = ws.Cells(rowNum, "L").Left
    Set btnReject = ws.Shapes.AddShape(msoShapeRoundedRectangle, cellLeft, cellTop, cellWidth, cellHeight)
    With btnReject
        .TextFrame.Characters.Text = "Reject"
        .Fill.ForeColor.RGB = RGB(255, 0, 0)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Name = "RejectBtn_" & rowNum
        .OnAction = "RejectAction"
    End With
End Sub

' -------------------------------
' Approve action
' -------------------------------
Sub ApproveAction()
    Dim ws As Worksheet
    Dim btnName As String
    Dim rowNum As Long
    Dim pwd As String
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    btnName = Application.Caller
    rowNum = CLng(Replace(btnName, "ApproveBtn_", ""))
    
    ' Prompt for password
    pwd = InputBox("Enter password to approve:", "Approval Required")
    If pwd <> "Pass8371!" Then
        MsgBox "Incorrect password! Action cancelled.", vbCritical
        Exit Sub
    End If
    
    ' Unprotect with password
    ws.Unprotect Password:="Pass8371!"
    
    ' Log approver
    ws.Cells(rowNum, "M").Value = Application.UserName
    ws.Cells(rowNum, "M").Interior.Color = RGB(198, 239, 206)
    
    ' Update status
    ws.Cells(rowNum, "N").Value = "Approved"
    ws.Cells(rowNum, "N").Interior.Color = RGB(198, 239, 206)
    
    ' Log date/time
    ws.Cells(rowNum, "O").Value = Now
    ws.Cells(rowNum, "O").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(rowNum, "O").Interior.Color = RGB(198, 239, 206)
    
    ' Highlight row
    ws.Range("C" & rowNum & ":J" & rowNum).Interior.Color = RGB(198, 239, 206)
    
    ' Disable buttons for this row
    With ws.Shapes("ApproveBtn_" & rowNum)
        .OnAction = ""
        .Fill.Transparency = 0.3
    End With
    With ws.Shapes("RejectBtn_" & rowNum)
        .OnAction = ""
        .Fill.Transparency = 0.3
    End With
    
    ' Lock entire row
    ws.Rows(rowNum).Locked = True
    
    ' Re-protect sheet
    ws.Protect Password:="Pass8371!", UserInterfaceOnly:=True
End Sub

' -------------------------------
' Reject action
' -------------------------------
Sub RejectAction()
    Dim ws As Worksheet
    Dim btnName As String
    Dim rowNum As Long
    Dim pwd As String
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    btnName = Application.Caller
    rowNum = CLng(Replace(btnName, "RejectBtn_", ""))
    
    ' Prompt for password
    pwd = InputBox("Enter password to reject:", "Approval Required")
    If pwd <> "Pass8371!" Then
        MsgBox "Incorrect password! Action cancelled.", vbCritical
        Exit Sub
    End If
    
    ' Unprotect with password
    ws.Unprotect Password:="Pass8371!"
    
    ' Log rejector
    ws.Cells(rowNum, "M").Value = Application.UserName
    ws.Cells(rowNum, "M").Interior.Color = RGB(255, 199, 206)
    
    ' Update status
    ws.Cells(rowNum, "N").Value = "Rejected"
    ws.Cells(rowNum, "N").Interior.Color = RGB(255, 199, 206)
    
    ' Log date/time
    ws.Cells(rowNum, "O").Value = Now
    ws.Cells(rowNum, "O").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(rowNum, "O").Interior.Color = RGB(255, 199, 206)
    
    ' Highlight row
    ws.Range("C" & rowNum & ":J" & rowNum).Interior.Color = RGB(255, 199, 206)
    
    ' Disable buttons for this row
    With ws.Shapes("ApproveBtn_" & rowNum)
        .OnAction = ""
        .Fill.Transparency = 0.3
    End With
    With ws.Shapes("RejectBtn_" & rowNum)
        .OnAction = ""
        .Fill.Transparency = 0.3
    End With
    
    ' Lock entire row
    ws.Rows(rowNum).Locked = True
    
    ' Re-protect sheet
    ws.Protect Password:="Pass8371!", UserInterfaceOnly:=True
End Sub



