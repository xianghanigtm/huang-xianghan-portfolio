VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim cell As Range
    
    Set ws = Me
    
    ' 1) If Column H changes (row 9 onwards) ? add Approve & Reject buttons
    If Not Intersect(Target, ws.Columns("H")) Is Nothing Then
        Application.EnableEvents = False
        For Each cell In Intersect(Target, ws.Columns("H"))
            If cell.Row >= 9 And cell.Value <> "" Then
                Call AddApproveRejectButtons(cell.Row)
            End If
        Next cell
        Application.EnableEvents = True
    End If
End Sub

