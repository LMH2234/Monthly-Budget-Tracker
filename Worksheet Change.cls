VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Only update if changes are in the log area (columns H to M, rows 2+)
    If Not Intersect(Target, Me.Range("H:M")) Is Nothing Then
        If Target.Row >= 2 Then
            ' Disable events to prevent multiple triggers
            Application.EnableEvents = False
            
            ' Wait a moment for all cells to be filled
            Application.OnTime Now + TimeValue("00:00:01"), "UpdateBudgetAll"
            
            Application.EnableEvents = True
        End If
    End If
End Sub

