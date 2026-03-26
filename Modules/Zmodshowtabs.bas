Attribute VB_Name = "Zmodshowtabs"
Option Explicit

Sub Afficheronglet()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets: ws.Visible = xlSheetVisible: Next ws
    ThisWorkbook.Worksheets(1).Activate
End Sub
