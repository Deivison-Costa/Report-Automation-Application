Attribute VB_Name = "FormatCleanerModule"
Option Explicit

Sub CleanAndFormatData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Removendo duplicatas da coluna A
    ws.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
    
    ' Formatando coluna B como número
    ws.Range("B:B").NumberFormat = "0.00"
End Sub

