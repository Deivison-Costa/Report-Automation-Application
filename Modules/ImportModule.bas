Attribute VB_Name = "ImportModule"
Option Explicit

Sub ImportDatas()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileDatas As String
    
    fileDatas = ThisWorkbook.Path & "\testData.csv"
    
    Set wb = Workbooks.Open(fileDatas)
    Set ws = wb.Worksheets(1)
    
    ' Movendo os dados para a planilha ativa do Excel
    ws.UsedRange.Copy ThisWorkbook.ActiveSheet.Range("A1")
    
    wb.Close SaveChanges:=False
End Sub
