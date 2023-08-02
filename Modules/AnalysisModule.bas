Attribute VB_Name = "AnalysisModule"
Option Explicit

Sub DataAnalysis()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Calculando a soma dos valores da coluna B
    Dim sumB As Double
    sumB = Application.WorksheetFunction.sum(ws.Range("B:B"))
    
    ' Encontrando o valor m�ximo na coluna B
    Dim maxB As Double
    maxB = Application.WorksheetFunction.Max(ws.Range("B:B"))
    
    ' Calculando a m�dia na coluna B
    Dim meanB As Double
    meanB = Application.WorksheetFunction.Average(ws.Range("B:B"))
End Sub

