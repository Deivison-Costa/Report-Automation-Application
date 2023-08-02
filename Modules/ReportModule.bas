Attribute VB_Name = "ReportModule"
Option Explicit

Sub GenerateReport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Inserindo os resultados das an�lises na planilha
    ws.Range("A10").Value = "Sum of the values:"
    ws.Range("B10").Value = sumB ' Vari�vel calculada na rotina de an�lise de dados
    
    ws.Range("A11").Value = "Max value:"
    ws.Range("B11").Value = maxB ' Vari�vel calculada na rotina de an�lise de dados
    
    ws.Range("A12").Value = "Mean:"
    ws.Range("B12").Value = meanB ' Vari�vel calculada na rotina de an�lise de dados
    
    ' Criando um relat�rio com o gr�fico criado anteriormente
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects(1)
    chartObj.Top = ws.Range("A14").Top
    chartObj.Left = ws.Range("A14").Left
End Sub

