Attribute VB_Name = "GraphicsModule"
Option Explicit

Sub CreateGraphics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Criando um gráfico de colunas com base nos dados da coluna B
    Dim chartObj As ChartObject
    Dim rngDatas As Range
    
    Set rngDatas = ws.Range("B:B")
    Set chartObj = ws.ChartObjects.Add(Left:=100, Width:=375, Top:=75, Height:=225)
    
    With chartObj.Chart
        .SetSourceData Source:=rngDatas
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Columns Graphic"
    End With
End Sub

