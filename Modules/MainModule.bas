Attribute VB_Name = "MainModule"
Option Explicit

Sub Main()
    On Error Resume Next
    
    ImportDatas
    CleanAndFormatData
    DataAnalysis
    CreateGraphics
    GenerateReport
    
    On Error GoTo 0
End Sub

