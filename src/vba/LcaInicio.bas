Attribute VB_Name = "LcaInicio"
Option Explicit

Public Const LCA_VERSION As String = "P03"

' Entry point for the initial workbook. No CPU simulation is implemented here.
Public Sub VerificarInstalacion()
    Dim inicio As Worksheet
    Set inicio = ThisWorkbook.Worksheets("Inicio")
    inicio.Range("B13").Value = "VBA OK - " & LCA_VERSION
End Sub

