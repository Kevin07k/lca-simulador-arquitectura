Attribute VB_Name = "LcaInicio"
Option Explicit

Public Const LCA_VERSION As String = "P03"

Private mSilent As Boolean
Private mLastPassed As Boolean
Private mLastReport As String

' Entry point for the initial workbook. No CPU simulation is implemented here.
Public Sub VerificarInstalacion()
    Dim inicio As Worksheet
    Set inicio = ThisWorkbook.Worksheets("Inicio")
    inicio.Range("B13").Value = "VBA OK - " & LCA_VERSION
End Sub

' Every Verificar* macro reports here: a MsgBox for a person, silent inside VerificarTodo/scripts.
Public Sub LcaReport(ByVal text As String, ByVal passed As Boolean)
    mLastPassed = passed
    mLastReport = text
    Debug.Print text
    If mSilent Then Exit Sub
    If passed Then
        MsgBox text, vbInformation, "LCA"
    Else
        MsgBox text, vbCritical, "LCA"
    End If
End Sub

' Result of the last suite that reported (used by panel buttons).
Public Function LcaLastPassed() As Boolean
    LcaLastPassed = mLastPassed
End Function

Public Function LcaLastReport() As String
    LcaLastReport = mLastReport
End Function

' Runs every engine regression (P04..P14) without dialogs and returns one line per suite.
' First line is "ALL PASS" or "SOME FAIL". Used by VerificarTodo and scripts/Test-Workbook.ps1.
Public Function LcaVerificarTodoTexto() As String
    Dim suites As Variant, i As Long, out As String, allOk As Boolean
    suites = Array("VerificarCpuAlu", "VerificarBusesP05", "VerificarRamP06", "VerificarMmuP07", _
                   "VerificarPagerP08", "VerificarCachesP09", "VerificarCachePaginacionP10", _
                   "VerificarEnsambladorP11", "VerificarCicloP12", "VerificarTecladoP13", "VerificarLcdP14")
    allOk = True
    mSilent = True
    For i = LBound(suites) To UBound(suites)
        mLastPassed = False
        mLastReport = CStr(suites(i)) & ": sin reporte / no report"
        On Error Resume Next
        Application.Run "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & CStr(suites(i))
        If Err.Number <> 0 Then
            mLastPassed = False
            mLastReport = CStr(suites(i)) & " ERROR: " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0
        out = out & IIf(mLastPassed, "PASS  ", "FAIL  ") & mLastReport & vbCrLf
        If Not mLastPassed Then allOk = False
    Next i
    mSilent = False
    LcaVerificarTodoTexto = IIf(allOk, "ALL PASS", "SOME FAIL") & vbCrLf & out
End Function

Public Sub VerificarTodo()
    Dim text As String
    text = LcaVerificarTodoTexto()
    If Left$(text, 8) = "ALL PASS" Then
        MsgBox text, vbInformation, "LCA - P04..P14"
    Else
        MsgBox text, vbCritical, "LCA - P04..P14"
    End If
End Sub

