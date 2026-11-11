Attribute VB_Name = "LcaMmuPanel"
Option Explicit

Private Const ROW_BASE As Long = LCA_BASE_MMU
Private Const COL_BASE As Long = LCA_COL_MMU
Private mStatus As String

Public Sub ConstruirPanelPaginas()
    Dim ws As Worksheet, i As Long
    Set ws = MmuSheet(True)
    LcaPrepareBand ws, ROW_BASE, 33, "P07_", COL_BASE
    LcaRowHeight ws.Rows(RA("1:33")), 24
    ws.Range(RA("A1:S33")).Interior.Color = RGB(15, 23, 42)
    With ws.Range(RA("A1:S33")).Font
        .Name = "Consolas"
        .Color = RGB(226, 232, 240)
        .Size = 11
    End With
    ws.Range(RA("B2:S3")).Merge
    With ws.Range(RA("B2"))
        .Value = "P07  /  TABLA DE PAGINAS + MMU"
        .Font.Name = "Segoe UI"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(125, 211, 252)
    End With
    ws.Range(RA("B4:S4")).Merge
    ws.Range(RA("B4")).Value = "512 bytes virtuales / 16 paginas de 32 bytes / 7 marcos asignables + 1 reservado E/S"
    ws.Range(RA("B6:D6")).Merge
    ws.Range(RA("B6")).Value = "VA DECIMAL (0-511)"
    ws.Range(RA("F6:G6")).Merge
    ws.Range(RA("F6")).Value = "FETCH / READ / WRITE"
    ws.Range(RA("B7:D8")).Merge
    ws.Range(RA("F7:G8")).Merge
    ws.Range(RA("B7:D8")).Interior.Color = RGB(230, 240, 250)
    ws.Range(RA("F7:G8")).Interior.Color = RGB(230, 240, 250)
    ws.Range(RA("B7:D8")).Font.Color = RGB(15, 23, 42)
    ws.Range(RA("F7:G8")).Font.Color = RGB(15, 23, 42)
    ws.Range(RA("B7")).Value = 300
    ws.Range(RA("F7")).Value = "READ"
    AddMmuButton ws, "H6:K7", "RESET PTE", "MmuPanelReset"
    AddMmuButton ws, "L6:O7", "EJEMPLO", "MmuDemoSetup"
    AddMmuButton ws, "H9:K10", "TRADUCIR", "MmuPanelTranslate"
    AddMmuButton ws, "L9:O10", "FETCH REAL", "MmuPanelFetch"
    AddMmuButton ws, "P9:S10", "ST REAL", "MmuPanelStore"
    ws.Range(RA("B10:G10")).Merge
    ws.Range(RA("B10:G10")).WrapText = True
    ws.Range(RA("B12")).Value = "VPN"
    ws.Range(RA("C12")).Value = "VA RANGE"
    ws.Range(RA("D12")).Value = "P"
    ws.Range(RA("E12")).Value = "PFN"
    ws.Range(RA("F12")).Value = "DIRTY"
    ws.Range(RA("G12")).Value = "REF"
    ws.Range(RA("H12")).Value = "RESERVED"
    ws.Range(RA("B12:H12")).Interior.Color = RGB(30, 64, 105)
    ws.Range(RA("B12:H28")).Borders.Color = RGB(51, 65, 85)
    ws.Range(RA("J12:S12")).Merge
    ws.Range(RA("J12")).Value = "ULTIMA TRAZA / LAST BUS CYCLE"
    ws.Range(RA("J13:S16")).Merge
    ws.Range(RA("J13:S16")).WrapText = True
    ws.Range(RA("J13:S16")).Interior.Color = RGB(30, 41, 59)
    ws.Range(RA("J18:S18")).Merge
    ws.Range(RA("J18")).Value = "FALLO RECUPERABLE / PAGE FAULT"
    ws.Range(RA("J19:S20")).Merge
    ws.Range(RA("J19:S20")).Interior.Color = RGB(30, 41, 59)
    ws.Range(RA("J29:S32")).Merge
    ws.Range(RA("J29:S32")).WrapText = True
    ws.Range(RA("J29")).Value = "VPN 15 es E/S reservada. PAGE_FAULT no asigna un marco: P08 cargara la pagina y reintentara. " & _
                           "FETCH y ST usan la RAM P06 y los buses P05; P12 ejecutara el opcode."
    MmuDemoSetup
    LcaGoto ws, ROW_BASE, COL_BASE
End Sub

Public Sub MostrarPanelPaginas()
    Dim ws As Worksheet
    Set ws = MmuSheet(False)
    If ws Is Nothing Then
        ConstruirPanelPaginas
    Else
        LcaGoto ws, ROW_BASE, COL_BASE
        MmuPanelRefresh
    End If
End Sub

Public Sub MmuPanelReset()
    MmuReset
    mStatus = "16 PTE ausentes; VPN 15 RESERVADA. EJEMPLO para cargar mapas de demostracion."
    MmuPanelRefresh
End Sub

Public Sub MmuDemoSetup()
    CpuReset
    RamReset
    MemoryPathReset
    MmuReset
    MmuMap 0, 0
    MmuMap 1, 1
    MmuMap 7, 3
    MmuMap 8, 4
    MmuMap 9, 2
    RamWriteUser 32, &H10
    RamWriteUser 33, 0
    RamWriteUser 34, &H2A
    RamWriteUser 35, 0
    CpuWriteControl LcaPC, 32
    CpuWriteRegister LcaAX, &H2A
    mStatus = "VPN 9 -> PFN 2: VA 300 -> PA 76. FETCH usa VA 32 -> PA 32; ST usa VA 300."
    PanelActualizar
    MmuPanelRefresh
End Sub

Public Sub MmuPanelTranslate()
    Dim ws As Worksheet, number As Double, va As Long, pa As Long, kind As Long
    Dim choice As String
    On Error GoTo failed
    Set ws = MmuSheet(False)
    If ws Is Nothing Then Err.Raise 5, "LcaMmuPanel", "ConstruirPanelPaginas primero"
    If IsError(ws.Range(RA("B7")).Value) Then Err.Raise 5, "LcaMmuPanel", "VA debe ser un entero"
    If Not IsNumeric(ws.Range(RA("B7")).Value) Then Err.Raise 5, "LcaMmuPanel", "VA debe ser un entero"
    number = CDbl(ws.Range(RA("B7")).Value)
    If number <> Fix(number) Or number < -1000000# Or number > 1000000# Then _
        Err.Raise 5, "LcaMmuPanel", "VA debe ser un entero cercano a 0-511"
    va = CLng(number)
    choice = UCase$(Trim$(CStr(ws.Range(RA("F7")).Value)))
    Select Case choice
        Case "FETCH": kind = MMU_FETCH
        Case "READ": kind = MMU_READ
        Case "WRITE": kind = MMU_WRITE
        Case Else: Err.Raise 5, "LcaMmuPanel", "Acceso: FETCH, READ o WRITE"
    End Select
    If MmuTryTranslate(va, kind, pa) Then
        mStatus = "VA " & Hex3(va) & " -> VPN " & CStr(MmuVpnOf(va)) & _
                  " / offset " & CStr(MmuOffsetOf(va)) & " -> PA " & Hex3(pa) & ". Sin transferencia."
    Else
        mStatus = "PAGE_FAULT: VPN " & CStr(MmuLastFaultVPN()) & " ausente; PA = -1. P08 atendera el fault."
    End If
    MmuPanelRefresh
    Exit Sub
failed:
    mStatus = "ERROR: " & Err.Description
    MmuPanelRefresh
End Sub

Public Sub MmuPanelFetch()
    On Error GoTo failed
    MmuFetchFour 32
    mStatus = "FETCH VA 020h -> PA 020h: IR=" & Hex3(CpuIrValue()) & "; VPN 1.ref=1; PC sin avance."
    PanelActualizar
    BusPanelActualizar
    MmuPanelRefresh
    Exit Sub
failed:
    mStatus = "ERROR FETCH: " & Err.Description
    MmuPanelRefresh
End Sub

Public Sub MmuPanelStore()
    On Error GoTo failed
    MmuStoreByte 300, LcaAX
    mStatus = "ST VA 12Ch -> PA 04Ch: RAM[76]=" & Hex3(RamReadUser(76)) & "; VPN 9.dirty/ref=1."
    PanelActualizar
    BusPanelActualizar
    MmuPanelRefresh
    Exit Sub
failed:
    mStatus = "ERROR ST: " & Err.Description
    MmuPanelRefresh
End Sub

Public Sub MmuPanelRefresh()
    Dim ws As Worksheet, vpn As Long, row As Long, entry As LcaPageEntry, n As Long
    Set ws = MmuSheet(False)
    If ws Is Nothing Then Exit Sub
    For vpn = 0 To 15
        entry = MmuPage(vpn)
        row = 13 + vpn
        CL(ws, ROW_BASE + row, 2).Value = vpn
        CL(ws, ROW_BASE + row, 3).Value = Hex3(vpn * 32) & " - " & Hex3(vpn * 32 + 31)
        CL(ws, ROW_BASE + row, 4).Value = Abs(CLng(entry.Present))
        CL(ws, ROW_BASE + row, 5).Value = IIf(entry.Present, CStr(entry.PFN), "-")
        CL(ws, ROW_BASE + row, 6).Value = Abs(CLng(entry.Dirty))
        CL(ws, ROW_BASE + row, 7).Value = Abs(CLng(entry.Referenced))
        CL(ws, ROW_BASE + row, 8).Value = IIf(entry.Reserved, "IO / RESERVED", "-")
        ws.Range(CL(ws, ROW_BASE + row, 2), CL(ws, ROW_BASE + row, 8)).Interior.Color = _
            IIf(entry.Reserved, RGB(90, 45, 57), IIf(entry.Present, RGB(30, 65, 84), RGB(30, 41, 59)))
    Next vpn
    ws.Range(RA("B10")).Value = mStatus
    n = BusTraceCount()
    If n > 0 Then
        ws.Range(RA("J13")).Value = BusTraceLine(n - 1)
    Else
        ws.Range(RA("J13")).Value = "High-Z / sin transferencias"
    End If
    If MmuLastFaultVPN() >= 0 Then
        ws.Range(RA("J19")).Value = "VA=" & Hex3(MmuLastFaultVA()) & " / VPN=" & MmuLastFaultVPN() & " / PA=-1"
    Else
        ws.Range(RA("J19")).Value = "Ninguno / None"
    End If
    MmuDetalleActualizar
End Sub

Private Function CpuIrValue() As Long
    Dim state As LcaCpuState
    state = CpuSnapshot()
    CpuIrValue = state.IR
End Function

Private Function Hex3(ByVal value As Long) As String
    Hex3 = Right$("000" & Hex$(value), 3) & "h"
End Function

Private Function MmuSheet(ByVal create As Boolean) As Worksheet
    Set MmuSheet = LcaBand("P07_MmuPanelReset", create)
End Function

' Row shift of this panel inside the shared LCA sheet.
Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

Private Sub AddMmuButton(ByVal ws As Worksheet, ByVal address As String, _
                         ByVal caption As String, ByVal macroName As String)
    Dim area As Range, button As Shape
    Set area = ws.Range(RA(address))
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width - 4, area.Height - 2)
    button.Name = "P07_" & macroName
    button.Fill.ForeColor.RGB = RGB(37, 99, 235)
    button.Line.Visible = msoFalse
    button.TextFrame2.TextRange.Text = caption
    button.TextFrame2.TextRange.Font.Name = "Segoe UI"
    button.TextFrame2.TextRange.Font.Size = 12
    button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    button.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    button.TextFrame2.VerticalAnchor = msoAnchorMiddle
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub

' Cell of this block: row is absolute (ROW_BASE already added), column is relative to the block.
Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function
