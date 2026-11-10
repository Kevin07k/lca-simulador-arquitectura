Attribute VB_Name = "LcaCachePath"
Option Explicit

' P10: integrated memory path  VA -> pager (residency first) -> L1I/L1D -> L2 -> RAM.
' Order for every CPU request (P02 section 3/4):
'   1. validate VA and access kind (P07);  2. resolve a PAGE_FAULT if the VPN is absent (P08),
'      which invalidates every cache line of a reused frame before writeback/load;
'   3. only then touch the caches (P09) and emit the real bus cycles (P05/P06).
' I/O (VPN 15 / E0..EFh) never reaches this path: it is rejected as PROTECTION in step 1.
Public Const LCA_PATH_FETCH As Long = 1
Public Const LCA_PATH_READ As Long = 2
Public Const LCA_PATH_WRITE As Long = 3
Public Const LCA_ERR_CACHE_PATH As Long = vbObjectError + 2101
Public Const LCA_P10_STEPS As Long = 13
Private Const TRACE_MAX As Long = 64

' Summary of the last request, for the panel and the tests (no worksheet dependency).
Public Type LcaPathStep
    Kind As Long
    VA As Long
    VPN As Long
    PA As Long                  ' first physical byte, -1 when the request failed before translation
    Value As Long               ' first byte fetched/read/written, -1 when none
    PageFault As Boolean
    PFN As Long                 ' frame used, -1 when not reached
    VictimVPN As Long           ' -1 when no victim (free frame or no fault)
    InvalidatedLines As Long
    Writeback As Boolean
    L1Hits As Long
    L1Misses As Long
    L2Hits As Long
    L2Misses As Long
    RamCycles As Long           ' external RAM bus cycles: 4 per line fill + 1 per ST
    Failed As Boolean
    Outcome As String
End Type

Private mReady As Boolean
Private mLast As LcaPathStep
Private mAccessCount As Long
Private mRamCycles As Long
Private mBase(0 To 8) As Long
Private mTrace(0 To TRACE_MAX - 1) As String
Private mTraceCount As Long

' Resets bus/control, RAM, caches, page table and swap. Callers reset the CPU separately
' (same convention as PagerReset) and seed the backing image with SwapWriteByte afterwards.
Public Sub CachePathReset()
    Dim i As Long
    MemoryPathReset
    PagerReset
    mAccessCount = 0
    mRamCycles = 0
    mTraceCount = 0
    For i = 0 To TRACE_MAX - 1
        mTrace(i) = vbNullString
    Next i
    mLast = EmptyStep(0, -1)
    mLast.Outcome = "RESET"
    mReady = True
End Sub

' Instruction fetch of four bytes through L1I. All four translations (and any page fault)
' complete before the first cache lookup; an L2 miss emits the four RAM fill cycles.
Public Sub CachePathFetchFour(ByVal va As Long)
    Dim pa As Long, i As Long, value As Byte
    Dim errNo As Long, errSrc As String, errText As String
    On Error GoTo failed
    BeginStep LCA_PATH_FETCH, va
    If va < 0 Or va > 252 Or (va Mod 4) <> 0 Then _
        Err.Raise LCA_ERR_CACHE_PATH, "LcaCachePath", "INVALID_FETCH_PC"
    pa = PagerResolve(va, MMU_FETCH)
    For i = 1 To 3
        If PagerResolve(va + i, MMU_FETCH) <> pa + i Then _
            Err.Raise LCA_ERR_CACHE_PATH, "LcaCachePath", "FETCH_NOT_CONTIGUOUS"
    Next i
    mLast.PA = pa
    ControlStartFetch va, pa
    For i = 0 To 3
        value = CacheRead(pa + i, LCA_CACHE_FETCH)
        If CacheLastFillPA() >= 0 Then ControlLineFill va + i, CacheLastFillPA()
        ControlReadCacheByte value
        ControlLatchFetchByte
        MmuCommitAccess va + i, MMU_FETCH, pa + i
        If i = 0 Then mLast.Value = value
    Next i
    EndStep 0
    Exit Sub
failed:
    errNo = Err.Number: errSrc = Err.Source: errText = Err.Description
    FailStep errText
    Err.Raise errNo, errSrc, errText
End Sub

' Data read (LD path) through L1D. Returns the byte delivered to MDR.
Public Function CachePathReadByte(ByVal va As Long) As Byte
    Dim pa As Long, value As Byte
    Dim errNo As Long, errSrc As String, errText As String
    On Error GoTo failed
    BeginStep LCA_PATH_READ, va
    pa = PagerResolve(va, MMU_READ)
    mLast.PA = pa
    value = CacheRead(pa, LCA_CACHE_DATA)
    If CacheLastFillPA() >= 0 Then ControlLineFill va, CacheLastFillPA()
    ControlReadDataByte va, pa, value
    MmuCommitAccess va, MMU_READ, pa
    mLast.Value = value
    EndStep 0
    CachePathReadByte = value
    Exit Function
failed:
    errNo = Err.Number: errSrc = Err.Source: errText = Err.Description
    FailStep errText
    Err.Raise errNo, errSrc, errText
End Function

' Data write (ST path): register -> MDR, write-allocate fill if L1D and L2 miss, one external
' WRITE cycle, then write-through to L1D/L2/RAM. The faulting ST is resolved before any bus cycle
' and is never replayed; dirty/ref change only after the RAM write succeeded.
Public Sub CachePathStoreByte(ByVal va As Long, ByVal sourceRegister As Long)
    Dim pa As Long, value As Byte, transferred As Long
    Dim errNo As Long, errSrc As String, errText As String
    On Error GoTo failed
    BeginStep LCA_PATH_WRITE, va
    value = CpuReadRegister(sourceRegister)   ' a bad register fails before any page fault
    pa = PagerResolve(va, MMU_WRITE)
    mLast.PA = pa
    ControlPrepareStore va, pa, sourceRegister
    If Not CacheContains(LCA_CACHE_L1D, pa) And Not CacheContains(LCA_CACHE_L2, pa) Then _
        ControlLineFill va, (pa \ LCA_CACHE_LINE_BYTES) * LCA_CACHE_LINE_BYTES
    transferred = ControlCommitStore()
    CacheWrite pa, transferred
    MmuCommitAccess va, MMU_WRITE, pa
    mLast.Value = transferred
    EndStep 1
    Exit Sub
failed:
    errNo = Err.Number: errSrc = Err.Source: errText = Err.Description
    FailStep errText
    Err.Raise errNo, errSrc, errText
End Sub

' --- Reproducible P10 scenario (shared by the panel and VerificarCachePaginacionP10) -------
' Backing image: VPN0 code 10 00 2A 00; data VPN8..14 byte = VPN + offset; AX = A5h.
' Data reads use offsets >= 4 so the dirty VPN8 line stays in L2 index 0 until step 11.

Public Sub CachePathScenarioReset()
    Dim vpn As Long, offset As Long
    CpuReset
    CachePathReset
    For vpn = 8 To 14
        For offset = 0 To 31
            SwapWriteByte vpn, offset, vpn + offset
        Next offset
    Next vpn
    SwapWriteByte 0, 0, &H10
    SwapWriteByte 0, 1, &H0
    SwapWriteByte 0, 2, &H2A
    SwapWriteByte 0, 3, &H0
    CpuWriteRegister LcaAX, &HA5
End Sub

' Runs scenario step 0..12 and returns its description. Step 12 raises PROTECTION (expected).
Public Function CachePathScenarioStep(ByVal stepIndex As Long) As String
    Dim value As Byte
    Select Case stepIndex
        Case 0, 10: CachePathFetchFour &H0
        Case 1: CachePathStoreByte &H100, LcaAX
        Case 2, 11: value = CachePathReadByte(&H100)
        Case 3, 6: value = CachePathReadByte(&H124)
        Case 4: value = CachePathReadByte(&H148)
        Case 5: value = CachePathReadByte(&H16C)
        Case 7: value = CachePathReadByte(&H190)
        Case 8: value = CachePathReadByte(&H1B4)
        Case 9: value = CachePathReadByte(&H1D8)
        Case 12: value = CachePathReadByte(&H1E8)
        Case Else: Err.Raise LCA_ERR_CACHE_PATH, "LcaCachePath", "SCENARIO_STEP"
    End Select
    CachePathScenarioStep = CachePathScenarioText(stepIndex)
End Function

Public Function CachePathScenarioText(ByVal stepIndex As Long) As String
    Select Case stepIndex
        Case 0: CachePathScenarioText = "FETCH 000h: fallo de pagina, VPN0 -> PFN0; L1I 1 miss + 3 hits, L2 llena 00-03h."
        Case 1: CachePathScenarioText = "ST 100h = A5h: VPN8 -> PFN1, write-allocate + write-through; RAM[20h] = A5h, dirty = 1."
        Case 2: CachePathScenarioText = "READ 100h: L1D HIT, sin ciclo de RAM / no RAM cycle."
        Case 3: CachePathScenarioText = "READ 124h: VPN9 -> PFN2."
        Case 4: CachePathScenarioText = "READ 148h: VPN10 -> PFN3."
        Case 5: CachePathScenarioText = "READ 16Ch: VPN11 -> PFN4; sobrescribe la linea 1 de L1D."
        Case 6: CachePathScenarioText = "READ 124h otra vez: L1D miss por conflicto, L2 HIT."
        Case 7: CachePathScenarioText = "READ 190h: VPN12 -> PFN5."
        Case 8: CachePathScenarioText = "READ 1B4h: VPN13 -> PFN6. Los siete marcos estan llenos."
        Case 9: CachePathScenarioText = "READ 1D8h: octava pagina; CLOCK expulsa VPN0 (limpia) e invalida su linea L1I."
        Case 10: CachePathScenarioText = "FETCH 000h: expulsa VPN8 (dirty) de PFN1, invalida su linea L2 (A5h) y escribe 32 B a swap. IR = 10h, no A5h."
        Case 11: CachePathScenarioText = "READ 100h: VPN8 vuelve desde swap a PFN2 y devuelve A5h (dato modificado recuperado)."
        Case 12: CachePathScenarioText = "READ 1E8h (LCD): PROTECTION. La E/S no pasa por paginacion ni cache."
        Case Else: CachePathScenarioText = "-"
    End Select
End Function

' --- Inspection --------------------------------------------------------------------------

Public Function CachePathLast() As LcaPathStep
    CachePathLast = mLast
End Function

Public Function CachePathAccessCount() As Long
    CachePathAccessCount = mAccessCount
End Function

Public Function CachePathRamCycles() As Long
    CachePathRamCycles = mRamCycles
End Function

Public Function CachePathReady() As Boolean
    CachePathReady = mReady
End Function

Public Function CachePathKindName(ByVal kind As Long) As String
    Select Case kind
        Case LCA_PATH_FETCH: CachePathKindName = "FETCH"
        Case LCA_PATH_READ: CachePathKindName = "READ"
        Case LCA_PATH_WRITE: CachePathKindName = "ST"
        Case Else: CachePathKindName = "-"
    End Select
End Function

Public Function CachePathTraceCount() As Long
    CachePathTraceCount = mTraceCount
    If CachePathTraceCount > TRACE_MAX Then CachePathTraceCount = TRACE_MAX
End Function

' index 0 = newest retained line.
Public Function CachePathTraceLine(ByVal index As Long) As String
    If index < 0 Or index >= CachePathTraceCount() Then _
        Err.Raise LCA_ERR_CACHE_PATH, "LcaCachePath", "TRACE_INDEX"
    CachePathTraceLine = mTrace((mTraceCount - 1 - index) Mod TRACE_MAX)
End Function

' --- Internals -----------------------------------------------------------------------------

Private Sub BeginStep(ByVal kind As Long, ByVal va As Long)
    Dim l1 As Long
    If Not mReady Then Err.Raise LCA_ERR_CACHE_PATH, "LcaCachePath", "CACHE_PATH_RESET_REQUIRED"
    mLast = EmptyStep(kind, va)
    If va >= 0 And va <= 511 Then mLast.VPN = va \ 32
    If kind = LCA_PATH_FETCH Then l1 = LCA_CACHE_L1I Else l1 = LCA_CACHE_L1D
    mBase(0) = PagerFaultCount()
    mBase(1) = PagerInvalidatedLines()
    mBase(2) = PagerWritebackCount()
    mBase(3) = CacheHits(l1)
    mBase(4) = CacheMisses(l1)
    mBase(5) = CacheHits(LCA_CACHE_L2)
    mBase(6) = CacheMisses(LCA_CACHE_L2)
    mBase(7) = CacheRamFills()
    mBase(8) = PagerEvictionCount()
End Sub

Private Sub EndStep(ByVal writeCycles As Long)
    Dim l1 As Long
    If mLast.Kind = LCA_PATH_FETCH Then l1 = LCA_CACHE_L1I Else l1 = LCA_CACHE_L1D
    With mLast
        .PageFault = PagerFaultCount() > mBase(0)
        If .PageFault And PagerEvictionCount() > mBase(8) Then .VictimVPN = PagerLastVictimVPN()
        .InvalidatedLines = PagerInvalidatedLines() - mBase(1)
        .Writeback = PagerWritebackCount() > mBase(2)
        .L1Hits = CacheHits(l1) - mBase(3)
        .L1Misses = CacheMisses(l1) - mBase(4)
        .L2Hits = CacheHits(LCA_CACHE_L2) - mBase(5)
        .L2Misses = CacheMisses(LCA_CACHE_L2) - mBase(6)
        .RamCycles = 4 * (CacheRamFills() - mBase(7)) + writeCycles
        If .PA >= 0 Then .PFN = .PA \ 32
        .Outcome = "OK"
    End With
    mAccessCount = mAccessCount + 1
    mRamCycles = mRamCycles + mLast.RamCycles
    AddTrace StepText(mLast)
End Sub

Private Sub FailStep(ByVal description As String)
    With mLast
        .Failed = True
        .Outcome = description
        If .PA >= 0 Then .PFN = .PA \ 32
    End With
    mAccessCount = mAccessCount + 1
    AddTrace "#" & CStr(mAccessCount) & " " & CachePathKindName(mLast.Kind) & " VA=" & Hex3(mLast.VA) & _
             " -> ERROR " & description & " (sin acceso a cache/RAM / no cache or RAM access)"
End Sub

Private Function EmptyStep(ByVal kind As Long, ByVal va As Long) As LcaPathStep
    Dim s As LcaPathStep
    s.Kind = kind
    s.VA = va
    s.VPN = -1
    s.PA = -1
    s.Value = -1
    s.PFN = -1
    s.VictimVPN = -1
    EmptyStep = s
End Function

Private Function StepText(ByRef s As LcaPathStep) As String
    Dim t As String, l1 As String
    t = "#" & CStr(mAccessCount) & " " & CachePathKindName(s.Kind) & " VA=" & Hex3(s.VA) & _
        " VPN" & CStr(s.VPN) & " -> "
    If s.PageFault Then
        t = t & "PAGE_FAULT PFN" & CStr(s.PFN)
        If s.VictimVPN >= 0 Then
            t = t & " (evict VPN" & CStr(s.VictimVPN) & ", inval " & CStr(s.InvalidatedLines) & " lin"
            If s.Writeback Then t = t & ", writeback 32 B"
            t = t & ")"
        Else
            t = t & " (libre / free)"
        End If
    Else
        t = t & "PTE hit PFN" & CStr(s.PFN)
    End If
    If s.Kind = LCA_PATH_FETCH Then l1 = "L1I" Else l1 = "L1D"
    t = t & " | PA=" & Hex2(s.PA) & " | " & l1 & " " & CStr(s.L1Hits) & "H/" & CStr(s.L1Misses) & "M"
    t = t & " L2 " & CStr(s.L2Hits) & "H/" & CStr(s.L2Misses) & "M | RAM " & CStr(s.RamCycles) & " ciclos"
    t = t & " | byte=" & Hex2(s.Value)
    StepText = t
End Function

Private Sub AddTrace(ByVal text As String)
    mTrace(mTraceCount Mod TRACE_MAX) = text
    mTraceCount = mTraceCount + 1
End Sub

Private Function Hex2(ByVal value As Long) As String
    If value < 0 Then
        Hex2 = "-"
    Else
        Hex2 = Right$("00" & Hex$(value), 2) & "h"
    End If
End Function

Private Function Hex3(ByVal value As Long) As String
    If value < 0 Then
        Hex3 = "-"
    Else
        Hex3 = Right$("000" & Hex$(value), 3) & "h"
    End If
End Function
