Attribute VB_Name = "LcaMemDemo"
Option Explicit

' P17: memory demo program written in the P11 assembly and run on the real CPU (P12) through the P10 path
' (MMU -> CLOCK pager -> L1I/L1D -> L2 -> RAM). It is built to show, in this order:
'   1) SCATTER  seven stores to seven different data pages (VPN 8,10,12,14,9,11,13): every one is a cold access
'               (page fault + L1D/L2 miss) and, together with the three code pages, the program touches 10
'               pages while only 7 frames are pageable, so the CLOCK has to evict (dirty pages are written back);
'   2) REPEAT   a four-pass loop that reads the same byte [100h]: the first pass recovers the data that was
'               modified in phase 1 (its page was evicted and now comes back from swap), the next passes hit
'               the cache. In L1I (2 lines) LD and JNZ share an index and evict each other: those refills hit L2;
'   3) RESIDENT a read of the last written page, still in RAM: no page fault, yet an L1D miss (the loop evicted its line);
'   4) COPY     the recovered byte is stored again at [101h] and the program halts.
' Expected values are derived by hand and asserted by VerificarMemoriaP17. No worksheet dependency.
Public Const MEMDEMO_STEPS As Long = 30
Public Const MEMDEMO_VA_RECOVER As Long = &H100      ' modified in phase 1, recovered in phase 2
Public Const MEMDEMO_VA_LAST As Long = &H1A0         ' last page written, still resident
Public Const MEMDEMO_LOG_MAX As Long = 64
Public Const LCA_ERR_MEMDEMO As Long = vbObjectError + 2181

Private mLoadId As Long
Private mN As Long
Private mPC(0 To MEMDEMO_LOG_MAX - 1) As Long
Private mKind(0 To MEMDEMO_LOG_MAX - 1) As String
Private mVA(0 To MEMDEMO_LOG_MAX - 1) As Long
Private mFault(0 To MEMDEMO_LOG_MAX - 1) As Long       ' page faults caused by this instruction (fetch + data)
Private mVictim(0 To MEMDEMO_LOG_MAX - 1) As Long      ' VPN evicted, -1 when none
Private mWb(0 To MEMDEMO_LOG_MAX - 1) As Long          ' dirty write-backs caused
Private mL1Hit(0 To MEMDEMO_LOG_MAX - 1) As Long       ' L1I + L1D hits of this instruction
Private mL1Miss(0 To MEMDEMO_LOG_MAX - 1) As Long
Private mL2Hit(0 To MEMDEMO_LOG_MAX - 1) As Long
Private mL2Miss(0 To MEMDEMO_LOG_MAX - 1) As Long
Private mDataHit(0 To MEMDEMO_LOG_MAX - 1) As Long     ' 1 hit, 0 miss, -1 no data access (L1D only)

Public Function MemDemoSource() As Variant
    MemDemoSource = Array( _
        "; MEMORIA: dispersion (faults y misses), repeticion (hits), mas paginas que marcos y recuperar un dato.", _
        "        MOV AX, #0A1h     ; --- fase 1: DISPERSION, siete paginas distintas ---", _
        "        ST [100h], AX     ; pagina 8   (este dato se recupera en la fase 2)", _
        "        MOV AX, #0A2h", _
        "        ST [140h], AX     ; pagina 10", _
        "        MOV AX, #0A3h", _
        "        ST [180h], AX     ; pagina 12", _
        "        MOV AX, #0A4h", _
        "        ST [1C0h], AX     ; pagina 14", _
        "        MOV AX, #0A5h", _
        "        ST [120h], AX     ; pagina 9", _
        "        MOV AX, #0A6h", _
        "        ST [160h], AX     ; pagina 11", _
        "        MOV AX, #0A7h", _
        "        ST [1A0h], AX     ; pagina 13   (la ultima: sigue residente)", _
        "        MOV CX, #04h      ; --- fase 2: REPETICION, cuatro lecturas del mismo byte ---", _
        "rep:    LD BX, [100h]     ; 1a vez: recupera A1h del swap; luego acierta en cache", _
        "        DEC CX", _
        "        JNZ rep", _
        "        LD DX, [1A0h]     ; --- fase 3: pagina residente, sin falta ---", _
        "        ST [101h], BX     ; --- fase 4: copia el dato recuperado ---", _
        "        HLT")
End Function

' Source text exactly as stored in programs/memoria/memoria.asm (LF line ends).
Public Function MemDemoAsmText() As String
    MemDemoAsmText = Join(MemDemoSource(), vbLf)
End Function

Public Function MemDemoFile() As String
    MemDemoFile = "memoria.asm"
End Function

' Assembles and loads the program; the whole machine is reset (pages, caches, counters, trace).
Public Function MemDemoLoad() As Boolean
    MemDemoLoad = LoaderLoadSource(MemDemoSource())
    If MemDemoLoad Then
        mLoadId = LoaderLoadCount()
        mN = 0
    End If
End Function

Public Function MemDemoLoaded() As Boolean
    MemDemoLoaded = LoaderLoaded() And mLoadId = LoaderLoadCount()
End Function

' One instruction with its memory effects recorded. Returns False when nothing ran (HLT done / error).
Public Function MemDemoStep() As Boolean
    Dim f0 As Long, w0 As Long, hi As Long, mi As Long
    Dim hd As Long, md As Long, hl As Long, ml As Long, pc As Long, i As Long
    If Not MemDemoLoaded() Then Err.Raise LCA_ERR_MEMDEMO, "LcaMemDemo", "LOAD_THE_MEMORY_PROGRAM_FIRST"
    pc = CpuSnapshot().PC
    f0 = PagerFaultCount(): w0 = PagerWritebackCount()
    hi = CacheHits(LCA_CACHE_L1I): mi = CacheMisses(LCA_CACHE_L1I)
    hd = CacheHits(LCA_CACHE_L1D): md = CacheMisses(LCA_CACHE_L1D)
    hl = CacheHits(LCA_CACHE_L2): ml = CacheMisses(LCA_CACHE_L2)
    If Not ExecStep() Then Exit Function
    MemDemoStep = True
    If mN >= MEMDEMO_LOG_MAX Then Exit Function
    i = mN
    mPC(i) = pc
    mKind(i) = ExecLastKind()
    mVA(i) = ExecLastDataVA()
    mFault(i) = PagerFaultCount() - f0
    mWb(i) = PagerWritebackCount() - w0
    If mFault(i) > 0 Then mVictim(i) = PagerLastVictimVPN() Else mVictim(i) = -1
    mL1Hit(i) = (CacheHits(LCA_CACHE_L1I) - hi) + (CacheHits(LCA_CACHE_L1D) - hd)
    mL1Miss(i) = (CacheMisses(LCA_CACHE_L1I) - mi) + (CacheMisses(LCA_CACHE_L1D) - md)
    mL2Hit(i) = CacheHits(LCA_CACHE_L2) - hl
    mL2Miss(i) = CacheMisses(LCA_CACHE_L2) - ml
    If mVA(i) >= 0 Then
        If CacheHits(LCA_CACHE_L1D) - hd > 0 Then mDataHit(i) = 1 Else mDataHit(i) = 0
    Else
        mDataHit(i) = -1
    End If
    mN = mN + 1
End Function

' Runs to HLT (or error / limit). Returns the instructions completed by this call.
Public Function MemDemoRun() As Long
    Do While MemDemoRunning() And MemDemoRun < 200
        If Not MemDemoStep() Then Exit Do
        MemDemoRun = MemDemoRun + 1
    Loop
End Function

Public Function MemDemoRunning() As Boolean
    MemDemoRunning = (ExecStatus() = EXEC_READY)
End Function

' --- Access log (one row per executed instruction) ------------------------------------------------
Public Function MemDemoLogCount() As Long
    MemDemoLogCount = mN
End Function

Public Function MemDemoLogPC(ByVal i As Long) As Long
    MemDemoLogPC = mPC(i)
End Function

Public Function MemDemoLogKind(ByVal i As Long) As String
    MemDemoLogKind = mKind(i)
End Function

Public Function MemDemoLogVA(ByVal i As Long) As Long
    MemDemoLogVA = mVA(i)
End Function

Public Function MemDemoLogFaults(ByVal i As Long) As Long
    MemDemoLogFaults = mFault(i)
End Function

Public Function MemDemoLogVictim(ByVal i As Long) As Long
    MemDemoLogVictim = mVictim(i)
End Function

Public Function MemDemoLogWritebacks(ByVal i As Long) As Long
    MemDemoLogWritebacks = mWb(i)
End Function

Public Function MemDemoLogL1Hits(ByVal i As Long) As Long
    MemDemoLogL1Hits = mL1Hit(i)
End Function

Public Function MemDemoLogL1Misses(ByVal i As Long) As Long
    MemDemoLogL1Misses = mL1Miss(i)
End Function

Public Function MemDemoLogL2Misses(ByVal i As Long) As Long
    MemDemoLogL2Misses = mL2Miss(i)
End Function

' 1 = the data access hit L1D, 0 = it missed, -1 = the instruction has no data access.
Public Function MemDemoLogDataHit(ByVal i As Long) As Long
    MemDemoLogDataHit = mDataHit(i)
End Function

' One compact line per instruction for the Logs sheet and the tests.
Public Function MemDemoLogLine(ByVal i As Long) As String
    Dim s As String
    s = "#" & CStr(i + 1) & " " & AsmHex3(mPC(i)) & " " & mKind(i)
    If mVA(i) >= 0 Then s = s & " VA=" & AsmHex3(mVA(i)) & IIf(mDataHit(i) = 1, " L1D-HIT", " L1D-MISS")
    If mFault(i) > 0 Then
        s = s & " FAULTx" & CStr(mFault(i))
        If mVictim(i) >= 0 Then s = s & " evicts VPN" & CStr(mVictim(i))
    End If
    If mWb(i) > 0 Then s = s & " WRITEBACK"
    MemDemoLogLine = s
End Function

' Name of the phase that owns an instruction address.
Public Function MemDemoPhase(ByVal pc As Long) As String
    If pc <= &H34 Then
        MemDemoPhase = "1 DISPERSION"
    ElseIf pc <= &H44 Then
        MemDemoPhase = "2 REPETICION"
    ElseIf pc = &H48 Then
        MemDemoPhase = "3 RESIDENTE"
    Else
        MemDemoPhase = "4 COPIA"
    End If
End Function

' Byte of the data space the way the program sees it: RAM if the page is resident, else swap.
Public Function MemDemoByte(ByVal va As Long) As Long
    Dim e As LcaPageEntry
    e = MmuPage(va \ 32)
    If e.Present Then
        MemDemoByte = RamReadUser(e.PFN * 32 + (va Mod 32))
    Else
        MemDemoByte = SwapReadByte(va \ 32, va Mod 32)
    End If
End Function

Public Function MemDemoExpectedValue(ByVal k As Long) As Long
    MemDemoExpectedValue = &HA1 + k       ' k = 0..6 in the order of the stores
End Function

Public Function MemDemoExpectedVA(ByVal k As Long) As Long
    MemDemoExpectedVA = Choose(k + 1, &H100, &H140, &H180, &H1C0, &H120, &H160, &H1A0)
End Function

' --- Expected counters (derived by hand, see docs/P17_MEMORY_DEMO.md) -----------------------------
' Index: 0 instructions, 1 page faults, 2 evictions, 3 write-backs, 4 L1D hits, 5 L1D misses, 6 L1I hits,
' 7 L1I misses, 8 L2 hits, 9 L2 misses, 10 the byte recovered at [100h].
Public Function MemDemoExpectedCounter(ByVal i As Long) As Long
    MemDemoExpectedCounter = Choose(i + 1, 30, 11, 4, 3, 3, 10, 93, 27, 6, 31, &HA1)
End Function

Public Function MemDemoObservedCounter(ByVal i As Long) As Long
    Select Case i
        Case 0: MemDemoObservedCounter = ExecSteps()
        Case 1: MemDemoObservedCounter = PagerFaultCount()
        Case 2: MemDemoObservedCounter = PagerEvictionCount()
        Case 3: MemDemoObservedCounter = PagerWritebackCount()
        Case 4: MemDemoObservedCounter = CacheHits(LCA_CACHE_L1D)
        Case 5: MemDemoObservedCounter = CacheMisses(LCA_CACHE_L1D)
        Case 6: MemDemoObservedCounter = CacheHits(LCA_CACHE_L1I)
        Case 7: MemDemoObservedCounter = CacheMisses(LCA_CACHE_L1I)
        Case 8: MemDemoObservedCounter = CacheHits(LCA_CACHE_L2)
        Case 9: MemDemoObservedCounter = CacheMisses(LCA_CACHE_L2)
        Case 10: MemDemoObservedCounter = CpuReadRegister(LcaBX)
    End Select
End Function

Public Function MemDemoCounterName(ByVal i As Long) As String
    MemDemoCounterName = Choose(i + 1, "instructions", "page faults", "evictions", "write-backs", "L1D hits", "L1D misses", _
                                "L1I hits", "L1I misses", "L2 hits", "L2 misses", "BX = [100h] recovered")
End Function

' True when every expected counter matches (valid after the program halts).
Public Function MemDemoCountersMatch() As Boolean
    Dim i As Long
    For i = 0 To 10
        If MemDemoObservedCounter(i) <> MemDemoExpectedCounter(i) Then Exit Function
    Next i
    MemDemoCountersMatch = True
End Function
