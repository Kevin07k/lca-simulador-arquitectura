Attribute VB_Name = "LcaMemTests"
Option Explicit

Private mChecks As Long

' P17 regression: the memory demo runs on the real CPU through MMU -> CLOCK pager -> L1I/L1D -> L2 -> RAM.
' Scattered stores cause page faults and cache misses, a repeated read causes cache hits, the program uses
' more pages than frames (the CLOCK evicts, dirty pages are written back) and the modified byte comes back
' from the swap. Every counter below was derived by hand (docs/P17_MEMORY_DEMO.md); STEP and RUN agree;
' the .asm file is the loaded image; the P17 block draws the same frames and counters as the engine.
Public Sub VerificarMemoriaP17()
    Dim n As Long, i As Long, k As Long, regs(0 To 3) As Long, stepCount As Long, hits As Long, faults As Long
    Dim stepFaults As Long, stepEvict As Long, stepWb As Long, stepL1D As Long, stepL1I As Long
    On Error GoTo failed
    mChecks = 0

    ' --- A: load ------------------------------------------------------------------------------------
    Check MemDemoLoad(), "A: the memory program assembles and loads"
    Check MemDemoLoaded() And LoaderInstructionCount() = 21, "A: 21 instructions (84 bytes = 3 code pages)"
    Check ExecStatus() = EXEC_READY And MemDemoLogCount() = 0, "A: machine READY, empty access log"
    Check PagerFaultCount() = 0 And CacheMisses(LCA_CACHE_L1D) = 0 And CacheHits(LCA_CACHE_L1I) = 0, "A: counters at zero"
    For i = 0 To 14
        If MmuPage(i).Present Then Check False, "A: page " & CStr(i) & " must not be resident before the first access"
    Next i
    mChecks = mChecks + 1

    ' --- B: the program runs to HLT in 30 instructions --------------------------------------------------
    n = MemDemoRun()
    Check n = MEMDEMO_STEPS And ExecSteps() = 30, "B: 14 stores/MOVs + MOV CX + 4 x 3 loop + LD + ST + HLT = 30 (got " & CStr(n) & ")"
    Check ExecStatus() = EXEC_HALTED And MemDemoLogCount() = 30, "B: halted with 30 log rows"
    Check CpuReadRegister(LcaCX) = 0, "B: CX counted four passes down to 0"

    ' --- C: scattered stores: every one faults on a new page and misses the cache --------------------
    For k = 0 To 6
        i = 2 * k + 1
        Check MemDemoLogKind(i) = "ST" And MemDemoLogVA(i) = MemDemoExpectedVA(k), "C: step " & CStr(i + 1) & " stores to " & AsmHex3(MemDemoExpectedVA(k))
        Check MemDemoLogFaults(i) = 1, "C: the store to page " & CStr(MemDemoExpectedVA(k) \ 32) & " is a page fault"
        Check MemDemoLogDataHit(i) = 0, "C: and an L1D miss (cold line)"
    Next k
    Check MemDemoExpectedVA(0) \ 32 = 8 And MemDemoExpectedVA(6) \ 32 = 13, "C: first page 8, last page 13 (scattered order 8,10,12,14,9,11,13)"

    ' --- D: more pages than frames: the CLOCK evicts, dirty pages are written back -------------------
    ' Pages touched: code 0,1,2 + data 8,9,10,11,12,13,14 = 10 > 7 frames; faults = 10 first touches + 1 refault = 11.
    Check PagerFaultCount() = 11, "D: 11 page faults (got " & CStr(PagerFaultCount()) & ")"
    Check PagerEvictionCount() = 11 - 7, "D: 4 evictions = faults - 7 frames"
    Check PagerWritebackCount() = 3, "D: 3 write-backs (dirty pages 8, 10 and 12)"
    Check MemDemoLogVictim(11) = 0 And MemDemoLogWritebacks(11) = 0, "D: the first eviction (step 12) takes code page 0, which is clean"
    Check MemDemoLogVictim(13) = 8 And MemDemoLogWritebacks(13) = 1, "D: step 14 evicts dirty page 8 and writes it back"
    Check MemDemoLogVictim(15) = 10 And MemDemoLogWritebacks(15) = 1, "D: step 16 evicts dirty page 10 and writes it back"
    Check MemDemoLogVictim(16) = 12 And MemDemoLogWritebacks(16) = 1, "D: step 17 evicts dirty page 12 for code page 2"
    stepFaults = 0
    For i = 0 To MemDemoLogCount() - 1
        stepFaults = stepFaults + MemDemoLogFaults(i)
    Next i
    Check stepFaults = PagerFaultCount(), "D: the per-instruction faults add up to the total"

    ' --- E: controlled repetition: four reads of [100h], the last three hit L1D ----------------------
    Check MemDemoLogKind(15) = "LD" And MemDemoLogVA(15) = &H100 And MemDemoLogFaults(15) = 1, "E: first LD [100h] faults: page 8 comes back from the swap"
    Check MemDemoLogDataHit(15) = 0, "E: and misses L1D"
    For k = 0 To 2
        i = 18 + 3 * k
        Check MemDemoLogKind(i) = "LD" And MemDemoLogVA(i) = &H100, "E: step " & CStr(i + 1) & " is LD [100h] again"
        Check MemDemoLogDataHit(i) = 1 And MemDemoLogFaults(i) = 0, "E: it hits L1D with no page fault"
    Next k
    Check CacheHits(LCA_CACHE_L1D) = 3, "E: L1D hits = 3"
    Check CacheMisses(LCA_CACHE_L1D) = 10, "E: L1D misses = 10 (7 stores + first LD + LD [1A0h] + ST [101h])"

    ' --- F: the modified data is recovered ---------------------------------------------------------------
    Check CpuReadRegister(LcaBX) = &HA1, "F: BX = A1h recovered from the page that had been evicted"
    Check MemDemoByte(&H100) = &HA1 And MemDemoByte(&H101) = &HA1, "F: [100h] = [101h] = A1h"
    For k = 0 To 6
        Check MemDemoByte(MemDemoExpectedVA(k)) = MemDemoExpectedValue(k), "F: " & AsmHex3(MemDemoExpectedVA(k)) & " holds " & AsmHex2(MemDemoExpectedValue(k)) & " in RAM or swap"
    Next k
    Check CpuReadRegister(LcaDX) = &HA7, "F: DX = A7h read from the resident page 13"

    ' --- G: a resident page can still miss the cache -------------------------------------------------------
    Check MemDemoLogKind(27) = "LD" And MemDemoLogVA(27) = &H1A0, "G: step 28 is LD [1A0h]"
    Check MemDemoLogFaults(27) = 0 And MmuPage(13).Present, "G: page 13 is resident: no page fault"
    Check MemDemoLogDataHit(27) = 0, "G: but its L1D line was displaced by the loop: an L1D miss"

    ' --- H: counters derived by hand ---------------------------------------------------------------------
    ' L1I: 30 fetches x 4 bytes = 120 accesses; 21 cold lines + the loop conflict (LD and JNZ share an L1I
    ' index, 3 passes x 2) = 27 misses, 93 hits. L2 is asked once per L1 miss: 27 + 10 = 37 = 6 hits (the loop
    ' refills) + 31 misses (21 code lines + 10 data lines, all cold).
    Check CacheMisses(LCA_CACHE_L1I) = 27 And CacheHits(LCA_CACHE_L1I) = 93, "H: L1I 93 hits / 27 misses"
    Check CacheHits(LCA_CACHE_L1I) + CacheMisses(LCA_CACHE_L1I) = 120, "H: 30 instructions x 4 bytes = 120 L1I accesses"
    Check CacheHits(LCA_CACHE_L2) = 6 And CacheMisses(LCA_CACHE_L2) = 31, "H: L2 6 hits / 31 misses"
    Check CacheHits(LCA_CACHE_L2) + CacheMisses(LCA_CACHE_L2) = CacheMisses(LCA_CACHE_L1I) + CacheMisses(LCA_CACHE_L1D), _
          "H: L2 is consulted exactly once per L1 miss"
    Check CacheRamFills() = 31, "H: 31 RAM line fills = L2 misses"
    For i = 0 To 10
        Check MemDemoObservedCounter(i) = MemDemoExpectedCounter(i), "H: " & MemDemoCounterName(i) & " = " & CStr(MemDemoExpectedCounter(i)) & _
              " (got " & CStr(MemDemoObservedCounter(i)) & ")"
    Next i
    Check MemDemoCountersMatch(), "H: all expected counters match"
    Check MemDemoPhase(0) = "1 DISPERSION" And MemDemoPhase(&H3C) = "2 REPETICION" And MemDemoPhase(&H48) = "3 RESIDENTE" And MemDemoPhase(&H50) = "4 COPIA", _
          "H: phase names by PC"

    ' --- I: STEP one by one and RUN give the same machine ------------------------------------------------
    For i = 0 To 3
        regs(i) = CpuReadRegister(i)
    Next i
    stepFaults = PagerFaultCount(): stepEvict = PagerEvictionCount(): stepWb = PagerWritebackCount()
    stepL1D = CacheMisses(LCA_CACHE_L1D): stepL1I = CacheHits(LCA_CACHE_L1I)
    Check MemDemoLoad(), "I: reload (full machine reset)"
    Check PagerFaultCount() = 0 And CacheHits(LCA_CACHE_L2) = 0 And ExecSteps() = 0, "I: counters are zero again after reloading"
    stepCount = 0
    Do While stepCount < 200
        If Not MemDemoStep() Then Exit Do
        stepCount = stepCount + 1
        If ExecStatus() = EXEC_HALTED Then Exit Do
    Loop
    Check stepCount = 30 And ExecStatus() = EXEC_HALTED, "I: STEP x 30 reaches HLT"
    For i = 0 To 3
        Check CpuReadRegister(i) = regs(i), "I: same " & Choose(i + 1, "AX", "BX", "CX", "DX")
    Next i
    Check PagerFaultCount() = stepFaults And PagerEvictionCount() = stepEvict And PagerWritebackCount() = stepWb, "I: same paging counters"
    Check CacheMisses(LCA_CACHE_L1D) = stepL1D And CacheHits(LCA_CACHE_L1I) = stepL1I, "I: same cache counters"
    Check MemDemoCountersMatch(), "I: the counters are reproducible"
    Check Not MemDemoStep(), "I: after HLT nothing runs"

    ' --- J: the .asm is the loaded image --------------------------------------------------------------------
    n = LoaderInstructionCount()
    Check AsmAssembleText(MemDemoAsmText()), "J: MemDemoAsmText assembles"
    Check AsmInstructionCount() = n, "J: same instruction count"
    For i = 0 To 4 * n - 1
        If AsmByte(i) <> LoaderImageByte(i) Then Check False, "J: byte " & AsmHex2(i) & " differs"
    Next i
    mChecks = mChecks + 1
    Check MemDemoFile() = "memoria.asm", "J: file name memoria.asm"

    ' --- K: the P17 block draws the same frames and counters as the engine (when it is built) ---------------
    If MemPanelConstruido() Then
        Check MemDemoLoad(), "K: reload"
        Call MemDemoRun
        MemPanelRefrescar
        Check MemPanelCoherente(), "K: drawn frames, pages and counters equal the engine"
        Check InStr(MemPanelEstado(), "TERMINADO") > 0, "K: the block shows HALTED"
    End If

    LcaReport "P17 PASS: " & CStr(mChecks) & " comprobaciones / assertions." & vbCrLf & _
              "Demo de memoria: dispersion (faults y misses), repeticion (hits), mas paginas que marcos, dato modificado recuperado; contadores esperados verificados.", True
    Exit Sub
failed:
    LcaReport "P17 FAIL after " & CStr(mChecks) & " checks: " & Err.Description, False
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2182, "LcaMemTests", message
End Sub
