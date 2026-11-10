Attribute VB_Name = "LcaCacheTests"
Option Explicit

Private mChecks As Long

' Desktop Excel verification for P09 using physical addresses only (P07 not required).
Public Sub VerificarCachesP09()
    Dim failure As Long
    Dim value As Byte
    Dim i As Long
    On Error GoTo failed
    mChecks = 0

    ' --- Section A: P02 manual fetch case: PA 40..43h = 10 00 2A 00 -----------------
    RamReset
    CacheReset
    RamWriteUser &H40, &H10
    RamWriteUser &H41, &H0
    RamWriteUser &H42, &H2A
    RamWriteUser &H43, &H0
    Check CacheRead(&H40, LCA_CACHE_FETCH) = &H10, "A: first fetch byte"
    Check CacheLastFillPA() = &H40 And CacheRamFills() = 1, "A: one four-byte line fill from 40h"
    Check CacheRead(&H41, LCA_CACHE_FETCH) = &H0 And CacheLastFillPA() = -1, "A: byte 41h hits, no RAM fill"
    Check CacheRead(&H42, LCA_CACHE_FETCH) = &H2A, "A: byte 42h"
    Check CacheRead(&H43, LCA_CACHE_FETCH) = &H0, "A: byte 43h"
    Counters "A fetch", 3, 1, 0, 0, 0, 1
    Check CacheRamFills() = 1, "A: fill happens once, not per byte"
    Check CacheTraceCount() = 4 And InStr(1, CacheTraceLine(0), "L1I MISS", vbBinaryCompare) > 0 _
        And InStr(1, CacheTraceLine(0), "L2 MISS fill 40h-43h", vbBinaryCompare) > 0, "A: miss/fill trace"
    Check InStr(1, CacheTraceLine(1), "L1I HIT", vbBinaryCompare) > 0, "A: hit trace"

    ' --- Section B: L1D read, same-line hits ------------------------------------------
    CacheReset
    RamWriteUser &H6C, 1
    RamWriteUser &H6D, 2
    RamWriteUser &H6E, 3
    RamWriteUser &H6F, 4
    Check CacheRead(&H6D, LCA_CACHE_DATA) = 2, "B: data miss returns RAM"
    Check CacheRead(&H6D, LCA_CACHE_DATA) = 2, "B: data repeat"
    Check CacheRead(&H6C, LCA_CACHE_DATA) = 1, "B: same line neighbour"
    Counters "B data", 0, 0, 2, 1, 0, 1

    ' --- Section C: write-through coherence on L1D hit -------------------------------
    Check CacheWrite(&H6D, &H55), "C: write hit"
    Check RamReadUser(&H6D) = &H55, "C: write-through reached RAM"
    Check CacheRead(&H6D, LCA_CACHE_DATA) = &H55, "C: L1D shows new value"
    Check CacheLineByte(LCA_CACHE_L2, 3, 1) = &H55, "C: resident L2 line updated"
    Counters "C write hit", 0, 0, 4, 1, 0, 1

    ' --- Section D: write-allocate on miss ----------------------------------------------
    Check CacheWrite(&H70, &H77), "D: write miss"
    Check CacheContains(LCA_CACHE_L1D, &H70) And CacheContains(LCA_CACHE_L2, &H70), "D: line allocated"
    Check RamReadUser(&H70) = &H77, "D: RAM updated after allocate"
    Check CacheRead(&H70, LCA_CACHE_DATA) = &H77, "D: allocated line hits"
    Counters "D write-allocate", 0, 0, 5, 2, 0, 2

    ' --- Section E: L1I direct-map conflict (2 lines) -----------------------------------
    CacheReset
    RamWriteUser &H48, &H33
    RamWriteUser &H40, &H11
    Check CacheRead(&H40, LCA_CACHE_FETCH) = &H11, "E: first line"
    Check CacheRead(&H48, LCA_CACHE_FETCH) = &H33, "E: conflicting line overwrites index 0"
    Check Not CacheContains(LCA_CACHE_L1I, &H40) And CacheContains(LCA_CACHE_L1I, &H48), "E: replacement by overwrite"
    Check CacheRead(&H40, LCA_CACHE_FETCH) = &H11, "E: evicted line returns from L2"
    Counters "E conflict", 0, 3, 0, 0, 1, 2

    ' --- Section F: L2 conflict (8 lines): PA 00h and 20h share L2 index 0 ----------
    CacheReset
    RamWriteUser &H0, &HA0
    RamWriteUser &H20, &HB0
    Check CacheRead(&H0, LCA_CACHE_DATA) = &HA0, "F: first data line"
    Check CacheRead(&H20, LCA_CACHE_DATA) = &HB0, "F: second data line"
    Check CacheRead(&H0, LCA_CACHE_DATA) = &HA0, "F: reload after L2 conflict"
    Counters "F L2 conflict", 0, 0, 0, 3, 0, 3

    ' --- Section G: fetch and data caches are separate ---------------------------------
    CacheReset
    RamWriteUser &H40, &H11
    Check CacheRead(&H40, LCA_CACHE_FETCH) = &H11, "G: fetch"
    Check CacheRead(&H40, LCA_CACHE_DATA) = &H11, "G: same PA as data"
    Counters "G fetch vs data", 0, 1, 0, 1, 1, 1

    ' --- Section H: write keeps L1I coherent without counting ------------------------
    CacheReset
    RamWriteUser &H40, 1
    RamWriteUser &H41, 2
    Check CacheRead(&H40, LCA_CACHE_FETCH) = 1, "H: L1I loaded"
    CacheWrite &H41, &H99
    Check CacheRead(&H41, LCA_CACHE_FETCH) = &H99, "H: fetch sees stored byte"
    Counters "H L1I coherence", 1, 1, 0, 1, 1, 1

    ' --- Section I: frame invalidation ---------------------------------------------------
    CacheReset
    RamWriteUser &H40, &H11
    RamWriteUser &H64, &H22
    CacheRead &H40, LCA_CACHE_FETCH
    CacheRead &H64, LCA_CACHE_DATA
    CacheInvalidateFrame 2
    Check Not CacheContains(LCA_CACHE_L1I, &H40) And Not CacheContains(LCA_CACHE_L2, &H40), "I: frame 2 invalidated"
    Check CacheContains(LCA_CACHE_L1D, &H64) And CacheContains(LCA_CACHE_L2, &H64), "I: frame 3 untouched"
    RamWritePhysical &H40, &H5A   ' replacement/loader changes RAM; stale copy must not be seen
    Check CacheRead(&H40, LCA_CACHE_FETCH) = &H5A, "I: no stale byte after invalidation"
    Check CacheRead(&H64, LCA_CACHE_DATA) = &H22, "I: other frame still hits"
    Counters "I invalidate", 0, 2, 1, 1, 0, 3

    ' --- Section J: reset ------------------------------------------------------------------
    CacheReset
    Counters "J reset", 0, 0, 0, 0, 0, 0
    Check Not CacheContains(LCA_CACHE_L1I, &H40) And Not CacheContains(LCA_CACHE_L2, &H64), "J: lines invalid"
    Check CacheTraceCount() = 0 And CacheRamFills() = 0, "J: trace and fills cleared"

    ' --- Section K: rejected requests change nothing ---------------------------------------
    RamReset
    CacheReset
    failure = 0
    On Error Resume Next
    value = CacheRead(&HE0, LCA_CACHE_DATA)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_RAM_PROTECTION, "K: I/O address is never cached"
    failure = 0
    On Error Resume Next
    value = CacheRead(&H100, LCA_CACHE_FETCH)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_RAM_ADDRESS, "K: PA 100h rejected"
    failure = 0
    On Error Resume Next
    value = CacheRead(&H40, 5)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_CACHE_ARGUMENT, "K: unknown access kind rejected"
    failure = 0
    On Error Resume Next
    Call CacheWrite(&H40, 256)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_BYTE And RamReadUser(&H40) = 0 And Not CacheContains(LCA_CACHE_L1D, &H40), _
        "K: invalid byte leaves caches and RAM untouched"
    failure = 0
    On Error Resume Next
    Call CacheWrite(&HEF, 1)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_RAM_PROTECTION And RamReadIO(&HEF) = &H20, "K: write cannot reach LCD buffer"
    Counters "K rejected", 0, 0, 0, 0, 0, 0

    ' --- Section L: capacities ------------------------------------------------------------------
    Check CacheLineCount(LCA_CACHE_L1I) * LCA_CACHE_LINE_BYTES = 8, "L: L1I is 8 bytes"
    Check CacheLineCount(LCA_CACHE_L1D) * LCA_CACHE_LINE_BYTES = 8, "L: L1D is 8 bytes"
    Check CacheLineCount(LCA_CACHE_L2) * LCA_CACHE_LINE_BYTES = 32, "L: L2 is 32 bytes"

    CachePanelActualizar
    LcaReport "P09 PASS: " & mChecks & " comprobaciones / assertions." & vbCrLf & _
           "Caches L1I/L1D/L2 verificadas con direcciones fisicas; P07/P10 aun no forman parte.", True
    Exit Sub
failed:
    LcaReport "P09 FAIL after " & mChecks & " checks: " & Err.Description, False
End Sub

' Hits/misses in the order L1I, L1D, L2 (hit, miss).
Private Sub Counters(ByVal label As String, ByVal i1h As Long, ByVal i1m As Long, _
                     ByVal d1h As Long, ByVal d1m As Long, ByVal l2h As Long, ByVal l2m As Long)
    Check CacheHits(LCA_CACHE_L1I) = i1h And CacheMisses(LCA_CACHE_L1I) = i1m, label & ": L1I counters " & _
        CacheHits(LCA_CACHE_L1I) & "/" & CacheMisses(LCA_CACHE_L1I)
    Check CacheHits(LCA_CACHE_L1D) = d1h And CacheMisses(LCA_CACHE_L1D) = d1m, label & ": L1D counters " & _
        CacheHits(LCA_CACHE_L1D) & "/" & CacheMisses(LCA_CACHE_L1D)
    Check CacheHits(LCA_CACHE_L2) = l2h And CacheMisses(LCA_CACHE_L2) = l2m, label & ": L2 counters " & _
        CacheHits(LCA_CACHE_L2) & "/" & CacheMisses(LCA_CACHE_L2)
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2092, "LcaCacheTests", message
End Sub
