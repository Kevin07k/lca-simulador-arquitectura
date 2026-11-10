Attribute VB_Name = "LcaCache"
Option Explicit

' P09: L1I (8 B), L1D (8 B) and unified L2 (32 B). Physical-address tags only.
' Direct mapped, four-byte lines, overwrite replacement, write-through + write-allocate.
' No worksheet dependency; RAM is reached only through the restricted user path of LcaRam.
Public Const LCA_CACHE_L1I As Long = 0
Public Const LCA_CACHE_L1D As Long = 1
Public Const LCA_CACHE_L2 As Long = 2
Public Const LCA_CACHE_FETCH As Long = 0
Public Const LCA_CACHE_DATA As Long = 1
Public Const LCA_CACHE_LINE_BYTES As Long = 4
Public Const LCA_ERR_CACHE_ARGUMENT As Long = vbObjectError + 2091
Private Const TRACE_MAX As Long = 64

Private mValid(0 To 2, 0 To 7) As Boolean
Private mTag(0 To 2, 0 To 7) As Long
Private mData(0 To 2, 0 To 7, 0 To 3) As Byte
Private mHits(0 To 2) As Long
Private mMisses(0 To 2) As Long
Private mRamFills As Long
Private mLastFillPA As Long
Private mLastEvent As String
Private mTrace(0 To TRACE_MAX - 1) As String
Private mTraceCount As Long

Public Function CacheLineCount(ByVal level As Long) As Long
    Select Case level
        Case LCA_CACHE_L1I, LCA_CACHE_L1D: CacheLineCount = 2
        Case LCA_CACHE_L2: CacheLineCount = 8
        Case Else: Err.Raise LCA_ERR_CACHE_ARGUMENT, "LcaCache", "CACHE_LEVEL"
    End Select
End Function

Public Function CacheLevelName(ByVal level As Long) As String
    Select Case level
        Case LCA_CACHE_L1I: CacheLevelName = "L1I"
        Case LCA_CACHE_L1D: CacheLevelName = "L1D"
        Case LCA_CACHE_L2: CacheLevelName = "L2"
        Case Else: Err.Raise LCA_ERR_CACHE_ARGUMENT, "LcaCache", "CACHE_LEVEL"
    End Select
End Function

' Reset: invalidate every line, zero counters and trace (P02 reset contract).
Public Sub CacheReset()
    CacheInvalidateAll
    Dim level As Long
    For level = 0 To 2
        mHits(level) = 0
        mMisses(level) = 0
    Next level
    mRamFills = 0
    mLastFillPA = -1
    mTraceCount = 0
    mLastEvent = "RESET"
End Sub

Public Sub CacheInvalidateAll()
    Dim level As Long, idx As Long
    For level = 0 To 2
        For idx = 0 To 7
            mValid(level, idx) = False
        Next idx
    Next level
End Sub

' Called before a PFN is reused (P08) or its backing is reloaded: no stale line survives.
' Returns how many valid lines were dropped (0 when the frame had none cached).
Public Function CacheInvalidateFrame(ByVal pfn As Long) As Long
    Dim level As Long, idx As Long, n As Long, lineNo As Long
    If pfn < 0 Or pfn >= LCA_RAM_FRAMES Then Err.Raise LCA_ERR_CACHE_ARGUMENT, "LcaCache", "CACHE_FRAME"
    For level = 0 To 2
        n = CacheLineCount(level)
        For idx = 0 To n - 1
            If mValid(level, idx) Then
                lineNo = mTag(level, idx) * n + idx
                If (lineNo * LCA_CACHE_LINE_BYTES) \ LCA_RAM_FRAME_SIZE = pfn Then
                    mValid(level, idx) = False
                    CacheInvalidateFrame = CacheInvalidateFrame + 1
                End If
            End If
        Next idx
    Next level
    AddTrace "INVALIDATE frame " & CStr(pfn) & " in L1I/L1D/L2 (" & CStr(CacheInvalidateFrame) & " lines)"
End Function

' Reads one byte requested by the CPU. kind: LCA_CACHE_FETCH (L1I) or LCA_CACHE_DATA (L1D).
' PA must be a user-memory address (00..DFh); I/O (E0..EFh) is never cached.
Public Function CacheRead(ByVal pa As Long, ByVal kind As Long) As Byte
    Dim l1 As Long, lineNo As Long, idx As Long, tag As Long, off As Long
    Dim l2state As String, text As String
    l1 = L1Of(kind)
    RamRequireUserAddress pa
    lineNo = pa \ LCA_CACHE_LINE_BYTES
    off = pa Mod LCA_CACHE_LINE_BYTES
    idx = lineNo Mod 2
    tag = lineNo \ 2
    mLastFillPA = -1
    If mValid(l1, idx) And mTag(l1, idx) = tag Then
        mHits(l1) = mHits(l1) + 1
        text = CacheLevelName(l1) & " HIT"
        l2state = "L2 -"
    Else
        mMisses(l1) = mMisses(l1) + 1
        l2state = ConsultL2(lineNo)
        CopyLine LCA_CACHE_L2, lineNo Mod 8, l1, idx
        mValid(l1, idx) = True
        mTag(l1, idx) = tag
        text = CacheLevelName(l1) & " MISS"
    End If
    CacheRead = mData(l1, idx, off)
    AddTrace AccessText("READ", kind, pa) & " -> " & text & " | " & l2state & _
             " | byte=" & Hex2(CacheRead)
End Function

' ST: write-allocate into L1D, then write-through to L2 (if resident) and RAM.
' L1I is kept coherent without counting an access (code is normally immutable).
Public Function CacheWrite(ByVal pa As Long, ByVal value As Long) As Boolean
    Dim lineNo As Long, idx As Long, tag As Long, off As Long
    Dim l2idx As Long, l2state As String, text As String
    RamRequireUserAddress pa
    CpuRequireByte value          ' validate everything before any cache or RAM mutation
    lineNo = pa \ LCA_CACHE_LINE_BYTES
    off = pa Mod LCA_CACHE_LINE_BYTES
    idx = lineNo Mod 2
    tag = lineNo \ 2
    mLastFillPA = -1
    If mValid(LCA_CACHE_L1D, idx) And mTag(LCA_CACHE_L1D, idx) = tag Then
        mHits(LCA_CACHE_L1D) = mHits(LCA_CACHE_L1D) + 1
        text = "L1D HIT"
        l2state = "L2 -"
    Else
        mMisses(LCA_CACHE_L1D) = mMisses(LCA_CACHE_L1D) + 1
        l2state = ConsultL2(lineNo)
        CopyLine LCA_CACHE_L2, lineNo Mod 8, LCA_CACHE_L1D, idx
        mValid(LCA_CACHE_L1D, idx) = True
        mTag(LCA_CACHE_L1D, idx) = tag
        text = "L1D MISS (write-allocate)"
    End If
    mData(LCA_CACHE_L1D, idx, off) = CByte(value)
    l2idx = lineNo Mod 8
    If mValid(LCA_CACHE_L2, l2idx) And mTag(LCA_CACHE_L2, l2idx) = lineNo \ 8 Then
        mData(LCA_CACHE_L2, l2idx, off) = CByte(value)
    End If
    If mValid(LCA_CACHE_L1I, idx) And mTag(LCA_CACHE_L1I, idx) = tag Then
        mData(LCA_CACHE_L1I, idx, off) = CByte(value)
    End If
    RamWriteUser pa, value
    CacheWrite = True
    AddTrace AccessText("WRITE", LCA_CACHE_DATA, pa) & "=" & Hex2(value) & " -> " & text & _
             " | " & l2state & " | write-through RAM"
End Function

' P10 snoop: a write that reached RAM outside CacheWrite (P08 uncached ST path) updates every
' resident copy of that byte. No counters, no allocation, so no line can keep an old value.
Public Sub CacheSnoopWrite(ByVal pa As Long, ByVal value As Long)
    Dim level As Long, n As Long, lineNo As Long, idx As Long
    RamRequireUserAddress pa
    CpuRequireByte value
    lineNo = pa \ LCA_CACHE_LINE_BYTES
    For level = 0 To 2
        n = CacheLineCount(level)
        idx = lineNo Mod n
        If mValid(level, idx) And mTag(level, idx) = lineNo \ n Then _
            mData(level, idx, pa Mod LCA_CACHE_LINE_BYTES) = CByte(value)
    Next level
End Sub

' Number of valid lines of one level that belong to a physical frame (panel/tests only).
Public Function CacheFrameLines(ByVal level As Long, ByVal pfn As Long) As Long
    Dim n As Long, idx As Long
    n = CacheLineCount(level)
    If pfn < 0 Or pfn >= LCA_RAM_FRAMES Then Err.Raise LCA_ERR_CACHE_ARGUMENT, "LcaCache", "CACHE_FRAME"
    For idx = 0 To n - 1
        If mValid(level, idx) Then
            If ((mTag(level, idx) * n + idx) * LCA_CACHE_LINE_BYTES) \ LCA_RAM_FRAME_SIZE = pfn Then _
                CacheFrameLines = CacheFrameLines + 1
        End If
    Next idx
End Function

' One L2 query per L1 miss; on L2 miss the whole four-byte line is read from RAM.
Private Function ConsultL2(ByVal lineNo As Long) As String
    Dim idx As Long, tag As Long, i As Long, base As Long
    idx = lineNo Mod 8
    tag = lineNo \ 8
    If mValid(LCA_CACHE_L2, idx) And mTag(LCA_CACHE_L2, idx) = tag Then
        mHits(LCA_CACHE_L2) = mHits(LCA_CACHE_L2) + 1
        ConsultL2 = "L2 HIT"
    Else
        mMisses(LCA_CACHE_L2) = mMisses(LCA_CACHE_L2) + 1
        base = lineNo * LCA_CACHE_LINE_BYTES
        For i = 0 To LCA_CACHE_LINE_BYTES - 1
            mData(LCA_CACHE_L2, idx, i) = RamReadUser(base + i)
        Next i
        mValid(LCA_CACHE_L2, idx) = True
        mTag(LCA_CACHE_L2, idx) = tag
        mRamFills = mRamFills + 1
        mLastFillPA = base
        ConsultL2 = "L2 MISS fill " & Hex2(base) & "-" & Hex2(base + 3)
    End If
End Function

Private Sub CopyLine(ByVal fromLevel As Long, ByVal fromIdx As Long, _
                     ByVal toLevel As Long, ByVal toIdx As Long)
    Dim i As Long
    For i = 0 To LCA_CACHE_LINE_BYTES - 1
        mData(toLevel, toIdx, i) = mData(fromLevel, fromIdx, i)
    Next i
End Sub

Private Function L1Of(ByVal kind As Long) As Long
    Select Case kind
        Case LCA_CACHE_FETCH: L1Of = LCA_CACHE_L1I
        Case LCA_CACHE_DATA: L1Of = LCA_CACHE_L1D
        Case Else: Err.Raise LCA_ERR_CACHE_ARGUMENT, "LcaCache", "CACHE_KIND"
    End Select
End Function

Private Function Hex2(ByVal value As Long) As String
    Hex2 = Right$("00" & Hex$(value), 2) & "h"
End Function

Private Function AccessText(ByVal verb As String, ByVal kind As Long, ByVal pa As Long) As String
    Dim label As String
    If kind = LCA_CACHE_FETCH Then label = "FETCH" Else label = verb
    AccessText = label & " PA=" & Hex2(pa)
End Function

Private Sub AddTrace(ByVal text As String)
    Dim i As Long
    mLastEvent = text
    If mTraceCount = TRACE_MAX Then
        For i = 1 To TRACE_MAX - 1
            mTrace(i - 1) = mTrace(i)
        Next i
        mTraceCount = TRACE_MAX - 1
    End If
    mTrace(mTraceCount) = text
    mTraceCount = mTraceCount + 1
End Sub

' Read-only inspection (no counters, no state change).
Public Function CacheHits(ByVal level As Long) As Long
    CacheLineCount level
    CacheHits = mHits(level)
End Function

Public Function CacheMisses(ByVal level As Long) As Long
    CacheLineCount level
    CacheMisses = mMisses(level)
End Function

Public Function CacheRamFills() As Long
    CacheRamFills = mRamFills
End Function

' Base PA of the last RAM line fill, or -1 when the last access caused none.
Public Function CacheLastFillPA() As Long
    CacheLastFillPA = mLastFillPA
End Function

Public Function CacheLineValid(ByVal level As Long, ByVal idx As Long) As Boolean
    RequireLine level, idx
    CacheLineValid = mValid(level, idx)
End Function

Public Function CacheLineTag(ByVal level As Long, ByVal idx As Long) As Long
    RequireLine level, idx
    CacheLineTag = mTag(level, idx)
End Function

Public Function CacheLineByte(ByVal level As Long, ByVal idx As Long, ByVal off As Long) As Byte
    RequireLine level, idx
    If off < 0 Or off > 3 Then Err.Raise LCA_ERR_CACHE_ARGUMENT, "LcaCache", "CACHE_OFFSET"
    CacheLineByte = mData(level, idx, off)
End Function

' True when the PA's line is resident in the level (no counters touched).
Public Function CacheContains(ByVal level As Long, ByVal pa As Long) As Boolean
    Dim n As Long, lineNo As Long
    n = CacheLineCount(level)
    RamRequirePhysicalAddress pa
    lineNo = pa \ LCA_CACHE_LINE_BYTES
    CacheContains = mValid(level, lineNo Mod n) And mTag(level, lineNo Mod n) = lineNo \ n
End Function

Public Function CacheLastEvent() As String
    CacheLastEvent = mLastEvent
End Function

Public Function CacheTraceCount() As Long
    CacheTraceCount = mTraceCount
End Function

Public Function CacheTraceLine(ByVal index As Long) As String
    If index < 0 Or index >= mTraceCount Then Err.Raise LCA_ERR_CACHE_ARGUMENT, "LcaCache", "CACHE_TRACE_INDEX"
    CacheTraceLine = mTrace(index)
End Function

Private Sub RequireLine(ByVal level As Long, ByVal idx As Long)
    If idx < 0 Or idx >= CacheLineCount(level) Then _
        Err.Raise LCA_ERR_CACHE_ARGUMENT, "LcaCache", "CACHE_LINE_INDEX"
End Sub
