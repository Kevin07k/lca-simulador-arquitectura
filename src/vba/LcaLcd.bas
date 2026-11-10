Attribute VB_Name = "LcaLcd"
Option Explicit

' P14: 8x1 LCD controller. The video buffer is the eight reserved RAM bytes E8..EFh (PFN 7),
' right after the keyboard queue E0..E7h (P13) and without overlapping it.
'   Reset: E8..EFh = 20h (eight spaces), cursor = 0.
'   OUT R: one external cycle IO=1 W=1 Ld_LCD=1 writes R to E8h+cursor; only after the bus edge is the
'          byte stored and the cursor advanced modulo 8. Character eight fills the last cell (cursor
'          wraps to 0); character nine replaces the first one. OUT never waits and never changes the
'          keyboard queue, the registers or the flags.
'   The display is always read back from RAM, so screen and RAM cannot disagree.
' No worksheet dependency.
Public Const LCA_LCD_BASE_PA As Long = &HE8
Public Const LCA_LCD_BASE_VA As Long = &H1E8
Public Const LCA_LCD_SIZE As Long = 8
Public Const LCA_LCD_BLANK As Long = &H20
Public Const LCA_ERR_LCD As Long = vbObjectError + 2141
Private Const TRACE_MAX As Long = 64

Private mReady As Boolean
Private mCursor As Long
Private mWrites As Long
Private mWraps As Long
Private mTrace(0 To TRACE_MAX - 1) As String
Private mTraceCount As Long

Public Sub LcdReset()
    Dim i As Long
    For i = 0 To LCA_LCD_SIZE - 1
        RamWriteIO LCA_LCD_BASE_PA + i, LCA_LCD_BLANK
    Next i
    mCursor = 0
    mWrites = 0
    mWraps = 0
    mTraceCount = 0
    For i = 0 To TRACE_MAX - 1
        mTrace(i) = vbNullString
    Next i
    mReady = True
    AddTrace "RESET LCD = 8 espacios / spaces, cursor=0"
End Sub

' OUT R. Returns the cell (0..7) that was written.
Public Function LcdOut(ByVal sourceRegister As Long) As Long
    Dim pa As Long, value As Byte, cell As Long
    RequireLcd
    cell = mCursor
    pa = LCA_LCD_BASE_PA + cell
    value = ControlLcdWrite(pa, sourceRegister)
    ' Commit point: the bus edge succeeded, so the video RAM and the cursor change together.
    RamWriteIO pa, value
    mCursor = (mCursor + 1) Mod LCA_LCD_SIZE
    If mCursor = 0 Then mWraps = mWraps + 1
    mWrites = mWrites + 1
    AddTrace "OUT " & RegisterText(sourceRegister) & " = " & ByteText(value) & " -> " & PaText(pa) & _
             " celda/cell " & CStr(cell) & " Ld_LCD=1 (cursor=" & CStr(mCursor) & ")"
    LcdOut = cell
End Function

' --- Inspection (no state change) ---------------------------------------------------------

Public Function LcdReady() As Boolean
    LcdReady = mReady
End Function

Public Function LcdCursor() As Long
    RequireLcd
    LcdCursor = mCursor
End Function

' Raw video byte of a cell 0..7, read from RAM E8h + cell.
Public Function LcdCell(ByVal cell As Long) As Byte
    RequireLcd
    If cell < 0 Or cell >= LCA_LCD_SIZE Then Err.Raise LCA_ERR_LCD, "LcaLcd", "LCD_CELL_RANGE"
    LcdCell = RamReadIO(LCA_LCD_BASE_PA + cell)
End Function

' Character shown by a cell: printable ASCII 20h..7Eh as is, anything else as "?".
Public Function LcdChar(ByVal cell As Long) As String
    LcdChar = DisplayChar(LcdCell(cell))
End Function

' The eight characters on screen, always derived from RAM.
Public Function LcdText() As String
    Dim i As Long
    For i = 0 To LCA_LCD_SIZE - 1
        LcdText = LcdText & LcdChar(i)
    Next i
End Function

Public Function LcdWrites() As Long
    LcdWrites = mWrites
End Function

Public Function LcdWraps() As Long
    LcdWraps = mWraps
End Function

Public Function LcdTraceCount() As Long
    LcdTraceCount = mTraceCount
    If LcdTraceCount > TRACE_MAX Then LcdTraceCount = TRACE_MAX
End Function

' index 0 = newest retained line.
Public Function LcdTraceLine(ByVal index As Long) As String
    If index < 0 Or index >= LcdTraceCount() Then Err.Raise LCA_ERR_LCD, "LcaLcd", "LCD_TRACE_INDEX"
    LcdTraceLine = mTrace((mTraceCount - 1 - index) Mod TRACE_MAX)
End Function

Public Function LcdDisplayChar(ByVal value As Long) As String
    CpuRequireByte value
    LcdDisplayChar = DisplayChar(value)
End Function

' --- Internals -----------------------------------------------------------------------------

Private Function DisplayChar(ByVal value As Long) As String
    If value >= &H20 And value <= &H7E Then DisplayChar = Chr$(value) Else DisplayChar = "?"
End Function

Private Sub RequireLcd()
    If Not mReady Then Err.Raise LCA_ERR_LCD, "LcaLcd", "LCD_RESET_REQUIRED"
End Sub

Private Sub AddTrace(ByVal text As String)
    mTrace(mTraceCount Mod TRACE_MAX) = text
    mTraceCount = mTraceCount + 1
End Sub

Private Function ByteText(ByVal value As Long) As String
    ByteText = Right$("00" & Hex$(value), 2) & "h"
    If value >= 32 And value <= 126 Then ByteText = ByteText & " '" & Chr$(value) & "'"
End Function

Private Function PaText(ByVal pa As Long) As String
    PaText = Right$("00" & Hex$(pa), 2) & "h"
End Function

Private Function RegisterText(ByVal code As Long) As String
    Select Case code
        Case LcaAX: RegisterText = "AX"
        Case LcaBX: RegisterText = "BX"
        Case LcaCX: RegisterText = "CX"
        Case LcaDX: RegisterText = "DX"
        Case Else: RegisterText = "R?"
    End Select
End Function
