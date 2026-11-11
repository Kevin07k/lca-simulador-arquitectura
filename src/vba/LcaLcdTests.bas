Attribute VB_Name = "LcaLcdTests"
Option Explicit

Private mChecks As Long

' Desktop Excel verification for P14: video RAM E8..EFh, OUT with Ld_LCD, cursor, 8th/9th character,
' RAM = screen, no overlap with the P13 keyboard queue.
Public Sub VerificarLcdP14()
    Dim failure As Long
    Dim i As Long, traceBefore As Long, text As String
    Dim cycle As LcaBusCycle
    On Error GoTo failed
    mChecks = 0

    ' --- Section A: reset ------------------------------------------------------------------------
    CpuReset
    RamReset
    MemoryPathReset
    KeyboardReset
    Call KeyboardPush(&H5A)                 ' one queued key that the LCD must never touch
    LcdReset
    Check LcdCursor() = 0 And LcdWrites() = 0, "A: cursor 0 after reset"
    For i = 0 To 7
        Check RamReadIO(&HE8 + i) = &H20 And LcdCell(i) = &H20, "A: cell " & CStr(i) & " is a space in RAM"
    Next i
    Check LcdText() = "        ", "A: screen shows eight spaces"
    Check LCA_LCD_BASE_PA = &HE8 And LCA_LCD_SIZE = 8 And LCA_KBD_BASE_PA + LCA_KBD_SIZE = LCA_LCD_BASE_PA, _
        "A: video E8..EFh starts right after the queue E0..E7h (no overlap)"
    Check RamReadIO(&HE0) = &H5A And KeyboardCount() = 1, "A: LCD reset leaves the keyboard queue intact"

    ' --- Section B: P02 example, OUT AX = 'A' ---------------------------------------------------
    CpuWriteRegister LcaAX, &H41
    CpuSetArithmeticFlags 0, True           ' ZF=1 CF=1: OUT must preserve them
    traceBefore = BusTraceCount()
    Check LcdOut(LcaAX) = 0, "B: OUT AX writes cell 0"
    Check RamReadIO(&HE8) = &H41 And LcdCursor() = 1, "B: E8h = 41h, cursor = 1"
    Check LcdText() = "A       ", "B: screen shows A"
    Check CpuReadRegister(LcaAX) = &H41, "B: AX unchanged by OUT"
    Check CpuFlagIsSet(LCA_FLAG_ZF) And CpuFlagIsSet(LCA_FLAG_CF) And Not CpuFlagIsSet(LCA_FLAG_SF), "B: flags preserved"
    Check BusTraceCount() = traceBefore + 1, "B: exactly one bus cycle for OUT"
    cycle = BusSnapshot(LCA_BUS_EXTERNAL)
    Check cycle.IOSignal And cycle.WriteSignal And Not cycle.ReadSignal, "B: IO=1 W=1 R=0"
    Check cycle.LdLCD And Not cycle.EnKBD, "B: Ld_LCD=1 En_KBD=0"
    Check cycle.EnableSignal And cycle.LoadSignal And cycle.ClockSignal, "B: Enable, Load and Clock set"
    Check cycle.PA = &HE8 And cycle.VA = &H1E8, "B: PA E8h / VA 1E8h"
    Check cycle.Driver = "AX" And cycle.Receiver = "LCD" And cycle.Data = &H41, "B: AX drives 41h into LCD"
    Check InStr(1, ControlLastWord(), "LCD=1", vbBinaryCompare) > 0, "B: control word shows LCD=1"

    ' --- Section C: eighth character fills the last cell, cursor wraps ---------------------------
    text = "BCDEFGH"
    For i = 1 To 7
        CpuWriteRegister LcaBX, Asc(Mid$(text, i, 1))
        Check LcdOut(LcaBX) = i, "C: OUT " & Mid$(text, i, 1) & " goes to cell " & CStr(i)
    Next i
    Check LcdText() = "ABCDEFGH", "C: eight characters fill the screen"
    Check RamReadIO(&HEF) = Asc("H") And LcdCursor() = 0 And LcdWraps() = 1, "C: 8th fills EFh, cursor wraps to 0"

    ' --- Section D: ninth character replaces the first -----------------------------------------
    CpuWriteRegister LcaCX, Asc("I")
    Check LcdOut(LcaCX) = 0, "D: ninth OUT goes to cell 0"
    Check LcdText() = "IBCDEFGH" And RamReadIO(&HE8) = Asc("I") And LcdCursor() = 1, "D: I replaces A, cursor 1"
    Check LcdWrites() = 9, "D: nine OUT counted"

    ' --- Section E: RAM and screen always agree ------------------------------------------------
    For i = 0 To 7
        Check LcdCell(i) = RamReadIO(&HE8 + i) And LcdChar(i) = Mid$(LcdText(), i + 1, 1), "E: cell " & CStr(i) & " = RAM"
    Next i
    CpuWriteRegister LcaDX, 7               ' non-printable byte: stored as is, shown as "?"
    Call LcdOut(LcaDX)
    Check RamReadIO(&HE9) = 7 And LcdChar(1) = "?", "E: non-printable 07h stored in RAM, shown as ?"

    ' --- Section F: rejected OUT changes nothing ------------------------------------------------
    traceBefore = BusTraceCount()
    failure = 0
    On Error Resume Next
    Call LcdOut(9)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_REGISTER And LcdCursor() = 2 And BusTraceCount() = traceBefore, _
        "F: invalid register, no bus cycle, cursor kept"
    ControlStartFetch 0, 0
    failure = 0
    On Error Resume Next
    Call LcdOut(LcaAX)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_SEQUENCE And LcdCursor() = 2 And RamReadIO(&HEA) = Asc("C"), _
        "F: OUT during a fetch is SEQUENCE_BUSY, video RAM unchanged"
    MemoryPathReset

    ' --- Section G: echo keyboard -> LCD (P13 + P14) --------------------------------------------
    Check KeyboardCount() = 1 And RamReadIO(&HE0) = &H5A, "G: LCD writes never touched the queue"
    Check KeyboardIn(LcaAX) And CpuReadRegister(LcaAX) = &H5A, "G: IN AX = 5Ah 'Z'"
    Check LcdOut(LcaAX) = 2 And LcdChar(2) = "Z", "G: OUT AX shows Z in cell 2"
    Check KeyboardIsEmpty() And RamReadIO(&HE0) = &H5A, "G: queue consumed once, slot byte kept"

    ' --- Section H: the drawn band matches RAM (when the LCA sheet is built) ---------------------
    LcdPanelActualizar
    Check LcdPanelCoherente(), "H: LCD band = RAM E8..EFh"

    LcaReport "P14 PASS: " & mChecks & " comprobaciones / assertions." & vbCrLf & _
           "Video E8-EFh, OUT con Ld_LCD=1, cursor mod 8, caracteres 8 y 9, RAM = pantalla. P12 decide el avance de PC.", True
    Exit Sub
failed:
    LcaReport "P14 FAIL after " & mChecks & " checks: " & Err.Description, False
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2142, "LcaLcdTests", message
End Sub
