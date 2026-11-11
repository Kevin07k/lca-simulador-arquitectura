Attribute VB_Name = "LcaCpu"
Option Explicit

Public Const LCA_FLAG_ZF As Long = 1
Public Const LCA_FLAG_CF As Long = 2
Public Const LCA_FLAG_SF As Long = 4
Public Const LCA_ERR_REGISTER As Long = vbObjectError + 2001
Public Const LCA_ERR_BYTE As Long = vbObjectError + 2002
Public Const LCA_ERR_ADDRESS As Long = vbObjectError + 2003
Public Const LCA_ERR_CONTROL As Long = vbObjectError + 2004
Public Const LCA_ERR_ALU As Long = vbObjectError + 2005

Public Enum LcaRegister
    LcaAX = 0
    LcaBX = 1
    LcaCX = 2
    LcaDX = 3
End Enum

Public Enum LcaControlRegister
    LcaPC = 0
    LcaMAR_V = 1
    LcaMAR_P = 2
    LcaMDR = 3
    LcaIR = 4
End Enum

Public Type LcaCpuState
    AX As Byte
    BX As Byte
    CX As Byte
    DX As Byte
    PC As Long
    MAR_V As Long
    MAR_P As Byte
    MDR As Byte
    IR As Byte
    flags As Byte
End Type

Private mState As LcaCpuState

' CPU-local reset. Other modules own memory, buses, cache and I/O reset.
Public Sub CpuReset()
    Dim cleared As LcaCpuState
    mState = cleared
End Sub

' Return a copy: callers cannot bypass validation by editing global state.
Public Function CpuSnapshot() As LcaCpuState
    CpuSnapshot = mState
End Function

Public Sub CpuRequireByte(ByVal value As Long)
    If value < 0 Or value > 255 Then
        Err.Raise LCA_ERR_BYTE, "LcaCpu", "BYTE_OUT_OF_RANGE"
    End If
End Sub

Public Function CpuReadRegister(ByVal registerCode As Long) As Byte
    Select Case registerCode
        Case LcaAX: CpuReadRegister = mState.AX
        Case LcaBX: CpuReadRegister = mState.BX
        Case LcaCX: CpuReadRegister = mState.CX
        Case LcaDX: CpuReadRegister = mState.DX
        Case Else
            Err.Raise LCA_ERR_REGISTER, "LcaCpu", "INVALID_REGISTER"
    End Select
End Function

Public Sub CpuWriteRegister(ByVal registerCode As Long, ByVal value As Long)
    Dim previous As Byte
    previous = CpuReadRegister(registerCode)
    CpuRequireByte value
    Select Case registerCode
        Case LcaAX: mState.AX = CByte(value)
        Case LcaBX: mState.BX = CByte(value)
        Case LcaCX: mState.CX = CByte(value)
        Case LcaDX: mState.DX = CByte(value)
    End Select
End Sub

' Range validation only. Fetch alignment/protection belongs to the P12 decoder.
Public Sub CpuWriteControl(ByVal registerCode As Long, ByVal value As Long)
    Select Case registerCode
        Case LcaPC, LcaMAR_V
            If value < 0 Or value > 511 Then
                Err.Raise LCA_ERR_ADDRESS, "LcaCpu", "VA_OUT_OF_RANGE"
            End If
        Case LcaMAR_P, LcaMDR, LcaIR
            CpuRequireByte value
        Case Else
            Err.Raise LCA_ERR_CONTROL, "LcaCpu", "INVALID_CONTROL_REGISTER"
    End Select
    Select Case registerCode
        Case LcaPC: mState.PC = value
        Case LcaMAR_V: mState.MAR_V = value
        Case LcaMAR_P: mState.MAR_P = CByte(value)
        Case LcaMDR: mState.MDR = CByte(value)
        Case LcaIR: mState.IR = CByte(value)
    End Select
End Sub

Public Function CpuFlagIsSet(ByVal mask As Long) As Boolean
    CpuFlagIsSet = ((CLng(mState.flags) And mask) <> 0)
End Function

Public Sub CpuSetArithmeticFlags(ByVal result As Long, ByVal carryOrBorrow As Boolean)
    Dim flags As Long
    CpuRequireByte result
    If result = 0 Then flags = flags Or LCA_FLAG_ZF
    If carryOrBorrow Then flags = flags Or LCA_FLAG_CF
    If result >= 128 Then flags = flags Or LCA_FLAG_SF
    mState.flags = CByte(flags)
End Sub
