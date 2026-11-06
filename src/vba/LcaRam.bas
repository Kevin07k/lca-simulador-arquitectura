Attribute VB_Name = "LcaRam"
Option Explicit

' Physical backing store: eight frames, 32 bytes each. No Excel worksheet dependency.
Public Const LCA_RAM_FRAMES As Long = 8
Public Const LCA_RAM_FRAME_SIZE As Long = 32
Public Const LCA_RAM_BYTES As Long = 256
Public Const LCA_ERR_RAM_ADDRESS As Long = vbObjectError + 2061
Public Const LCA_ERR_RAM_PROTECTION As Long = vbObjectError + 2062

Private mFrames(0 To 7, 0 To 31) As Byte

Public Sub RamReset()
    Dim frame As Long, offset As Long
    For frame = 0 To LCA_RAM_FRAMES - 1
        For offset = 0 To LCA_RAM_FRAME_SIZE - 1
            mFrames(frame, offset) = 0
        Next offset
    Next frame
    ' P02: LCD video buffer starts as eight spaces; P13/P14 own device behavior.
    For offset = 8 To 15
        mFrames(7, offset) = &H20
    Next offset
End Sub

Public Function RamFrameOf(ByVal pa As Long) As Long
    RamRequirePhysicalAddress pa
    RamFrameOf = pa \ LCA_RAM_FRAME_SIZE
End Function

Public Function RamOffsetOf(ByVal pa As Long) As Long
    RamRequirePhysicalAddress pa
    RamOffsetOf = pa Mod LCA_RAM_FRAME_SIZE
End Function

' Raw physical access is for loader, page-frame maintenance and verification.
' CPU requests MUST use the restricted user/I/O paths below and the bus adapter.
Public Function RamReadPhysical(ByVal pa As Long) As Byte
    RamRequirePhysicalAddress pa
    RamReadPhysical = mFrames(pa \ LCA_RAM_FRAME_SIZE, pa Mod LCA_RAM_FRAME_SIZE)
End Function

Public Sub RamWritePhysical(ByVal pa As Long, ByVal value As Long)
    RamRequirePhysicalAddress pa
    CpuRequireByte value
    mFrames(pa \ LCA_RAM_FRAME_SIZE, pa Mod LCA_RAM_FRAME_SIZE) = CByte(value)
End Sub

Public Sub RamRequirePhysicalAddress(ByVal pa As Long)
    If pa < 0 Or pa >= LCA_RAM_BYTES Then _
        Err.Raise LCA_ERR_RAM_ADDRESS, "LcaRam", "PHYSICAL_ADDRESS_OUT_OF_RANGE"
End Sub

' Frames 0..6 only. P07 translates/validates the virtual address first.
Public Sub RamRequireUserAddress(ByVal pa As Long)
    RamRequirePhysicalAddress pa
    If pa >= &HE0 Then Err.Raise LCA_ERR_RAM_PROTECTION, "LcaRam", "RAM_USER_PROTECTION"
End Sub

Public Function RamReadUser(ByVal pa As Long) As Byte
    RamRequireUserAddress pa
    RamReadUser = RamReadPhysical(pa)
End Function

Public Sub RamWriteUser(ByVal pa As Long, ByVal value As Long)
    RamRequireUserAddress pa
    RamWritePhysical pa, value
End Sub

' Only the I/O controller (P13/P14) may call these for E0..EFh.
Public Function RamReadIO(ByVal pa As Long) As Byte
    RequireIoAddress pa
    RamReadIO = RamReadPhysical(pa)
End Function

Public Sub RamWriteIO(ByVal pa As Long, ByVal value As Long)
    RequireIoAddress pa
    RamWritePhysical pa, value
End Sub

Private Sub RequireIoAddress(ByVal pa As Long)
    RamRequirePhysicalAddress pa
    If pa < &HE0 Or pa > &HEF Then _
        Err.Raise LCA_ERR_RAM_PROTECTION, "LcaRam", "RAM_IO_PROTECTION"
End Sub
