Attribute VB_Name = "LcaLoader"
Option Explicit

' P11 loader. The whole program is assembled and validated BEFORE the current image is replaced;
' a failed assembly leaves the loaded program and the execution state untouched.
' The 256-byte code image goes to the backing store of VPN 0..7 (virtual pages); data VPN 8..14
' start at zero. Pages reach RAM on demand through the P08/P10 paging path.
Public Const LCA_ERR_LOADER As Long = vbObjectError + 2111

Private mImage(0 To 255) As Byte
Private mLoaded As Boolean
Private mInstrCount As Long
Private mLoadCount As Long

' Assembles and, only if valid, loads the program and resets execution state.
Public Function LoaderLoadSource(ByVal source As Variant) As Boolean
    If Not AsmAssemble(source) Then Exit Function
    LoaderLoadAssembled
    LoaderLoadSource = True
End Function

Public Function LoaderLoadText(ByVal text As String) As Boolean
    If Not AsmAssembleText(text) Then Exit Function
    LoaderLoadAssembled
    LoaderLoadText = True
End Function

' Loads the result of the last successful AsmAssemble.
Public Sub LoaderLoadAssembled()
    Dim i As Long
    If Not AsmSucceeded() Then Err.Raise LCA_ERR_LOADER, "LcaLoader", "ASSEMBLE_VALID_PROGRAM_FIRST"
    For i = 0 To LCA_ASM_CODE_BYTES - 1
        mImage(i) = AsmByte(i)
    Next i
    mInstrCount = AsmInstructionCount()
    mLoaded = True
    mLoadCount = mLoadCount + 1
    LoaderResetExecution
End Sub

' P02 reset / validated load:
'  - CPU: AX..DX, FLAGS, MAR/MDR/IR cleared and PC = 000h (entry point);
'  - control idle, both buses High-Z (MemoryPathReset via CachePathReset);
'  - RAM 00..DFh cleared, LCD buffer eight spaces (RamReset); caches invalidated and counters zeroed;
'  - every PTE non-present, CLOCK hand 0, paging counters zeroed;
'  - backing image restored: loaded code in VPN 0..7, zero data in VPN 8..14.
' Keyboard queue/cursor reset belongs to P13/P14 when those modules exist.
Public Sub LoaderResetExecution()
    Dim vpn As Long, offset As Long
    If Not mLoaded Then Err.Raise LCA_ERR_LOADER, "LcaLoader", "NO_PROGRAM_LOADED"
    CpuReset
    CachePathReset
    For vpn = 0 To 7
        For offset = 0 To 31
            SwapWriteByte vpn, offset, mImage(vpn * 32 + offset)
        Next offset
    Next vpn
End Sub

Public Function LoaderLoaded() As Boolean
    LoaderLoaded = mLoaded
End Function

Public Function LoaderInstructionCount() As Long
    LoaderInstructionCount = mInstrCount
End Function

' Code pages that hold at least one program instruction (the rest are HLT fill).
Public Function LoaderCodePages() As Long
    LoaderCodePages = (mInstrCount * 4 + 31) \ 32
End Function

Public Function LoaderLoadCount() As Long
    LoaderLoadCount = mLoadCount
End Function

Public Function LoaderImageByte(ByVal address As Long) As Byte
    If address < 0 Or address > 255 Then Err.Raise LCA_ERR_LOADER, "LcaLoader", "IMAGE_ADDRESS"
    LoaderImageByte = mImage(address)
End Function

' Ready-to-type example: sums 5+4+3+2+1 into AX with a loop and stores it at 100h.
Public Function LoaderSampleProgram() As Variant
    LoaderSampleProgram = Array( _
        "; P11 ejemplo: AX = 5+4+3+2+1 = 0Fh, guardado en [100h]", _
        "        MOV AX, #00h      ; acumulador", _
        "        MOV CX, #05h      ; contador", _
        "bucle:  ADD AX, CX", _
        "        DEC CX", _
        "        JNZ bucle         ; repetir mientras CX <> 0", _
        "        ST [100h], AX", _
        "        LD BX, [100h]", _
        "fin:    HLT")
End Function
