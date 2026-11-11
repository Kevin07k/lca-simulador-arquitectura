Attribute VB_Name = "LcaSwap"
Option Explicit

' Backing image: exactly 15 pageable VPNs, 32 bytes each. VPN 15 is I/O.
Public Const LCA_SWAP_PAGES As Long = 15
Public Const LCA_SWAP_PAGE_BYTES As Long = 32
Public Const LCA_ERR_SWAP_ADDRESS As Long = vbObjectError + 2081

Private mImage(0 To 14, 0 To 31) As Byte

Public Sub SwapReset()
    Dim vpn As Long, offset As Long
    For vpn = 0 To 14
        For offset = 0 To 31
            mImage(vpn, offset) = 0
        Next offset
    Next vpn
End Sub

' Loader calls SwapWriteByte after reset, before executing the program.
' Pager also uses it to preserve each byte of an evicted dirty page.
Public Sub SwapWriteByte(ByVal vpn As Long, ByVal offset As Long, ByVal value As Long)
    RequireSwapAddress vpn, offset
    CpuRequireByte value
    mImage(vpn, offset) = CByte(value)
End Sub

Public Function SwapReadByte(ByVal vpn As Long, ByVal offset As Long) As Byte
    RequireSwapAddress vpn, offset
    SwapReadByte = mImage(vpn, offset)
End Function

Private Sub RequireSwapAddress(ByVal vpn As Long, ByVal offset As Long)
    If vpn < 0 Or vpn >= LCA_SWAP_PAGES Then _
        Err.Raise LCA_ERR_SWAP_ADDRESS, "LcaSwap", "SWAP_VPN_OUT_OF_RANGE_OR_RESERVED"
    If offset < 0 Or offset >= LCA_SWAP_PAGE_BYTES Then _
        Err.Raise LCA_ERR_SWAP_ADDRESS, "LcaSwap", "SWAP_OFFSET_OUT_OF_RANGE"
End Sub
