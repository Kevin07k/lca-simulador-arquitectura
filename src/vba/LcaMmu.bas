Attribute VB_Name = "LcaMmu"
Option Explicit

' P07: address validation and page table only. P08 owns faults, swap and replacement.
Public Const MMU_FETCH As Long = 1
Public Const MMU_READ As Long = 2
Public Const MMU_WRITE As Long = 3
Public Const LCA_ERR_MMU_RANGE As Long = vbObjectError + 2071
Public Const LCA_ERR_MMU_PROTECTION As Long = vbObjectError + 2072
Public Const LCA_ERR_PAGE_FAULT As Long = vbObjectError + 2073
Public Const LCA_ERR_MMU_MAPPING As Long = vbObjectError + 2074
Public Const LCA_ERR_MMU_STATE As Long = vbObjectError + 2075

Public Type LcaPageEntry
    VPN As Long
    Present As Boolean
    PFN As Long                   ' -1 when absent, including reserved VPN 15
    Dirty As Boolean
    Referenced As Boolean
    Reserved As Boolean
End Type

Private mPresent(0 To 15) As Boolean
Private mPFN(0 To 15) As Long
Private mDirty(0 To 15) As Boolean
Private mRef(0 To 15) As Boolean
Private mFrameOwner(0 To 6) As Long
Private mLastFaultVA As Long
Private mLastFaultVPN As Long
Private mInitialized As Boolean

Public Sub MmuReset()
    Dim vpn As Long, pfn As Long
    For vpn = 0 To 15
        mPresent(vpn) = False
        mPFN(vpn) = -1
        mDirty(vpn) = False
        mRef(vpn) = False
    Next vpn
    For pfn = 0 To 6
        mFrameOwner(pfn) = -1
    Next pfn
    mLastFaultVA = -1
    mLastFaultVPN = -1
    mInitialized = True
End Sub

Public Function MmuPage(ByVal vpn As Long) As LcaPageEntry
    Dim entry As LcaPageEntry
    EnsureInitialized
    RequireVpn vpn
    entry.VPN = vpn
    entry.Present = mPresent(vpn)
    entry.PFN = mPFN(vpn)
    entry.Dirty = mDirty(vpn)
    entry.Referenced = mRef(vpn)
    entry.Reserved = (vpn = 15)
    MmuPage = entry
End Function

Public Function MmuVpnOf(ByVal va As Long) As Long
    RequireVa va
    MmuVpnOf = va \ 32
End Function

Public Function MmuOffsetOf(ByVal va As Long) As Long
    RequireVa va
    MmuOffsetOf = va Mod 32
End Function

' P08 calls this only after loading a page into a free/evicted frame and invalidating stale cache lines.
Public Sub MmuMap(ByVal vpn As Long, ByVal pfn As Long)
    EnsureInitialized
    RequireUserVpn vpn
    If pfn < 0 Or pfn > 6 Then Err.Raise LCA_ERR_MMU_MAPPING, "LcaMmu", "PFN_RESERVED_OR_INVALID"
    If mPresent(vpn) Then
        If mPFN(vpn) = pfn Then Exit Sub
        Err.Raise LCA_ERR_MMU_MAPPING, "LcaMmu", "VPN_ALREADY_MAPPED"
    End If
    If mFrameOwner(pfn) <> -1 Then Err.Raise LCA_ERR_MMU_MAPPING, "LcaMmu", "FRAME_IN_USE"
    mPresent(vpn) = True
    mPFN(vpn) = pfn
    mFrameOwner(pfn) = vpn
    mDirty(vpn) = False
    mRef(vpn) = False
End Sub

' P08 must write back dirty data and invalidate physical cache lines before calling MmuUnmap.
Public Sub MmuUnmap(ByVal vpn As Long)
    Dim pfn As Long
    EnsureInitialized
    RequireUserVpn vpn
    If Not mPresent(vpn) Then Exit Sub
    If mDirty(vpn) Then Err.Raise LCA_ERR_MMU_STATE, "LcaMmu", "DIRTY_PAGE_NEEDS_WRITEBACK"
    pfn = mPFN(vpn)
    mFrameOwner(pfn) = -1
    mPresent(vpn) = False
    mPFN(vpn) = -1
    mDirty(vpn) = False
    mRef(vpn) = False
End Sub

Public Sub MmuClearDirty(ByVal vpn As Long)
    RequireResident vpn
    mDirty(vpn) = False
End Sub

Public Sub MmuClearReference(ByVal vpn As Long)
    RequireResident vpn
    mRef(vpn) = False
End Sub

' Returns False for an absent page and PA=-1; a valid PFN=0 instead returns True.
' Invalid addresses/permissions raise fatal errors. No page is loaded here.
Public Function MmuTryTranslate(ByVal va As Long, ByVal accessType As Long, ByRef pa As Long) As Boolean
    Dim vpn As Long
    EnsureInitialized
    pa = -1
    RequireAccess va, accessType
    vpn = va \ 32
    If Not mPresent(vpn) Then
        mLastFaultVA = va
        mLastFaultVPN = vpn
        Exit Function
    End If
    pa = mPFN(vpn) * 32 + (va Mod 32)
    mLastFaultVA = -1
    mLastFaultVPN = -1
    MmuTryTranslate = True
End Function

Public Function MmuTranslate(ByVal va As Long, ByVal accessType As Long) As Long
    Dim pa As Long
    If Not MmuTryTranslate(va, accessType, pa) Then _
        Err.Raise LCA_ERR_PAGE_FAULT, "LcaMmu", "PAGE_FAULT VPN=" & CStr(mLastFaultVPN)
    MmuTranslate = pa
End Function

' Ref/dirty change only AFTER the bus/RAM operation completed successfully.
Public Sub MmuCommitAccess(ByVal va As Long, ByVal accessType As Long, ByVal pa As Long)
    Dim translated As Long, vpn As Long
    translated = MmuTranslate(va, accessType)
    If translated <> pa Then Err.Raise LCA_ERR_MMU_STATE, "LcaMmu", "STALE_TRANSLATION"
    vpn = va \ 32
    mRef(vpn) = True
    If accessType = MMU_WRITE Then mDirty(vpn) = True
End Sub

Public Function MmuLastFaultVA() As Long
    EnsureInitialized
    MmuLastFaultVA = mLastFaultVA
End Function

Public Function MmuLastFaultVPN() As Long
    EnsureInitialized
    MmuLastFaultVPN = mLastFaultVPN
End Function

Private Sub RequireAccess(ByVal va As Long, ByVal accessType As Long)
    RequireVa va
    Select Case accessType
        Case MMU_FETCH
            If va > &HFF Then Err.Raise LCA_ERR_MMU_PROTECTION, "LcaMmu", "FETCH_PROTECTION"
        Case MMU_READ, MMU_WRITE
            If va < &H100 Or va > &H1DF Then _
                Err.Raise LCA_ERR_MMU_PROTECTION, "LcaMmu", "DATA_OR_IO_PROTECTION"
        Case Else
            Err.Raise LCA_ERR_MMU_STATE, "LcaMmu", "INVALID_ACCESS_TYPE"
    End Select
End Sub

Private Sub RequireVa(ByVal va As Long)
    If va < 0 Or va > 511 Then Err.Raise LCA_ERR_MMU_RANGE, "LcaMmu", "VA_OUT_OF_RANGE"
End Sub

Private Sub RequireVpn(ByVal vpn As Long)
    If vpn < 0 Or vpn > 15 Then Err.Raise LCA_ERR_MMU_RANGE, "LcaMmu", "VPN_OUT_OF_RANGE"
End Sub

Private Sub RequireUserVpn(ByVal vpn As Long)
    RequireVpn vpn
    If vpn = 15 Then Err.Raise LCA_ERR_MMU_PROTECTION, "LcaMmu", "RESERVED_IO_PAGE"
End Sub

Private Sub RequireResident(ByVal vpn As Long)
    EnsureInitialized
    RequireUserVpn vpn
    If Not mPresent(vpn) Then Err.Raise LCA_ERR_MMU_STATE, "LcaMmu", "PAGE_NOT_PRESENT"
End Sub

Private Sub EnsureInitialized()
    If Not mInitialized Then MmuReset
End Sub
