Attribute VB_Name = "LcaAlu"
Option Explicit

Public Enum LcaAluOperation
    LcaADD = &H20
    LcaSUB = &H21
    LcaCMP = &H22
    LcaINC = &H23
    LcaDEC = &H24
End Enum

' Arithmetic primitive only: no PC advance, fetch, bus or worksheet access.
' Binary operations receive a register code; unary operations ignore sourceCode.
' P12 validates canonical instruction fields before calling this primitive.
Public Function AluExecute(ByVal operation As Long, ByVal destinationCode As Long, _
                           Optional ByVal sourceCode As Long = 0) As Byte
    Dim leftValue As Long
    Dim rightValue As Long
    Dim wideResult As Long
    Dim result As Long
    Dim carryOrBorrow As Boolean

    ' Validate everything before committing register or flag changes.
    Select Case operation
        Case LcaADD, LcaSUB, LcaCMP, LcaINC, LcaDEC
        Case Else
            Err.Raise LCA_ERR_ALU, "LcaAlu", "INVALID_ALU_OPERATION"
    End Select
    leftValue = CLng(CpuReadRegister(destinationCode))
    Select Case operation
        Case LcaADD, LcaSUB, LcaCMP
            rightValue = CLng(CpuReadRegister(sourceCode))
        Case LcaINC, LcaDEC
            rightValue = 1
    End Select

    Select Case operation
        Case LcaADD, LcaINC
            wideResult = leftValue + rightValue
            carryOrBorrow = (wideResult > 255)
        Case LcaSUB, LcaCMP, LcaDEC
            wideResult = leftValue - rightValue
            carryOrBorrow = (wideResult < 0)
    End Select
    result = wideResult And &HFF&
    If operation <> LcaCMP Then CpuWriteRegister destinationCode, result
    CpuSetArithmeticFlags result, carryOrBorrow
    AluExecute = CByte(result)
End Function
