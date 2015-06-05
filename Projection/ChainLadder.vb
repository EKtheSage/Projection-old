Imports ExcelDna.Integration

Public Module ChainLadder
    'This module creates functions to handle calculations of ATA factors

    'The function AtaAlg takes in a range
    <ExcelFunction(Description:="Create ATA factors, specify type of factors(volume wtd, straight avg, etc), number of points")> _
    Public Function AtaAlg(ByVal rng As Object, ATAType As String, numPoints As Integer) As Double
        'Takes in a cell, from bottom of the range
        'algorithms: wtd avg, simp avg, h/l avg, seasonal factors, least square
        'if it is simple ata, do cell.offset(-1, 1) for number of points / cell.offset
        Return AtaAlg
    End Function

    'The function ATA calculates individual ATA factors
    <ExcelFunction(Description:="Calculate individual ATA factors", IsMacroType:=True)> _
    Public Function ATA(<ExcelArgument(AllowReference:=True)> ByVal rng As Object) As Double
        'Take in a cell, find the cell.offset(0, 1), 
        'if it is not a blank, then do cell.offset(0, 1)/cell
        'if it is blank, return nothing DBNull.Value, or nullable type?
        'if the input cell is blank or zero, return nothing.
        Dim curAge As Object, nextAge As Object

        curAge = offsetFunction(rng, 0, 0)
        nextAge = offsetFunction(rng, 0, 1)
        ATA = nextAge / curAge
        Return ATA
    End Function

    'Helper functions
    Private Function referenceToRange(ByVal xlRef As ExcelReference) As Object
        Dim strAddress As String = XlCall.Excel(XlCall.xlfReftext, xlRef, True)
        referenceToRange = ExcelDnaUtil.Application.Range(strAddress)
    End Function

    Private Function offsetFunction(<ExcelArgument(AllowReference:=True)> ByVal rng As Object, row As Integer, column As Integer) As Double
        Dim offsetRef As ExcelReference
        offsetRef = CType(XlCall.Excel(XlCall.xlfOffset, rng, row, column), ExcelReference)
        Return CType(offsetRef.GetValue(), Double)
    End Function

End Module
