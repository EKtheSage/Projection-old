Imports ExcelDna.Integration

Public Module ChainLadder
    'This module creates functions to handle calculations of ATA factors
    <ExcelFunction(Description:="Create ATA factors, specify type of factors(volume wtd, straight avg, etc), number of points")> _
    Public Function ATA(ByVal rng As Object, ATAType As String, numPoints As Integer) As Double
        Return ATA
    End Function
End Module
