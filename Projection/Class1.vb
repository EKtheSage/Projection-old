Imports ExcelDna.Integration

Public Module MyFunctions

    <ExcelFunction(Description:="My first .NET function")> _
    Public Function HelloDna(name As String) As String
        Return "Hello " & name
    End Function

End Module