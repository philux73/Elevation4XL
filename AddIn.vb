Imports System.Linq.Expressions
Imports ExcelDna.Integration
Imports ExcelDna.IntelliSense
Imports ExcelDna.Registration


Public Class AddIn
    Implements IExcelAddIn

    ''' <summary>
    ''' Gets the <see href="https://www.opentopodata.org/">Open Topo Data </see>HTTP client.
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property Client As OpenTopoDataHttpClient

    ''' <summary>
    ''' Code here will run every time the add-in is loaded.
    ''' </summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        _Client = New OpenTopoDataHttpClient()
        RegisterFunctions()
        IntelliSenseServer.Install() 'Must run after function registration
    End Sub

    ''' <summary>
    ''' Code here will run when the add-in is removed in the Add-Ins dialog, but not when Excel closes normally.
    ''' </summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        Client.Dispose()
        IntelliSenseServer.Uninstall()
    End Sub

    ''' <summary>
    ''' Get all the ExcelFunction functions, process and register
    ''' </summary>
    Private Sub RegisterFunctions()
        ExcelRegistration.GetExcelFunctions.
            ProcessAsyncRegistrations(nativeAsyncIfAvailable:=False).
            ProcessParameterConversions(GetPostAsyncReturnConversionConfig).
            RegisterFunctions()
    End Sub

    ''' <summary>
    ''' This conversion replaces the default #N/A return value of async functions with the #GETTING_DATA value.
    ''' </summary>
    ''' <remarks>
    ''' Source: <see href="https://github.com/Excel-DNA/Registration/blob/master/Source/Samples/Registration.Sample/ExampleAddIn.cs">https://github.com/Excel-DNA/Registration/blob/master/Source/Samples/Registration.Sample/ExampleAddIn.cs</see><br/>
    ''' </remarks>
    ''' <returns></returns>
    Private Shared Function GetPostAsyncReturnConversionConfig() As ParameterConversionConfiguration
        Return New ParameterConversionConfiguration().
            AddReturnConversion(Function(type, customAttributes)
                                    Return If(type <> GetType(Object), Nothing,
                                    (CType((Function(returnValue As Object) If(returnValue.Equals(ExcelError.ExcelErrorNA),
                                                                                            ExcelError.ExcelErrorGettingData,
                                                                                            returnValue)),
                                                                                            Expression(Of Func(Of Object, Object)))))

                                End Function)
    End Function

End Class