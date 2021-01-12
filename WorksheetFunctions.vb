Imports ExcelDna.Integration

''' <summary>
''' Module containing all Excel User Defined Functions.
''' </summary>
Public Module WorksheetFunctions

    <ExcelFunction(Name:="ELEVATION", Description:="Returns the elevation of a location", HelpTopic:="https://www.opentopodata.org/")>
    Public Async Function GetElevationAsync(<ExcelArgument(Description:="is the latitude of the location (WGS-84 format).")> latitude As Double,
                                            <ExcelArgument(Description:="is the longitude of the location (WGS-84 format).")> longitude As Double,
                                            <ExcelArgument(Description:="is an optional number assigned to a digital elevation model dataset: for ASTER use 0 or omit, for ETOPO1 use 1, for EU-DEM use 2, for Mapzen use 3, for NED use 4, for NZ DEM use 5, for SRTM use 6, for EMOD bathymetry use 7, for GEBCO bathymetry use 8.")> dataset As Object,
                                            <ExcelArgument(Description:="is an optional number assigned to the interpolation method: for Nearest use 0, for Bilinear use 1 or omit, for Cubic use 2.")> interpolation As Object) As Task(Of Object)
        If TypeOf dataset Is ExcelMissing Then dataset = OpenTopoDataHttpClient.DigitalElevationModel.aster30m
        If TypeOf interpolation Is ExcelMissing Then interpolation = OpenTopoDataHttpClient.Interpolation.bilinear
        Try
            With AddIn.Client
                Dim myResponseId As Integer = .GetResponseId(latitude, longitude, dataset, interpolation)
                Await .Downloads(myResponseId)
                Return .GetElevation(myResponseId, latitude, longitude)
            End With
        Catch ex As ArgumentException
            Return ExcelError.ExcelErrorValue
        Catch ex As Exception
            Return ExcelError.ExcelErrorNull 'Unfortunataly cannot use ExcelErrorNA cause would be converted to ExcelErrorGettingData.
        End Try
    End Function

End Module
