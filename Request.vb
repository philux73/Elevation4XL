''' <summary>
''' Povides a class for recording a request from Excel.
''' </summary>
Public Class Request

    Private Const unassigned = -1

    Public ReadOnly Property Dataset As OpenTopoDataHttpClient.DigitalElevationModel
    Public ReadOnly Property Interpolation As OpenTopoDataHttpClient.Interpolation
    Public ReadOnly Property Latitude As Double
    Public ReadOnly Property Longitude As Double
    Public Property ResponseId As Integer

    Public Sub New(dataset As OpenTopoDataHttpClient.DigitalElevationModel, latitude As Double, longitude As Double, interpolation As OpenTopoDataHttpClient.Interpolation)
        If [Enum].IsDefined(GetType(OpenTopoDataHttpClient.DigitalElevationModel), dataset) Then
            _Dataset = dataset
        Else
            Throw New ArgumentException(String.Format(My.Resources.E001, dataset, "dataset"), "dataset")
        End If
        If [Enum].IsDefined(GetType(OpenTopoDataHttpClient.Interpolation), interpolation) Then
            _Interpolation = interpolation
        Else
            Throw New ArgumentException(String.Format(My.Resources.E001, interpolation, "interpolation"), "interpolation")
        End If
        If latitude >= -90 And latitude <= 90 Then
            _Latitude = latitude
        Else
            Throw New ArgumentException(String.Format(My.Resources.E001, latitude, "latitude"), "latitude")
        End If
        If longitude >= -180 And longitude <= 180 Then
            _Longitude = longitude
        Else
            Throw New ArgumentException(String.Format(My.Resources.E001, longitude, "longitude"), "longitude")
        End If
        ResponseId = unassigned
    End Sub

End Class
