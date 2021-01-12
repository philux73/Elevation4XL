Public Class Location
    Public Property Lat As Double
    Public Property Lng As Double
End Class

Public Class Result
    Public Property Elevation As Double?
    Public Property Location As Location
End Class

Public Class Root
    Public Property Results As List(Of Result)
    Public Property Status As String
End Class
