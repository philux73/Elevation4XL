''' <summary>
''' Provides a class for recording a response to multiple Excel requests. 
''' </summary>
Public Class Response

    Private Const notApplicable As Integer = -1

    Private _data As Root

    ''' <summary>
    ''' The request URI which produced this response.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Question As String
    ''' <summary>
    ''' The deserialized HTTP response.
    ''' </summary>
    ''' <returns></returns>
    Public Property Data As Root
        Get
            Return _data
        End Get
        Set(value As Root)
            If value.Results IsNot Nothing Then
                Countdown = value.Results.Count
            Else
                Countdown = notApplicable
            End If
            _data = value
        End Set
    End Property
    ''' <summary>
    ''' The number of remaining Excel requests which still need to retrieve data from this response.
    ''' </summary>
    ''' <remarks>When value equals 0 and Response.Data is not null, the response can be cleared safely.</remarks>
    ''' <returns></returns>
    Public Property Countdown As Integer

    Public Sub New(question As String)
        _Question = question
    End Sub

End Class
