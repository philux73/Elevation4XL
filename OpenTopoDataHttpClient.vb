Imports System.Collections.Concurrent
Imports System.Net
Imports System.Net.Http
Imports System.Text.Json
Imports System.Threading

''' <summary>
''' Provides a class for sending HTTP requests and receiving HTTP responses to/from the <see href="https://www.opentopodata.org/">Open Topo Data API</see>.
''' </summary>
Public Class OpenTopoDataHttpClient
    Inherits HttpClient

    ''' <summary>
    ''' String constant: base Address of the Open Topo Data API.
    ''' </summary>
    Private Const apiBaseAddress As String = "https://api.opentopodata.org/v1/"
    ''' <summary>
    ''' Constant number: API's maximum number of locations per HTTP request.
    ''' </summary>
    Private Const apiMaxLocationsPerRequest As Integer = 100
    ''' <summary>
    ''' Constant number: API's maximum number of HTTP requests per second.
    ''' </summary>
    Private Const apiMaxThroughput As Double = 1

    ''' <summary>
    ''' Gets the list of active requests.
    ''' </summary>
    Private ReadOnly _requests As List(Of Request)
    ''' <summary>
    ''' Gest the dictionary of active responses.
    ''' </summary>
    Private ReadOnly _responses As Dictionary(Of Integer, Response)
    ''' <summary>
    ''' Gets the request processing lock.
    ''' </summary>
    ''' <remarks>Use this lock for  ensuring that requests are not assigned to a response for which the download process has started.</remarks>
    Private ReadOnly _requestLock As Object
    ''' <summary>
    ''' Gets the API lock.
    ''' </summary>
    ''' <remarks>Use this lock for ensuring that HTTP requests are processed sequentially.</remarks>
    Private ReadOnly _apiLock As Object
    ''' <summary>
    ''' Gets the stopwatch for throttling API requests.
    ''' </summary>
    Private ReadOnly _stopwatch As Stopwatch
    ''' <summary>
    ''' Gets or sets the current response ID.
    ''' </summary>
    Private _responseId As Integer

    ''' <summary>
    ''' Gets the dictionary of active Download tasks.
    ''' </summary>
    Public ReadOnly Property Downloads As Dictionary(Of Integer, Task)

    ''' <summary>
    ''' Enumeration of allowed digital elevation model datasets.
    ''' </summary>
    Public Enum DigitalElevationModel
        aster30m
        etopo1
        eudem25m
        mapzen
        ned10m
        nzdem8m
        srtm90m
        emod2018
        gebco2020
    End Enum
    ''' <summary>
    ''' Enumeration of allowed interpolation methods.
    ''' </summary>
    Public Enum Interpolation
        nearest
        bilinear
        cubic
    End Enum

    ''' <summary>
    ''' Initializes a new instance of the <see cref="OpenTopoDataHttpClient">OpenTopoDataHttpClient</see>.
    ''' </summary>
    Sub New()
        MyBase.New()
        BaseAddress = New Uri(apiBaseAddress)
        DefaultRequestHeaders.UserAgent.TryParseAdd(String.Format("{0}/{1}", My.Application.Info.ProductName, My.Application.Info.Version))
        _requests = New List(Of Request)
        _responses = New Dictionary(Of Integer, Response)
        Downloads = New Dictionary(Of Integer, Task)
        _responseId = -1
        _requestLock = New Object
        _apiLock = New Object
        _stopwatch = New Stopwatch
    End Sub

    ''' <summary>
    ''' Creates a request and assigns a response ID to a request. A new download task is started the first time a response ID is assigned.
    ''' </summary>
    ''' <remarks>
    ''' A response ID is assigned to one or more requests.<br/>
    ''' </remarks>
    ''' <param name="latitude">Latitude (WGS-84 format) of the request.</param>
    ''' <param name="longitude">Longitude (WGS-84 format) of the request.</param>
    ''' <param name="dataset">Digital elevation model dataset of the request.</param>
    ''' <param name="interpolation">Interpolation method of the request.</param>
    ''' <returns>The response ID assigned to the request.</returns>
    Public Function GetResponseId(latitude As Double, longitude As Double, dataset As DigitalElevationModel, interpolation As Interpolation) As Integer
        Dim thisRequest = New Request(dataset, latitude, longitude, interpolation)
        SyncLock _requestLock
            _requests.Add(thisRequest)
            If CountRequests(dataset, interpolation) Mod apiMaxLocationsPerRequest = 1 Then
                _responseId += 1
                thisRequest.ResponseId = _responseId
                Downloads.Add(thisRequest.ResponseId, Task.Run(Sub() Download(thisRequest.ResponseId)))
            Else
                thisRequest.ResponseId = GetMaxResponseId(dataset, interpolation)
            End If
        End SyncLock
        Return thisRequest.ResponseId
    End Function

    ''' <summary>
    ''' Returns the elevation for a given request.
    ''' </summary>
    ''' <param name="responseId">The response ID assigned to the request.</param>
    ''' <param name="latitude">The latitude (WGS-84 format) of the request.</param>
    ''' <param name="longitude">The longitude (WGS-84 format) of the request.</param>
    ''' <returns>Elevation expressed as a double precision number.</returns>
    Public Function GetElevation(responseId, latitude, longitude) As Double
        Dim result As Double?
        If _responses(responseId).Data.Results IsNot Nothing Then
            result = (From response In _responses(responseId).Data.Results
                      Where response.Location.Lat = latitude And response.Location.Lng = longitude
                      Select response.Elevation).FirstOrDefault
        End If
        If result IsNot Nothing Then
            _responses(responseId).Countdown -= 1
            If _responses(responseId).Countdown = 0 Then _responses.Remove(responseId)
            Return result
        Else
            Throw New InvalidOperationException
        End If
    End Function

    ''' <summary>
    ''' Download responses for all requests assigned to a given response ID.
    ''' </summary>
    ''' <param name="responseId"> The response ID of the requests for which to download responses.</param>
    Private Sub Download(responseId As Integer)
        DelayBeforeAcquiringData(responseId) 'becasue we want to bundle as many requests as possible.
        SyncLock _requestLock
            _responses.Add(responseId, New Response(CreateQuestion(responseId)))
            _requests.RemoveAll(Function(item) item.ResponseId = responseId)
        End SyncLock
        AcquireData(responseId)
    End Sub

    ''' <summary>
    ''' Delays until no more requests assigned to a given response ID are added or until the maximum locations per HTTP request is reached.
    ''' </summary>
    ''' <param name="responseId">The response ID of the requests for which to delay.</param>
    Private Sub DelayBeforeAcquiringData(responseId As Integer)
        Dim previousCount As Integer = GetRequests(responseId).Count
        Dim currentCount As Integer
        While currentCount < apiMaxLocationsPerRequest
            Thread.Sleep(10)
            currentCount = GetRequests(responseId).Count
            If currentCount = previousCount Then
                Exit While
            Else
                previousCount = currentCount
            End If
        End While
    End Sub

    ''' <summary>
    ''' Sends a HTTP request to the API and deserializes the response.
    ''' </summary>
    ''' <param name="responseId">The response ID of the requests for which to send the HTTP request.</param>
    Private Sub AcquireData(responseId As Integer)
        Try
            With _responses(responseId)
                .Data = JsonSerializer.Deserialize(Of Root)(GetString(.Question), New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True})
            End With
        Catch ex As Exception
            Throw
        Finally
            Downloads.Remove(responseId)
        End Try
    End Sub

    ''' <summary>
    ''' Returns the HTTP request URI of a given response ID.
    ''' </summary>
    ''' <param name="responseId">The response ID of the requests for which to create the HTTP request URI.</param>
    ''' <returns>HTTP request URI.</returns>
    Private Function CreateQuestion(responseId As Integer) As String
        Dim myRequests As List(Of Request) = GetRequests(responseId)
        Dim result As String = String.Format("{0}?locations=", [Enum].GetName(GetType(DigitalElevationModel), myRequests.Last.Dataset))
        For Each item In myRequests
            result = String.Format("{0}{1}{2},{3}", result, If(result.EndsWith("="), "", "|"), item.Latitude, item.Longitude)
        Next
        Return String.Format("{0}&interpolation={1}", result, [Enum].GetName(GetType(Interpolation), myRequests.Last.Interpolation))
    End Function


    ''' <summary>
    ''' Sends a GET request to the Open Topo Data API and returns the HTTP response body as a string.
    ''' </summary>
    ''' <param name="requestUri">The HTTP request URI to send.</param>
    ''' <returns>The response body as a string.</returns>
    Private Function GetString(requestUri As String) As String
        Dim myHttpResponse As HttpResponseMessage
        SyncLock _apiLock
            For trial As Integer = 0 To 2
                Thread.Sleep(GetWaitingTimeBeforeSend(trial)) 'To avoid 429 error: Too many requests
                _stopwatch.Start()
                myHttpResponse = SendAsync(New HttpRequestMessage(HttpMethod.Get, requestUri), New CancellationToken).GetAwaiter.GetResult
                If myHttpResponse.IsSuccessStatusCode Then Exit For
            Next
        End SyncLock
        If Not myHttpResponse.IsSuccessStatusCode Then Throw New HttpRequestException(String.Format(My.Resources.E002, myHttpResponse.StatusCode, [Enum].GetName(GetType(HttpStatusCode), myHttpResponse.StatusCode)))
        Return myHttpResponse.Content.ReadAsStringAsync.Result
    End Function

    '''' <summary>
    '''' Returns the waiting time before the next HTTP request can be send in order to avoid a 429 error (too many requests).
    '''' </summary>
    '''' <returns>The waiting time in milliseconds.</returns>
    Private Function GetWaitingTimeBeforeSend(trial As Integer) As Integer
        Dim tillNextRequest As Integer
        If _stopwatch.IsRunning Then
            Dim threshold As Integer = Math.Round(1000 / apiMaxThroughput) * (1 + trial) ^ 2
            _stopwatch.Stop()
            Dim ellapsedTime As Integer = _stopwatch.ElapsedMilliseconds
            _stopwatch.Reset()
            tillNextRequest = Math.Max(0, threshold - ellapsedTime)
        End If
        Return tillNextRequest
    End Function

    ''' <summary>
    ''' Returns the number of requests for a given dataset and interpolation method.
    ''' </summary>
    ''' <param name="dataset">The digital elevation model dataset for which to count requests.</param>
    ''' <param name="interpolation">The interpolation method for which to count requests.</param>
    ''' <returns>The number of requests matching the input arguments.</returns>
    Private Function CountRequests(dataset As DigitalElevationModel, interpolation As Interpolation) As Integer
        Dim result = (From request In _requests
                      Where request.Dataset = dataset And request.Interpolation = interpolation
                      Select request).Count
        Return result
    End Function

    ''' <summary>
    ''' Returns the maximum Response ID number for a given dataset and interpolation method.
    ''' </summary>
    ''' <param name="dataset">The digital elevation model dataset for which to get the max response ID.</param>
    ''' <param name="interpolation">The interpolation method for which to get the max response ID.</param>
    ''' <returns>The maximum Response ID number matching the input arguments.</returns>
    Private Function GetMaxResponseId(dataset As DigitalElevationModel, interpolation As Interpolation) As Integer
        Dim results = From request In _requests
                      Where request.Dataset = dataset And request.Interpolation = interpolation
                      Select request.ResponseId
        Return results.Max
    End Function

    ''' <summary>
    ''' Returns the requests assigned to a given reponse ID.
    ''' </summary>
    ''' <param name="responseId">The response ID for which to return the requests.</param>
    ''' <returns>A list of requests matching the input argument.</returns>
    Private Function GetRequests(responseId As Integer) As List(Of Request)
        Dim results = From request In GetCopyOfRequests()
                      Where request.ResponseId = responseId
                      Select request
        Return results.ToList
    End Function

    ''' <summary>
    ''' Copies the entire list of requests into a one-dimensional array and returns it.
    ''' </summary>
    ''' <returns>A one-dimensioanl array of requests.</returns>
    Private Function GetCopyOfRequests() As Request()
        Dim results() As Request
        SyncLock _requestLock
            ReDim results(_requests.Count - 1)
            _requests.CopyTo(results)
        End SyncLock
        Return results
    End Function
End Class