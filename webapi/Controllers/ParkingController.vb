Imports System.Net
Imports System.Web.Http
Imports Microsoft.VisualBasic
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net.Http
Public Class ParkingController
    Inherits ApiController
    Private jsonstring As String = String.Empty
    Dim RC, RM, user_access_token, cards, value As String

    
    <HttpPost>
    Public Function MemberCheck() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result

    End Function




    ' GET api/<controller>
    Public Function GetValues() As IEnumerable(Of String)
        Return New String() {"value1", "value2"}
    End Function

    ' GET api/<controller>/5
    Public Function GetValue(ByVal id As Integer) As String
        Return "value"
    End Function

    ' POST api/<controller>
    Public Sub PostValue(<FromBody()> ByVal value As String)

    End Sub

    ' PUT api/<controller>/5
    Public Sub PutValue(ByVal id As Integer, <FromBody()> ByVal value As String)

    End Sub

    ' DELETE api/<controller>/5
    Public Sub DeleteValue(ByVal id As Integer)

    End Sub
End Class
