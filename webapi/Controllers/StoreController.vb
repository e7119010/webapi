Imports System.Net
Imports System.Web.Http
Imports System.Net.Http
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class StoreController
    Inherits ApiController
    Private jsonstring As String = String.Empty
    Dim RC, RM, user_access_token, cards, value As String
    Dim _trans As New Trans

    ' POST api/<controller>
    <HttpPost>
    Public Function Parking_no() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        Dim code As String = String.Empty
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        code = jsonsource("code").ToString()

        jsonstring = "{""RC"":1,""RM"":""成功"",""results"":{""parking_no"":{""motocycle"":"""",""car"":""""}}}"
        Return _trans.TransToJson(jsonstring)
    End Function

    ' PUT api/<controller>/5
    Public Sub PutValue(ByVal id As Integer, <FromBody()> ByVal value As String)

    End Sub

    ' DELETE api/<controller>/5
    Public Sub DeleteValue(ByVal id As Integer)

    End Sub
End Class
