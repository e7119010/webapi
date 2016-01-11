Imports System.Net
Imports System.Net.Http

Public Class Trans

    Public Function TransToJson(ByVal _json As String) As HttpResponseMessage
        Dim rep As New HttpResponseMessage() With {.Content = New StringContent(_json)}
        rep.Content.Headers.ContentType = New Headers.MediaTypeHeaderValue("application/json")
        Return rep
    End Function
End Class
