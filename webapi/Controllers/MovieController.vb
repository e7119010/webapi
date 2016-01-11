Imports System.Net
Imports System.Web.Http
Imports System.Net.Http
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Xml
<RoutePrefix("appapi/v1/movie")>
Public Class MovieController
    Inherits ApiController
    Private jsonstring As String = String.Empty
    Dim RC, RM, user_access_token, cards, value As String
    Dim _trans As New Trans


    <HttpPost, Route("ambassadormovie")>
    Public Function ambassadormovie() As HttpResponseMessage
        Dim ftprsp As FtpWebRequest = FtpWebRequest.Create("ftp://ftpmovie.ambassador.com.tw/XmlSourceMobile/MoviesInfo20110303.xml")
        ftprsp.Credentials = New Net.NetworkCredential("FTPBusiness_Taroko", "taroko9700")
        ftprsp.UseBinary = True
        ftprsp.Method = WebRequestMethods.Ftp.DownloadFile

        Dim _xml As New XmlDocument
        _xml.Load(ftprsp.GetResponse().GetResponseStream())
        jsonstring = JsonConvert.SerializeXmlNode(_xml)
        Return _trans.TransToJson(jsonstring)
    End Function

    <HttpPost, Route("ambassadorshowtime")>
    Public Function ambassadorshowtime() As HttpResponseMessage
        Dim ftprsp As FtpWebRequest = FtpWebRequest.Create("ftp://ftpmovie.ambassador.com.tw/XmlSourceMobile/ShowtimesInfo20130801.xml")
        ftprsp.Credentials = New Net.NetworkCredential("FTPBusiness_Taroko", "taroko9700")
        ftprsp.UseBinary = True
        ftprsp.Method = WebRequestMethods.Ftp.DownloadFile

        Dim _xml As New XmlDocument
        _xml.Load(ftprsp.GetResponse().GetResponseStream())
        jsonstring = JsonConvert.SerializeXmlNode(_xml)
        Return _trans.TransToJson(jsonstring)
    End Function
    <HttpPost, Route("ambassadortrailer")>
    Public Function ambassadortrailer() As HttpResponseMessage
        Dim ftprsp As FtpWebRequest = FtpWebRequest.Create("ftp://ftpmovie.ambassador.com.tw/XmlSourceMobile/TrailersInfo20110303.xml")
        ftprsp.Credentials = New Net.NetworkCredential("FTPBusiness_Taroko", "taroko9700")
        ftprsp.UseBinary = True
        ftprsp.Method = WebRequestMethods.Ftp.DownloadFile

        Dim _xml As New XmlDocument
        _xml.Load(ftprsp.GetResponse().GetResponseStream())
        jsonstring = JsonConvert.SerializeXmlNode(_xml)
        Return _trans.TransToJson(jsonstring)
    End Function
    <HttpPost, Route("ambassadormedia")>
    Public Function ambassadormedia() As HttpResponseMessage
        Dim ftprsp As FtpWebRequest = FtpWebRequest.Create("ftp://ftpmovie.ambassador.com.tw/XmlSourceMobile/TrailersMediaLink20130801.xml")
        ftprsp.Credentials = New Net.NetworkCredential("FTPBusiness_Taroko", "taroko9700")
        ftprsp.UseBinary = True
        ftprsp.Method = WebRequestMethods.Ftp.DownloadFile

        Dim _xml As New XmlDocument
        _xml.Load(ftprsp.GetResponse().GetResponseStream())
        jsonstring = JsonConvert.SerializeXmlNode(_xml)
        Return _trans.TransToJson(jsonstring)
    End Function

End Class

<RoutePrefix("appapi/v3/movie")>
Public Class MovieV2Controller
    Inherits ApiController
    Private jsonstring As String = String.Empty
    Dim RC, RM, user_access_token, cards, value As String
    Dim _trans As New Trans


    ' POST api/<controller>
    <HttpPost, Route("search")>
    Public Function Search() As HttpResponseMessage
        Dim _xml As New XmlDocument
        _xml.Load("http://www.vscinemas.com.tw/rss/movies/07TCMS.xml")
        jsonstring = JsonConvert.SerializeXmlNode(_xml)
        Return _trans.TransToJson(jsonstring)
    End Function

   
End Class
