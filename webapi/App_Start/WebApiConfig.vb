Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web.Http

Public Module WebApiConfig
    Public Sub Register(ByVal config As HttpConfiguration)
        ' Web API configuration and services

        ' Web API routes
        config.MapHttpAttributeRoutes()

        'config.Routes.MapHttpRoute(
        '    name:="v1-movie",
        '    routeTemplate:="appapi/v1/movie/{action}",
        '    defaults:=New With {.action = RouteParameter.Optional, .controller = "MovieV1"}
        ')


        config.Routes.MapHttpRoute(
            name:="appapi",
            routeTemplate:="appapi/{controller}/{action}",
            defaults:=New With {.action = RouteParameter.Optional}
        )
    End Sub
End Module
