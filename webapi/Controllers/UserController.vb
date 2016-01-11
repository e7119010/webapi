Imports System.Net
Imports System.Web.Http
Imports Microsoft.VisualBasic
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net.Http

Public Class UserController
    Inherits ApiController
    Private jsonstring As String = String.Empty
    Dim RC, RM, user_access_token, cards, value As String

    
    Public Function TransToJson(ByVal _json As String) As HttpResponseMessage
        Dim rep As New HttpResponseMessage() With {.Content = New StringContent(_json)}
        rep.Content.Headers.ContentType = New Headers.MediaTypeHeaderValue("application/json")
        Return rep
    End Function
   
    <HttpPost>
    Public Function Login() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        Dim login_type, channel As Integer
        Dim account, password As String
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        login_type = CInt(jsonsource("login_type"))
        channel = CInt(jsonsource("channel"))
        account = jsonsource("account").ToString()
        password = jsonsource("password").ToString()
        Select Case login_type
            Case 1 '身分證字號
                Dim tt As New Ado()
                If tt.IsCardMember(account) Then
                    If tt.CheckAPPStatus(account) Then
                        If tt.CheckPasswordbyPID(account, password) Then
                            Dim _token As String = tt.GetTokenByPID(account)
                            Dim _carddata As String = tt.GetCardData(account)
                            jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":{""user_access_token"":""" + _
                                _token + """,""cards"":" + _carddata + "}}"
                        Else
                            jsonstring = "{""rcrm"":{""RC"":-401.0001,""RM"":""密碼錯誤""},""results"":null}"
                        End If
                    Else
                        jsonstring = "{""rcrm"":{""RC"":-401.0003,""RM"":""APP尚未開通""},""results"":null}"
                    End If
                Else
                    jsonstring = "{""rcrm"":{""RC"":-404.0001,""RM"":""會員資料不存在""},""results"":null}"
                End If
            Case 2 '會員卡號
                Dim tt As New Ado()
                If tt.CheckPasswordByCard(account, password) Then
                    Dim _cardno As String = String.Empty
                    Dim _token As String = String.Empty
                    Dim _carddata As String = String.Empty
                    Dim _pid As String = String.Empty
                    If account.Split("|")(0).ToString() = "A101" Then '實體會員卡
                        _cardno = account.Split("|")(3).ToString()
                        _pid = tt.GetPIDByCard(_cardno)
                        _token = tt.GetTokenByPID(_pid, _cardno)
                        _carddata = tt.GetDataByCard(_pid, _cardno)
                        jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":{""user_access_token"":""" + _
                            _token + """,""cards"":" + _carddata + "}}"
                    Else '手機條碼
                        _token = account.Split("|")(1).ToString()
                        _pid = tt.GetPIDByToken(_token)
                        _carddata = tt.GetCardData(_pid)
                        jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":{""user_access_token"":""" + _
                            _token + """,""cards"":" + _carddata + "}}"
                    End If
                    
                Else
                    jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""},""results"":null}"
                End If
        End Select
        Return Me.TransToJson(jsonstring)
    End Function

    <HttpPost>
    Public Function Logout() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token, device_uuid As String
        Dim channel As Integer
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token")
        device_uuid = jsonsource1("device_uuid")
        channel = CInt(jsonsource("channel"))
        Dim tt As New Ado()
        If tt.IsTokenLive(user_access_token) Then
            If tt.VoidToken(user_access_token) Then
                jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""}}"
            Else
                jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""}}"
            End If
        Else
            jsonstring = "{""rcrm"":{""RC"":0,""RM"":""""},""results"":null}"
        End If
        Return Me.TransToJson(jsonstring)
    End Function

    Public Function Verify_identity() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        Dim user_access_token, device_uuid, password As String
        Dim channel As Integer
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token")
        device_uuid = jsonsource1("device_uuid")
        password = jsonsource("password")
        channel = CInt(jsonsource("channel"))
        Dim tt As New Ado()
        If tt.IsTokenLive(user_access_token) Then
            If tt.CheckPasswordbyToken(user_access_token, password) Then
                jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""}}"
            Else
                jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""}}"
            End If
        Else
            jsonstring = "{""rcrm"":{""RC"":0,""RM"":""""},""results"":null}"
        End If
        Return Me.TransToJson(jsonstring)
    End Function

    ''' <summary>
    ''' 會員開通
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Bound() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        Dim account As String
        Dim channel As Integer
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        account = jsonsource("account").ToString()
        channel = CInt(jsonsource("channel"))
        Dim tt As New Ado()
        If tt.OpenAppStatus(account) Then
            Dim _token As String = String.Empty
            _token = tt.GetTokenByPID(account)
            Return Me.TransToJson("{""rcrm"":{""RC"":-401.0006,""RM"":""成功""},""results"":{""user_access_token"":""" + _token + """}}")
        Else
            Return Me.TransToJson("{""rcrm"":{""RC"":-1,""RM"":""失敗""}}")
        End If
    End Function

    ''' <summary>
    ''' 會員取消開通(測試用)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Unbound() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        Dim account As String
        Dim channel As Integer
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        account = jsonsource("account").ToString()
        channel = CInt(jsonsource("channel"))
        Dim tt As New Ado()
        If tt.CloseAppStatus(account) Then
            Return Me.TransToJson("{""rcrm"":{""RC"":1,""RM"":""成功""}}")
        Else
            Return Me.TransToJson("{""rcrm"":{""RC"":-1,""RM"":""失敗""}}")
        End If
    End Function

    Public Function Update_Password() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim old_password As String = String.Empty
        Dim new_password As String = String.Empty
        Dim channel As Integer = 1
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token")
        device_uuid = jsonsource1("device_uuid").ToString()
        old_password = IIf(String.IsNullOrEmpty(jsonsource("old_password")), "", jsonsource("old_password"))
        new_password = jsonsource("new_password")
        channel = CInt(jsonsource("channel"))
        If new_password <> String.Empty Then
            Dim tt As New Ado()
            If tt.IsTokenLive(user_access_token) Then
                If tt.ChangePassword(user_access_token, old_password, new_password) Then
                    jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""}}"
                Else
                    jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""}}"
                End If
            Else
                jsonstring = "{""rcrm"":{""RC"":0,""RM"":""""},""results"":null}"
            End If
        Else
            jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""密碼不可為空白""}}"
        End If
        Return Me.TransToJson(jsonstring)
    End Function

    ''' <summary>
    ''' 忘記密碼
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Retrieve_Password() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        Dim identifier As String = String.Empty
        Dim birth As String = String.Empty
        Dim mobile As String = String.Empty
        Dim channel As Integer = 1
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        identifier = jsonsource("identifier").ToString()
        birth = jsonsource("birth").ToString.Replace("/", "-")
        mobile = jsonsource("mobile").ToString()
        channel = CInt(jsonsource("channel"))
        Dim tt As New Ado()
        If tt.VerifyMember(identifier, birth, mobile) Then
            Dim _token As String = String.Empty
            _token = tt.GetTokenByPID(identifier)
            jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":{""user_access_token"":""" + _token + """}}"
        Else
            jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""}}"
        End If
        Return Me.TransToJson(jsonstring)
    End Function

    ''' <summary>
    ''' 查詢積點紀錄
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Query_Bonus_Summary() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim channel As Integer = 1
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        channel = CInt(jsonsource("channel"))
        Dim tt As New Ado()
        If tt.IsTokenLive(user_access_token) Then
            Dim _bonus As String = String.Empty
            _bonus = tt.GetBonusByYear(user_access_token)
            If _bonus <> "" Then
                jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":{""bonus_summary"":" + _bonus + "}}"
            Else
                jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""},""results"":null}"
            End If
        Else
            jsonstring = "{""rcrm"":{""RC"":0,""RM"":""""},""results"":null}"
        End If
        Return Me.TransToJson(jsonstring)
    End Function

    Public Function Record_bonus() As HttpResponseMessage '補登發票
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim receipt_number As String = String.Empty
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        receipt_number = jsonsource("receipt_number").ToString()
        Dim tt As New Ado()
        Dim _status As String = String.Empty
        _status = tt.ReInputBonus(receipt_number, user_access_token)
        Select Case _status
            Case "ALIVE" '資料已存在
                Return Me.TransToJson("{""rcrm"":{""RC"":-1,""RM"":""資料已存在""}}")
            Case "SUCCESS" '補登成功
                Return Me.TransToJson("{""rcrm"":{""RC"":1,""RM"":""補登成功""}}")
            Case "FALSE" '無此發票資料
                Return Me.TransToJson("{""rcrm"":{""RC"":-1,""RM"":""無此發票資料""}}")
            Case "ERROR" '輸入資料驗證錯誤
                Return Me.TransToJson("{""rcrm"":{""RC"":-1,""RM"":""輸入資料驗證錯誤""}}")
        End Select

    End Function

    Public Function UnRecord_bonus() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim receipt_number As String = String.Empty
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        receipt_number = jsonsource("receipt_number").ToString()
        Dim tt As New Ado()
        Dim _status As String = String.Empty
        _status = tt.UNREINPUTBOUNS(receipt_number, user_access_token)
        Select Case _status
            Case "SUCCESS" '補登成功
                Return Me.TransToJson("{""rcrm"":{""RC"":1,""RM"":""作廢成功""}}")
            Case "FALSE" '無此發票資料
                Return Me.TransToJson("{""rcrm"":{""RC"":-1,""RM"":""作廢失敗""}}")
        End Select

    End Function
    Public Function Process_Point() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim type As String = String.Empty
        Dim point As Integer = 0
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        type = jsonsource("type").ToString.ToUpper()
        point = CInt(jsonsource("point"))
        Dim tt As New Ado()
        Select Case type.ToUpper()
            Case "A" '增加
                Dim _status As String = String.Empty
                _status = tt.MinusMEMBERBOUNS(point, user_access_token)
                Select Case _status
                    Case "SUCCESS"
                        Return Me.TransToJson("{""rcrm"":{""RC"":1,""RM"":""成功""}}")
                    Case "FALSE"
                        Return Me.TransToJson("{""rcrm"":{""RC"":-403.1702,""RM"":""點數異動失敗，無法增加點數""}}")
                End Select
            Case "D" '扣除
                If point <= tt.GetBonusByToken(user_access_token) Then
                    Dim _status As String = String.Empty
                    _status = tt.MinusMEMBERBOUNS(point * (-1), user_access_token)
                    Select Case _status
                        Case "SUCCESS"
                            Return Me.TransToJson("{""rcrm"":{""RC"":1,""RM"":""成功""}}")
                        Case "FALSE"
                            Return Me.TransToJson("{""rcrm"":{""RC"":-403.1703,""RM"":""點數異動失敗，無法扣除點數""}}")
                    End Select
                Else
                    Return Me.TransToJson("{""rcrm"":{""RC"":-403.1701,""RM"":""點數異動檢查失敗，會員點數不足扣除""}}")
                End If

            Case "C" '查詢
                If point <= tt.GetBonusByToken(user_access_token) Then
                    Return Me.TransToJson("{""rcrm"":{""RC"":1,""RM"":""成功""}}")
                Else
                    Return Me.TransToJson("{""rcrm"":{""RC"":-403.1701,""RM"":""點數異動檢查失敗，會員點數不足扣除""}}")
                End If
        End Select

        
    End Function
    
    Public Function Process_RulePoint() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim rule As String = String.Empty
        Dim type As String = String.Empty
        Dim point As Integer = 0
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        'Dim jsonsource2 As Object = JsonConvert.DeserializeObject(jsonsource("bonus_item").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        type = jsonsource("type").ToString()
        point = CInt(jsonsource("point"))

        Dim tt As New Ado()
        Select Case type.ToUpper()
            Case "D" '扣除
                If point <= tt.GetBonusByToken(user_access_token) Then
                    Dim _status As String = String.Empty
                    _status = tt.MinusMEMBERBOUNS(point * (-1), user_access_token, JsonConvert.DeserializeObject(jsonsource("bonus_item").ToString()).ToString())
                    Select Case _status
                        Case "ALIVE", "NONE"
                            Return Me.TransToJson("{""rcrm"":{""RC"":-403.1703,""RM"":""點數異動失敗，無法扣除點數""}}")
                        Case Else
                            Return Me.TransToJson("{""rcrm"":{""RC"":1,""RM"":""成功""},""coponname"":""" + _status.Split("|")(3).ToString() +
                                                  """,""COPONNO"":""" + _status.Split("|")(2).ToString() + """,""QRCODE"":""" + _status.Replace(_status.Split("|")(3).ToString(), "***") +
                                                  """,""remark"":""1.此優惠券僅限大魯閣新時代使用。|2.此優惠券不得重複列印使用。" +
                                                  "|3.優惠券詳情及其他注意事項依現場活動辦法。|4.大魯閣新時代保有最終修訂之權利。""}")
                    End Select
                Else

                End If

            Case "C" '查詢
                If point <= tt.GetBonusByToken(user_access_token) Then
                    Return Me.TransToJson("{""rcrm"":{""RC"":1,""RM"":""成功""}}")
                Else
                    Return Me.TransToJson("{""rcrm"":{""RC"":-403.1701,""RM"":""點數異動檢查失敗，會員點數不足扣除""}}")
                End If
        End Select


    End Function
    ''' <summary>
    ''' 查詢積點明細
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Query_Bonus_History() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim query_start_date As String = String.Empty
        Dim query_end_date As String = String.Empty
        Dim bonus_type As String = String.Empty
        Dim channel As Integer = 1
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        query_start_date = jsonsource("query_start_date").ToString()
        query_end_date = jsonsource("query_end_date").ToString()
        bonus_type = jsonsource("bonus_type").ToString()
        channel = CInt(jsonsource("channel"))
        Dim tt As New Ado()
        Dim _tempstring As String = String.Empty
        Dim _pid As String = String.Empty
        Select Case channel
            Case 1 'app
                If tt.IsTokenLive(user_access_token) Then
                    _tempstring = tt.GetBonusList(user_access_token, query_start_date.Replace("/", "-"), query_end_date.Replace("/", "-"))
                    If _tempstring = "error" Then
                        jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""},""results"":null}"
                    Else
                        jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":{""bonus_history"":" + _tempstring + "}}"
                    End If
                Else
                    jsonstring = "{""rcrm"":{""RC"":0,""RM"":""""},""results"":null}"
                End If
            Case 2 'kiosk
                If tt.IsTokenLive(user_access_token) Then
                    _tempstring = tt.GetBonusList_Kiosk(user_access_token, query_start_date.Replace("/", "-"), query_end_date.Replace("/", "-"))
                    If _tempstring = "error" Then
                        jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""},""results"":null}"
                    Else
                        jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":{""bonus_history"":" + _tempstring + "}}"
                    End If
                Else
                    jsonstring = "{""rcrm"":{""RC"":0,""RM"":""""},""results"":null}"
                End If
        End Select

        Return Me.TransToJson(jsonstring)
    End Function

    Public Function Query_consumption_history() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim query_start_date As String = String.Empty
        Dim query_end_date As String = String.Empty
        Dim query_store As String = String.Empty
        Dim channel As Integer = 1
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        query_start_date = jsonsource("query_start_date").ToString()
        query_end_date = jsonsource("query_end_date").ToString()
        query_store = jsonsource("query_store").ToString()
        channel = CInt(jsonsource("channel"))
        Dim _tempstring As String = String.Empty
        Dim _pid As String = String.Empty
        Dim tt As New Ado()
        If tt.IsTokenLive(user_access_token) Then

            _tempstring = tt.GetInvoiceMaster(user_access_token, query_start_date.Replace("/", "-"), query_end_date.Replace("/", "-"))
            If _tempstring = "error" Then
                jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""},""results"":null}"
            Else
                jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":{""consumption_history"":" + _tempstring + "}}"
            End If
        Else
            jsonstring = "{""rcrm"":{""RC"":0,""RM"":""""},""results"":null}"
        End If
        Return Me.TransToJson(jsonstring)
    End Function

    ''' <summary>
    ''' 查詢消費明細
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Query_consumption_history_detail() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim transaction_id As String = String.Empty
        Dim transaction_time As String = String.Empty
        Dim query_store As String = String.Empty
        Dim channel As Integer = 1
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        transaction_id = jsonsource("transaction_id").ToString()
        transaction_time = jsonsource("transaction_time").ToString()
        query_store = jsonsource("query_store").ToString()
        channel = CInt(jsonsource("channel"))
        Dim main() As String
        Dim itemlist As String = String.Empty
        Dim paylist As String = "null"
        Dim _pid As String = String.Empty
        Dim tt As New Ado()
        If tt.IsTokenLive(user_access_token) Then
            _pid = tt.GetPIDByToken(user_access_token)
            main = tt.GetInvoiceMaster(_pid, transaction_time.Substring(0, 10).Replace("/", "-"), query_store, transaction_id)
            itemlist = tt.GetInvoiceDetail(_pid, transaction_time.Substring(0, 10).Replace("/", "-"), query_store, transaction_id)
            'paylist = tt.GetPayDetail(_pid, transaction_time.Substring(0, 10).Replace("/", "-"), query_store, transaction_id)
            If main(0) = "error" OrElse main(0) = "" Then
                jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""},""results"":null}"
            Else
                				
                jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":{""consumption_detail"":{" + main(0) + """consumption_item"":[" + itemlist + "]," +
                    """consumption_price_total"":""" + main(2) + """,""option_consumption_item"":" + paylist + ",""price_total"":""" + main(1) + """}}}"
            End If
        Else
            jsonstring = "{""rcrm"":{""RC"":0,""RM"":""""},""results"":null}"
        End If
        Return Me.TransToJson(jsonstring)
    End Function
    Public Function VipInfo() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        Dim tt As New Ado()
        Dim _tempstring As String = String.Empty
        If tt.IsTokenLive(user_access_token) Then
            _tempstring = tt.GetMemberData(user_access_token)
            If _tempstring <> "" Then
                jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":{""user_info"":" + _tempstring + "}}"
            Else
                jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""}}"
            End If
        Else
            jsonstring = "{""rcrm"":{""RC"":0,""RM"":""失敗""}}"
        End If
        Return Me.TransToJson(jsonstring)
    End Function
    Public Function Update_info() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        Dim tt As New Ado()
        Dim _tempstring As String = String.Empty

        _tempstring = tt.UpdateMemberData(user_access_token, jsonsource("user_info").ToString())
        If _tempstring = "SUCCESS" Then
            jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""}}"
        ElseIf _tempstring = "FALSE" Then
            jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""資料更新失敗失敗""}}"
        Else
            jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""必要欄位不得空白""}}"
        End If
        Return Me.TransToJson(jsonstring)
    End Function
    Public Function Register() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim tt As New Ado()
        Dim _tempstring As String = String.Empty
        _tempstring = tt.AddMemberdata(jsonsource("user_info").ToString)
        '增加身分檢查是否已申辦
        Select Case _tempstring
            Case "success"
                jsonstring = "{""rcrm"":{""RC"":1,""RM"":""您的申請資料已成功送出""}}"
            Case "space"
                jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗-必填欄位未填寫""}}"
            Case Else
                jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""}}"
        End Select
        Return Me.TransToJson(jsonstring)
    End Function

    Public Function Bonus_item() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        Dim tt As New Ado()
        Dim _tempstring As String = String.Empty
        _tempstring = tt.GetBonusitemlist(user_access_token)
        If _tempstring <> "" Then
            jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":{""bonus_item"":" + _tempstring + "}}"
        Else
            jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""}}"
        End If
        Return Me.TransToJson(jsonstring)
    End Function

    Public Function Token_Update() As HttpResponseMessage
        value = Request.Content.ReadAsStringAsync().Result
        jsonstring = ""
        Dim user_access_token As String = String.Empty
        Dim device_uuid As String = String.Empty
        Dim push_token As String = String.Empty
        Dim jsonsource As Object = JsonConvert.DeserializeObject(value)
        Dim jsonsource1 As Object = JsonConvert.DeserializeObject(jsonsource("credential").ToString)
        user_access_token = jsonsource1("user_access_token").ToString()
        device_uuid = jsonsource1("device_uuid").ToString()
        push_token = jsonsource("push_token").ToString()
        Dim tt As New Ado()
        Dim _tempstring As String = String.Empty
        _tempstring = tt.SetPushToken(user_access_token, device_uuid, push_token)
        If _tempstring = "SUCCESS" Then
            jsonstring = "{""rcrm"":{""RC"":1,""RM"":""成功""},""results"":null}"
        Else
            jsonstring = "{""rcrm"":{""RC"":-1,""RM"":""失敗""},""results"":null}"
        End If
        Return Me.TransToJson(jsonstring)

    End Function
End Class

