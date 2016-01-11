Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


Public Class Ado
    Private _conn, _txgposconn As SqlConnection
    Private _jsonstring As String

    Public Sub New()
        _conn = New SqlConnection("Data Source=132.147.168.10;Initial Catalog=CRM;User ID=sa;Password=Admin1793;")
        _txgposconn = New SqlConnection("Data Source=132.147.135.207;Initial Catalog=pos_svr;User ID=sa;Password=Admin1793;")
        '_conn = New SqlConnection("Data Source=132.147.135.155;Initial Catalog=CRM;User ID=sa;Password=Admin1793;")
        _jsonstring = String.Empty
    End Sub
    Public Function GetCardNobyPID(ByVal _PID As String) As String
        If _conn.State = ConnectionState.Closed Then _conn.Open()
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = String.Format("SELECT CARDNO FROM CRM_CARDDATA WHERE PID='{0}' AND VOID_CODE=0 AND CANCEL_STATUS=0", _PID)
                If _cmd.ExecuteScalar = Nothing Then
                    Return ""
                Else
                    Return _cmd.ExecuteScalar.ToString()
                End If
            End Using
        Catch ex As Exception
            Return "ERROE"
        Finally
            _conn.Close()
        End Try
    End Function
    Public Function GetCardData(ByVal _PID As String) As String
        Dim _dt As New DataTable
        _jsonstring = ""
        Using sqlapt As New SqlDataAdapter("", _conn)
            sqlapt.SelectCommand.CommandText = _
                "SELECT top 1 (CARDNO)card_number,(CATEGORYNO)card_kind,('會員卡')card_name,(NULL)card_description," +
                "(CASE WHEN CANCEL_STATUS=1 THEN '已作廢' ELSE (CASE WHEN OPEN_STATUS=1 THEN '已開卡' ELSE '未開卡' END) END)card_status," +
                "CONVERT(VARCHAR(10),SYSTEMDATE,111)card_start_date,(CASE WHEN CANCEL_STATUS =1 THEN CONVERT(VARCHAR(10),LAST_DATE,111) ELSE NULL END)card_end_date," +
                "CONVERT(VARCHAR(10),SYSTEMDATE,111)card_apply_date,(NULL)card_image," +
                "(SELECT ('A102|'+TOKEN+'|'+CONVERT(VARCHAR,TIME_BEGIN,120)+'|'+CONVERT(VARCHAR,TIME_END,120)) FROM CRM_CUSTOMERTOKEN WHERE PID=A.PID " +
                String.Format("AND TIME_END>=GETDATE() AND VOID_CODE=0 )card_note FROM CRM_CARDDATA A WHERE PID='{0}' AND VOID_CODE=0 ORDER BY SYSTEMDATE DESC", _PID)
            sqlapt.Fill(_dt)
            If _dt.Rows.Count > 0 Then
                _jsonstring = JsonConvert.SerializeObject(_dt).ToString()

            Else
                _jsonstring = "{""rcrm"":{""RC"":-404.0001,""RM"":""會員資料不存在""},""results"":null}"
            End If
        End Using
        Return _jsonstring
    End Function
    Public Function IsCardMember(ByVal _PID As String) As Boolean
        Dim status As Boolean
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = String.Format("SELECT CARDNO FROM CRM_CARDDATA WHERE PID='{0}' AND VOID_CODE=0 AND CANCEL_STATUS=0", _PID)
                If _cmd.ExecuteScalar = Nothing Then
                    status = False
                Else
                    status = True
                End If
            End Using
        Catch ex As Exception
            status = False
        Finally
            _conn.Close()
        End Try

        Return status
    End Function
    Public Function CheckAPPStatus(ByVal _PID As String) As Boolean
        Dim app_status As Boolean
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = String.Format("SELECT APP_STATUS FROM CRM_CUSTOMERPLUS WHERE PID='{0}' AND VOID_CODE=0", _PID)
                If _cmd.ExecuteScalar = Nothing Then
                    app_status = False
                Else
                    app_status = CBool(_cmd.ExecuteScalar)
                End If
            End Using
        Catch ex As Exception
            app_status = False
        Finally
            _conn.Close()
        End Try

        Return app_status
    End Function
    Public Function VoidToken(ByVal _token As String) As Boolean
        Dim status As Boolean = False
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = String.Format("UPDATE CRM_CUSTOMERTOKEN SET VOID_CODE=1 WHERE TOKEN='{0}'", _token)
                If _cmd.ExecuteNonQuery = 1 Then status = True
            End Using
        Catch ex As Exception

        Finally
            _conn.Close()
        End Try
        Return status
    End Function
    Public Function CheckPasswordbyPID(ByVal _pid As String, ByVal _PWD As String) As Boolean
        Dim _dt As New DataTable
        Dim status As Boolean = False
        Try
            Using sqlapt As New SqlDataAdapter("", _conn)
                sqlapt.SelectCommand.CommandText = String.Format("SELECT * FROM CRM_CUSTOMERPLUS WHERE PID='{0}' AND VOID_CODE=0 AND PASS_WORD='{1}' ", _pid, _PWD)
                sqlapt.Fill(_dt)
                If _dt.Rows.Count > 0 Then
                    status = True
                End If
            End Using
        Catch ex As Exception

        Finally
            _conn.Close()
        End Try
        Return status
    End Function

    Public Function CheckPasswordbyToken(ByVal _TOKEN As String, ByVal _PWD As String) As Boolean
        Dim _dt As New DataTable
        Dim _pid As String = String.Empty
        _pid = Me.GetPIDByToken(_TOKEN)
        Dim status As Boolean = Me.CheckPasswordbyPID(_pid, _PWD)
        Return status
    End Function
    Public Function CheckPasswordByCard(ByVal _CARD As String, ByVal _PWD As String) As Boolean
        Dim status As Boolean = False
        If _conn.State = ConnectionState.Closed Then _conn.Open()
        Try
            Using COMM As New SqlCommand("", _conn)
                COMM.CommandText = String.Format("EXEC CRM.DBO.CRM_MEMBERCHECKLIVE 'LOGIN',0,'{0}','{1}'", _CARD, _PWD)
                If COMM.ExecuteScalar.ToString() = "IN" Then '確認
                    status = True
                Else '失敗
                    status = False
                End If
            End Using
        Catch ex As Exception
            status = False
        Finally
            _conn.Close()
        End Try
        Return status
    End Function
    Public Function GetDataByCard(ByVal _PID As String, ByVal _CARD As String) As String
        Dim _dt As New DataTable
        _jsonstring = ""
        Using sqlapt As New SqlDataAdapter("", _conn)
            sqlapt.SelectCommand.CommandText = _
                "SELECT (CARDNO)card_number,(CATEGORYNO)card_kind,('會員卡_曼麗卡')card_name,(NULL)card_description," +
                "(CASE WHEN CANCEL_STATUS=1 THEN '已作廢' ELSE '' END)card_status," +
                "CONVERT(VARCHAR(10),SYSTEMDATE,111)card_start_date,(CASE WHEN CANCEL_STATUS =1 THEN CONVERT(VARCHAR(10),LAST_DATE,111) ELSE NULL END)card_end_date," +
                "CONVERT(VARCHAR(10),SYSTEMDATE,111)card_apply_date,(NULL)card_image," +
                String.Format("(SELECT ('A102|'+TOKEN+'|'+CONVERT(VARCHAR,TIME_BEGIN,120)+'|'+CONVERT(VARCHAR,TIME_END,120)) FROM CRM_CUSTOMERTOKEN WHERE PID=A.PID AND TIME_END>=GETDATE() AND VOID_CODE=0 )card_note FROM CRM_CARDDATA A WHERE PID='{0}' AND CARDNO='{1}' AND VOID_CODE=0 ", _PID, _CARD)
            sqlapt.Fill(_dt)
            If _dt.Rows.Count > 0 Then
                _jsonstring = JsonConvert.SerializeObject(_dt).ToString()
            Else
                _jsonstring = "{""rcrm"":{""RC"":-404.0001,""RM"":""會員資料不存在""},""results"":null}"
            End If
        End Using
        Return _jsonstring
    End Function

    Public Function VerifyMember(ByVal _PID As String, ByVal _BIRTH As String, ByVal _MOBILE As String) As Boolean
        Dim app_status As Boolean
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = String.Format("SELECT 1 FROM CRM_CUSTOMER WHERE PID='{0}' AND BIRTHDAY='{1}' AND LINK_MTEL='{2}' AND VOID_CODE=0", _PID, _BIRTH, _MOBILE)
                If _cmd.ExecuteScalar = Nothing Then
                    app_status = False
                Else
                    app_status = CBool(_cmd.ExecuteScalar)
                End If
            End Using
        Catch ex As Exception
            app_status = False
        Finally
            _conn.Close()
        End Try

        Return app_status
    End Function
    Public Function ChangePassword(ByVal _TOKEN As String, ByVal _OLDPASSWORD As String, ByVal _NEWPASSWORD As String) As Boolean
        Dim _pid As String = Me.GetPIDByToken(_TOKEN)
        Dim status As Boolean = False
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                If _OLDPASSWORD = "" Then
                    _cmd.CommandText = String.Format("UPDATE CRM_CUSTOMERPLUS SET PASS_WORD='{0}' WHERE PID='{1}'", _NEWPASSWORD, _pid)
                Else
                    If Me.CheckPasswordbyPID(_pid, _OLDPASSWORD) Then
                        If _conn.State = ConnectionState.Closed Then _conn.Open()
                        _cmd.CommandText = String.Format("UPDATE CRM_CUSTOMERPLUS SET PASS_WORD='{0}' WHERE PID='{1}'", _NEWPASSWORD, _pid)
                    End If
                End If
                If _cmd.ExecuteNonQuery >= 1 Then status = True
            End Using
        Catch ex As Exception
            status = False
        Finally
            _conn.Close()
        End Try
        Return status
    End Function
    Public Function OpenAppStatus(ByVal _PID As String) As Boolean
        Dim status As Boolean = False
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = String.Format("UPDATE CRM_CUSTOMERPLUS SET APP_STATUS=1 WHERE PID='{0}' AND APP_STATUS=0 ", _PID)
                If _cmd.ExecuteNonQuery >= 1 Then status = True
            End Using
        Catch ex As Exception

        Finally
            _conn.Close()
        End Try
        Return status
    End Function
    Public Function CloseAppStatus(ByVal _PID As String) As Boolean
        Dim status As Boolean = False
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = String.Format("UPDATE CRM_CUSTOMERPLUS SET APP_STATUS=0 WHERE PID='{0}' AND APP_STATUS=1 ", _PID)
                If _cmd.ExecuteNonQuery >= 1 Then status = True
            End Using
        Catch ex As Exception

        Finally
            _conn.Close()
        End Try
        Return status
    End Function
    Public Function IsTokenLive(ByVal _TOKEN As String) As Boolean
        Dim status As Boolean = False
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = String.Format("SELECT TOKEN FROM CRM_CUSTOMERTOKEN WHERE TOKEN='{0}' AND TIME_END>=GETDATE() AND VOID_CODE=0 ", _TOKEN)
                status = Not IsNothing(_cmd.ExecuteScalar)
                If status Then
                    _cmd.CommandText = String.Format("UPDATE CRM_CUSTOMERTOKEN SET TIME_END=DATEADD(hh,3,GETDATE()) WHERE TOKEN='{0}'", _TOKEN)
                    _cmd.ExecuteNonQuery()
                End If
            End Using
        Catch ex As Exception
        Finally
            _conn.Close()
        End Try
        Return status
    End Function
    Public Function GetPIDByToken(ByVal _TOKEN As String) As String
        Dim _pid As String = String.Empty
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = String.Format("SELECT PID FROM CRM_CUSTOMERTOKEN WHERE TOKEN='{0}' AND TIME_END>=GETDATE() AND VOID_CODE=0 ", _TOKEN)
                If _cmd.ExecuteScalar <> Nothing Then
                    _pid = _cmd.ExecuteScalar.ToString()
                End If

            End Using
        Catch ex As Exception
        Finally
            _conn.Close()
        End Try
        Return _pid
    End Function
    Public Function GetPIDByCard(ByVal _CARD As String) As String
        Dim status As String = String.Empty
        If _conn.State = ConnectionState.Closed Then _conn.Open()
        Try
            Using COMM As New SqlCommand("", _conn)
                COMM.CommandText = String.Format("SELECT PID FROM CRM_CARDDATA WHERE CARDNO='{0}' AND VOID_CODE=0 ", _CARD)
                status = COMM.ExecuteScalar.ToString()
            End Using
        Catch ex As Exception
            status = ""
        Finally
            _conn.Close()
        End Try
        Return status
    End Function
    Public Function GetTokenByPID(ByVal _PID As String) As String
        Dim _token As String = String.Empty
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = String.Format("SELECT TOKEN FROM CRM_CUSTOMERTOKEN WHERE PID='{0}' AND TIME_END>=GETDATE() AND VOID_CODE=0 ", _PID)
                If _cmd.ExecuteScalar = Nothing Then
                    _token = Me.NewToken(_PID).ToString
                Else
                    _token = _cmd.ExecuteScalar.ToString()
                End If

            End Using
        Catch ex As Exception
        Finally
            _conn.Close()
        End Try
        Return _token
    End Function

    Public Function GetTokenByPID(ByVal _PID As String, ByVal _CARDNO As String) As String
        Dim _token As String = String.Empty
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = String.Format("SELECT TOKEN FROM CRM_CUSTOMERTOKEN WHERE PID='{0}' AND TIME_END>=GETDATE() AND VOID_CODE=0 ", _PID)
                If _cmd.ExecuteScalar = Nothing Then
                    _token = Me.NewToken(_PID, _CARDNO).ToString
                Else
                    _token = _cmd.ExecuteScalar.ToString()
                End If

            End Using
        Catch ex As Exception
        Finally
            _conn.Close()
        End Try
        Return _token
    End Function
    Public Function NewToken(ByVal _pid As String, Optional ByVal _cardno As String = "") As String
        Dim _guid As String = String.Empty
        Dim _token As String = String.Empty
        
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using _cmd As New SqlCommand("", _conn)
                _cmd.CommandText = "SELECT NEWID()"
                _guid = _cmd.ExecuteScalar.ToString()
                _token = _guid.Replace("-", "").ToString()
                _cmd.CommandText = String.Format("INSERT INTO CRM_CUSTOMERTOKEN(PID,TOKEN,GUID,DEVICE_UUID,TIME_BEGIN,TIME_END,VOID_CODE,API_ACTION) VALUES('{0}','{1}','{2}','','{3}','{4}',0,'{5}')", _pid, _token, _guid, Now.ToString("yyyy-MM-dd HH:MM:ss"), Now.AddHours(3).ToString("yyyy-MM-dd HH:MM:ss"), _cardno)
                _cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            _token = ""
        Finally
            _conn.Close()
        End Try
        Return _token
    End Function

    Public Function GetBonusByYear(ByVal _TOKEN As String) As String
        Dim _pid As String = Me.GetPIDByToken(_TOKEN)
        Dim usedpoint As Decimal = 0
        If _conn.State = ConnectionState.Closed Then _conn.Open()
        Using cmd As New SqlCommand("", _conn)
            cmd.CommandText = _
                "SELECT SUM(purchase_bouns)bonus FROM [publish_data].DBO.POS_PURCHASE_RECORD " +
                String.Format("WHERE CARDNO IN (SELECT CARDNO FROM CRM..CRM_CARDDATA WHERE PID='{0}') ", _pid) +
                "AND VOID_CODE=0 AND purchase_bouns>0 AND CONVERT(int,LEFT(SYS_DATE,4))>=DATEPART(year,GETDATE())-1 GROUP BY LEFT(SYS_DATE,4) "
        End Using
        Dim _dt As New DataTable
        _jsonstring = ""
        Using sqlapt As New SqlDataAdapter("", _conn)
            sqlapt.SelectCommand.CommandText = _
                "SELECT SUM(purchase_bouns)bonus,(convert(varchar,convert(int,LEFT(sys_date,4))+1)+'/12/31')end_date " +
                "FROM [publish_data].DBO.POS_PURCHASE_RECORD " +
                String.Format("WHERE CARDNO IN (SELECT CARDNO FROM CRM..CRM_CARDDATA WHERE PID='{0}') ", _pid) +
                "AND VOID_CODE=0 AND purchase_bouns>0 AND CONVERT(int,LEFT(SYS_DATE,4))>=DATEPART(year,GETDATE())-1 GROUP BY LEFT(SYS_DATE,4) "

            sqlapt.Fill(_dt)
            If _dt.Rows.Count > 0 Then
                _jsonstring = JsonConvert.SerializeObject(_dt).ToString()
            End If
        End Using
        Return _jsonstring

    End Function
    Public Function GetBonusByToken(ByVal _token As String) As Decimal
        Dim _nuc As Decimal = 0
        Dim _pid As String = Me.GetPIDByToken(_token)
        Dim _cardno As String = Me.GetCardNobyPID(_pid)
        Try
            If _txgposconn.State = ConnectionState.Closed Then _txgposconn.Open()
            Using _cmd As New SqlCommand("", _txgposconn)
                _cmd.CommandText = String.Format("exec dbo.CRM_GETMEMBERBOUNS '{0}'", _cardno)

                _nuc = CDec(_cmd.ExecuteScalar())
            End Using
        Catch ex As Exception
            _nuc = 0
        Finally
            _conn.Close()
        End Try
        Return _nuc
    End Function
    Public Function GetBonusList(ByVal _TOKEN As String, ByVal _SDATE As String, ByVal _EDATE As String) As String
        Dim _pid As String = Me.GetPIDByToken(_TOKEN)
        Dim _dt As New DataTable
        _jsonstring = ""
        Try
            Using sqlapt As New SqlDataAdapter("", _conn)
                sqlapt.SelectCommand.CommandText = _
                    "SELECT (CONVERT(varchar,SYSTEMDATE,111)+' '+convert(varchar(5),systemdate,114))transaction_time," +
                    "(CASE WHEN DATEDIFF(year,SYSTEMDATE,GETDATE())>=2 THEN '' ELSE (CONVERT(varchar,DATEPART(YEAR,SYSTEMDATE)+1)+'/12/31') END)bonus_end_day," +
                    "(SELECT Comp_SName FROM ERP..BS_Company where ComCode=A.COMPNO)department_name," +
                    "(SELECT SUBSTRING(goods_name,1,CHARINDEX('(',goods_name)-1) FROM [trkdc-einvsrv].[trkdctxg].[dbo].vms_barcode WHERE COMPNO=A.COMPNO and venderno=A.VNDRNO)store_name," +
                    "(CASE WHEN A.PURCHASE_BOUNS>0 AND A.VOID_CODE=0 THEN '集點' WHEN A.PURCHASE_BOUNS>0 AND A.VOID_CODE<>0 THEN '集點作廢' " +
                    "WHEN A.PURCHASE_BOUNS<0 AND A.VOID_CODE=0 THEN '折抵' WHEN A.PURCHASE_BOUNS<0 AND A.VOID_CODE<>0 THEN '折抵作廢' ELSE '' END)bonus_title," +
                    "(CASE WHEN A.PURCHASE_BOUNS>0 AND A.VOID_CODE=0 THEN '1' WHEN A.PURCHASE_BOUNS>0 AND A.VOID_CODE<>0 THEN '3' " +
                    "WHEN A.PURCHASE_BOUNS<0 AND A.VOID_CODE=0 THEN '2' WHEN A.PURCHASE_BOUNS<0 AND A.VOID_CODE<>0 THEN '4' ELSE '' END)bonus_type," +
                    "CONVERT(VARCHAR,ABS(PURCHASE_BOUNS))bonus_count,(B.INVOICE_NO)receipt_number " +
                    "FROM PUBLISH_DATA.DBO.POS_PURCHASE_RECORD A INNER JOIN [trkdc-einvsrv].[trkdctxg].DBO.POS_TRAN_MAIN B ON B.COMPNO=A.COMPNO AND B.TRAN_DATE=A.SYS_DATE " +
                    "AND B.ECR_NO=A.ECR_NO AND B.TRAN_COUNT=A.TRANS_VOL " +
                String.Format("WHERE CARDNO IN (SELECT CARDNO FROM CRM..CRM_CARDDATA WHERE PID='{0}' AND VOID_CODE=0) AND A.SYS_DATE BETWEEN '{1}' AND '{2}' ORDER BY 1,4", _pid, _SDATE, _EDATE)

                sqlapt.Fill(_dt)
                If _dt.Rows.Count > 0 Then
                    _jsonstring = JsonConvert.SerializeObject(_dt).ToString()
                Else
                    _jsonstring = ""
                End If
            End Using
        Catch ex As Exception
            _jsonstring = "error"
        End Try

        Return _jsonstring
    End Function

    Public Function GetBonusList_Kiosk(ByVal _TOKEN As String, ByVal _SDATE As String, ByVal _EDATE As String) As String
        Dim _pid As String = Me.GetPIDByToken(_TOKEN)
        Dim _dt As New DataTable
        _jsonstring = ""
        Try
            Using sqlapt As New SqlDataAdapter("", _conn)
                sqlapt.SelectCommand.CommandText = _
                    "SELECT (CONVERT(varchar,SYSTEMDATE,111)+' '+convert(varchar(5),systemdate,114))transaction_time," +
                    "(CASE WHEN DATEDIFF(year,SYSTEMDATE,GETDATE())>=2 THEN '' ELSE (CONVERT(varchar,DATEPART(YEAR,SYSTEMDATE)+1)+'/12/31') END)bonus_end_day," +
                    "(SELECT Comp_SName FROM ERP..BS_Company where ComCode=A.COMPNO)department_name," +
                    "(SELECT SUBSTRING(goods_name,1,CHARINDEX('(',goods_name)-1) FROM [trkdc-einvsrv].[trkdctxg].[dbo].vms_barcode WHERE COMPNO=A.COMPNO and venderno=A.VNDRNO)store_name," +
                    "(CASE WHEN A.PURCHASE_BOUNS>0 AND A.VOID_CODE=0 THEN '集點' WHEN A.PURCHASE_BOUNS>0 AND A.VOID_CODE<>0 THEN '集點作廢' " +
                    "WHEN A.PURCHASE_BOUNS<0 AND A.VOID_CODE=0 THEN '折抵' WHEN A.PURCHASE_BOUNS<0 AND A.VOID_CODE<>0 THEN '折抵作廢' ELSE '' END)bonus_title," +
                    "(CASE WHEN A.PURCHASE_BOUNS>0 AND A.VOID_CODE=0 THEN '1' WHEN A.PURCHASE_BOUNS>0 AND A.VOID_CODE<>0 THEN '3' " +
                    "WHEN A.PURCHASE_BOUNS<0 AND A.VOID_CODE=0 THEN '2' WHEN A.PURCHASE_BOUNS<0 AND A.VOID_CODE<>0 THEN '4' ELSE '' END)bonus_type," +
                    "CONVERT(VARCHAR,ABS(PURCHASE_BOUNS))bonus_count,(B.INVOICE_NO)receipt_number,a.purchase_amount " +
                    "FROM PUBLISH_DATA.DBO.POS_PURCHASE_RECORD A INNER JOIN [trkdc-einvsrv].[trkdctxg].DBO.POS_TRAN_MAIN B ON B.COMPNO=A.COMPNO AND B.TRAN_DATE=A.SYS_DATE " +
                    "AND B.ECR_NO=A.ECR_NO AND B.TRAN_COUNT=A.TRANS_VOL " +
                String.Format("WHERE CARDNO IN (SELECT CARDNO FROM CRM..CRM_CARDDATA WHERE PID='{0}' AND VOID_CODE=0) AND A.SYS_DATE BETWEEN '{1}' AND '{2}' ORDER BY 1,4", _pid, _SDATE, _EDATE)

                sqlapt.Fill(_dt)
                If _dt.Rows.Count > 0 Then
                    _jsonstring = JsonConvert.SerializeObject(_dt).ToString()
                Else
                    _jsonstring = ""
                End If
            End Using
        Catch ex As Exception
            _jsonstring = "error"
        End Try

        Return _jsonstring
    End Function

    Public Function GetInvoiceMaster(ByVal _TOKEN As String, ByVal _SDATE As String, ByVal _EDATE As String) As String
        Dim _pid = Me.GetPIDByToken(_TOKEN)
        Dim _dt As New DataTable
        _jsonstring = ""
        Try
            Using sqlapt As New SqlDataAdapter("", _conn)
                sqlapt.SelectCommand.CommandText = _
                    "SELECT ((SELECT QNO FROM ERP..BS_Argument_Detail WHERE QKind='BS_COMPNO' AND QName=A.COMPNO)+REPLACE(SUBSTRING(B.TRAN_DATE,6,5),'-','')+RIGHT(A.ECR_NO,4)+CONVERT(VARCHAR,B.TRAN_COUNT))TRANSACTION_ID," +
                    "(CONVERT(varchar,SYSTEMDATE,111)+' '+convert(varchar(5),systemdate,114))transaction_time," +
                    "(SELECT Comp_SName FROM ERP..BS_Company where ComCode=A.COMPNO)department_name," +
                    "(SELECT SUBSTRING(goods_name,1,CHARINDEX('(',goods_name)-1) FROM [trkdc-einvsrv].[trkdctxg].[dbo].vms_barcode WHERE COMPNO=A.COMPNO and venderno=A.VNDRNO)store_name," +
                    "(A.COMPNO)STORE_ID,CONVERT(VARCHAR,CONVERT(INT,B.TOTAL_AMOUNT))consumption_amount,(B.INVOICE_NO)receipt_number,(0)receipt_win," +
                    "(CASE WHEN a.VOID_CODE=1 THEN '作廢' ELSE CONVERT(VARCHAR,PURCHASE_BOUNS) END)point_description " +
                    "FROM PUBLISH_DATA.DBO.POS_PURCHASE_RECORD A INNER JOIN [trkdc-einvsrv].[trkdctxg].DBO.POS_TRAN_MAIN B ON B.COMPNO=A.COMPNO AND B.TRAN_DATE=A.SYS_DATE " +
                    "AND B.ECR_NO=A.ECR_NO AND B.TRAN_COUNT=A.TRANS_VOL " +
                String.Format("WHERE CARDNO IN (SELECT CARDNO FROM CRM..CRM_CARDDATA WHERE PID='{0}' AND VOID_CODE=0) AND A.SYS_DATE BETWEEN '{1}' AND '{2}' order by 2,4 ", _pid, _SDATE, _EDATE)
                sqlapt.Fill(_dt)
                If _dt.Rows.Count > 0 Then
                    _jsonstring = JsonConvert.SerializeObject(_dt).ToString()
                Else
                    _jsonstring = ""
                End If
            End Using
        Catch ex As Exception
            _jsonstring = "error"
        End Try

        Return _jsonstring


    End Function

    Public Function GetInvoiceMaster(ByVal _PID As String, ByVal _SDATE As String, ByVal _COMPNO As String, ByVal _TRANCOUNT As String) As String()
        Dim _dt As New DataTable
        Dim returnstr(2) As String
        _jsonstring = ""
        Try
            Using sqlapt As New SqlDataAdapter("", _conn)
                sqlapt.SelectCommand.CommandText = _
                    "SELECT (b.tran_count)transaction_id, (CONVERT(varchar,b.sys_time,111)+' '+convert(varchar(5),B.sys_time,114))transaction_time," +
                    "(SELECT Comp_SName FROM ERP..BS_Company where ComCode=A.COMPNO)department_name," +
                    "ISNULL((SELECT IMAGE_CHT_NAME+'('+IMAGE_EN_NAME+')' FROM ERP..VMS_VENDER_CONTRACT_IMAGE WHERE (VENDERNO+IMAGENO)=B.VENDER_NO),'')store_name," +
                    "(A.COMPNO)store_id," +
                    "(SELECT ISNULL(SUM(PAY_AMOUNT),0) FROM ERP..POS_PAY_AMOUNT WHERE COMPNO=A.COMPNO AND TRAN_DATE=A.TRAN_DATE AND ECR_NO=A.ECR_NO " +
                    "AND TRAN_COUNT=A.TRAN_COUNT AND VOID_CODE=0 AND PAY_TYPE IN('C101','C102','C103','C105','C201'))consumption_amount,(B.TOTAL_AMOUNT)consumption_price_total," +
                    "(B.INVOICE_NO)receipt_number,(0)receipt_win " +
                    "FROM ERP..POS_CUSTOMER A INNER JOIN ERP..POS_TRAN_MAIN B ON B.COMPNO=A.COMPNO AND B.TRAN_DATE=A.TRAN_DATE AND B.TRAN_COUNT=A.TRAN_COUNT " +
                String.Format("WHERE A.customer_no IN(SELECT CARDNO FROM CRM..CRM_CARDDATA WHERE PID='{0}') AND A.TRAN_DATE BETWEEN '{1}' AND '{1}' AND A.TRAN_COUNT='{2}' AND A.COMPNO='{3}' ", _PID, _SDATE, _TRANCOUNT, _COMPNO)
                sqlapt.Fill(_dt)
                If _dt.Rows.Count > 0 Then
                    _jsonstring = """transaction_time"":""" + _dt.Rows(0)("transaction_time").ToString() + """,""department_name"":""" + _dt.Rows(0)("department_name").ToString() + """," +
                        """store_name"":""" + _dt.Rows(0)("store_name").ToString() + """,""consumption_amount"":""" + _dt.Rows(0)("consumption_amount").ToString() + """," +
                        """receipt_number"":""" + _dt.Rows(0)("receipt_number").ToString() + """,""receipt_win"":" + _dt.Rows(0)("receipt_win").ToString() + ","
                    returnstr(0) = _jsonstring
                    returnstr(1) = _dt.Rows(0)("consumption_price_total").ToString()
                    returnstr(2) = _dt.Rows(0)("consumption_amount").ToString()
                Else
                    _jsonstring = ""
                    returnstr(0) = _jsonstring
                    returnstr(1) = ""
                    returnstr(2) = ""
                End If
            End Using
        Catch ex As Exception
            _jsonstring = "error"
            returnstr(0) = _jsonstring
            returnstr(1) = ""
            returnstr(2) = ""
        End Try

        Return returnstr


    End Function
    Public Function GetInvoiceDetail(ByVal _PID As String, ByVal _DATE As String, ByVal _COMPNO As String, ByVal _TRANCOUNT As String) As String
        Dim _dt As New DataTable
        _jsonstring = ""
        Try
            Using sqlapt As New SqlDataAdapter("", _conn)
                sqlapt.SelectCommand.CommandText = _
                    "select (B.GOODS_NAME)name,CONVERT(bit,1)show_tax_info,(CASE WHEN TAX_TYPE='1' THEN '應稅' ELSE '免稅' END)tax,convert(bit,1)show_unit_price_info," +
                    "convert(varchar,convert(int,GOODS_PRICE))unit_price,convert(bit,1)show_unit_info,convert(varchar,GOODS_QTY)unit," +
                    "convert(varchar,convert(int,GOODS_AMOUNT))amount " +
                    "from ERP..POS_CUSTOMER A INNER JOIN ERP..POS_TRAN_GOODS B ON B.COMPNO=A.COMPNO AND B.TRAN_DATE=A.TRAN_DATE " +
                    "AND B.ECR_NO=A.ECR_NO AND B.TRAN_COUNT=A.TRAN_COUNT " +
                    String.Format("WHERE A.customer_no IN(SELECT CARDNO FROM CRM..CRM_CARDDATA WHERE PID='{0}') ", _PID) +
                    String.Format("AND A.TRAN_DATE BETWEEN '{0}' AND '{0}' AND A.TRAN_COUNT={1}", _DATE, _TRANCOUNT)
                sqlapt.Fill(_dt)
                If _dt.Rows.Count > 0 Then
                    _jsonstring = JsonConvert.SerializeObject(_dt).ToString()
                Else
                    _jsonstring = ""
                End If
            End Using
        Catch ex As Exception
            _jsonstring = "error"
        End Try

        Return _jsonstring


    End Function
    Public Function GetPayDetail(ByVal _PID As String, ByVal _DATE As String, ByVal _COMPNO As String, ByVal _TRANCOUNT As String) As String
        Dim _dt As New DataTable
        _jsonstring = ""
        Try
            Using sqlapt As New SqlDataAdapter("", _conn)
                sqlapt.SelectCommand.CommandText = _
                    "SELECT (B.PAY_NAME)name,CONVERT(bit,0)show_tax_info,('')tax,CONVERT(bit,0)show_unit_price_info,('')unit_price,CONVERT(bit,0)show_unit_info,('')unit,CONVERT(VARCHAR,pay_amount*(-1))amount " +
                    "FROM ERP..POS_CUSTOMER A INNER JOIN ERP..POS_PAY_AMOUNT B ON B.COMPNO =A.COMPNO AND B.TRAN_DATE =A.TRAN_DATE AND B.ECR_NO =A.ECR_NO AND B.TRAN_COUNT=A.TRAN_COUNT " +
                    String.Format("WHERE A.CUSTOMER_NO IN(SELECT CARDNO FROM CRM_CARDDATA WHERE PID='{0}') AND A.COMPNO='{1}' ", _PID, _COMPNO) +
                    String.Format("AND A.tran_date='{0}' AND A.TRAN_COUNT={1} AND pay_type NOT IN('C100')", _DATE, _TRANCOUNT)
                    
                sqlapt.Fill(_dt)
                If _dt.Rows.Count > 0 Then
                    _jsonstring = JsonConvert.SerializeObject(_dt).ToString()
                Else
                    _jsonstring = ""
                End If
            End Using
        Catch ex As Exception
            _jsonstring = "error"
        End Try

        Return _jsonstring


    End Function
    ''' <summary>
    ''' 檢查發票號碼是否符合規定
    ''' </summary>
    ''' <param name="_barcode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckIsUniform(ByRef _barcode As String) As Boolean
        Dim CheckReg As New Regex("[a-zA-Z][a-zA-Z][0-9]{8}")
        Return CheckReg.IsMatch(_barcode)
    End Function
    Public Function ReInputBonus(ByVal _BARCODE As String, ByVal _TOKEN As String) As String
        Dim status As String = String.Empty
        Dim compno, uniform As String
        Dim ecrno, trandate As String
        Dim trancount As Integer

        compno = _BARCODE.Split("|")(0).ToString()
        'pos_eno = _BARCODE.Split("|")(1).ToString()
        uniform = _BARCODE.Split("|")(2).ToString()

        If Not Me.CheckIsUniform(uniform) Then
            Return "ERROR"
            Exit Function
        End If



        Try
            If _txgposconn.State = ConnectionState.Closed Then _txgposconn.Open()
            Using COMM As New SqlCommand("", _txgposconn)
                COMM.CommandText = String.Format("EXEC [DBO].CRM_REINPUTBOUNS_APP '{0}','{1}'", uniform, _TOKEN)
                status = COMM.ExecuteScalar.ToString()
            End Using
        Catch ex As Exception
            status = ""
        Finally
            _txgposconn.Close()
        End Try
        Return status
    End Function
    Public Function UNREINPUTBOUNS(ByVal _BARCODE As String, ByVal _TOKEN As String) As String
        Dim status As String = String.Empty
        Dim uniform As String = String.Empty
        uniform = _BARCODE.Split("|")(2).ToString()

        If Not Me.CheckIsUniform(uniform) Then
            Return "ERROR"
            Exit Function
        End If
        Try
            If _txgposconn.State = ConnectionState.Closed Then _txgposconn.Open()
            Using COMM As New SqlCommand("", _txgposconn)
                COMM.CommandText = String.Format("EXEC [DBO].CRM_UNREINPUTBOUNS '{0}','{1}'", uniform, _TOKEN)
                status = COMM.ExecuteScalar.ToString()
            End Using
        Catch ex As Exception
            status = ""
        Finally
            _conn.Close()
        End Try
        Return status
    End Function
    Public Function MinusMEMBERBOUNS(ByVal _BONUS As Integer, ByVal _TOKEN As String) As String
        Dim status As String = String.Empty
        If _txgposconn.State = ConnectionState.Closed Then _txgposconn.Open()
        Try
            Using COMM As New SqlCommand("", _txgposconn)
                COMM.CommandText = String.Format("EXEC [POS_SVR].[DBO].CRM_MinusMEMBERBOUNS '{0}','{1}'", _TOKEN, _BONUS)
                status = COMM.ExecuteScalar.ToString()
            End Using
        Catch ex As Exception
            status = ""
        Finally
            _conn.Close()
        End Try
        Return status
    End Function
    Public Function MinusMEMBERBOUNS(ByVal _BONUS As Integer, ByVal _TOKEN As String, ByVal _RESULT As String) As String
        Dim status As String = String.Empty
        Dim _pid As String = Me.GetPIDByToken(_TOKEN)
        Dim _bonuspoint As Integer = Me.GetBonusByToken(_TOKEN)
        Dim _dt As New traderule
        _dt = JsonConvert.DeserializeObject(Of traderule)(_RESULT)
        If _pid = "" Then
            status = "NONE"
        Else
            If _txgposconn.State = ConnectionState.Closed Then _txgposconn.Open()
            Try
                Using COMM As New SqlCommand("", _txgposconn)
                    COMM.CommandText = String.Format("EXEC [POS_SVR].[DBO].CRM_EXCHANGE_PRINT2 '{0}','{1}','{2}',{3},'{4}',{5},'{6}','{7}'",
                                                     _dt.compno, Today.ToString("yyyy-MM-dd"), "APP", CInt(Now.ToString("mmss")), "APP", _bonuspoint, _pid, _dt.trade_id + _dt.trade_seq)
                    status = COMM.ExecuteScalar.ToString()
                    'status = status.Replace(status.Split("|")(3).ToString(), "***")
                End Using
            Catch ex As Exception
                status = "NONE"
            Finally
                _conn.Close()
            End Try
        End If
        Return status
    End Function
    Public Function GetBonusitemlist(ByVal _token As String) As String
        Dim _dt As New DataTable
        Try
            Using sqlapt As New SqlDataAdapter("", _txgposconn)
                sqlapt.SelectCommand.CommandText = "SELECT A.COMPNO,A.TRADE_ID,A.TRADE_SEQ,A.EXCHANGE_KIND,A.EXCHANGE_NAME,A.INV_AMOUNT,A.CHANGE_QTY,A.CHANGE_NUMBER," +
                                                   "C.TRADE_SDATE,C.TRADE_EDATE,(CASE WHEN C.EXCHANGE_SDATE='TODAY' THEN CONVERT(VARCHAR(10),GETDATE(),120) ELSE C.EXCHANGE_SDATE END)EXCHANGE_SDATE," +
                                                   "(CASE WHEN C.EXCHANGE_EDATE='TODAY' THEN CONVERT(VARCHAR(10),GETDATE(),120) ELSE C.EXCHANGE_EDATE END)EXCHANGE_EDATE," +
                                                   "(D.FUNCTION_DATA)VENDERNO,E.TENANT_SNA,(CASE WHEN E.FLOOR_NOS LIKE '6%' THEN REPLACE(E.FLOOR_NOS,'6','B') ELSE E.FLOOR_NOS END)FLOOR_NOS " +
                                                   "FROM CRM_TRADE_DETAIL A INNER JOIN CRM_TRADE_FUNCTION B ON B.TRADE_ID=A.TRADE_ID AND B.TRADE_SEQ=A.TRADE_SEQ " +
                                                   "AND B.FUNCTION_CATEGORY='SEND' AND FUNCTION_KIND='VENDER' AND FUNCTION_DATA ='KIOSK' " +
                                                   "INNER JOIN CRM_TRADE_MAIN C ON C.TRADE_ID=A.TRADE_ID AND CONVERT(VARCHAR(10),GETDATE(),120) BETWEEN C.TRADE_SDATE AND C.TRADE_EDATE " +
                                                   "LEFT JOIN CRM_TRADE_FUNCTION D ON D.TRADE_ID=A.TRADE_ID AND D.TRADE_SEQ=A.TRADE_SEQ " +
                                                   "AND D.FUNCTION_CATEGORY='RECOVER' AND D.FUNCTION_KIND='VENDER' " +
                                                   "LEFT JOIN [TRKDCTXG-DB].ORACLE_DB.dbo.TENANT_MN E ON E.TENANT_COD=D.FUNCTION_DATA"

                sqlapt.Fill(_dt)
                If _dt.Rows.Count > 0 Then
                    _jsonstring = JsonConvert.SerializeObject(_dt).ToString()
                Else
                    _jsonstring = ""
                End If
            End Using
        Catch ex As Exception
            _jsonstring = ""
        Finally
            _conn.Close()
        End Try
        Return _jsonstring
    End Function
    Public Function GetMemberData(ByVal _TOKEN As String) As String
        If Me.IsTokenLive(_TOKEN) Then
            Dim _dt As New DataTable
            Try
                Using sqlapt As New SqlDataAdapter("", _conn)
                    sqlapt.SelectCommand.CommandText = "SELECT (B.PNAME)name,(b.sex)gender,(b.pid)identifier,left(b.birthday,4)birth_year," +
                                                       "substring(b.birthday,6,2)birth_month,substring(b.birthday,9,2)birth_day,(b.link_tel)tel,(b.link_mtel)mobile," +
                                                       "(b.zipno)postcode,(b.adds)address,(b.link_email)email,c.marry_type,(c.edudata)edu_data," +
                                                       "(case when charindex(',',c.job,len(c.job)-1)>0 then substring(c.job,1,len(c.job)-1) else c.job end)job," +
                                                       "c.year_income,(case when charindex(',',c.interest,len(c.interest)-1)>0 then substring(c.interest,1,len(c.interest)-1) else c.interest end)interest " +
                                                       "FROM CRM_CUSTOMERTOKEN A INNER JOIN CRM_CUSTOMER B ON B.PID=A.PID INNER JOIN CRM_CUSTOMERPLUS C ON C.PID=A.PID " +
                                                       String.Format("WHERE A.TOKEN='{0}' AND A.VOID_CODE=0 ", _TOKEN)

                    sqlapt.Fill(_dt)
                    If _dt.Rows.Count > 0 Then
                        _jsonstring = JsonConvert.SerializeObject(_dt).ToString().Replace("[", "").Replace("]", "")
                    Else
                        _jsonstring = ""
                    End If
                End Using
            Catch ex As Exception
                _jsonstring = ""
            Finally
                _conn.Close()
            End Try
        Else
            _jsonstring = ""
        End If
        Return _jsonstring
    End Function
    Public Function UpdateMemberData(ByVal _token As String, ByVal _userinfo As String) As String
        Dim status As String = "FALSE"
        Dim _dt As New member
        _dt = JsonConvert.DeserializeObject(Of member)(_userinfo)
        Dim _pid As String = GetPIDByToken(_token)
        If IsNothing(_dt.pname) OrElse IsNothing(_dt.sex) OrElse IsNothing(_dt.pid) OrElse IsNothing(_dt.birth_year) OrElse IsNothing(_dt.birth_month) _
        OrElse IsNothing(_dt.birth_day) OrElse IsNothing(_dt.link_tel) OrElse IsNothing(_dt.link_mtel) OrElse IsNothing(_dt.adds) Then
            status = "SPACE"
        Else
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Dim _trans As SqlTransaction = _conn.BeginTransaction()
            Try
                Using _comm As New SqlCommand("", _conn, _trans)
                    _comm.CommandText = "UPDATE CRM_CUSTOMER SET PNAME=@PNAME,SEX=@SEX,BIRTHDAY=@BIRTHDAY,ZIPNO=@ZIPNO,ADDS=@ADDS,LINK_TEL=@LINK_TEL,LINK_MTEL=@LINK_MTEL,LINK_EMAIL=@LINK_EMAIL,USERNO='APP',LAST_DATE=GETDATE() WHERE PID=@PID "
                    _comm.Parameters.AddWithValue("@PID", _dt.pid)
                    _comm.Parameters.AddWithValue("@PNAME", _dt.pname)
                    _comm.Parameters.AddWithValue("@SEX", _dt.sex)
                    _comm.Parameters.AddWithValue("@BIRTHDAY", _dt.birth_year + "-" + _dt.birth_month + "-" + _dt.birth_day)
                    _comm.Parameters.AddWithValue("@ZIPNO", IIf(_dt.zipno Is Nothing, DBNull.Value, _dt.zipno))
                    _comm.Parameters.AddWithValue("@ADDS", _dt.adds)
                    _comm.Parameters.AddWithValue("@LINK_TEL", _dt.link_tel)
                    _comm.Parameters.AddWithValue("@LINK_MTEL", _dt.link_mtel)
                    _comm.Parameters.AddWithValue("@LINK_EMAIL", _dt.link_email)
                    _comm.ExecuteNonQuery()
                    _comm.Parameters.Clear()
                    _comm.CommandText = "UPDATE CRM_CUSTOMERPLUS SET USERNO='APP',LAST_DATE=GETDATE()," +
                        "INTEREST=@INTEREST,JOB=@JOB,MARRY_TYPE=@MARRY_TYPE,YEAR_INCOME=@YEAR_INCOME,EDUDATA=@EDUDATA WHERE PID=@PID"
                    _comm.Parameters.AddWithValue("@PID", _dt.pid)
                    _comm.Parameters.AddWithValue("@INTEREST", IIf(_dt.INTEREST Is Nothing, DBNull.Value, _dt.INTEREST))
                    _comm.Parameters.AddWithValue("@JOB", IIf(_dt.JOB Is Nothing, DBNull.Value, _dt.JOB))
                    _comm.Parameters.AddWithValue("@MARRY_TYPE", IIf(_dt.MARRY_TYPE Is Nothing, DBNull.Value, _dt.MARRY_TYPE))
                    _comm.Parameters.AddWithValue("@YEAR_INCOME", IIf(_dt.YEAR_INCOME Is Nothing, DBNull.Value, _dt.YEAR_INCOME))
                    _comm.Parameters.AddWithValue("@EDUDATA", IIf(_dt.edudata Is Nothing, DBNull.Value, _dt.edudata))
                    _comm.ExecuteNonQuery()
                    _trans.Commit()
                End Using
                status = "SUCCESS"
            Catch ex As Exception
                If _trans IsNot Nothing Then _trans.Rollback()
                status = "FALSE"
            Finally
                _conn.Close()
            End Try
        End If
        Return status
    End Function
    Public Function AddMemberdata(ByVal _userinfo As String) As String
        Dim status As String = "false"
        Dim _dt As New member
        _dt = JsonConvert.DeserializeObject(Of member)(_userinfo)
        If IsNothing(_dt.pname) OrElse IsNothing(_dt.sex) OrElse IsNothing(_dt.pid) OrElse IsNothing(_dt.birth_year) OrElse IsNothing(_dt.birth_month) _
        OrElse IsNothing(_dt.birth_day) OrElse IsNothing(_dt.link_tel) OrElse IsNothing(_dt.link_mtel) OrElse IsNothing(_dt.adds) Then
            status = "space"
        Else
            If _conn.State = ConnectionState.Closed Then _conn.Open()

            ' _conn.BeginTransaction("insert")
            Dim _trans As SqlTransaction = _conn.BeginTransaction()

            Try


                Using _comm As New SqlCommand("", _conn, _trans)

                    _comm.CommandText = "INSERT INTO CRM_CUSTOMER(PID,PNAME,SEX,BIRTHDAY,ZIPNO,ADDS,LINK_TEL,LINK_MTEL,LINK_EMAIL,VOID_CODE,USERNO,SYSTEMDATE,LAST_DATE) " +
                        "VALUES(@PID,@PNAME,@SEX,@BIRTHDAY,@ZIPNO,@ADDS,@LINK_TEL,@LINK_MTEL,@LINK_EMAIL,0,'APP',GETDATE(),NULL) "
                    _comm.Parameters.AddWithValue("@PID", _dt.pid)
                    _comm.Parameters.AddWithValue("@PNAME", _dt.pname)
                    _comm.Parameters.AddWithValue("@SEX", _dt.sex)
                    _comm.Parameters.AddWithValue("@BIRTHDAY", _dt.birth_year + "-" + _dt.birth_month + "-" + _dt.birth_day)
                    _comm.Parameters.AddWithValue("@ZIPNO", IIf(_dt.zipno Is Nothing, "", _dt.zipno))
                    _comm.Parameters.AddWithValue("@ADDS", _dt.adds)
                    _comm.Parameters.AddWithValue("@LINK_TEL", _dt.link_tel)
                    _comm.Parameters.AddWithValue("@LINK_MTEL", _dt.link_mtel)
                    _comm.Parameters.AddWithValue("@LINK_EMAIL", _dt.link_email)
                    _comm.ExecuteNonQuery()
                    _comm.Parameters.Clear()
                    _comm.CommandText = "INSERT INTO CRM_CUSTOMERPLUS(PID,EINVO_STATUS,DM_REVCODE,APP_STATUS,PASS_WORD,VOID_CODE,USERNO,SYSTEMDATE,LAST_DATE," +
                        "INTEREST,JOB,MARRY_TYPE,CHILD_NUM,YEAR_INCOME,EDUDATA) VALUES(@PID,0,0,1,@PASS_WORD,0,'APP',GETDATE(),NULL,@INTEREST,@JOB,@MARRY_TYPE,0,@YEAR_INCOME,@EDUDATA)"
                    _comm.Parameters.AddWithValue("@PID", _dt.pid)
                    _comm.Parameters.AddWithValue("@INTEREST", IIf(_dt.INTEREST Is Nothing, DBNull.Value, _dt.INTEREST))
                    _comm.Parameters.AddWithValue("@JOB", IIf(_dt.JOB Is Nothing, DBNull.Value, _dt.JOB))
                    _comm.Parameters.AddWithValue("@PASS_WORD", _dt.birth_year + _dt.birth_month + _dt.birth_day)
                    _comm.Parameters.AddWithValue("@MARRY_TYPE", IIf(_dt.MARRY_TYPE Is Nothing, DBNull.Value, _dt.MARRY_TYPE))
                    _comm.Parameters.AddWithValue("@YEAR_INCOME", IIf(_dt.YEAR_INCOME Is Nothing, DBNull.Value, _dt.YEAR_INCOME))
                    _comm.Parameters.AddWithValue("@EDUDATA", IIf(_dt.edudata Is Nothing, DBNull.Value, _dt.edudata))

                    _comm.ExecuteNonQuery()
                    _trans.Commit()
                End Using
                status = "success"
            Catch ex As Exception
                If _trans IsNot Nothing Then _trans.Rollback()
                status = "false"
            Finally
                _conn.Close()
            End Try
        End If
        Return status


    End Function

    Public Function SetPushToken(ByVal token As String, ByVal uuid As String, ByVal pushcode As String) As String
        Dim status As String = "FALSE"
        Dim pid As String = Me.GetPIDByToken(token)
        If _conn.State = ConnectionState.Closed Then _conn.Open()
        Try
            Using COMM As New SqlCommand("", _conn)
                COMM.CommandText = String.Format("INSERT INTO CRM_APP_PUSH(TRAN_DATE,TOKEN,PID,PUSH_CODE,DEVICE_UUID,VOID_CODE,SYS_TIME) VALUES('{0}','{1}','{2}','{3}','{4}',0,getdate())",
                                                 Today.ToString("yyyy-MM-dd"), token, pid, pushcode, uuid)
                COMM.ExecuteScalar()
                'status = status.Replace(status.Split("|")(3).ToString(), "***")
            End Using
            status = "SUCCESS"
        Catch ex As Exception
            status = "FALSE"
        Finally
            _conn.Close()
        End Try
        Return status
    End Function
End Class
Public Class member
    <JsonProperty("NAME")>
    Public pname As String
    <JsonProperty("gender")>
    Public sex As String
    <JsonProperty("identifier")>
    Public pid As String
    Public birth_year As String
    Public birth_month As String
    Public birth_day As String
    <JsonProperty("tel")>
    Public link_tel As String
    <JsonProperty("mobile")>
    Public link_mtel As String
    <JsonProperty("postcode")>
    Public zipno As String
    <JsonProperty("address")>
    Public adds As String
    <JsonProperty("email")>
    Public link_email As String
    Public MARRY_TYPE As String
    <JsonProperty("edu_data")>
    Public edudata As String
    Public JOB As String
    Public YEAR_INCOME As String
    Public INTEREST As String

End Class

Public Class traderule

    Public compno As String
    Public trade_id As String
    Public trade_seq As String
    Public change_number As String
    Public venderno As String

End Class