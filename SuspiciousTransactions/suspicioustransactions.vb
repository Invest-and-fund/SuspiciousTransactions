Imports System.Net.Mail
Imports System.Data.SqlClient
Module SuspiciousTransactions
    Dim LenderActivated(20000) As Integer

    Public Function fnDBStringField(ByVal sField) As String
        If IsDBNull(sField) Then
            fnDBStringField = " "
        Else
            fnDBStringField = Trim(CStr(sField))
        End If
    End Function
    Public Function fnDBIntField(ByVal sField) As String
        If IsDBNull(sField) Then
            fnDBIntField = 0
        Else
            fnDBIntField = CInt(sField)
        End If
    End Function

    Public Sub SendSimpleMail(sEmail As String, sSubject As String, sBody As String)
        Dim sPW As String = Configuration.ConfigurationManager.AppSettings("DailyReporterPW")
        Dim sUSR As String = Configuration.ConfigurationManager.AppSettings("DailyReporterUSR")
        Dim MyMailMessage As New MailMessage() With {
            .From = New MailAddress(sUSR),
            .Subject = sSubject,
            .IsBodyHtml = True,
            .Body = "<table><tr><td>" + sBody + "</table></td></tr>"
        }
        MyMailMessage.To.Add(sEmail)

        Dim SMTPServer As New SmtpClient("smtp.office365.com") With {
            .Credentials = New System.Net.NetworkCredential(sUSR, sPW),
            .Port = 587,
            .EnableSsl = True
        }

        Try
            SMTPServer.Send(MyMailMessage)
        Catch ex As Exception
            SendErrorMessage(ex)
        End Try
        SMTPServer = Nothing
        MyMailMessage = Nothing
    End Sub


    Public Function ExecuteIT() As String
        Dim sUsers, MySQL, MySQLFB, sHTML, s As String

        Dim dt, dt1, dt2, dt3 As New DataTable
        Dim FBSQLEnv As String = System.Configuration.ConfigurationManager.AppSettings("RunFBSQL")
        sUsers = Configuration.ConfigurationManager.AppSettings("EmailList")
        Dim iPeriod As Integer = System.Configuration.ConfigurationManager.AppSettings("Period")
        Dim iDayPeriod As Integer = System.Configuration.ConfigurationManager.AppSettings("DayPeriod")
        Dim startdate As Date = DateAdd(DateInterval.Day, iPeriod, Now())
        Dim enddate As Date = Now()
        Dim depositdate As Date

        MySQLFB = "select  u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID, ft.DATECREATED, ft.amount, ft.transtype from FIN_TRANS ft, USERS u, Accounts A
                where ft.DATECREATED  > @p1
                and a.ACCOUNTID = ft.ACCOUNTID
                and u.USERID = a.USERID
                and u.USERTYPE = 0
                and u.USERID not in (10,20,30)
                and ft.TRANSTYPE in (1100,1102,1023,1022,1200,1401)
                order by u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID, ft.DATECREATED, ft.amount, ft.transtype "

        MySQL = "select  u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID, ft.DATECREATED, ft.amount, ft.transtype from FIN_TRANS ft, USERS u, Accounts A
                where ft.DATECREATED  > DATEADD(dd, @p1, GETDATE())
                and a.ACCOUNTID = ft.ACCOUNTID
                and u.USERID = a.USERID
                and u.USERTYPE = 0
                and u.USERID not in (10,20,30)
                and ft.TRANSTYPE in (1100,1102,1023,1022,1200,1400)
                order by u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID, ft.DATECREATED, ft.amount, ft.transtype "
        Try
            If FBSQLEnv = "FB" Then
                Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
                Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQLFB, New FirebirdSql.Data.FirebirdClient.FbConnection(Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString))
                Adaptor.SelectCommand.Parameters.Clear()
                Adaptor.SelectCommand.Parameters.Add("@p1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = Now.AddDays(iPeriod)

                Adaptor.Fill(dt)
                Adaptor.Dispose()
            Else
                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    Try
                        Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                        Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                        con.Open()
                        cmd.Parameters.Clear()
                        With cmd.Parameters
                            .Add(New SqlParameter("@p1", iPeriod))
                        End With
                        adapter.SelectCommand = cmd
                        adapter.Fill(dt)
                        con.Close()
                        con.Dispose()
                    Catch ex As Exception
                    Finally

                    End Try
                End Using
            End If
            MySQLFB = "select count(*) as icount, u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID from FIN_TRANS ft, USERS u, Accounts A
                where ft.DATECREATED  >  @p1
                and a.ACCOUNTID = ft.ACCOUNTID
                and u.USERID = a.USERID
                and u.USERTYPE = 0
                and u.USERID not in (10,20,30)
                and ft.TRANSTYPE in (1100)
                group by u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID"

            MySQL = "select count(*) as icount, u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID from FIN_TRANS ft, USERS u, Accounts A
                where ft.DATECREATED  > DATEADD(dd, @p1, GETDATE())
                and a.ACCOUNTID = ft.ACCOUNTID
                and u.USERID = a.USERID
                and u.USERTYPE = 0
                and u.USERID not in (10,20,30)
                and ft.TRANSTYPE in (1100)
                group by u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID"

            If FBSQLEnv = "FB" Then
                Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
                Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQLFB, New FirebirdSql.Data.FirebirdClient.FbConnection(Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString))
                Adaptor.SelectCommand.Parameters.Clear()
                Adaptor.SelectCommand.Parameters.Add("@p1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = Now.AddDays(iPeriod)

                Adaptor.Fill(dt1)
                Adaptor.Dispose()
            Else
                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    Try
                        Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                        Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                        con.Open()
                        cmd.Parameters.Clear()
                        With cmd.Parameters
                            .Add(New SqlParameter("@p1", iPeriod))
                        End With
                        adapter.SelectCommand = cmd
                        adapter.Fill(dt1)
                        con.Close()
                        con.Dispose()
                    Catch ex As Exception
                    Finally

                    End Try
                End Using
            End If

            MySQLFB = "select  u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID, ft.DATECREATED, ft.amount, ft.transtype  
                 from FIN_TRANS ft, USERS u, Accounts A
                where ft.DATECREATED  >  @p1
                and a.ACCOUNTID = ft.ACCOUNTID
                and u.USERID = a.USERID
                and u.USERTYPE = 0
                and u.USERID not in (10,20,30)
                and ft.TRANSTYPE in (1100)
                and exists(select * from FIN_TRANS ft1
                where ft1.DATECREATED  between ft.DATECREATED and ft.DATECREATED +5
                and ft.ACCOUNTID = ft1.ACCOUNTID
                and ft1.TRANSTYPE in (1102))"

            MySQL = "select  u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID, ft.DATECREATED, ft.amount, ft.transtype  
                 from FIN_TRANS ft, USERS u, Accounts A
                where ft.DATECREATED  > DATEADD(dd, @p1, getdate())
                and a.ACCOUNTID = ft.ACCOUNTID
                and u.USERID = a.USERID
                and u.USERTYPE = 0
                and u.USERID not in (10,20,30)
                and ft.TRANSTYPE in (1100)
                and exists(select * from FIN_TRANS ft1
                where ft1.DATECREATED  between ft.DATECREATED and dateadd(d, 5,ft.DATECREATED)
                and ft.ACCOUNTID = ft1.ACCOUNTID
                and ft1.TRANSTYPE in (1102))"

            If FBSQLEnv = "FB" Then
                Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
                Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQLFB, New FirebirdSql.Data.FirebirdClient.FbConnection(Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString))
                Adaptor.SelectCommand.Parameters.Clear()
                Adaptor.SelectCommand.Parameters.Add("@p1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = Now.AddDays(iPeriod)

                Adaptor.Fill(dt2)
                Adaptor.Dispose()
            Else
                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    Try
                        Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                        Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                        con.Open()
                        cmd.Parameters.Clear()
                        With cmd.Parameters
                            .Add(New SqlParameter("@p1", iPeriod))
                        End With
                        adapter.SelectCommand = cmd
                        adapter.Fill(dt2)
                        con.Close()
                        con.Dispose()
                    Catch ex As Exception
                    Finally

                    End Try
                End Using
            End If
            MySQLFB = "select  u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID, ft.DATECREATED, ft.amount, ft.transtype  
                 from FIN_TRANS ft, USERS u, Accounts A
                where ft.DATECREATED  >  @p1
                and a.ACCOUNTID = ft.ACCOUNTID
                and u.USERID = a.USERID
                and u.USERTYPE = 0
                and u.USERID not in (10,20,30)
                and ft.TRANSTYPE in (1100)
                and not exists(select * from FIN_TRANS ft1
                where ft1.DATECREATED  between ft.DATECREATED and ft.DATECREATED + 10
                and ft.ACCOUNTID = ft1.ACCOUNTID
                and ft1.TRANSTYPE in (1200,1401))"

            MySQL = "select  u.userid, u.companyname,u.firstname, u.lastname,ft.ACCOUNTID, ft.DATECREATED, ft.amount, ft.transtype  
                 from FIN_TRANS ft, USERS u, Accounts A
                where ft.DATECREATED  > DATEADD(dd, @p1, getdate())
                and a.ACCOUNTID = ft.ACCOUNTID
                and u.USERID = a.USERID
                and u.USERTYPE = 0
                and u.USERID not in (10,20,30)
                and ft.TRANSTYPE in (1100)
                and not exists(select * from FIN_TRANS ft1
                where ft1.DATECREATED  between ft.DATECREATED and dateadd(d, @p2,ft.DATECREATED)
                and ft.ACCOUNTID = ft1.ACCOUNTID
                and ft1.TRANSTYPE in (1200,1401));"

            If FBSQLEnv = "FB" Then
                Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
                Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(MySQLFB, New FirebirdSql.Data.FirebirdClient.FbConnection(Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString))
                Adaptor.SelectCommand.Parameters.Clear()
                Adaptor.SelectCommand.Parameters.Add("@p1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = Now.AddDays(iPeriod)
                Adaptor.SelectCommand.Parameters.Add("@p2", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = iDayPeriod
                Adaptor.Fill(dt3)
                Adaptor.Dispose()
            Else
                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    Try
                        Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                        Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                        con.Open()
                        cmd.Parameters.Clear()
                        With cmd.Parameters
                            .Add(New SqlParameter("@p1", iPeriod))
                            .Add(New SqlParameter("@p2", iDayPeriod))
                        End With
                        adapter.SelectCommand = cmd
                        adapter.Fill(dt3)
                        con.Close()
                        con.Dispose()
                    Catch ex As Exception
                    Finally

                    End Try
                End Using
            End If

            sHTML = "<html><body><head>
                <style>
                table {
                    font-family: arial, sans-serif;
                    border-collapse: collapse;
                    width: 100%;
                }

                td, th {
                    border: 1px solid #dddddd;
                    text-align: left;
                    padding: 8px;
                }

                tr:nth-child(even) {
                    background-color: #dddddd;
                }
                </style>
                </head>
                <table>
                  <tr>
                    <th style='font-size:30px' colspan=5>Suspicious Transactions Report</th>
                  </tr>
                  <tr>
                    <th style='font-size:15px' align='center' colspan=5>Report Period - Between " & startdate.ToString("dd/MM/yyyy") & " & " & enddate.ToString("dd/MM/yyyy") & "</th>
                  </tr>
                  <tr>
                    <th colspan=5></th>
                  </tr>
                  <tr>
                    <th style='font-size:20px' colspan=5>Deposits > £100K </th>
                  </tr>
                  <tr>
                    <th>Transaction Date</th>
                    <th>Individ/Co.</th>
                    <th>Name</th>
                    <th>Account</th>
                    <th>Amount</th>
                    <th>Transaction Type</th>
                  </tr>"

            For Each ThisRow As DataRow In dt.Rows
                If fnDBIntField(ThisRow("amount")) > 10000000 And fnDBStringField(ThisRow("transtype")) = 1100 Then
                    Dim iUserID As Integer = fnDBIntField(ThisRow("UserID"))
                    'If iActiv <> LenderActivated(iUserID) Then
                    sHTML &= "<tr><td>" & fnDBStringField(ThisRow("DATECREATED")) & "</td>"
                    sHTML &= "<td>" & fnDBStringField(ThisRow("companyname")) & "</td>"
                    sHTML &= "<td>" & fnDBStringField(ThisRow("Firstname")) & " " & fnDBStringField(ThisRow("Lastname")) & "</td>"
                    sHTML &= "<td>" & fnDBStringField(ThisRow("ACCOUNTID")) & "</td>"
                    sHTML &= "<td>" & PenceToCurrencyStringPounds(ThisRow("amount")) & "</td>"
                    sHTML &= "<td>" & fnDBStringField(ThisRow("transtype")) & " - Deposit</td>"
                    sHTML &= vbNewLine

                    ' End If
                End If

            Next
            sHTML &= "<tr>
                    <th colspan=5></th>
                  </tr>
                  <tr>
                    <th style='font-size:20px' colspan=5>3 or more Deposits and no subsequent Bid or Buy within 10 days</th>
                  </tr>
                  <tr>
                    <th>Transaction Date</th>
                    <th>Individ/Co.</th>
                    <th>Name</th>
                    <th>Account</th>
                    <th>Amount</th>
                    <th>Transaction Type</th>
                  </tr>"
            Dim icount As Integer
            icount = dt.Rows.Count

            For Each ThisRow1 As DataRow In dt1.Rows
                If fnDBIntField(ThisRow1("icount")) > 2 Then
                    For Each ThisRow As DataRow In dt3.Rows
                        If fnDBIntField(ThisRow1("UserID")) = fnDBStringField(ThisRow("UserID")) And fnDBStringField(ThisRow("transtype")) = 1100 Then
                            Dim iUserID As Integer = fnDBIntField(ThisRow("UserID"))
                            'If iActiv <> LenderActivated(iUserID) Then
                            sHTML &= "<tr><td>" & fnDBStringField(ThisRow("DATECREATED")) & "</td>"
                            sHTML &= "<td>" & fnDBStringField(ThisRow("companyname")) & "</td>"
                            sHTML &= "<td>" & fnDBStringField(ThisRow("Firstname")) & " " & fnDBStringField(ThisRow("Lastname")) & "</td>"
                            sHTML &= "<td>" & fnDBStringField(ThisRow("ACCOUNTID")) & "</td>"
                            sHTML &= "<td>" & PenceToCurrencyStringPounds(ThisRow("amount")) & "</td>"
                            sHTML &= "<td>" & fnDBStringField(ThisRow("transtype")) & " - Deposit</td>"
                            sHTML &= vbNewLine

                            ' End If
                        End If
                    Next
                    sHTML &= "<tr>
                    <th colspan=5></th>
                  </tr>"
                End If

            Next
            sHTML &= "<tr>
                    <th colspan=5></th>
                  </tr>
                  <tr>
                    <th style='font-size:20px' colspan=5>Deposits and subsequent Withdrawal within 5 days or less </th>
                  </tr>
                  <tr>
                    <th>Transaction Date</th>
                    <th>Individ/Co.</th>
                    <th>Name</th>
                    <th>Account</th>
                    <th>Amount</th>
                    <th>Transaction Type</th>
                  </tr>"

            icount = dt2.Rows.Count

            For Each ThisRow1 As DataRow In dt2.Rows
                Dim newdate, newdate1 As Date
                newdate = fnDBDateField(ThisRow1("DATECREATED"))
                For Each ThisRow As DataRow In dt.Rows

                    If fnDBIntField(ThisRow1("UserID")) = fnDBStringField(ThisRow("UserID")) And fnDBStringField(ThisRow("transtype")) = 1100 And fnDBDateField(ThisRow1("DATECREATED")) = fnDBDateField(ThisRow("DATECREATED")) Then
                        depositdate = fnDBDateField(ThisRow1("DATECREATED"))
                        newdate1 = fnDBDateField(ThisRow("DATECREATED"))
                        Dim iUserID As Integer = fnDBIntField(ThisRow("UserID"))
                        sHTML &= "<tr><td>" & fnDBStringField(ThisRow("DATECREATED")) & "</td>"
                        sHTML &= "<td>" & fnDBStringField(ThisRow("companyname")) & "</td>"
                        sHTML &= "<td>" & fnDBStringField(ThisRow("Firstname")) & " " & fnDBStringField(ThisRow("Lastname")) & "</td>"
                        sHTML &= "<td>" & fnDBStringField(ThisRow("ACCOUNTID")) & "</td>"
                        sHTML &= "<td>" & PenceToCurrencyStringPounds(ThisRow("amount")) & "</td>"
                        sHTML &= "<td>" & fnDBStringField(ThisRow("transtype")) & " - Deposit</td>"
                        sHTML &= vbNewLine
                    End If
                    If fnDBIntField(ThisRow1("UserID")) = fnDBStringField(ThisRow("UserID")) And fnDBStringField(ThisRow("transtype")) = 1102 And fnDBDateField(ThisRow("DATECREATED")) <= DateAdd(DateInterval.Day, 5, depositdate) Then
                        Dim iUserID As Integer = fnDBIntField(ThisRow("UserID"))
                        newdate1 = fnDBDateField(ThisRow("DATECREATED"))
                        sHTML &= "<tr><td>" & fnDBStringField(ThisRow("DATECREATED")) & "</td>"
                        sHTML &= "<td>" & fnDBStringField(ThisRow("companyname")) & "</td>"
                        sHTML &= "<td>" & fnDBStringField(ThisRow("Firstname")) & " " & fnDBStringField(ThisRow("Lastname")) & "</td>"
                        sHTML &= "<td>" & fnDBStringField(ThisRow("ACCOUNTID")) & "</td>"
                        sHTML &= "<td>" & PenceToCurrencyStringPounds(ThisRow("amount")) & "</td>"
                        sHTML &= "<td>" & fnDBStringField(ThisRow("transtype")) & " - Withdrawal</td>"
                        sHTML &= vbNewLine
                    End If
                Next

                sHTML &= "<tr>
                    <th colspan=5></th>
                  </tr>"
                depositdate = "01/01/2000"
            Next

            sHTML &= "<tr>
                    <th colspan=5></th>
                  </tr>
                  <tr>
                    <th style='font-size:20px' colspan=5>Withdrawal > £100K  </th>
                  </tr>
                  <tr>
                    <th>Transaction Date</th>
                    <th>Individ/Co.</th>
                    <th>Name</th>
                    <th>Account</th>
                    <th>Amount</th>
                    <th>Transaction Type</th>
                  </tr>"

            For Each ThisRow As DataRow In dt.Rows
                If fnDBIntField(ThisRow("amount")) > 10000000 And fnDBStringField(ThisRow("transtype")) = 1102 Then
                    Dim iUserID As Integer = fnDBIntField(ThisRow("UserID"))
                    'If iActiv <> LenderActivated(iUserID) Then
                    sHTML &= "<tr><td>" & fnDBStringField(ThisRow("DATECREATED")) & "</td>"
                    sHTML &= "<td>" & fnDBStringField(ThisRow("companyname")) & "</td>"
                    sHTML &= "<td>" & fnDBStringField(ThisRow("Firstname")) & " " & fnDBStringField(ThisRow("Lastname")) & "</td>"
                    sHTML &= "<td>" & fnDBStringField(ThisRow("ACCOUNTID")) & "</td>"
                    sHTML &= "<td>" & PenceToCurrencyStringPounds(ThisRow("amount")) & "</td>"
                    sHTML &= "<td>" & fnDBStringField(ThisRow("transtype")) & " - Withdrawal</td>"
                    sHTML &= vbNewLine

                    ' End If
                End If

            Next
            sHTML &= "</table></body></html>"

            SendSimpleMail(sUsers, "Suspicious Transactions Report", sHTML)

            ExecuteIT = sHTML

            MySQL = " Insert into REPORTRUN (REPORTNAME,SUCCESS,EMAILSENT,MSGDATA) VALUES (@P1,@P2,@P3,@P4)"

            If FBSQLEnv = "FB" Then
                Dim strConn As String
                Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
                Dim Cmd As New FirebirdSql.Data.FirebirdClient.FbCommand
                strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
                MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
                MyConn.Open()
                Cmd.Connection = MyConn
                Cmd.CommandText = MySQL
                Cmd.Parameters.Clear()
                Cmd.Parameters.Add("p1", FirebirdSql.Data.FirebirdClient.FbDbType.Char).Value = "SUSPICIOUSTRANSACTIONS"
                Cmd.Parameters.Add("p2", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = 1
                Cmd.Parameters.Add("p3", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = 1
                Cmd.Parameters.Add("p4", FirebirdSql.Data.FirebirdClient.FbDbType.Char).Value = ""
                Cmd.ExecuteNonQuery()

                MyConn.Close()
                MyConn = Nothing
            Else
                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    con.Open()

                    Dim command As SqlCommand = con.CreateCommand()
                    command.Connection = con

                    Try
                        command.CommandText = MySQL
                        With command.Parameters
                            .Add(New SqlParameter("@p1", "SUSPICIOUSTRANSACTIONS"))
                            .Add(New SqlParameter("@p2", 1))
                            .Add(New SqlParameter("@p3", 1))
                            .Add(New SqlParameter("@p4", ""))

                        End With
                        command.ExecuteNonQuery()
                    Catch ex As Exception

                    End Try
                End Using
            End If
        Catch ex As Exception
            s = ex.Message
            SendErrorMessage(ex)
            ExecuteIT = 0
        End Try
    End Function

    Sub Main()
        ExecuteIT()
    End Sub

    Sub SendErrorMessage(ByVal ThisException As Exception)
        Dim errorPW As String = Configuration.ConfigurationManager.AppSettings("ErrorPW")
        Dim errorUSR As String = Configuration.ConfigurationManager.AppSettings("ErrorUSR")
        Dim mm As New MailMessage() With {
            .From = New MailAddress(errorUSR),
            .Subject = "An Error Has Occurred!",
            .IsBodyHtml = True,
            .Priority = MailPriority.High
        }
        mm.To.Add("web@investandfund.com")

        mm.Body =
            "<html>" & vbCrLf &
            "<body>" & vbCrLf &
            "<h1>An Error Has Occurred!</h1>" & vbCrLf &
            "<table cellpadding=""5"" cellspacing=""0"" border=""1"">" & vbCrLf &
            ItemFormat("Time of Error", DateTime.Now.ToString("dd/MM/yyyy HH:mm:sss"))

        Try
            mm.Body += ItemFormat("Exception Type", ThisException.GetType().ToString())
        Catch ex As Exception
            mm.Body += ItemFormat("Exception Type", "Could not get exception type")
        End Try

        Try
            mm.Body += ItemFormat("Message", ThisException.Message)
        Catch ex As Exception
            mm.Body += ItemFormat("Message", "Could not get message")
        End Try

        Try
            mm.Body += ItemFormat("File Name", "suspicioustransactions.vb")
        Catch ex As Exception
            mm.Body += ItemFormat("File Name", "Could not get file name")
        End Try

        Try
            mm.Body += ItemFormat("Line Number", New StackTrace(ThisException, True).GetFrame(0).GetFileLineNumber)
        Catch ex As Exception
            mm.Body += ItemFormat("Line Number", "Could not get line number")
        End Try

        mm.Body +=
            "</table>" & vbCrLf &
            "</body>" & vbCrLf &
            "</html>"

        Dim smtp As New SmtpClient("smtp.office365.com") With {
            .Credentials = New System.Net.NetworkCredential(errorUSR, errorPW),
            .EnableSsl = True,
            .Port = 587
        }
        smtp.Send(mm)
        Dim MySQL As String
        Dim FBSQLEnv As String = System.Configuration.ConfigurationManager.AppSettings("RunFBSQL")
        MySQL = " Insert into REPORTRUN (REPORTNAME,SUCCESS,EMAILSENT,MSGDATA) VALUES (@P1,@P2,@P3,@P4)"

        If FBSQLEnv = "FB" Then
            Dim strConn As String
            Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
            Dim Cmd As New FirebirdSql.Data.FirebirdClient.FbCommand
            strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
            MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
            MyConn.Open()
            Cmd.Connection = MyConn
            Cmd.CommandText = MySQL
            Cmd.Parameters.Clear()
            Cmd.Parameters.Add("p1", FirebirdSql.Data.FirebirdClient.FbDbType.Char).Value = "SUSPICIOUSTRANSACTIONS"
            Cmd.Parameters.Add("p2", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = 0
            Cmd.Parameters.Add("p3", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = 0
            Cmd.Parameters.Add("p4", FirebirdSql.Data.FirebirdClient.FbDbType.Char).Value = ThisException.Message
            Cmd.ExecuteNonQuery()

            MyConn.Close()
            MyConn = Nothing
        Else
            Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                con.Open()

                Dim command As SqlCommand = con.CreateCommand()
                command.Connection = con

                Try
                    command.CommandText = MySQL
                    With command.Parameters
                        .Add(New SqlParameter("@p1", "SUSPICIOUSTRANSACTIONS"))
                        .Add(New SqlParameter("@p2", 0))
                        .Add(New SqlParameter("@p3", 0))
                        .Add(New SqlParameter("@p4", ThisException.Message))

                    End With
                    command.ExecuteNonQuery()
                Catch ex As Exception

                End Try
            End Using
        End If
    End Sub

    Public Function ItemFormat(ByVal Title As String, ByVal Message As String) As String
        Return "  <tr>" & vbCrLf &
                "  <tdtext-align: right;font-weight: bold"">" & Title & ":</td>" & vbCrLf &
                "  <td>" & Message & "</td>" & vbCrLf &
                "  </tr>" & vbCrLf
    End Function
    Public Function PenceToCurrencyStringPounds(ByVal sField) As String
        Dim rVal As Double
        Dim iPence As Integer

        If IsDBNull(sField) Then
            rVal = 0.0
        Else
            Try
                iPence = CInt(sField)
                rVal = iPence / 100
            Catch ex As Exception
                rVal = 0.0
            End Try
        End If
        PenceToCurrencyStringPounds = "£" & Format(rVal, "###,###,##0.00")
    End Function
    Public Function fnDBDateField(ByVal sField) As DateTime
        If IsDBNull(sField) Then
            fnDBDateField = DateTime.MinValue
        Else
            fnDBDateField = CDate(sField)
        End If
    End Function
End Module
