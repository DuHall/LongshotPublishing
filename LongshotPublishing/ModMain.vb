Imports System.Net.Mail
Imports System.Net.Mail.SmtpClient

Module ModMain

    Dim g_EmailHostIPAddress As String = ConfigurationManager.AppSettings("EmailIPAddress")
    Dim g_EmailPort As String = ConfigurationManager.AppSettings("EmailPort")
    Dim g_EmailUserName As String = ConfigurationManager.AppSettings("EmailUserName")
    Dim g_EmailPassword As String = ConfigurationManager.AppSettings("EmailPassword")
    Dim g_EmailFromAddress As String = ConfigurationManager.AppSettings("EmailFromAddress")
    Dim g_EmailEnabled As Boolean = IIf(UCase(ConfigurationManager.AppSettings("EmailEnabled")) = "TRUE", True, False)
    Dim g_SiteLocation As String = ConfigurationManager.AppSettings("siteLocation")
    Public g_EfilesReferralsBaseDirectory As String = ConfigurationManager.AppSettings("EfilesReferralsBaseDirectory")
    Public g_PDFTempDirectory As String = ConfigurationManager.AppSettings("TempPDFDir")
    Public g_StringArrayOuterSplitParameter = "||"
    Public g_StringArrayValueSplitParameter = "~~"
    Public g_intUserRecid = -1
    Public g_strFormObjectTypes = "input+select+textarea"
    Public g_Debug As Boolean = False

    Public Sub g_RetrieveSessions(ByRef txtSessions As TextBox)
        g_RetrieveSessions(txtSessions.Text)
    End Sub

    Public Sub g_RetrieveSessions(ByRef txtSessions As HiddenField)
        g_RetrieveSessions(txtSessions.Value)
    End Sub

    Public Sub g_RetrieveSessions(ByRef txtSessions As String)

        If Trim(txtSessions = "") Then
            ' no Sessions string area sent to client form
        Else
            Dim arrStrSessionVariables() As String = Split(txtSessions, "^^")


            ' restore session variables
            For Each strSessionVariable As String In arrStrSessionVariables
                Dim strSessionVariablePair() As String = Split(strSessionVariable, "||")
                System.Web.HttpContext.Current.Session(strSessionVariablePair(0)) = strSessionVariablePair(1)
            Next
        End If

    End Sub

    Public Sub g_SendSessions(ByRef txtSessions As TextBox)
        txtSessions.Text = g_SendSessions(txtSessions.Text)
    End Sub

    Public Sub g_SendSessions(ByRef txtSessions As HiddenField)
        txtSessions.Value = g_SendSessions(txtSessions.Value)
    End Sub

    Public Function g_SendSessions(ByRef txtSessions As String) As String

        '  Only do this if user is not signed on
        Dim strSessionsName As String = ""
        Dim strDelimiter As String = ""

        For Each txtFieldName As String In System.Web.HttpContext.Current.Session.Keys
            If txtFieldName.ToUpper = "RELAYMESSAGE" Then
            Else
                strSessionsName = strSessionsName & strDelimiter & txtFieldName & "||" & System.Web.HttpContext.Current.Session(txtFieldName)
                strDelimiter = "^^"
            End If
        Next

        Return strSessionsName

    End Function

    Public Sub g_SendEmail(ByVal ToAddress As String, ByVal Subject As String, ByVal Message As String)
        Debug.WriteLine("Email Module -- To: " & ToAddress & "  Subject: " & Subject & "   Message: " & Message)
        If ConfigurationManager.AppSettings("EmailEnabled") = True Then
            Dim Mail As New System.Net.Mail.MailMessage
            Mail.Subject = Subject
            If ToAddress = "" Then
                Debug.Print("ModMain (g_SendEmail): No email address provided. Can't send it.")
            Else
                For Each strEmailAddress As String In Split(ToAddress, ";")
                    Mail.To.Add(strEmailAddress)
                Next

                Mail.From = New System.Net.Mail.MailAddress(g_EmailFromAddress)
                Mail.Body = Message

                Dim strHTMLCheck As String = UCase(Message)
                Mail.IsBodyHtml = strHTMLCheck.ToUpper.Contains("<BODY") Or strHTMLCheck.ToUpper.Contains("<TABLE") Or strHTMLCheck.ToUpper.Contains("<DIV") Or strHTMLCheck.ToUpper.Contains("<BR") Or strHTMLCheck.ToUpper.Contains("<P")
                Dim SMTPServer As New System.Net.Mail.SmtpClient()

                SMTPServer.Timeout = 100000
                SMTPServer.Host = g_EmailHostIPAddress
                SMTPServer.Port = g_EmailPort
                SMTPServer.EnableSsl = False
                ''SMTPServer.Credentials = New System.Net.NetworkCredential(g_EmailUserName, g_EmailPassword)
                Debug.Print(ToAddress & " - " & Subject)
                If g_EmailEnabled Then
                    SMTPServer.Send(Mail)
                End If
            End If
        End If
    End Sub

    Public Function g_ArchiveData(ByVal ActiveTableName As String,
                              ByVal ArchiveTableName As String,
                              ByVal WherePhrase_DoNotIncludeTheWordWhere As String,
                              ByVal FormName As String) As Integer

        ' function returns # of records archived

        Dim strSQL As String = "Select * from " & ActiveTableName & " where " & WherePhrase_DoNotIncludeTheWordWhere
        Dim tblActiveData As DataTable = g_IO_Execute_SQL(strSQL, False)

        ' get field information for archive table so as to know auto increment filed
        Dim strAutoIncFields As String = ""
        Dim strDelim As String = ""
        Dim tblArchiveTable As DataTable = g_getTableColumnInfo(ArchiveTableName)
        For Each rowColumn In tblArchiveTable.Rows
            If rowColumn("AutoInc") = 1 Then
                strAutoIncFields &= strDelim & UCase(rowColumn("FieldName"))
                strDelim = ","
            End If
        Next
        Dim arrAutoIncFields() As String = Split(strAutoIncFields, ",")


        'move this over to the history table
        Dim nvcInsert As New NameValueCollection
        For Each rowItem As DataRow In tblActiveData.Rows
            For Each colFields As DataColumn In tblActiveData.Columns
                Dim blnIncludeColumn As Boolean = True
                If IsDBNull(rowItem(colFields.ColumnName)) Then
                    blnIncludeColumn = False
                    'skip adding this column since it's NULL
                Else
                    Dim strColumnName As String = UCase(colFields.ColumnName)
                    For Each strArrColumnName As String In arrAutoIncFields
                        If strArrColumnName = strColumnName Then
                            blnIncludeColumn = False
                            Exit For
                        End If
                    Next
                End If
                If blnIncludeColumn Then
                    nvcInsert(colFields.ColumnName) = rowItem(colFields.ColumnName)
                End If
            Next
            g_IO_SQLInsert(ArchiveTableName, nvcInsert, FormName)
        Next

        ' delete the records from the active
        g_IO_SQLDelete(ActiveTableName, WherePhrase_DoNotIncludeTheWordWhere)

        Return tblActiveData.Rows.Count
    End Function

    Function GetNextDate(ByVal d As DayOfWeek, Optional ByVal StartDate As Date = Nothing) As Date
        If StartDate = DateTime.MinValue Then
            StartDate = Now
            For p As Integer = 1 To 7
                If StartDate.AddDays(p).DayOfWeek = d Then Return StartDate.AddDays(p)
            Next
        Else
            Return Date.Now
        End If
    End Function

    Public Sub loadUserPrivilegesToSession()
        ''This routine will set all user items into session variables for quick checks while getting rid of unnecessary SQL I/O.
        If System.Web.HttpContext.Current.Session("User_ID") = "" Then
            ''Do nothing
        Else
            ''System.Web.HttpContext.Current.Session("User_Name")
            Dim user As String = System.Web.HttpContext.Current.Session("User_ID").ToString.ToUpper

            Dim strSql As String = "Select [user_id], [action], [authorization] from vw_sys_user_privileges where user_ID = '" & user & "' order by action"
            Dim tblResult As DataTable = g_IO_Execute_SQL(strSql, False)
            If tblResult.Rows.Count > 0 Then
                For Each row In tblResult.Rows
                    System.Web.HttpContext.Current.Session(row("action")) = row("authorization")
                Next
            Else
                ''Do nothing
            End If
        End If
    End Sub

    Public Function checkUserSessionPrivilege(ByRef action As String)

        If System.Web.HttpContext.Current.Session("User_ID") = "" Then
            Return False
        Else
            If Not IsNothing(System.Web.HttpContext.Current.Session(action)) Then
                If System.Web.HttpContext.Current.Session(action).ToString.ToUpper = "ALLOW" Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If

        End If
    End Function

    Public Function checkUserPrivilege(ByRef action As String)
        If System.Web.HttpContext.Current.Session("User_ID") = "" Then
            Return False
        Else
            ''System.Web.HttpContext.Current.Session("User_Name")
            Dim user As String = System.Web.HttpContext.Current.Session("User_ID").ToString.ToUpper

            Dim strSql As String = "Select [Authorization] from vw_sys_user_privileges where user_ID = '" & user & "' and action = '" & action & "'"
            Dim tblResult As DataTable = g_IO_Execute_SQL(strSql, False)
            If tblResult.Rows.Count > 0 Then
                If tblResult.Rows(0)(0) = "ALLOW" Then
                    Return True
                Else
                    Return False
                End If
            Else
                'Doesn't exist in the database, add it. 
                strSql = "Insert into SYS_USER_PRIVILEGES (user_id, action_id) VALUES ('" & user & "', '" & action & "')"
                g_IO_Execute_SQL(strSql, False)
                Return False
            End If
        End If
    End Function

    Public Function checkUserReportPrivilege(ByRef rptName As String)
        If System.Web.HttpContext.Current.Session("User_Name") = "" Then
            Return False
        Else
            ''System.Web.HttpContext.Current.Session("User_Name")
            Dim user As String = System.Web.HttpContext.Current.Session("User_Name")

            Dim strSql As String = "Select [Authorization] from vw_sys_user_report_privileges where user_name = '" & user & "' and report_name = '" & rptName & "'"
            Dim tblResult As DataTable = g_IO_Execute_SQL(strSql, False)
            If tblResult.Rows.Count > 0 Then
                If tblResult.Rows(0)(0) = "ALLOW" Then
                    Return True
                Else
                    Return False
                End If
            Else
                'Doesn't exist in the database, add it. 
                'strSql = "Insert into vw_sys_user_report_privileges (user_id, action_id) VALUES ('" & user & "', '" & rptName & "')"
                'g_IO_Execute_SQL(strSql, False)
                Return False
            End If
        End If
    End Function

    Public Sub AuditPrivileges()
        ''This routine audits the sys_user_privileges and enters missing entries.
        'get list of users
        Dim strSql As String = "Select * from sys_users"
        Dim tblUsers As DataTable = g_IO_Execute_SQL(strSql, False)

        'get list of Actions
        strSql = "Select * from Actions"
        Dim tblActions As DataTable = g_IO_Execute_SQL(strSql, False)

        For Each rowUser In tblUsers.Rows
            Dim user_name As String = rowUser("user_name")
            Dim user_id As String = rowUser("user_id")

            'iterate through the actions and see if an entry exists in the database.
            For Each rowAction In tblActions.Rows
                Dim Action As String = rowAction("Action")
                Dim Actions_id As String = rowAction("Actions_ID")
                strSql = "Select * from vw_sys_user_privileges where user_id = '" & user_id & "' and action = '" & Action & "'"
                If g_IO_Execute_SQL(strSql, False).Rows.Count > 0 Then
                Else
                    'create the entry
                    strSql = "Insert into sys_user_privileges (user_id, action_id) VALUES ('" & user_id & "'," & Actions_id & ")"
                    g_IO_Execute_SQL(strSql, False)
                End If
            Next
        Next
    End Sub

    Public Sub AuditReportPrivileges()
        ''This routine audits the SYS_USER_REPORT_PRIVILEGES and enters missing entries.
        'get list of users
        Dim strSql As String = "Select * from sys_users"
        Dim tblUsers As DataTable = g_IO_Execute_SQL(strSql, False)

        'get list of Actions
        strSql = "Select * from Report_Listing"
        Dim tblActions As DataTable = g_IO_Execute_SQL(strSql, False)

        For Each rowUser In tblUsers.Rows
            Dim user_name As String = rowUser("user_name")
            Dim user_id As String = rowUser("user_id")

            'iterate through the actions and see if an entry exists in the database.
            For Each rowReport In tblActions.Rows
                Dim Report_Name As String = rowReport("Report_Name")
                Dim Report_Listing_ID As Integer = rowReport("Report_Listing_ID")
                strSql = "Select * from VW_REPORT_PRIVILEGES where user_id = '" & user_id & "' and Report_Name = '" & Report_Name & "'"
                If g_IO_Execute_SQL(strSql, False).Rows.Count > 0 Then
                Else
                    'create the entry
                    strSql = "Insert into SYS_USER_REPORT_PRIVILEGES (user_id, Report_Listing_ID) VALUES ('" & user_id & "'," & Report_Listing_ID & ")"
                    g_IO_Execute_SQL(strSql, False)
                End If
            Next
        Next
    End Sub

    Public Function validateUser()
        Dim User As String = ""
        Dim strSql As String = ""
        Dim tblResults As DataTable = Nothing

        ''''''''''VERIFY USER''''''''''''
        If IsNothing(System.Web.HttpContext.Current.Session("User_Name")) Then
            User = HttpContext.Current.Request.LogonUserIdentity.Name
            strSql = "Select user_name, user_id, portal, default_page from [dbo].[SYS_USERS] where user_id = '" & User & "'"
            tblResults = g_IO_Execute_SQL(strSql, False)

            If tblResults.Rows.Count > 0 Then
                System.Web.HttpContext.Current.Session("User_Name") = tblResults.Rows(0)("user_name")
                System.Web.HttpContext.Current.Session("User_ID") = tblResults.Rows(0)("user_id")
                System.Web.HttpContext.Current.Session("Default_Portal") = tblResults.Rows(0)("portal")
                System.Web.HttpContext.Current.Session("Default_Page") = tblResults.Rows(0)("Default_Page")
                loadUserPrivilegesToSession()
                Return True
            Else
                Return False
            End If
        Else
            Return True
            Exit Function
        End If
    End Function

    Public Sub createDefaultRow(ByRef tbl As DataTable, ByVal strDefault As String, ByVal strTextField As String, ByVal strValueField As String)
        ''Create the default
        Dim defaultRow As DataRow = tbl.NewRow
        defaultRow(strTextField) = strDefault
        defaultRow(strValueField) = "-1"
        tbl.Rows.InsertAt(defaultRow, 0)
    End Sub

    Public Function g_getRows_FailedBackupsLastDay()
        Dim strsql As String = "exec udp_portal_rptGetFailedBackupsListLast24Hrs"  ''"SELECT    sld.server_name , A.database_name , ( SELECT TOP 1 last_backup_date FROM [V700GV158].[DBA].[dbo].[udt_Database_FileSizes] B WHERE     A.database_name = B.database_name ) AS last_backup_date, sms.mirroring_state_desc FROM [V700GV158].[DBA].[dbo].[udt_Database_FileSizes] A INNER JOIN dbo.vw_server_listing_detail sld ON A.server_name = sld.server_name LEFT outer JOIN DBO.VW_SERVER_MIRROR_STATUS SMS ON A.server_name = SMS.server_name AND A.database_name = SMS.database_name WHERE     LOWER(A.database_name) NOT IN ( 'master', 'model', 'tempdb', 'dba', 'msdb' ) AND last_backup_date < DATEADD(DAY, -1, GETDATE())  AND A.fileid = 1 ORDER BY  sld.server_name;"
        ''and sld.CRITICALITY = 'PRODUCTION' " & _
        Dim tblResults As DataTable = g_IO_Execute_SQL(strsql, False)
        Return tblResults
    End Function

    Public Function g_getCountFailedBackupsLastDay()
        Dim strSqL As String = "select count(sld.server_name) as Count " &
                "from [V700GV158].[DBA].[dbo].[udt_Database_FileSizes] A " &
                "inner join dbo.vw_server_listing_detail sld on A.server_name = sld.SERVER_NAME " &
                "LEFT OUTER JOIN dbo.VW_SERVER_MIRROR_STATUS SMS ON A.server_name = SMS.SERVER_NAME AND A.database_name = SMS.database_name " &
                "where " &
                "lower(A.database_name) not in ('master','model','tempdb','dba', 'msdb') " &
                "and A.fileid = 1 " &
                "and last_backup_date < dateadd(day,-1,getdate())"
        Dim tblResults As DataTable = g_IO_Execute_SQL(strSqL, False)
        Return tblResults
    End Function

    Public Function g_getCount_FailedSqlJobs()
        Dim strsql As String = "Select Count(*) as Count from RECENT_FAILED_JOBS"
        Dim tblResults As DataTable = g_IO_Execute_SQL(strsql, False)
        Return tblResults
    End Function

    Public Function g_globalSearch(ByVal searchString As String, ByVal searchType As String)
        Dim strResult As String = ""

        If Trim(searchString) = "" Then
            strResult = ""
            Return strResult
        End If

        Dim strSearchWord As String = Trim(searchString).Replace("'", "''")
        'clear the previous search history
        searchString = ""
        ''Begin search with through the servers
        Dim strSql As String = "Select server_id, server_name from vw_server_listing_detail where server_name like '%" & strSearchWord & "%' order by server_name"
        Dim tblresults As DataTable = g_IO_Execute_SQL(strSql, False)
        Dim blnEntryFound As Boolean = False
        searchString &= "<h2>Search Results</h2><table style=""width: 100%; padding: 3px; border: thin solid black; background-color: white;""><tr style=""background-color: #cdd4ff;""><th style=""border: thin solid black; padding: 5px;"">Type</th><th style=""border: thin solid black; padding: 5px;"">&nbsp;</th></tr>"
        If tblresults.Rows.Count > 0 Then
            ''It's a server, Display this as a result
            blnEntryFound = True
            For Each row In tblresults.Rows
                searchString &= "<tr><td style=""border: thin solid black; padding: 5px;"">Server</td><td style=""border: thin solid black; padding: 5px;""><a href=""serverdetail.aspx?id=" & row("server_id") & """>" & row("Server_Name") & "</a></td></tr>"
            Next
        End If

        ''Now search through the APPS
        strSql = "SELECT  SERVER_ID,SERVER_NAME,APPLICATION_NAME,CRITICALITY FROM [dbo].vw_server_application WHERE APPLICATION_NAME LIKE '%" & strSearchWord & "%' AND SERVER_NAME IN (" &
            "Select DISTINCT SERVER_NAME  FROM    [dbo].[vw_server_listing_detail] WHERE   ARCHIVED = 0 ) ORDER BY CRITICALITY , SERVER_NAME;"
        tblresults = g_IO_Execute_SQL(strSql, False)
        If tblresults.Rows.Count > 0 Then
            ''It's a database, display them.
            blnEntryFound = True
            For Each row In tblresults.Rows
                searchString &= "<tr><td style=""border: thin solid black; padding: 5px;"">APP</td><td style=""border: thin solid black; padding: 5px;"">" &
                    "<a href=""serverdetail.aspx?id=" & row("server_id") & """>" & row("Server_Name") & "</a> - > " & row("APPLICATION_NAME") & " - > " & row("CRITICALITY") & "</td></tr>"
            Next
        End If

        ''Now search through the Users
        strSql = "Select sam_account_name, name, office_phone FROM [T_DESKTOP_PORTAL].[dbo].[USERS] where name like '%" & strSearchWord & "%'  or sam_account_name like '%" & strSearchWord & "%' order by name"
        tblresults = g_IO_Execute_SQL(strSql, False)
        If tblresults.Rows.Count > 0 Then
            ''It's a database, display them.
            blnEntryFound = True
            For Each row In tblresults.Rows
                searchString &= "<tr><td style=""border: thin solid black; padding: 5px;"">USER</td><td style=""border: thin solid black; padding: 5px;"">" & row("NAME").ToString.ToUpper() & " -> " & row("SAM_ACCOUNT_NAME").ToString.ToUpper() & " -> " & row("OFFICE_PHONE").ToString.Replace(".", "-") & "</td></tr>"
            Next
        End If

        ''Check the search type
        If searchType = "FULL" Then
            ''Now search throught the Databases Names
            strSql = "Select * from vw_db_file_sizes where database_name like '%" & strSearchWord & "%' and fileid = '1' order by database_name"
            tblresults = g_IO_Execute_SQL(strSql, False)
            If tblresults.Rows.Count > 0 Then
                ''It's a database, display them.
                blnEntryFound = True
                For Each row In tblresults.Rows
                    searchString &= "<tr><td style=""border: thin solid black; padding: 5px;"">Database</td><td style=""border: thin solid black; padding: 5px;""><a href=""serverdetail.aspx?id=" & row("server_id") & """>" & row("server_name") & "</a> -> " & row("database_name") & " -> Size: " & row("size_mb") & " MB</td></tr>"
                Next
            End If

            ''Now search through the DNS Aliases
            strSql = "Select a.owner_name, b.server_name, a.ip_address FROM v700gv191.P_INFRA_Server_RPT.dbo.udt_INFRA_server_DNS as A inner join dbo.SERVERS As B on a.SERVER_NAME = b.SERVER_NAME where a.ip_address like '%" & strSearchWord & "%' or a.owner_name like '%" & strSearchWord & "%' or a.SERVER_NAME like '%" & strSearchWord & "%' order by owner_name"
            tblresults = g_IO_Execute_SQL(strSql, False)
            If tblresults.Rows.Count > 0 Then
                ''It's a database, display them.
                blnEntryFound = True
                For Each row In tblresults.Rows
                    Dim ipaddress As String = IIf((IsDBNull(row("ip_address")) Or row("ip_address") = ""), "SERVER REF", row("ip_address"))
                    searchString &= "<tr><td style=""border: thin solid black; padding: 5px;"">DNS</td><td style=""border: thin solid black; padding: 5px;"">" & row("OWNER_NAME") & " - > " & row("SERVER_NAME") & " - > " & ipaddress & "</td></tr>"
                Next
            End If

            ''Now search throught the Databases FILENAMES
            strSql = "Select * from vw_db_file_sizes where name like '%" & strSearchWord & "%'  and fileid = '1' order by name"
            tblresults = g_IO_Execute_SQL(strSql, False)
            If tblresults.Rows.Count > 0 Then
                ''It's a database, display them.
                blnEntryFound = True
                For Each row In tblresults.Rows
                    searchString &= "<tr><td style=""border: thin solid black; padding: 5px;"">Database FileName</td><td style=""border: thin solid black; padding: 5px;""><a href=""serverdetail.aspx?id=" & row("server_id") & """>" & row("server_name") & "</a> -> " & row("name") & " -> Size: " & row("size_mb") & " MB</td></tr>"
                Next
            End If





        End If



        If blnEntryFound = False Then
            searchString &= "<tr><td colspan=""2"">No results returned</td></tr>"
        End If

        searchString &= "</table>"
        Return searchString
    End Function

End Module
