Imports System.Collections.Specialized

Module ModIO

    '
    '**********************************Local Variable Definition**********************************
    '
    ' these variables are defined in web.config
    Public g_intNumberOfRetriesToAccessDatabase As Integer = ConfigurationManager.AppSettings("inNumberOfRetriesToAccessDatabase")
    Public g_ConnectionToUse As String = ConfigurationManager.AppSettings("ConnectionType")
    Public g_defaultEmail As String = ConfigurationManager.AppSettings("defaultEmail")
    Public g_siteLocation As String = ConfigurationManager.AppSettings("siteLocation")
    Public g_contactPhone As String = ConfigurationManager.AppSettings("contactPhone")
    Public g_ConnectionSchema As String = ConfigurationManager.AppSettings("schema")
    '
    Public g_blnAbort As Boolean = False
    '
    '_____local IO varaibles
    '
    Private m_nvcTables As New NameValueCollection               ' key=tablename, data=index to nvcFields 
    Private m_nvcFields() As NameValueCollection                 ' (=index#fromnvcTables),Key=FieldName,Data=index to nvcFieldAttributes
    Private m_nvcFieldAttributes() As NameValueCollection        ' (=index#fromnvcFields),Key=AttributeName,Data=AttributeValue
    Private m_IntLastRecid As Integer                          ' passes recid from insert down to recid request (in MSSQL only)

    '
    '**********************************Public Sub-Routines for Data IO**********************************
    '
    Public Function g_IO_PrepareSQL_String_A(ByVal strSQL As String) As String
        Dim blnInsert As Boolean = False
        If strSQL.ToUpper.IndexOf("INSERT ") = 0 Then
            Try
                ' convert MySQL True/False references to 0/1
                Dim strFields As String = strSQL.Substring(0, strSQL.ToUpper.IndexOf(" VALUES"))
                Dim strValues As String = strSQL.Substring(strSQL.ToUpper.IndexOf(" VALUES"))
                strValues = strValues.Replace(", True", ", 1")
                strValues = strValues.Replace(",True", ",1")
                strValues = strValues.Replace(", TRUE", ", 1")
                strValues = strValues.Replace(",TRUE", ",1")
                strValues = strValues.Replace(", False", ", 0")
                strValues = strValues.Replace(",False", ",0")
                strValues = strValues.Replace(", FALSE", ", 0")
                strValues = strValues.Replace(",FALSE", ",0")
                strSQL = strFields & strValues

                strSQL &= ";Select @@IDENTITY as ID;"    ' on MSSQL inserts must retrieve the new RECID now
                blnInsert = True
            Catch : End Try
        ElseIf strSQL.ToUpper.IndexOf("UPDATE ") = 0 Then
            Try
                ' convert MySQL True/False references to 0/1
                Dim strFields As String = strSQL.Substring(0, strSQL.ToUpper.IndexOf(" SET"))
                Dim strValues As String = strSQL.Substring(strSQL.ToUpper.IndexOf(" SET"))
                strValues = strValues.Replace("= True", "=1")
                strValues = strValues.Replace("=True", "=1")
                strValues = strValues.Replace("= TRUE", "= 1")
                strValues = strValues.Replace("=TRUE", "=1")
                strValues = strValues.Replace("= False", "= 0")
                strValues = strValues.Replace("=False", "=0")
                strValues = strValues.Replace("= FALSE", "= 0")
                strValues = strValues.Replace("=FALSE", "=0")
                strSQL = strFields & strValues
            Catch : End Try
        ElseIf strSQL.ToUpper.IndexOf("SELECT ") = 0 Then
            If strSQL.ToUpper.IndexOf("CONCAT(") = -1 Then
            Else
                Call convertConcatStatement(strSQL)
            End If
        End If
        Return strSQL
    End Function

    Public Function g_IO_Execute_SQL(ByVal strSQL As String, ByRef blnReturnErrorCode As Boolean) As DataTable

        'System.Web.HttpContext.Current.Response.Write(strSQL & "<br>")

        If g_ConnectionToUse = "MYSQL" Then
            Return IO_Execute_MYSQL(strSQL, blnReturnErrorCode)
        Else

            Dim blnInsert As Boolean = False
            If strSQL.ToUpper.IndexOf("INSERT ") = 0 Then
                Try
                    ' convert MySQL True/False references to 0/1
                    Dim strFields As String = strSQL.Substring(0, strSQL.ToUpper.IndexOf(" VALUES"))
                    Dim strValues As String = strSQL.Substring(strSQL.ToUpper.IndexOf(" VALUES"))
                    strValues = strValues.Replace(", True", ", 1")
                    strValues = strValues.Replace(",True", ",1")
                    strValues = strValues.Replace(", TRUE", ", 1")
                    strValues = strValues.Replace(",TRUE", ",1")
                    strValues = strValues.Replace(", False", ", 0")
                    strValues = strValues.Replace(",False", ",0")
                    strValues = strValues.Replace(", FALSE", ", 0")
                    strValues = strValues.Replace(",FALSE", ",0")
                    strSQL = strFields & strValues

                    strSQL &= ";Select @@IDENTITY as ID;"    ' on MSSQL inserts must retrieve the new RECID now
                    blnInsert = True
                Catch : End Try
            ElseIf strSQL.ToUpper.IndexOf("UPDATE ") = 0 Then
                Try
                    ' convert MySQL True/False references to 0/1
                    Dim strFields As String = strSQL.Substring(0, strSQL.ToUpper.IndexOf(" SET"))
                    Dim strValues As String = strSQL.Substring(strSQL.ToUpper.IndexOf(" SET"))
                    strValues = strValues.Replace("= True", "=1")
                    strValues = strValues.Replace("=True", "=1")
                    strValues = strValues.Replace("= TRUE", "= 1")
                    strValues = strValues.Replace("=TRUE", "=1")
                    strValues = strValues.Replace("= False", "= 0")
                    strValues = strValues.Replace("=False", "=0")
                    strValues = strValues.Replace("= FALSE", "= 0")
                    strValues = strValues.Replace("=FALSE", "=0")
                    strSQL = strFields & strValues
                Catch : End Try
            ElseIf strSQL.ToUpper.IndexOf("SELECT ") = 0 Then
                If strSQL.ToUpper.IndexOf("CONCAT(") = -1 Then
                Else
                    Call convertConcatStatement(strSQL)
                End If
            End If

            Try
                Dim tblTemp As DataTable = IO_Execute_MSSQL(strSQL, blnReturnErrorCode)
                If strSQL.ToUpper.IndexOf("SELECT ") = 0 Then
                    ' trim all text entries
                    For Each rowTemp As DataRow In tblTemp.Rows
                        For i = 0 To tblTemp.Columns.Count - 1
                            Try : rowTemp(i) = Trim(rowTemp(i)) : Catch : End Try
                        Next
                    Next
                ElseIf blnInsert Then
                    System.Web.HttpContext.Current.Session("NewMSSQLRECID") = tblTemp.Rows(0)("ID")

                End If
                Return tblTemp
            Catch ex As Exception
                System.Web.HttpContext.Current.Session("SQLERROR") = ex.Message
            End Try
        End If

    End Function

    Private Sub convertConcatStatement(ByRef strSQL As String)

        Dim strEndOfFieldMarker As String = ","
        Dim blnThisCharIsAQuote As Boolean = False
        Dim intIndex As Integer = 0
        Dim strFieldValue As String = ""
        Dim blnEndOfFieldFound As Boolean = False
        Dim intLargestRowIndex As Integer = 0
        Dim blnConversionComplete As Boolean = False



        Dim intStart As Integer = strSQL.ToUpper.IndexOf("CONCAT(")
        Dim strFrontSQL As String = Left(strSQL, intStart) & "("
        Dim strBackSQL As String = Mid(strSQL, intStart + 8)
        Dim I As Integer = 1

        Do Until blnConversionComplete
            Dim strCurrentChar As String = Mid(strBackSQL, I, 1)

            blnEndOfFieldFound = False

            If strCurrentChar = ")" And (strEndOfFieldMarker = "," Or strEndOfFieldMarker = "") Then

                Exit Do
            ElseIf strEndOfFieldMarker = "" Then
                ' end the middle of evaluating a string

                Select Case strCurrentChar
                    Case ","
                        strEndOfFieldMarker = ","
                        strFrontSQL &= " + "
                    Case "'"
                        strEndOfFieldMarker = "'"
                End Select
            ElseIf strEndOfFieldMarker = "'" Then

                ' this is a string under evaluation
                If Mid(strBackSQL, I, 2) = "''" Then
                    'this is a quote within the string
                    blnThisCharIsAQuote = True
                    I += 1   ' this quote takes up two spaces
                Else
                    blnThisCharIsAQuote = False
                End If

                If blnThisCharIsAQuote Then
                    strFrontSQL &= "'"
                Else
                    ' is this the end of the string?
                    If strCurrentChar = strEndOfFieldMarker Then
                        strFrontSQL &= "'"
                        strEndOfFieldMarker = ""
                        intIndex += 1
                    Else
                        ' build current field char by char
                        strFrontSQL &= strCurrentChar
                    End If

                End If


            Else
                ' this should be a field name (might not be if there are spaces before the starting quote of a string)
                If strCurrentChar = "'" Then
                    ' oops, this is a string after all
                    strEndOfFieldMarker = "'"
                    strFrontSQL &= "'"
                ElseIf strCurrentChar = strEndOfFieldMarker Then
                    strEndOfFieldMarker = ","
                    strFrontSQL &= " + "
                    intIndex += 1
                Else
                    strFrontSQL &= strCurrentChar
                End If
            End If

            I += 1
        Loop

        strSQL = strFrontSQL & Mid(strBackSQL, I)

    End Sub

    Public Function g_IO_ReadPageOfRecords(ByVal strSQL As String, ByVal intSkip As Integer, ByVal intNumberOfRecords As Integer, ByRef blnReturnErrorCode As Boolean) As DataTable

        If g_ConnectionToUse = "MYSQL" Then
            strSQL &= " limit " & intSkip & "," & intNumberOfRecords
        Else
            Dim strOrderBy As String = UCase(strSQL.Substring(UCase(strSQL).IndexOf(" ORDER BY ")))  ' extract Order By clause
            Dim strOrderByReversed As String = strOrderBy.Replace(" ASC", " %%").Replace(" DESC", " ASC").Replace("%%", "ASC")
            '' ''strSQL = strSQL.Substring(0, UCase(strSQL).IndexOf(" ORDER BY "))   ' remove Order By clause
            strSQL = "Select * from (" &
                " Select top " & intNumberOfRecords & " * from (" &
                strSQL.Substring(0, UCase(strSQL).IndexOf("SELECT") + 6) &
                " Top " & intSkip + intNumberOfRecords & " " &
                strSQL.Substring(UCase(strSQL).IndexOf("SELECT") + 6) &
                ") as Dummy1 " & strOrderByReversed &
                ") as Dummy2 " & strOrderBy
        End If

        Return g_IO_Execute_SQL(strSQL, blnReturnErrorCode)
    End Function

    Public Function g_IO_GetLastRecId() As Integer
        Dim intLastRecId As Integer
        If g_ConnectionToUse = "MYSQL" Then
            intLastRecId = g_IO_Execute_SQL("select last_insert_id() as recid", False).Rows(0)("RECID")
        Else
            intLastRecId = System.Web.HttpContext.Current.Session("NewMSSQLRECID")
        End If
        Return intLastRecId
    End Function
    '
    '**********************************Private Functions - SQL Executes**********************************
    '
    '
    '_____ sql final executions functions 
    ' 
    Private Function IO_Execute_MYSQL(ByVal strSQL As String, ByRef ReturnCode As Boolean) As DataTable

        ' '' '' build sql command for this query
        '' ''Dim blnProcessWasSuccessful As Boolean = True
        '' ''Dim blnRetry As Boolean = True

        '' ''Dim tblDatatable As New DataTable           ' result table

        ' '' ''  Do Until Not blnRetry
        '' ''Dim g_strConnectionStringMYSQL As New MySql.Data.MySqlClient.MySqlConnection

        '' ''g_strConnectionStringMYSQL = New MySql.Data.MySqlClient.MySqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ToString)

        '' ''Dim cmdSQL As New MySql.Data.MySqlClient.MySqlCommand
        '' ''cmdSQL.CommandText = strSQL
        '' ''cmdSQL.Connection = g_strConnectionStringMYSQL

        '' ''Dim tblTableAdapter As New Global.MySql.Data.MySqlClient.MySqlDataAdapter
        '' ''tblTableAdapter.SelectCommand = cmdSQL
        ' '' '' fill table 
        '' ''Dim strErrorCode As String = ""
        ' '' ''For i = 1 To g_intNumberOfRetriesToAccessDatabase
        '' ''Try
        '' ''    ' fill table 
        '' ''    tblTableAdapter.Fill(tblDatatable)
        '' ''    ReturnCode = blnProcessWasSuccessful
        '' ''    Return tblDatatable
        '' ''    '  Exit For
        '' ''Catch ex As Exception
        '' ''    '' ''Select Case Microsoft.VisualBasic.Left(ex.Message, 5)
        '' ''    '' ''    ' case 1 = timeout -- no connection friendly message -- retry for number of retries specified
        '' ''    '' ''    ' case 2 = syntax assumed - friendly message contact programmer
        '' ''    '' ''    Case "Unkno"
        '' ''    '' ''        ' probably a SQL syntax error
        '' ''    '' ''    Case Else
        '' ''    '' ''        ' assume that the service is down, give time to repair and try again
        '' ''    '' ''        Call g_SYS_Wait(30, "MySQL: Reconnecting to database.  Attempt # " & i & " of " & _
        '' ''    '' ''                        g_intNumberOfRetriesToAccessDatabase & ".  Retry in 30 seconds.")
        '' ''    '' ''End Select

        '' ''    blnProcessWasSuccessful = False
        '' ''    strErrorCode = ex.Message
        '' ''End Try
        ' '' ''Next
        ' '' ''If blnProcessWasSuccessful Then
        ' '' ''Else
        ' '' ''    If ReturnCode Then
        ' '' ''        ReturnCode = False
        ' '' ''        blnRetry = False
        ' '' ''        '' ''Else
        ' '' ''        '' ''    MsgBox("Error processing to MySQL." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & strSQL & _
        ' '' ''        '' ''           Chr(13) & Chr(10) & Chr(13) & Chr(10) & " -- " & _
        ' '' ''        '' ''           strErrorCode, MsgBoxStyle.Critical, "MySQL Database Communication")
        ' '' ''    End If
        ' '' ''End If
        ' '' ''Loop

        Return Nothing

    End Function

    Private Function IO_Execute_MSSQL(ByVal strSQL As String, ByRef ReturnCode As Boolean) As DataTable

        'System.Web.HttpContext.Current.Response.Write(ConfigurationManager.ConnectionStrings("ConnectionString").ToString & "<br>")

        Try

            '            System.Web.HttpContext.Current.Response.Write("IO_Execute: 1<br>")
            Dim adapt As New SqlClient.SqlDataAdapter

            Dim blnProcessWasSuccessful As Boolean = True
            Dim blnRetry As Boolean = True
            Dim tblDatatable As New DataTable
            '           System.Web.HttpContext.Current.Response.Write("IO_Execute: 2<br>")

            Dim g_strConnectionStringMSSQL As New System.Data.SqlClient.SqlConnection

            ' System.Web.HttpContext.Current.Response.Write("IO_Execute: 3<br>")
            g_strConnectionStringMSSQL = New System.Data.SqlClient.SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ToString)

            'System.Web.HttpContext.Current.Response.Write("IO_Execute: 4<br>")

            Dim command As New SqlClient.SqlCommand(strSQL, g_strConnectionStringMSSQL)
            Debug.WriteLine(strSQL)
            'System.Web.HttpContext.Current.Response.Write("IO_Execute: 5<br>")

            command.CommandTimeout = 600
            adapt.SelectCommand = command

            Dim strErrorCode As String = ""
            For i = 1 To g_intNumberOfRetriesToAccessDatabase
                Debug.WriteLine(strSQL)
                adapt.Fill(tblDatatable)
                'System.Web.HttpContext.Current.Response.Write("IO_Execute: 6<br>")

                blnRetry = False
                ReturnCode = True
                Return tblDatatable
                Exit For
            Next
        Catch ex As Exception
            ReturnCode = False
            Debug.WriteLine(ex.Message)
            'System.Web.HttpContext.Current.Response.Write("IO_Execute: 7<br>")
            System.Web.HttpContext.Current.Session("SystemErrorMessage") = ex.Message

            'System.Web.HttpContext.Current.Response.Write(ConfigurationManager.ConnectionStrings("ConnectionString").ToString & "<br>")
            'System.Web.HttpContext.Current.Response.Write(ex.Message & "<br>")
        End Try
        Return Nothing
    End Function
    Public Function g_getTableColumnInfo(ByRef TableName As String) As DataTable
        Dim tblColumns As New DataTable
        tblColumns.Columns.Add("FieldName", GetType(System.String))
        tblColumns.Columns.Add("Type", GetType(System.String))
        tblColumns.Columns.Add("AllowNull", GetType(System.String))
        tblColumns.Columns.Add("Default", GetType(System.String))
        tblColumns.Columns.Add("Index", GetType(System.String))
        tblColumns.Columns.Add("AutoInc", GetType(System.String))
        Dim rowColumns As DataRow = Nothing

        If g_ConnectionToUse = "MYSQL" Then
            Dim tblColumnInfo As DataTable = IO_Execute_MYSQL("Show Columns from " & TableName, False)
            For Each rowFieldInfoExtracted In tblColumnInfo.Rows
                rowColumns = tblColumns.NewRow
                rowColumns("FieldName") = rowFieldInfoExtracted("Field")
                rowColumns("Type") = getSQLDataType(rowFieldInfoExtracted("Type"))
                rowColumns("AllowNull") = IIf(rowFieldInfoExtracted("Null") = "YES", "1", "0")
                rowColumns("Default") = IIf(IsDBNull(rowFieldInfoExtracted("Default")), "", rowFieldInfoExtracted("Default"))
                rowColumns("Index") = IIf(rowFieldInfoExtracted("Key") = "", "0", "1")
                rowColumns("AutoInc") = IIf(rowFieldInfoExtracted("Extra") = "auto_increment", "1", "0")
                tblColumns.Rows.Add(rowColumns)
            Next
        Else
            'MSSQL
            Dim tblColumnInfo As DataTable = IO_Execute_MSSQL("Select Column_Name, Data_Type, Is_Nullable, " &
                                                  "Column_Default from Information_Schema.columns where " &
                                                  "Table_Catalog = '" & g_ConnectionSchema & "' and Table_Name ='" & TableName & "'",
                                                  False)
            If g_blnAbort Then
                Return Nothing
            End If
            Dim tblColumnKeyInfo As DataTable = IO_Execute_MSSQL("Select Column_Name from Information_Schema.KEY_COLUMN_USAGE where " &
                                                    "Table_Catalog = '" & g_ConnectionSchema & "' and Table_Name ='" & TableName & "'",
                                                    False)
            If g_blnAbort Then
                Return Nothing
            End If
            ' 12/27/2011.cpb.
            Dim tblIsIdentity As DataTable = IO_Execute_MSSQL("SELECT c2.Name,c2.Is_Identity FROM sysobjects c1 inner join sys.columns c2 on c1.id = c2.object_id where c1.name like '" &
                                                              TableName & "' AND c2.Is_Identity = 1", False)
            If g_blnAbort Then
                Return Nothing
            End If

            For Each rowFieldInfoExtracted In tblColumnInfo.Rows
                rowColumns = tblColumns.NewRow
                rowColumns("FieldName") = rowFieldInfoExtracted("Column_Name")
                rowColumns("Type") = getSQLDataType(rowFieldInfoExtracted("Data_Type"))
                rowColumns("AllowNull") = IIf(rowFieldInfoExtracted("Is_Nullable") = "YES", "1", "0")
                rowColumns("Default") = IIf(IsDBNull(rowFieldInfoExtracted("Column_Default")), "",
                                            rowFieldInfoExtracted("Column_Default"))
                For Each rowKeyFieldExtracted In tblColumnKeyInfo.Rows
                    rowColumns("Index") = "0"     ' default to false
                    If rowColumns("FieldName") = rowKeyFieldExtracted("Column_Name") Then
                        rowColumns("Index") = "1"
                    End If
                Next
                '12/28/2011.cpb.set autoinc based on auto inc field from table..mssql can only have 1 auto inc field per table

                If tblIsIdentity.Rows.Count > 0 AndAlso rowFieldInfoExtracted("Column_Name") = tblIsIdentity(0)("Name") Then
                    rowColumns("AutoInc") = 1
                Else
                    rowColumns("AutoInc") = 0
                End If
                tblColumns.Rows.Add(rowColumns)
            Next
        End If

        Return tblColumns

    End Function
    Private Function getSQLDataType(ByVal DataType As String) As String
        Dim strString As String = "CHAR,TEXT,BLOB"
        Dim strNumber As String = "INT,DECI,DOUB,FLOA,BINA,MONEY,IDENT"
        Dim strDATE As String = "DATE"
        Dim strTime As String = "TIME"
        Dim strBoolean As String = "TINYINT,BIT"
        Dim strTypeDetermined As String = ""
        Do
            DataType = UCase(DataType)
            Try : DataType = Microsoft.VisualBasic.Left(DataType, DataType.IndexOf("(")) : Catch : End Try
            For Each strType As String In Split(strBoolean, ",")
                If DataType.Contains(strType) Then
                    strTypeDetermined = "BOO"
                    Exit Do
                End If
            Next
            For Each strType As String In Split(strString, ",")

                If DataType.Contains(strType) Then
                    strTypeDetermined = "STR"
                    Exit Do
                End If
            Next
            For Each strType As String In Split(strNumber, ",")
                If DataType.Contains(strType) Then
                    strTypeDetermined = "NUM"
                    Exit Do
                End If
            Next
            For Each strType As String In Split(strDATE, ",")
                If DataType.Contains(strType) Then
                    strTypeDetermined = "DAT"
                    Exit Do
                End If
            Next
            For Each strType As String In Split(strTime, ",")
                If DataType.Contains(strType) Then
                    strTypeDetermined = "TIM"
                    Exit Do
                End If
            Next
            strTypeDetermined = "STR"
            Exit Do
        Loop
        Return strTypeDetermined

    End Function

    Private Function getFieldAttributeValue(ByVal TableName As String, ByVal FieldName As String,
                                ByVal AttributeName As String) As String

        Call EvaluateTableEntries(TableName)

        If m_nvcFields(m_nvcTables(TableName))(FieldName) Is Nothing Then
            ' MsgBox("Invalid field name in request", MsgBoxStyle.Information, "Table Field Attribute Evaluation")
            Return "UKN"
        Else
            If m_nvcFieldAttributes(m_nvcFields(m_nvcTables(TableName))(FieldName))(AttributeName) Is Nothing Then
                MsgBox("Invalid attribuate name in request", MsgBoxStyle.Information, "Table Field Attribute Evaluation")
                Return Nothing
            Else
                Return m_nvcFieldAttributes(m_nvcFields(m_nvcTables(TableName))(FieldName))(AttributeName)
            End If
        End If
    End Function
    Private Sub EvaluateTableEntries(ByVal TableName As String)

        If m_nvcTables(TableName) Is Nothing Then

            Dim tblColumns = g_getTableColumnInfo(TableName)  ' MySQL or MSSQL

            Dim intFieldsIndex As Integer = 0
            Dim intFieldAttributesIndex As Integer = 0

            Try : intFieldsIndex = m_nvcFields.Count : Catch : End Try
            ReDim Preserve m_nvcFields(intFieldsIndex)
            m_nvcFields(intFieldsIndex) = New NameValueCollection
            m_nvcTables(TableName) = intFieldsIndex
            For Each rowColumn In tblColumns.Rows
                Try : intFieldAttributesIndex = m_nvcFieldAttributes.Count : Catch : End Try
                m_nvcFields(intFieldsIndex)(UCase(rowColumn("FieldName"))) = intFieldAttributesIndex
                ReDim Preserve m_nvcFieldAttributes(intFieldAttributesIndex)
                m_nvcFieldAttributes(intFieldAttributesIndex) = New NameValueCollection
                m_nvcFieldAttributes(intFieldAttributesIndex)("TYPE") = rowColumn("TYPE")
                m_nvcFieldAttributes(intFieldAttributesIndex)("NULL") = rowColumn("AllowNull")
                m_nvcFieldAttributes(intFieldAttributesIndex)("DEFAULT") = rowColumn("Default")
                m_nvcFieldAttributes(intFieldAttributesIndex)("INDEX") = IIf(IsDBNull(rowColumn("Index")), 0, rowColumn("Index"))
                m_nvcFieldAttributes(intFieldAttributesIndex)("AUTOINC") = rowColumn("AutoInc")
            Next

        End If
    End Sub
    '**********************************Maintain Tables Referenced from DataBase**********************************
    ' 
    ' returns a list of fields for the requested table as a name value collection
    Private Function getTableFieldList(ByVal TableName As String) As NameValueCollection
        Call EvaluateTableEntries(TableName)
        If m_nvcFields(m_nvcTables(TableName)) Is Nothing Then
            MsgBox("Invalid table name in request", MsgBoxStyle.Information, "Table List Evaluation")
        End If
        Return m_nvcFields(m_nvcTables(TableName))
    End Function

    '**********************************SQL Insert**********************************
    ' data inserts via nvc - 4 overloads
    Public Sub g_IO_SQLInsert(ByRef TableName As String, ByRef nvcFieldValues As NameValueCollection,
                              ByVal FormName As String)
        Call IO_Execute_SQLInsert(TableName, nvcFieldValues, FormName, False, True)
    End Sub
    Public Sub g_IO_SQLInsert(ByRef TableName As String, ByRef nvcFieldValues As NameValueCollection,
                              ByVal FormName As String, ByRef blnReturnError As Boolean)
        Call IO_Execute_SQLInsert(TableName, nvcFieldValues, FormName, blnReturnError, True)
    End Sub
    Public Sub g_IO_SQLInsert(ByRef TableName As String, ByRef nvcFieldValues As NameValueCollection,
                              ByVal FormName As String, ByRef blnReturnError As Boolean, ByRef blnAudit As Boolean)
        Call IO_Execute_SQLInsert(TableName, nvcFieldValues, FormName, blnReturnError, blnAudit)
    End Sub
    Public Sub g_IO_SQLInsert(ByRef TableName As String, ByRef TableDataRow As DataRow, ByVal FormName As String,
                              ByRef blnReturnError As Boolean, ByRef blnAudit As Boolean)
        Dim nvcnvcFieldValues As New NameValueCollection

        ' extract all the columns in this table from the database (just read one record and extract the info from the datatable)
        Dim tblDataTable As DataTable

        If g_ConnectionToUse = "MYSQL" Then
            tblDataTable = g_IO_Execute_SQL("Select * from " & TableName & " Limit 1", False)
        Else
            tblDataTable = g_IO_Execute_SQL("select TOP 1 * from " & TableName, False)
        End If

        For Each colTableColumn As DataColumn In tblDataTable.Columns
            ' loop on column list retrieved from database
            Dim strColumnName As String = colTableColumn.ColumnName
            Try
                nvcnvcFieldValues(strColumnName) = TableDataRow(strColumnName)
            Catch ex As Exception
            End Try
        Next

        If nvcnvcFieldValues.Count > 0 Then
            Call IO_Execute_SQLInsert(TableName, nvcnvcFieldValues, FormName, blnReturnError, blnAudit)
        End If

    End Sub
    '
    '**********************************SQL Delete**********************************
    Public Sub g_IO_SQLDelete(ByRef TableName As String, ByVal WherePhrase As String)
        Call IO_Execute_SQLDelete(TableName, WherePhrase, Nothing, False, True)
    End Sub

    Public Sub g_IO_SQLDelete(ByRef TableName As String, ByVal WherePhrase As String,
                             ByVal FormName As String, ByRef blnReturnError As Boolean)
        Call IO_Execute_SQLDelete(TableName, WherePhrase, FormName, blnReturnError, True)
    End Sub

    Public Sub g_IO_SQLDelete(ByRef TableName As String, ByVal WherePhrase As String,
                              ByVal FormName As String, ByRef blnReturnError As Boolean, ByRef blnAudit As Boolean)
        IO_Execute_SQLDelete(TableName, WherePhrase, FormName, blnReturnError, blnAudit)
    End Sub


    Private Function IO_Execute_SQLInsert(ByRef TableName As String, ByRef nvcFieldValues As NameValueCollection,
                                      ByVal FormName As String, ByRef ReturnACompletionCode As Boolean,
                                      ByVal AuditInsert As Boolean) As Boolean
        ' this routine will receive from the programmer a list of columns in a table with associated values 
        ' to be inserted into a table (see parameter list above)

        Dim blnRetry As Boolean = True

        ' define audit variables to be built during insert
        Dim nvcAuditKeyValues As New NameValueCollection

        Dim m_nvcConvertedInserts As New NameValueCollection '  take Column collection array from programmer and make all keys capital

        ' make all keys upper case
        For Each strcolumn As String In nvcFieldValues.AllKeys
            If IsDBNull(nvcFieldValues(strcolumn)) Or IsNothing(nvcFieldValues(strcolumn)) Then
                nvcFieldValues.Remove(strcolumn)
            Else
                m_nvcConvertedInserts(UCase(strcolumn)) = nvcFieldValues(strcolumn)
            End If
        Next

        ' start to build INSERT SQL command
        Dim strInsertSQL As String = ""
        Dim strDelimiter As String = " "

        Do Until Not blnRetry

            blnRetry = False

            strInsertSQL = "Insert into " & TableName & " "

            Dim strFields As String = " ("              ' list of columns sent by programmer
            Dim strValues As String = " VALUES ("       ' list of associated values sent by programmer

            ' column info from memory variables
            Dim nvcFields As NameValueCollection
            nvcFields = getTableFieldList(TableName)

            ' loop through column names extract from database and build INSERT (also check to see if programmer provided bogus columns)
            ' any columns extracted not listed by programmer will be defaulted to an appropriate value
            Dim strFieldsDelimiter As String = ""
            Dim strValuesDelimiter As String = ""
            Dim strValue As String = ""

            For Each strColumnName As String In nvcFields.Keys
                Dim strColumnType As String = getFieldAttributeValue(TableName, strColumnName, "TYPE")
                Dim strColumnKey As String = getFieldAttributeValue(TableName, strColumnName, "INDEX")
                ' determine if field was supplied by programmer
                If m_nvcConvertedInserts(strColumnName) Is Nothing Then

                    ' was not in list of fields supplied by programmer - get default value
                    Dim strColumnDefault As String = getFieldAttributeValue(TableName, strColumnName, "DEFAULT")
                    Dim strColumnNull As String = getFieldAttributeValue(TableName, strColumnName, "NULL")
                    Dim strColumnAutoInc As String = getFieldAttributeValue(TableName, strColumnName, "AUTOINC")
                    If strColumnDefault <> "" Or strColumnNull = "1" Or strColumnAutoInc = "1" Then
                    Else
                        ' column in data file not supplied by programer insert data
                        ' does not have default and can not be null -- fill it for programmer
                        If g_Debug Then
                            MsgBox("Programmer You Did Not supply a field that can not be left null and does not have default value.", MsgBoxStyle.Exclamation, "Programmer-NOTICE")
                        Else
                            Select Case UCase(strColumnType)
                                Case "STR"
                                    ' this is a string (I hope)
                                    strValues &= strValuesDelimiter & "''"
                                    strValue = ""
                                Case "DAT"
                                    ' this is a date
                                    strValues &= strValuesDelimiter & "'1970-01-01'"
                                    strValue = "1970-01-01"
                                Case "TIM"
                                    ' this is a date
                                    strValues &= strValuesDelimiter & "'00:00:00'"
                                    strValue = "00:00:00"
                                Case "BOO"
                                    ' this is a boolean
                                    strValues &= strValuesDelimiter & False
                                    strValue = "0"
                                Case "NUM"
                                    ' this is a number so don't include quotes
                                    strValues &= strValuesDelimiter & 0
                                    strValue = "0"
                                Case Else
                                    ' this is a string (I hope)
                                    strValues &= strValuesDelimiter & "''"
                                    strValue = ""
                            End Select
                            nvcFieldValues(strColumnName) = strValue
                            If strColumnKey = "1" Then
                                nvcAuditKeyValues(strColumnName) = strValue
                            End If
                        End If
                        ' strFields &= strFieldsDelimiter & strColumnName
                        'strFieldsDelimiter = ", "
                        'strValuesDelimiter = ", "
                    End If
                Else
                    Try
                        ' create column list and values list to be assembled later into Insert SQL statement
                        '       if column being extracted from table is not in programmer's list then will error to catch
                        strValue = m_nvcConvertedInserts(strColumnName)
                        Select Case strColumnType
                            Case "STR"
                                ' STRING TYPE
                                strValues &= strValuesDelimiter & "'" & strValue.Replace("'", "''") & "'"
                            Case "DAT"
                                If IsDate(strValue) Then
                                    ' this is a date
                                    Dim dteDate As DateTime = strValue
                                    If strValue.IndexOf(":") > -1 Then
                                        strValue = Format(dteDate, "yyyy-MM-dd HH:mm:ss")
                                        strValues &= strValuesDelimiter & "'" & strValue & "'"
                                    Else
                                        strValue = Format(dteDate, "yyyy-MM-dd")
                                        strValues &= strValuesDelimiter & "'" & strValue & "'"
                                    End If
                                Else
                                    strValue = "null"
                                    strValues &= strValuesDelimiter & strValue
                                End If
                                nvcFieldValues(strColumnName) = strValue   ' convert date format for future use
                            Case "TIM"
                                ' this is a date
                                strValues &= strValuesDelimiter & "'" & strValue & "'"
                            Case "BOO"
                                ' this is a boolean
                                strValues &= strValuesDelimiter & strValue
                            Case "NUM"
                                If IsNumeric(strValue) Then
                                    strValues &= strValuesDelimiter & strValue.Replace(",", "")
                                Else
                                    strValues &= strValuesDelimiter & "0"
                                End If
                            Case Else
                                ' Don't know what this is but assume it needs quotes
                                strValues &= strValuesDelimiter & "'" & strValue.Replace("'", "''") & "'"
                        End Select
                        strValuesDelimiter = ", "
                        strFields &= strFieldsDelimiter & strColumnName
                        strFieldsDelimiter = ", "
                        If strColumnKey = "1" Then
                            nvcAuditKeyValues(strColumnName) = strValue
                        End If

                        ' remove from user supplied list to indicate that this entry is handled
                        m_nvcConvertedInserts.Remove(UCase(strColumnName))  ' any columns left in this array at end of process will cause an error

                    Catch ex As Exception
                        MsgBox("Dear Programmer -- You have data problems -- Invalid Data Sent", MsgBoxStyle.Critical, "OOPS!")
                    End Try

                    'strFields &= strFieldsDelimiter & strColumnName
                    'strFieldsDelimiter = ", "
                    'strValuesDelimiter = ", "

                    ' remove from user supplied list to indicate that this entry is handled (since it is empty just let FoxPro initialize it)
                    m_nvcConvertedInserts.Remove(UCase(strColumnName))  ' any columns left in this array at end of process will cause an error
                End If

            Next
            strInsertSQL &= strFields & ")" & strValues & ")"

            ' were there any programmer supplied columns not used?
            If m_nvcConvertedInserts.Count = 0 Then
                'array is empty so all columns used
                blnRetry = False
            Else
                Dim strUnusedColumnsProvided As String = ""
                strDelimiter = ""
                For Each strcolumn As String In m_nvcConvertedInserts.AllKeys
                    m_nvcConvertedInserts(UCase(strcolumn)) = nvcFieldValues(strcolumn)
                    strUnusedColumnsProvided &= strDelimiter & strcolumn
                    strDelimiter = ", "
                Next
                MsgBox("The following fields were provided but don't seem to exist in " & TableName & " (Could be database communication error): " & Chr(13) & Chr(10) & Chr(13) & Chr(10) & strUnusedColumnsProvided, MsgBoxStyle.RetryCancel, "Unused Input")
            End If
        Loop

        ' write out to SQL 
        Dim LastRecId As Integer = Nothing

        Try

            If g_ConnectionToUse = "MYSQL" Then
                g_IO_Execute_SQL(strInsertSQL, ReturnACompletionCode)
            Else
                ' MSSQL must retrieve the new RECID now, 12/23/11 cpb
                strInsertSQL &= ";Select @@IDENTITY as ID;"
                Dim tblTemp As DataTable = g_IO_Execute_SQL(strInsertSQL, ReturnACompletionCode)
                LastRecId = tblTemp.Rows(0)("ID")
            End If

            'g_IO_Execute_SQL(strInsertSQL, ReturnACompletionCode)
            ' got to get command into data table to get the recid out back out


        Catch ex As Exception
            MsgBox("Error attempting to write: " & TableName & ".", MsgBoxStyle.Information, "Unable To Write To Table")
        End Try

        ' EVERY TABLE MUST HAVE RECID
        ' get record back to write the recid for audit data
        Dim strSelectSQL As String = ""
        If g_ConnectionToUse = "MYSQL" Then
            strSelectSQL = "Select Last_Insert_Id() as RecId from " & TableName
            LastRecId = g_IO_Execute_SQL(strSelectSQL, False).Rows(0)("RecId")
        Else
            ' 12/23/11 cpb handled above
            ''''strSelectSQL = "SELECT IDENT_CURRENT('" & TableName & "') as RecId"
            '''''  strSelectSQL = "SELECT SCOPE_IDENTITY() as RecId" --- '''THIS RETURNS NULL...BUT THE ABOVE IS NOT LIMITED TO OWN SESSION???NOT SURE ABOUT THIS???
        End If

        'Dim LastRecId As Integer = g_IO_Execute_SQL(strSelectSQL).Rows(0)("RecId")
        nvcAuditKeyValues("RecId") = LastRecId

        ' audit the insert
        'If AuditInsert And ReturnACompletionCode Then
        '    Audit_Data(TableName, FormName, nvcFieldValues, nvcAuditKeyValues, "", "I")
        'End If
        Return ReturnACompletionCode

    End Function
    '_____ sql delete functions 
    Private Function IO_Execute_SQLDelete(ByRef TableName As String, ByVal WherePhrase As String,
                                         ByVal FormName As String, ByRef ReturnACompletionCode As Boolean,
                                         ByVal AuditDelete As Boolean) As Boolean

        If AuditDelete Then
            ' get column info from memory variables to build list of keys for audit write
            Dim nvcFields As NameValueCollection
            nvcFields = getTableFieldList(TableName)

            ' build table of records to be deleted
            Dim tblLiveDatabaseRecordsBeingDeleted As DataTable = Nothing
            Dim strSQLToReadRecordsBeingDeleted As String = "SELECT * FROM " & TableName &
                            IIf(Trim(WherePhrase) = "", "", " WHERE " & WherePhrase)
            tblLiveDatabaseRecordsBeingDeleted = g_IO_Execute_SQL(strSQLToReadRecordsBeingDeleted, False)

            ' write out audit entry for each data row being deleted
            For Each rowLiveDatabaseRecordsBeingDeleted As DataRow In tblLiveDatabaseRecordsBeingDeleted.Rows

                ' get key and data values for deleted keys
                Dim nvcAuditKeyValues As New NameValueCollection
                Dim nvcAuditFieldValues As New NameValueCollection
                For Each field In nvcFields.Keys
                    If IsDBNull(rowLiveDatabaseRecordsBeingDeleted(field)) Then

                    Else

                        nvcAuditFieldValues(field) = rowLiveDatabaseRecordsBeingDeleted(field)
                        Dim strColumnKey As String = getFieldAttributeValue(TableName, field, "INDEX")
                        If strColumnKey = "1" Then
                            nvcAuditKeyValues(field) = rowLiveDatabaseRecordsBeingDeleted(field)
                        End If
                    End If
                Next

                'Audit_Data(TableName, FormName, nvcAuditFieldValues, nvcAuditKeyValues, _
                '           IIf(Trim(WherePhrase) = "", "", WherePhrase), "D")
            Next
        End If

        ' perform delete
        Dim strSQLToDeleteRecords As String = "Delete FROM " & TableName &
                        IIf(Trim(WherePhrase) = "", "", " WHERE " & WherePhrase)
        g_IO_Execute_SQL(strSQLToDeleteRecords, ReturnACompletionCode)


        Return ReturnACompletionCode

    End Function


End Module
