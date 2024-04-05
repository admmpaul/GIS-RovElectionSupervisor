'///////////////////////////////////////////////////////////////////////////////////
'//
'//                         COPYRIGHT [2007-2008] Weston Solutions INC.
'//                           ALL RIGHTS RESERVED.
'//
'//                    WESTON SOLUTIONS CONFIDENTIAL PROPRIETARY
'//
'// This file and the information contained in it is Weston Solutions Proprietary 
'// Confidential, and shall not be used, or published, or disclosed, or 
'// disseminated outside of Weston Solutions in whole or in part without Weston’s 
'// consent. This document contains trade secrets of Weston Solutions. Reverse 
'// engineering of any or all of the information in this document is prohibited.  
'// This copyright notice does not imply publication of this document.
'//
'///////////////////////////////////////////////////////////////////////////////////
'//
'//   FILE NAME		    : LoadXLSData.vb
'//   Namespace		    : ElectionSupervisor
'//   CLASS NAME(S)	    : LoadXLSData
'//   ORIGINATOR		: Keith Palmer (WESTON SOLUTIONS INC.)
'//   DATE OF ORIGIN	: 02/07/08
'//   Current Version #	: 1.0
'//
'//   Purpose           : Class logic for importing Excel spreadsheet and upload to DB.
'//
'///////////////////////////////////////////////////////////////////////////////////

Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Text
Imports System.IO

Public Class LoadXLSData

    Dim sSqlConn As String
    Dim sDBPrefix As String
    Dim connection As SqlConnection
    Dim command As SqlCommand

    Public Sub SetAttributes(ByVal sConnStr As String, ByVal sDBPre As String)
        sSqlConn = sConnStr
        sDBPrefix = sDBPre
    End Sub

    Public Function LoadR700_03Data(ByVal excelfilename As String) As Integer

        Dim SQL As String = ""
        LoadR700_03Data = 1

        'Check for existing table.  If it exists then delete and create.
        SQL = "CREATE TABLE " & sDBPrefix & "R700_03 (PRECINCT Int, VBM Int, BALLOTTYPE Int, REGISTEREDTOTAL Int, ABSENTEETOTAL Int);"
        CreateTable(sDBPrefix & "R700_03", SQL)

        'Load in the spreadsheet.
        Dim dtTest As DataTable = ReadDataFromExcel(excelfilename)

        If (dtTest Is Nothing) Then Return 0

        Dim dcTemp As DataColumn = New DataColumn

        Dim iPrecinct As Integer = -1
        Dim bAddAV As Boolean = False
        Dim iBallotType As Integer
        Dim iTotalVoters As Integer = 0
        Dim iAbsentVoters As Integer = 0
        Dim iLastCol As Integer = 0
        Dim sbSQL As StringBuilder = New StringBuilder

        Dim iColCnt As Integer = dtTest.Columns.Count - 1
        Dim myRow As DataRow
        Dim iPrecinctNum As Integer
        Dim iIsVBM As Integer = 0
        Dim iRowCnt As Integer = 0

        For Each myRow In dtTest.Rows
            iRowCnt = iRowCnt + 1

            'Load in the precinct number and ballot type.
            If SafeDBNull(myRow.Item(0)) <> "" And SafeDBNull(myRow.Item(2)) <> "" Then
                'KEP - Vote by mail precincts now have a "9" in front. 
                'If IsNumeric(myRow.Item(0)) Or Left(myRow.Item(0), 1) = "V" Then
                If IsNumeric(myRow.Item(0)) Or Left(SafeDBNull(myRow.Item(6)), 1) = "Y" Then

                    If (iPrecinctNum <> myRow.Item(3) And iPrecinctNum > -1) Then
                        sbSQL.Append("INSERT INTO " & sDBPrefix & "R700_03 VALUES (" & iPrecinctNum & ", " & iIsVBM & ", " & iBallotType & ", " & iTotalVoters & ", " & iAbsentVoters & "); ")
                        iPrecinctNum = 0
                        iIsVBM = 0
                        iBallotType = 0
                        iTotalVoters = 0
                        iAbsentVoters = 0
                        iLastCol = 0
                    End If

                    'KEP - Vote by mail precincts now identify by a "Y" in item 6. 
                    'If Left(myRow.Item(0), 1) = "V" Or Len(myRow.Item(0)) > 6 Then iIsVBM = 1
                    If Left(SafeDBNull(myRow.Item(6)), 1) = "Y" Then iIsVBM = 1

                    iPrecinctNum = CType(myRow.Item(3), Integer)
                    iBallotType = CType(myRow.Item(2), Integer)
                End If
            End If

            If (iRowCnt > (dtTest.Rows.Count - 14) And iPrecinctNum > 0) Then
                sbSQL.Append("INSERT INTO " & sDBPrefix & "R700_03 VALUES (" & iPrecinctNum & ", " & iIsVBM & ", " & iBallotType & ", " & iTotalVoters & ", " & iAbsentVoters & "); ")
                iPrecinctNum = 0
                iIsVBM = 0
                iBallotType = 0
                iTotalVoters = 0
                iAbsentVoters = 0
                iLastCol = 0
            End If

            'If the iTotalVoters value has been populated then the next line in the xls will be for the absentee voters.
            If iTotalVoters > 0 And bAddAV Then
                iAbsentVoters = CType(myRow.Item(iLastCol), Integer) + iAbsentVoters
                bAddAV = False
            End If

            If UCase(SafeDBNull(myRow.Item(1))) = "TOTAL" And iPrecinctNum > 0 Then
                Dim i As Integer
                For i = 0 To iColCnt
                    If UCase(SafeDBNull(myRow.Item(i))) = "" Then
                        iLastCol = i - 1
                        Exit For
                    End If
                Next
                If iLastCol = 0 Then iLastCol = iColCnt
                iTotalVoters = CType(myRow.Item(iLastCol), Integer) + iTotalVoters
                bAddAV = True
            End If
        Next myRow

        If sbSQL.Length > 0 Then InsertRecs(sbSQL)

    End Function

    Public Function LoadR700_09Data(ByVal excelfilename As String, ByVal sOverwrite As String, ByVal sElectionName As String) As Integer

        Dim SQL As String = ""
        LoadR700_09Data = 1

        If sOverwrite = "OVERWRITE" Then
            'Check for existing table.  If it exists then delete and create.
            SQL = "CREATE TABLE " & sDBPrefix & "R700_09 (PRECINCT Int, DISSOLVE VARCHAR(10), BALLOTTYPE Int, REGULARVOTERTOTAL Int, ABSENTEEVOTERTOTAL Int, REGISTEREDTOTAL Int, ELECTIONNAME VARCHAR(50), DATESTAMP DateTime, USERNAME VARCHAR(50)); CREATE NONCLUSTERED INDEX IDXPrecinct ON " & sDBPrefix & "R700_09 (PRECINCT ASC);"
            CreateTable(sDBPrefix & "R700_09", SQL)
        End If

        'Load in the spreadsheet.
        Dim dtTest As DataTable = ReadDataFromExcel(excelfilename)

        If (dtTest Is Nothing) Then Return 0

        Dim dcTemp As DataColumn = New DataColumn

        Dim iPrecinct As Integer = -1
        Dim iBallotType As Integer = 0
        Dim iRegularVoters As Integer = 0
        Dim iAbsentVoters As Integer = 0
        Dim iTotalVoters As Integer = 0
        Dim sbSQL As StringBuilder = New StringBuilder

        Dim iColCnt As Integer = dtTest.Columns.Count - 1
        Dim myRow As DataRow        

        Try            
            For Each myRow In dtTest.Rows
                'Load in the precinct number and ballot type and skip the total count lines.
                If SafeDBNull(myRow.Item(0)) <> "" And SafeDBNull(myRow.Item(1)) <> "" And SafeDBNull(myRow.Item(14)) <> "" Then
                    iPrecinct = CType(myRow.Item(0), Integer)
                    iBallotType = CType(myRow.Item(1), Integer)
                    iRegularVoters = CType(myRow.Item(12), Integer)
                    iAbsentVoters = CType(myRow.Item(13), Integer)
                    iTotalVoters = CType(myRow.Item(14), Integer)
                    sbSQL.Append("INSERT INTO " & sDBPrefix & "R700_09 VALUES (" & iPrecinct & ", NULL, " & iBallotType & ", " & iRegularVoters & ", " & iAbsentVoters & ", " & iTotalVoters & ", '" & sElectionName & "', NULL, NULL); ")
                    iPrecinct = 0
                    iBallotType = 0
                    iTotalVoters = 0
                    iAbsentVoters = 0
                End If
            Next myRow
        Catch ex As Exception
            MsgBox("Problem loading the R700_09 data: " & ex.Message)
        End Try

        If sbSQL.Length > 0 Then InsertRecs(sbSQL)

    End Function

    Public Function LoadR701_01Data(ByVal excelfilename As String) As Integer

        Dim SQL As String = ""
        LoadR701_01Data = 1

        'Check for existing table.  If it exists then delete and create.
        SQL = "CREATE TABLE " & sDBPrefix & "R701_01 (CONSOLPRECINCT Int, HOMEPRECINCT Int);"

        CreateTable(sDBPrefix & "R701_01", SQL)

        'Load in the spreadsheet.
        Dim dtTest As DataTable = ReadDataFromExcel(excelfilename)

        If (dtTest Is Nothing) Then Return 0

        Dim dcTemp As DataColumn = New DataColumn

        Dim iConsolPrecinct As Integer = 0
        Dim iHomePrecinct As Integer = 0
        Dim sbSQL As StringBuilder = New StringBuilder

        Dim iColCnt As Integer = dtTest.Columns.Count - 1
        Dim myRow As DataRow
        Dim i As Integer = 0
        For Each myRow In dtTest.Rows
            If i > 2073 Then
                Dim sTest As String = ""
            End If
            i = i + 1
            'Load in the precinct number.
            If SafeDBNull(myRow.Item(0)) <> "" And SafeDBNull(myRow.Item(2)) <> "" And iConsolPrecinct > 0 Then
                If IsNumeric(myRow.Item(0)) Then
                    If iConsolPrecinct <> CType(myRow.Item(0), Integer) Then
                        sbSQL.Append("INSERT INTO " & sDBPrefix & "R701_01 VALUES (" & iConsolPrecinct & ", " & iHomePrecinct & "); ")
                        iConsolPrecinct = CType(myRow.Item(0), Integer)
                        iHomePrecinct = CType(myRow.Item(2), Integer)
                    End If
                End If
            Else
                If SafeDBNull(myRow.Item(2)) <> "" Then
                    If iHomePrecinct <> CType(myRow.Item(2), Integer) And iConsolPrecinct > 0 Then
                        sbSQL.Append("INSERT INTO " & sDBPrefix & "R701_01 VALUES (" & iConsolPrecinct & ", " & iHomePrecinct & "); ")
                        iHomePrecinct = CType(myRow.Item(2), Integer)
                    End If
                Else
                    sbSQL.Append("INSERT INTO " & sDBPrefix & "R701_01 VALUES (" & iConsolPrecinct & ", " & iHomePrecinct & "); ")
                End If
            End If

            'Load in the values for the first row.
            If iConsolPrecinct = 0 Then
                iConsolPrecinct = CType(myRow.Item(0), Integer)
                iHomePrecinct = CType(myRow.Item(2), Integer)
            End If

        Next myRow

        If sbSQL.Length > 0 Then InsertRecs(sbSQL)

    End Function

    Private Function ReadDataFromExcel(ByVal excelfilename As String) As DataTable
        ReadDataFromExcel = Nothing
        Dim dt As New DataTable
        Dim da As OleDbDataAdapter
        Dim conn As OleDbConnection
        Try

            conn = New OleDbConnection( _
                  "provider=Microsoft.Jet.OLEDB.4.0; " & _
                  "data source=" & excelfilename & "; " & _
                  "Extended Properties='Excel 8.0;HDR=NO;IMEX=1'")

            da = New OleDbDataAdapter("SELECT * FROM [Sheet1$]", conn)

            conn.Open()

            da.Fill(dt)

            ReadDataFromExcel = dt

        Catch ex As Exception
            MsgBox("Could not open " & excelfilename & ".  Please enter the file name again.")
        Finally
            da = Nothing
            If conn.State = ConnectionState.Open Then conn.Close()
        End Try

    End Function

    Private Function SafeDBNull(ByVal value As Object) As String
        If IsDBNull(value) Then
            Return String.Empty
        Else
            Return CType(value, String)
        End If
    End Function

    Public Sub InsertRecs(ByVal sbRecs As StringBuilder)

        connection = New SqlConnection
        connection.ConnectionString = sSqlConn

        command = New SqlCommand(sbRecs.ToString, connection)
        command.Connection.Open()

        Try
            command.ExecuteNonQuery()
        Catch ex As SqlException
            MsgBox(ex.Message)
        Finally
            command.Connection.Close()
            connection.Close()
        End Try

    End Sub

    Private Sub CreateTable(ByVal sTableName As String, ByVal sSQL As String)

        Dim SQL As String
        connection = New SqlConnection
        connection.ConnectionString = sSqlConn

        'Drop the table.  Trap exception in case it doesn't exist.
        SQL = "DROP TABLE " & sTableName
        command = New SqlCommand(SQL, connection)
        command.Connection.Open()
        Try
            command.ExecuteNonQuery()
        Catch ex As Exception

        End Try

        'Now create the table.
        command = New SqlCommand(sSQL, connection)
        'command.Connection.Open()
        Try
            command.ExecuteNonQuery()
        Catch ex As SqlException
            MsgBox(ex.Message)
        Finally
            command.Connection.Close()
            connection.Close()
        End Try

    End Sub

    Public Shared Function IsNumeric(ByVal sText As String) As Boolean

        If Double.TryParse(sText, Globalization.NumberStyles.AllowDecimalPoint) Then
            Return True
        Else
            Return False
        End If

    End Function


    Public Shared Function ValueType(ByVal sText As String) As Boolean

        If Double.TryParse(sText, Globalization.NumberStyles.AllowDecimalPoint) Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function LoadPollingPlaceTbl(ByVal txtfilename As String) As Integer


        Dim SQL As String = ""
        'Check for existing table.  If it exists then delete and create.

        Dim sbSQL As StringBuilder = New StringBuilder

        'SQL = "CREATE TABLE " & sDBPrefix & "PollingPlaceInput ( " & _
        '        "poll_election_id int, election_id int, poll_id int, status nvarchar(255), facility_type nvarchar(255), reason_id int, " & _
        '        "division nvarchar(255), users_id int, sup_district int, billing_code nvarchar(255), manual_precinct nvarchar(255), precinct_id int, " & _
        '        "gis_x int, gis_y int, rating int, location_line_1 nvarchar(255), location_line_2 nvarchar(255), directions_cid int, phone_poll nvarchar(15), " & _
        '        "poll_type nvarchar(255), source nvarchar(255), permanent_poll nvarchar(255), by_mail nvarchar(255), useable_areas int, poll_length int,	poll_width int, " & _
        '        "served int, polls_open nvarchar(255), polls_close nvarchar(255), street_id int, house_number int, house_fraction nvarchar(255), pre_dir nvarchar(255), " & _
        '        "street nvarchar(255), type_ nvarchar(255), post_dir nvarchar(255), building_number nvarchar(255), apartment_number nvarchar(255), city nvarchar(255), " & _
        '        "state nvarchar(255), zip int, contact nvarchar(255), voter_id int, care_of nvarchar(255), mail_street nvarchar(255), mail_city nvarchar(255), " & _
        '        "mail_state nvarchar(255), mail_country nvarchar(255), mail_zip nvarchar(255), phone_1 nvarchar(15), phone_2 nvarchar(15), fax nvarchar(15), " & _
        '        "email nvarchar(255), poll_owner nvarchar(255), voter_id_owner int, care_of_owner nvarchar(255), mail_street_owner nvarchar(255), mail_city_owner nvarchar(255), " & _
        '        "mail_state_owner nvarchar(255), mail_country_owner nvarchar(255), mail_zip_owner nvarchar(255), phone_1_owner nvarchar(15), " & _
        '        "phone_2_owner nvarchar(15), fax_owner nvarchar(15), email_owner nvarchar(255), poll_custodian nvarchar(255), voter_id_custodian int, " & _
        '        "care_of_custodian nvarchar(255), mail_street_custodian nvarchar(255), mail_city_custodian nvarchar(255), mail_state_custodian nvarchar(255), " & _
        '        "mail_country_custodian nvarchar(255), mail_zip_custodian nvarchar(255), phone_1_custodian numeric(38, 8), phone_2_custodian numeric(38, 8), " & _
        '        "fax_custodian numeric(38, 8), email_custodian nvarchar(255), contract_date nvarchar(255), contract_date_original nvarchar(255), " & _
        '        "payroll_number nvarchar(255), newspaper int, key_instructions_cid int, miles int, map nvarchar(255), tb_guide_page nvarchar(255), " & _
        '        "tb_guide_suffix nvarchar(255), tb_reference nvarchar(255), depot_id int, agency int, route int, sequence int, return_center int, use_fee numeric(38, 8), " & _
        '        "staff_fee numeric(38, 8), other_fee_1 numeric(38, 8), other_fee_2 numeric(38, 8), other_fee_3 numeric(38, 8), pay_exact numeric(38, 8), " & _
        '        "accessible nvarchar(255), survey_date nvarchar(255), survey_rate nvarchar(255), devices_cid int, parking nvarchar(255), parking_disabled nvarchar(255), " & _
        '        "paths_entrance nvarchar(255), ramps_elevator nvarchar(255), voting_area nvarchar(255), flag_1 nvarchar(255), flag_2 nvarchar(255), flag_3 nvarchar(255), " & _
        '        "flag_4 nvarchar(255), flag_5 nvarchar(255), booths int, tables int, chairs int, light int, signs int, precincts int, voters int, supplies_cid int, " & _
        '        "requested nvarchar(255), returned nvarchar(255), available nvarchar(255), confirmed nvarchar(255), status_delivery nvarchar(255), " & _
        '        "status_returned nvarchar(255), status_problem nvarchar(255), trblshtr_area_id int, comment_id int, poll_election_trans_id int, ltd nvarchar(255), " & _
        '        "timestamp_ nvarchar(255), delivery_date nvarchar(255), web_site nvarchar(255), gen_field nvarchar(255), consolidation_id int, election_id1 int, " & _
        '        "consolidation int, consolidation_name nvarchar(255), serial_number int, serial_chk int, poll_election_id1 int, votes_by_mail nvarchar(255), " & _
        '        "poll_change_cards nvarchar(255), alpha_split nvarchar(255), delivery_status nvarchar(255), receive_status nvarchar(255), comment_id1 int, " & _
        '        "consolidation_trans_id int, ltd1 nvarchar(255), timestamp1 nvarchar(255), division1 int, billing_code1 nvarchar(255), newspaper1 nvarchar(255), " & _
        '        "depot_id1 int, agency1 nvarchar(255), route1 nvarchar(255), sequence1 int, delivery_date1 nvarchar(255), return_center1 nvarchar(255), " & _
        '        "trouble_shooter_id int, phone_assigned nvarchar(255), Addr_Geocode nvarchar(50) NULL);"

        Dim iRtn As Integer = ElectionLayersWindow.DeleteTable("Consol_PollingPlace_Export")

        SQL = "CREATE TABLE " & sDBPrefix & "Consol_PollingPlace_Export ( " & _
        "poll_election_id int, election_id int, poll_id int, precinct_id int, gis_x int, gis_y int," & _
        "location_line_1 nvarchar(255), location_line_2 nvarchar(255), street_id int, house_number int, house_fraction nvarchar(255), pre_dir nvarchar(255), " & _
        "street nvarchar(255), type_ nvarchar(255), post_dir nvarchar(255), building_number nvarchar(255), apartment_number nvarchar(255), city nvarchar(255), " & _
        "state nvarchar(255), zip int, consolidation int, Addr_Geocode nvarchar(50) NULL); CREATE NONCLUSTERED INDEX IDXPollID ON " & sDBPrefix & "Consol_PollingPlace_Export (POLL_ID ASC);"

        CreateTable(sDBPrefix & "Consol_PollingPlace_Export", SQL)

        'Load in the file.
        Dim iCnt As Integer = 0
        Using reader As StreamReader = New StreamReader(txtfilename)
            While reader.Peek <> -1
                Dim sRow() As String = reader.ReadLine.Split(ControlChars.Tab)
                If iCnt > 0 Then

                    'sbSQL.Append("INSERT INTO " & sDBPrefix & "PollingPlaceInput VALUES (" & _
                    '             sRow(0) & ", " & sRow(1) & ", " & sRow(2) & ", '" & sRow(3) & "', '" & sRow(4) & "', " & sRow(5) & _
                    '             ", '" & sRow(6) & "', " & sRow(7) & ", " & sRow(8) & ", '" & sRow(9) & "', '" & sRow(10) & "', " & sRow(11) & _
                    '             ", " & sRow(12) & ", " & sRow(13) & ", " & sRow(14) & ", '" & sRow(15) & "', '" & sRow(16) & "', " & sRow(17) & ", '" & sRow(18) & _
                    '             "', '" & sRow(19) & "', '" & sRow(20) & "', '" & sRow(21) & "', '" & sRow(22) & "', " & sRow(23) & ", " & sRow(24) & ", " & sRow(25) & _
                    '             ", " & sRow(26) & ", '" & sRow(27) & "', '" & sRow(28) & "', " & sRow(29) & ", " & sRow(30) & ", '" & sRow(31) & "', '" & sRow(32) & _
                    '             "', '" & sRow(33) & "', '" & sRow(34) & "', '" & sRow(35) & "', '" & sRow(36) & "', '" & sRow(37) & "', '" & sRow(38) & _
                    '             "', '" & sRow(39) & "', " & sRow(40) & ", '" & sRow(41) & "', " & sRow(42) & ", '" & sRow(43) & "', '" & sRow(44) & "', '" & sRow(45) & _
                    '             "', '" & sRow(46) & "', '" & sRow(47) & "', '" & sRow(48) & "', '" & sRow(49) & "', '" & sRow(50) & "', '" & sRow(51) & _
                    '             "', '" & sRow(52) & "', '" & sRow(53) & "', " & sRow(54) & ", '" & sRow(55) & "', '" & sRow(56) & "', '" & sRow(57) & _
                    '             "', '" & sRow(58) & "', '" & sRow(59) & "', '" & sRow(60) & "', '" & sRow(61) & "'" & _
                    '"); ")
                    'Remove records that don't have an address.
                    If SafeDBNull(sRow(29)) <> "0" Then
                        Dim sSiteAddr As String = Trim(sRow(30)) & " " & Trim(sRow(31))
                        sSiteAddr = Trim(sSiteAddr) & " " & Trim(sRow(32))
                        sSiteAddr = Trim(sSiteAddr) & " " & Trim(sRow(33))
                        sSiteAddr = Trim(sSiteAddr) & " " & Trim(sRow(34))
                        sSiteAddr = Trim(sSiteAddr) & " " & Trim(sRow(35))
                        sSiteAddr = Trim(sSiteAddr) & ", " & Trim(sRow(38))
                        sbSQL.Append("INSERT INTO " & sDBPrefix & "Consol_PollingPlace_Export VALUES (" & _
                                     sRow(0) & ", " & sRow(1) & ", " & sRow(2) & ", " & sRow(11) & ", " & sRow(12) & ", " & sRow(13) & _
                                     ", '" & sRow(15).Replace("'", "") & "', '" & sRow(16).Replace("'", "") & "', " & sRow(29) & ", " & sRow(30) & ", '" & sRow(31) & "', '" & sRow(32) & _
                                     "', '" & sRow(33) & "', '" & sRow(34) & "', '" & sRow(35) & "', '" & sRow(36) & "', '" & sRow(37) & "', '" & sRow(38) & _
                                     "', '" & sRow(39) & "', " & sRow(40) & ", " & sRow(137) & ", '" & sSiteAddr & "'); ")
 
                    End If
                End If
                iCnt = iCnt + 1
            End While
        End Using

        If sbSQL.Length > 0 Then InsertRecs(sbSQL)

        'Now register the table with the geodatabase.

        ElectionLayersWindow.RegisterTableWithDB("Consol_PollingPlace_Export")

        Return 0
    End Function

End Class

