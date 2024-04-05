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
'//   ORIGINATOR		: Richard Chamberlain (WESTON SOLUTIONS INC.)
'//   DATE OF ORIGIN	: 02/07/08
'//   Current Version #	: 1.0
'//
'//   Purpose           : Class logic for importing Excel spreadsheet and upload to DB.
'//
'///////////////////////////////////////////////////////////////////////////////////

Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Text

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
        SQL = "CREATE TABLE " & sDBPrefix & "R700_03 (PRECINCT Int, CONSOL Int, BALLOTTYPE Int, REGISTEREDTOTAL Int, ABSENTEETOTAL Int);"
        CreateTable(sDBPrefix & "R700_03", SQL)

        'Load in the spreadsheet.
        Dim dtTest As DataTable = ReadDataFromExcel(excelfilename)

        If (dtTest Is Nothing) Then Return 0

        Dim dcTemp As DataColumn = New DataColumn

        Dim iPrecinct As Integer = 0
        Dim iConsol As Integer = 0
        Dim iBallotType As Integer
        Dim iTotalVoters As Integer = -1
        Dim iAbsentVoters As Integer = 0
        Dim iLastCol As Integer = 0
        Dim sbSQL As StringBuilder = New StringBuilder

        Dim iColCnt As Integer = dtTest.Columns.Count - 1
        Dim myRow As DataRow
        Dim iPrecinctNum As Integer
        For Each myRow In dtTest.Rows

            'Load in the precinct number and ballot type.
            If SafeDBNull(myRow.Item(0)) <> "" And SafeDBNull(myRow.Item(2)) <> "" Then
                If IsNumeric(myRow.Item(0)) Or Left(myRow.Item(0), 1) = "V" Then
                    If Left(myRow.Item(0), 1) = "V" Then
                        iPrecinctNum = CType("9" & Mid(myRow.Item(0), 2), Integer)
                    Else
                        iPrecinctNum = CType(myRow.Item(0), Integer)
                    End If
                    If iPrecinct <> iPrecinctNum Then
                        iPrecinct = iPrecinctNum
                        If Len(CType(myRow.Item(0), String)) > 6 Then
                            iConsol = CType(Mid(myRow.Item(0), 2), Integer)
                        Else
                            iConsol = iPrecinctNum
                        End If
                        iBallotType = CType(myRow.Item(2), Integer)
                        iTotalVoters = -1
                        iAbsentVoters = 0
                    End If
                End If
            End If

            'If the iTotalVoters value has been populated then the next line in the xls will be for the absentee voters.
            If iTotalVoters > -1 Then
                iAbsentVoters = CType(myRow.Item(iLastCol), Integer)
                sbSQL.Append("INSERT INTO " & sDBPrefix & "R700_03 VALUES (" & iPrecinct & ", " & iConsol & ", " & iBallotType & ", " & iTotalVoters & ", " & iAbsentVoters & "); ")
                iPrecinct = 0
                iConsol = 0
                iBallotType = 0
                iTotalVoters = -1
                iAbsentVoters = 0
                iLastCol = 0
            End If

            If UCase(SafeDBNull(myRow.Item(1))) = "TOTAL" And iPrecinct > 0 Then
                Dim i As Integer
                For i = 0 To iColCnt
                    If UCase(SafeDBNull(myRow.Item(i))) = "" Then
                        iLastCol = i - 1
                        Exit For
                    End If
                Next
                If iLastCol = 0 Then iLastCol = iColCnt
                iTotalVoters = CType(myRow.Item(iLastCol), Integer)
            End If
        Next myRow

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
            If SafeDBNull(myRow.Item(0)) <> "" And iConsolPrecinct > 0 Then
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
                  "Extended Properties=Excel 8.0;")

            da = New OleDbDataAdapter("SELECT * FROM [Sheet1$]", conn)

            conn.Open()

            da.Fill(dt)

            ReadDataFromExcel = dt

        Catch ex As Exception
            MsgBox("Could not open " & excelfilename & ".  Please enter the file name again.")
        Finally
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
        End Try

    End Sub

    Public Shared Function IsNumeric(ByVal sText As String) As Boolean

        If Double.TryParse(sText, Globalization.NumberStyles.AllowDecimalPoint) Then
            Return True
        Else
            Return False
        End If

    End Function

End Class
