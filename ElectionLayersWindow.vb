'///////////////////////////////////////////////////////////////////////////////////
'//
'//                         COPYRIGHT [2014] PSOMAS INC.
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
'//   FILE NAME		    : ElectionLayersWindow.vb
'//   Namespace		    : ElectionSupervisor
'//   CLASS NAME(S)	    : ElectionLayersWindow
'//   ORIGINATOR		: Keith Palmer (PSOMAS, Inc.)
'//   DATE OF ORIGIN	: 02/07/08
'//   MAJOR UPDATE      : 02/18/2014
'//   Current Version #	: 2.0
'//
'//   Purpose           : Window and logic for Precinct layer export to Historic 
'//                       and Presnt layers.  Also upload logic for Rov spreadsheets.
'//
'///////////////////////////////////////////////////////////////////////////////////

Imports System.Text
Imports System.Data.SqlClient
Imports System.Security
Imports System.IO
Imports System.Security.Principal.WindowsIdentity
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.DataManagementTools
Imports ESRI.ArcGIS.esriSystem

Public Class ElectionLayersWindow

    Private pAOLicenseInitializer As New LicenseInitializer
    Private myAppStructure As ElectionAppStruct

    Public Sub New()

        ' Initialize, check license, create objects and read ini
        Try

            ' This call is required by the Windows Form Designer.
            InitializeComponent()

        Catch ex As Exception
            Dim strError As String
            Const ERR_PROC As String = "New()"
            strError = "ErrProcedure=" & ERR_PROC & vbCrLf & "ErrNumber=" & Err.Number & vbCrLf & "ErrDescription=" & Err.Description
            UpdateStatus("Initialization failed.")
            MsgBox(strError)
        End Try

    End Sub

    ' upload button logic
    Private Sub UploadBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UploadBtn.Click

        Dim sMsg As String = ""
        If rbOverwrite.Checked = True Then
            sMsg = "Overwriting existing data.  Be sure you have backed up Consolidation_Editor feature dataset before continuing."
            Dim result As Integer = MessageBox.Show(sMsg, "Upload data?", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Yes Then
                UploadProcess("OVERWRITE")
            End If
        Else
            UploadProcess("APPEND")
        End If

    End Sub

    Private Sub UploadProcess(ByVal sAppendOverwrite As String)

        ' Set cursor
        Me.Cursor = Cursors.WaitCursor

        If R700_03TB.Text = "" Or (R701_01TB.Text = "" And sAppendOverwrite = "OVERWRITE") Then
            MsgBox("You need to select both spreadsheets before continuing")
            ' Set cursor
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        Dim LoadXLSFunc As LoadXLSData = New LoadXLSData()
        LoadXLSFunc.SetAttributes(myAppStructure.sConnString, myAppStructure.sDSPrefix)

        If InStr(tbElectionName.Text, "'") > 0 Then
            MsgBox("Sorry.  You can not have apostrophes in the election name.")
            Exit Sub
        End If

        Dim iReturn700_09 As Integer = LoadXLSFunc.LoadR700_09Data(R700_03TB.Text, sAppendOverwrite, tbElectionName.Text)
        If iReturn700_09 = 0 Then
            UpdateStatus("The upload of the r700_03 table failed.")
            MsgBox("The upload of the r700_03 table failed.")
            ' Set cursor
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        If sAppendOverwrite = "OVERWRITE" Then
            Dim iReturn701_01 As Integer = LoadXLSFunc.LoadR701_01Data(R701_01TB.Text)

            ' upload status
            If iReturn701_01 = 0 Then
                UpdateStatus("The upload of the r701_01 table failed.")
                MsgBox("The upload of the r701_01 table failed.")
            End If

            ' call the export process logic
            ExportProcess()

        End If

        UpdateStatus("Upload of tables succeeded.")
        MsgBox("Upload of tables succeeded.")

        ' Set cursor
        Me.Cursor = Cursors.Default

    End Sub

    ' Call to Export the feature classes, add fields and update tables
    Private Sub ExportProcess()

        Dim iRtn As Integer
        Dim sHistFCIni, sPresFCIni As String
        Dim sHistFC, sPresFC As String
        Dim sHistFCAdded, sPresFCAdded As String

        ' Set cursor
        Me.Cursor = Cursors.WaitCursor

        ' check the election gis detail table
        If tbElectionName.Text > "" Then
            'iRtn = CheckElecGISDetail(tbElectionName.Text)
        Else
            MsgBox("Please enter an Election Name.")
            tbElectionName.Focus()
            Me.Cursor = Cursors.Default
            Return
        End If

        '' Duplicate election name found
        'If iRtn = -1 Then
        '    tbElectionName.Focus()
        '    Me.Cursor = Cursors.Default
        '    Return
        'Else
        '    ' now set the election name 
        '    sElecName = tbElectionName.Text
        'End If

        ''ESRI License Initializer generated code.
        Dim bLicense As Boolean = pAOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeEngineGeoDB, esriLicenseProductCode.esriLicenseProductCodeStandard, esriLicenseProductCode.esriLicenseProductCodeAdvanced}, _
        New esriLicenseExtensionCode() {})
        If Not bLicense Then
            MsgBox("An ArcGIS license is not available.  Try later or ask somebody to release a license for the next 5 minutes.")
            Exit Sub
        End If

        ' New pExportToFC class for exporting and validating
        Dim pExportToFC As ExportSDEFeatureClass = New ExportSDEFeatureClass

        'Delete existing layer
        Try
            'UpdateStatus("D of history layer stopped.")
            ' Delete historical table since this failed
            iRtn = pExportToFC.TruncateFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, myAppStructure.sTrgtConsol)
            Me.Cursor = Cursors.Default
        Catch ex As Exception

        End Try

        ' create table names
        sHistFCIni = myAppStructure.sTrgtHist

        ' Check that tables of both are populated
        If sHistFCIni > "" Then
            sHistFC = myAppStructure.sDSPrefix & sHistFCIni
        Else
            MsgBox("The naming of the Historical table has failed.")
            Me.Cursor = Cursors.Default
            Return
        End If

        'Delete existing layer
        Try
            'UpdateStatus("D of history layer stopped.")
            ' Delete historical table since this failed
            iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC)
            Me.Cursor = Cursors.Default
        Catch ex As Exception

        End Try

        ' Process historical feature class name
        ' Export to history table
        Try
            UpdateStatus("Historical table export started.")
            iRtn = pExportToFC.ExportTo(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sSrcDataset, myAppStructure.sDSPrefix & myAppStructure.sSrcFeatPrec, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC & "_t")
        Catch ex As Exception

        End Try

        ' Add history table fields
        If iRtn = 0 Then
            UpdateStatus("Historical table export complete, now adding fields.")
            iRtn = AddHistFields(myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC & "_t")
        Else
            UpdateStatus("Historical table export stopped.")
            Me.Cursor = Cursors.Default
            Return
        End If


        ' Start populate history table process
        If iRtn = 0 Then
            UpdateStatus("Addition of fields complete, now updating Historical table.")
            iRtn = UpdateHistRecords(sHistFC & "_t")
        Else
            UpdateStatus("Addition of fields stopped.")
            ' Delete historical table since this failed
            iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC & "_t")
            Me.Cursor = Cursors.Default
            Return
        End If

        If iRtn = 0 Then
            UpdateStatus("Dissolving the history layer based on the consolname field.")
            iRtn = DissolveHistTemp(sHistFC & "_t", sHistFC & "_t1")
        Else
            UpdateStatus("Dissolve of history layer stopped.")
            ' Delete historical table since this failed
            iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC & "_t1")
            Me.Cursor = Cursors.Default
            Return
        End If

        If iRtn = 0 Then
            ' now set the history feature class as successfully added
            iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC & "_t")
            sHistFCAdded = sHistFC
            UpdateStatus("Update of Historical table complete.")
        Else
            UpdateStatus("Process for Historical table failed.")
            ' Delete historical table since this failed
            iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC)
            Me.Cursor = Cursors.Default
            Return
        End If

        If iRtn = 0 Then
            UpdateStatus("Dissolving the history layer based on the consolname field.")
            iRtn = DissolveHistTemp(sHistFC & "_t1", sHistFC)
        Else
            UpdateStatus("Dissolve of history layer stopped.")
            ' Delete historical table since this failed
            iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC & "_t")
            Me.Cursor = Cursors.Default
            Return
        End If

        If iRtn = 0 Then
            ' now set the history feature class as successfully added
            iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC & "_t1")
            sHistFCAdded = sHistFC
            UpdateStatus("Update of Historical table complete.")
        Else
            UpdateStatus("Process for Historical table failed.")
            ' Delete historical table since this failed
            iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC)
            Me.Cursor = Cursors.Default
            Return
        End If


        sPresFCIni = myAppStructure.sTrgtPres

        ' Check that tables of both are populated
        If sPresFCIni > "" Then
            sPresFC = myAppStructure.sDSPrefix & sPresFCIni
        Else
            MsgBox("The naming of the Historical table has failed.")
            Me.Cursor = Cursors.Default
            Return
        End If

        'Delete existing precinct_present layer
        Try            
            iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sPresFC)
            Me.Cursor = Cursors.Default            
        Catch ex As Exception

        End Try

        ' Export to present table
        Try
            UpdateStatus("Present table export started.")
            iRtn = pExportToFC.ExportTo(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sSrcDataset, myAppStructure.sDSPrefix & myAppStructure.sSrcFeatPrec, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sPresFC)
            iRtn = pExportToFC.IndexField(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sPresFC, "PRECINCT", "IDXPRECINCT")
        Catch ex As Exception

        End Try
 
        '' Add present table fields
        'If iRtn = 0 Then
        '    UpdateStatus("Present table export complete, now adding fields.")
        '    iRtn = AddPresFields(myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sPresFC)
        'Else
        '    UpdateStatus("Present table export stopped.")
        '    Me.Cursor = Cursors.Default
        '    Return
        'End If

        ' Start populate present table process
        'If iRtn = 0 Then
        '    UpdateStatus("Addition of fields complete, now updating Present table.")
        '    iRtn = UpdatePresRecords(sPresFC)
        'Else
        '    UpdateStatus("Addition of fields stopped.")
        '    ' Delete Present table since this failed
        '    iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sPresFC)
        '    ' Also delete historical table since this failed
        '    iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC)
        '    Me.Cursor = Cursors.Default
        '    Return
        'End If

        If iRtn = 0 Then
            ' now set the present feature class as successfully added
            sPresFCAdded = sPresFC
            UpdateStatus("Update of Present table complete.")
        Else
            UpdateStatus("Process for Present table failed.")
            ' Delete Present table since this failed
            iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sPresFC)
            ' Also delete historical table since this failed
            iRtn = pExportToFC.DeleteFrom(myAppStructure.pConnPropSet, myAppStructure.sDSPrefix & myAppStructure.sOutputDataset, sHistFC)
            Me.Cursor = Cursors.Default
            Return
        End If

        ' Access status
        Select Case iRtn
            Case 0
                UpdateStatus("Processing complete.")
                'MsgBox("Processing complete.")
            Case -1
                UpdateStatus("Error in processing.")
                MsgBox("Error in processing.")
        End Select

        ' Set cursor
        Me.Cursor = Cursors.Default

        ' Nothing the export class and others
        pExportToFC = Nothing

    End Sub

    ' Add field logic for History feature class in SDE
    Public Function AddHistFields(ByVal sSDEFDS As String, ByVal sNewFC As String) As Integer

        ' New SDE workspace factory and workspace, open with propset
        Dim pSdeFactWS As IWorkspaceFactory = New SdeWorkspaceFactoryClass
        Dim pSDEWorkspace As IWorkspace

        pSDEWorkspace = pSdeFactWS.Open(myAppStructure.pConnPropSet, 0)
        If Not pSDEWorkspace.Exists Then
            Return -1
        Else
            'continue
        End If

        ' Check to make sure feature class does not already exist
        Dim pDataset As IDataset
        Dim pEnumDS As IEnumDataset
        Dim pFeatureDSToCheck As IFeatureDataset
        Dim pSDEworkspace2 As IWorkspace2
        Dim pSDEFWorkspace As IFeatureWorkspace
        Dim pFClass As IFeatureClass

        ' set workspace2
        pSDEworkspace2 = pSDEWorkspace

        ' Make sure feature class now exists and get feature class itself
        If pSDEworkspace2.NameExists(esriDatasetType.esriDTFeatureClass, sNewFC) Then

            'Open the workspace and set feature dataset to get feature class from
            pSDEFWorkspace = pSDEWorkspace
            pFeatureDSToCheck = pSDEFWorkspace.OpenFeatureDataset(sSDEFDS)
            pEnumDS = pFeatureDSToCheck.Subsets
            pEnumDS.Reset()

            'Loop through the dataset getting the feature class
            pFClass = Nothing
            pDataset = pEnumDS.Next
            Do Until pDataset Is Nothing
                If pDataset.Type = esriDatasetType.esriDTFeatureClass Then
                    If UCase(pDataset.Name) = UCase(sNewFC) Then
                        pFClass = pDataset
                    End If
                End If
                pDataset = pEnumDS.Next
            Loop
        Else
            MsgBox("History table is missing.") 'This feature class doesn't exist
            Return -1
        End If

        If pFClass Is Nothing Then
            MsgBox("History table is missing.") 'This feature class object doesn't exist
            Return -1
        Else
            'Add fields
            Dim pField, pField2 As IField
            Dim pFieldEdit, pFieldEdit2 As IFieldEdit
            pFieldEdit = New Field
            pFieldEdit2 = New Field

            pFieldEdit.AliasName_2 = "consolprecinct"
            pFieldEdit.Name_2 = "consolprecinct"
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeSingle
            pFieldEdit.Precision_2 = 10
            pFieldEdit.Editable_2 = True
            pField = CType(pFieldEdit, IField)

            pFClass.AddField(pField)

            pFieldEdit2.AliasName_2 = "homeprecinct"
            pFieldEdit2.Name_2 = "homeprecinct"
            'pFieldEdit2.Type_2 = esriFieldType.esriFieldTypeDouble
            pFieldEdit2.Type_2 = esriFieldType.esriFieldTypeSingle
            pFieldEdit2.Precision_2 = 10
            pFieldEdit2.Editable_2 = True
            pField2 = CType(pFieldEdit2, IField)

            pFClass.AddField(pField2)

            pField = Nothing
            pFieldEdit = Nothing
            pField2 = Nothing
            pFieldEdit2 = Nothing
        End If

        ' Nothing the connections, workspaces, and others
        pSdeFactWS = Nothing
        pSDEWorkspace = Nothing
        pSDEFWorkspace = Nothing
        pSDEworkspace2 = Nothing
        pFClass = Nothing
        pDataset = Nothing
        pEnumDS = Nothing
        pFeatureDSToCheck = Nothing

        Return 0

    End Function

    ' Update history table fields
    Public Function UpdateHistRecords(ByVal sFCName As String) As Integer

        Dim sSQL As String
        Dim pConnection As SqlConnection = New SqlConnection
        Dim pCommand As SqlCommand
        Dim pDataReader As SqlDataReader

        ' connection string and open connection
        pConnection.ConnectionString = myAppStructure.sConnString
        pConnection.Open()

        ' select SQL statement
        sSQL = "SELECT CONSOLPRECINCT, HOMEPRECINCT FROM " & myAppStructure.sDSPrefix & "R701_01"

        ' new SQLcommand with select SQL, datareader
        pCommand = New SqlCommand(sSQL, pConnection)
        pDataReader = pCommand.ExecuteReader()

        ' new stringbuilder
        Dim sb As StringBuilder = New StringBuilder()

        ' read with datareader
        While pDataReader.Read
            sb.Append("Update " & sFCName & " SET CONSOLPRECINCT = " & pDataReader(0) & _
                                            ", " & "HOMEPRECINCT = " & pDataReader(1))
            sb.Append(" WHERE ID = " & pDataReader(1)) '- KEP - 3/2/2012
            'sb.Append(" WHERE PRECINCT = " & pDataReader(1))
            sb.Append(" ")
        End While

        ' close datareader and connection
        pDataReader.Close()
        pConnection.Close()

        If sb.ToString = "" Then
            MsgBox("The R701_01 table doesn't contain any data.  Be sure to load the xls file before creating the election layers.")
            Return -1
            pCommand = Nothing
        End If

        ' new SQLcommand with stringbuilder string
        pCommand = New SqlCommand(sb.ToString, pConnection)
        pCommand.Connection.Open()

        ' Execute SQLcommand to do update
        Try
            pCommand.ExecuteNonQuery()
        Catch ex As SqlException
            If ex.Number > 0 Then
                MsgBox(ex.Message)
            End If

            ' Nothing thse
            pCommand = Nothing
            pConnection.Close()
            pDataReader.Close()

            Return -1
        End Try

        ' close command connection
        pCommand.Connection.Close()

        ' Nothing the commands, connections and others
        pCommand = Nothing
        pConnection.Close()
        pDataReader.Close()

        ' Success
        Return 0

    End Function

    Public Function DissolveHistTemp(ByVal sDissInName As String, ByVal sDissOutName As String) As Integer

        ' Intialize the Geoprocessor 
        Dim GP As Geoprocessor = New Geoprocessor
        Dim sDissInString As String
        Dim sDissOutString As String

        ' Set the OverwriteOutput setting to True
        GP.OverwriteOutput = True

        ' New Dissolve class
        Dim pDissClass As ESRI.ArcGIS.DataManagementTools.Dissolve = New ESRI.ArcGIS.DataManagementTools.Dissolve()
        'Dim pCopyClass As ESRI.ArcGIS.DataManagementTools.Copy = New ESRI.ArcGIS.DataManagementTools.Copy()

        ' create the dissolve strings
        sDissInString = myAppStructure.sSDEString & "\" & myAppStructure.sDSPrefix & myAppStructure.sOutputDataset & "\" & sDissInName
        sDissOutString = myAppStructure.sSDEString & "\" & myAppStructure.sDSPrefix & myAppStructure.sOutputDataset & "\" & sDissOutName

        ' set dissolve attributes
        pDissClass.dissolve_field = "consolprecinct"
        pDissClass.in_features = sDissInString
        pDissClass.out_feature_class = sDissOutString

        ' dissolve
        Try
            GP.Execute(pDissClass, Nothing)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        GP = Nothing

        ' Success
        Return 0

    End Function

    ' Add field logic for Present feature class in SDE
    Public Function AddPresFields(ByVal sSDEFDS As String, ByVal sNewFC As String) As Integer

        ' New SDE workspace factory and workspace, open with propset
        Dim pSdeFactWS As IWorkspaceFactory = New SdeWorkspaceFactoryClass
        Dim pSDEWorkspace As IWorkspace

        pSDEWorkspace = pSdeFactWS.Open(myAppStructure.pConnPropSet, 0)
        If Not pSDEWorkspace.Exists Then
            Return -1
        Else
            'continue
        End If

        ' Check to make sure feature class does not already exist
        Dim pDataset As IDataset
        Dim pEnumDS As IEnumDataset
        Dim pFeatureDSToCheck As IFeatureDataset
        Dim pSDEworkspace2 As IWorkspace2
        Dim pSDEFWorkspace As IFeatureWorkspace
        Dim pFClass As IFeatureClass

        ' set workspace2
        pSDEworkspace2 = pSDEWorkspace

        ' Make sure feature class now exists and get feature class itself
        If pSDEworkspace2.NameExists(esriDatasetType.esriDTFeatureClass, sNewFC) Then

            'Open the workspace and set feature dataset to get feature class from
            pSDEFWorkspace = pSDEWorkspace
            pFeatureDSToCheck = pSDEFWorkspace.OpenFeatureDataset(sSDEFDS)
            pEnumDS = pFeatureDSToCheck.Subsets
            pEnumDS.Reset()

            'Loop through the dataset getting the feature class
            pFClass = Nothing
            pDataset = pEnumDS.Next
            Do Until pDataset Is Nothing
                If pDataset.Type = esriDatasetType.esriDTFeatureClass Then
                    If UCase(pDataset.Name) = UCase(sNewFC) Then
                        pFClass = pDataset
                    End If
                End If
                pDataset = pEnumDS.Next
            Loop
        Else
            MsgBox("Present table is missing.") 'This feature class doesn't exist
            Return -1
        End If

        If pFClass Is Nothing Then
            MsgBox("Present table is missing.") 'This feature class object doesn't exist
            Return -1
        Else
            'Add fields
            Dim pField As IField
            Dim pFieldEdit As IFieldEdit
            pFieldEdit = New Field

            pFieldEdit.AliasName_2 = "Dissolve"
            pFieldEdit.Name_2 = "Dissolve"
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
            pFieldEdit.Editable_2 = True
            pFieldEdit.Length_2 = 10
            pField = CType(pFieldEdit, IField)

            pFClass.AddField(pField)

            pField = Nothing
            pFieldEdit = Nothing
        End If

        ' Nothing the connections, workspaces, and others
        pSdeFactWS = Nothing
        pSDEWorkspace = Nothing
        pSDEFWorkspace = Nothing
        pSDEworkspace2 = Nothing
        pFClass = Nothing
        pDataset = Nothing
        pEnumDS = Nothing
        pFeatureDSToCheck = Nothing

        Return 0

    End Function

    ' Close the window
    Private Sub CloseBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CloseBtn.Click

        Try
            pAOLicenseInitializer.ShutdownApplication()

            ' Nothing the class object, data types
            pAOLicenseInitializer = Nothing
            myAppStructure = Nothing

        Catch ex As Exception

        End Try

        Me.Close()

    End Sub

    ' update any status labels, bars and text within them
    Public Sub UpdateStatus(ByVal sStatus As String)
        TSSLabel.Text = sStatus
        Application.DoEvents()
    End Sub

    Private Sub R701_01TB_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles R701_01TB.DoubleClick
        OpenFileDialog1.Filter = "xls files (*.xls)|*.xls"
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then R701_01TB.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub R700_03TB_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles R700_03TB.DoubleClick
        OpenFileDialog1.Filter = "xls files (*.xls)|*.xls"
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then R700_03TB.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub tbPollingPlaceFile_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbPollingPlaceFile.DoubleClick
        OpenFileDialog1.Filter = "txt files (*.txt)|*.txt"
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then tbPollingPlaceFile.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub UploadPollingPlaceBtn_Click(sender As System.Object, e As System.EventArgs) Handles btnUploadPollingPlaceFile.Click

        If tbPollingPlaceFile.Text = "" Then
            MsgBox("You need to select a polling place file before continuing")
            ' Set cursor
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor

        Dim LoadXLSFunc As LoadXLSData = New LoadXLSData()
        LoadXLSFunc.SetAttributes(myAppStructure.sConnString, myAppStructure.sDSPrefix)

        Dim iReturnPollingPlace As Integer = LoadXLSFunc.LoadPollingPlaceTbl(tbPollingPlaceFile.Text)
        'If iReturnPollingPlace = 0 Then
        '    UpdateStatus("The upload of the Polling Place table failed.")
        '    MsgBox("The upload of the Polling Place table failed.")
        '    ' Set cursor
        '    Me.Cursor = Cursors.Default
        '    Exit Sub
        'End If

        'Now to compare the Polling Location SDE layer with the polling location table to find the records that don't have a match.
        Dim sSQL As String
        Dim pConnection As SqlConnection = New SqlConnection
        Dim dt As DataTable = New DataTable()

        ' connection string and open connection
        pConnection.ConnectionString = myAppStructure.sConnString
        pConnection.Open()

        ' select SQL statement

        sSQL = "SELECT PPI.poll_id, PPI.precinct_id, PPI.location_line_1, PPI.location_line_2, PPI.street_id, PPI.consolidation, PPI.Addr_Geocode, PPI.city, PPI.zip " & _
                "FROM " & myAppStructure.sDSPrefix & "Consol_PollingPlace_Export As PPI LEFT JOIN " & myAppStructure.sDSPrefix & myAppStructure.sPollLocationLayer & _
                " AS PP ON PPI.poll_id = PP.POLLSTATION_ID WHERE(PP.POLLSTATION_ID Is Null)"

        Dim da As SqlDataAdapter = New SqlDataAdapter(sSQL, pConnection)

        da.Fill(dt)

        Dim sOutputFolder As String = ""
        Dim curtime As DateTime = DateTime.Now
        Dim format As String = "MMddyy_HHmmss"
        sOutputFolder = System.IO.Path.GetDirectoryName(tbPollingPlaceFile.Text) & "\PollingLocationList_" & curtime.ToString(format) & ".csv"

        Dim bWorked As Boolean = TableToCSV(dt, sOutputFolder, True)
        If bWorked Then
            MsgBox("Upload of polling place table is complete.  Output non-match list is located at " & sOutputFolder)
        Else
            MsgBox("Something went wrong with the upload and output of the polling location file.")
        End If
        
        dt = Nothing
        da.Dispose()
        da = Nothing
        pConnection.Close()

        Me.Cursor = Cursors.Default

    End Sub

    Private Function TableToCSV(ByVal sourceTable As DataTable, ByVal filePathName As String, Optional ByVal HasHeader As Boolean = False) As Boolean
        'Writes a datatable back into a csv 
        Try
            Dim sb As New System.Text.StringBuilder
            If HasHeader Then
                Dim nameArray(200) As Object
                Dim i As Integer = 0
                For Each col As DataColumn In sourceTable.Columns
                    nameArray(i) = CType(col.ColumnName, Object)
                    i += 1
                Next col
                ReDim Preserve nameArray(i - 1)
                sb.AppendLine(String.Join(",", System.Array.ConvertAll(Of Object, String)(nameArray, _
                                Function(o As Object) If(o.ToString.Contains(","), _
                                ControlChars.Quote & o.ToString & ControlChars.Quote, o))))
            End If
            For Each dr As DataRow In sourceTable.Rows
                sb.AppendLine(String.Join(",", System.Array.ConvertAll(Of Object, String)(dr.ItemArray, _
                                Function(o As Object) If(o.ToString.Contains(","), _
                                ControlChars.Quote & o.ToString & ControlChars.Quote, o.ToString))))
            Next
            System.IO.File.WriteAllText(filePathName, sb.ToString)
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Return False
        End Try
    End Function

    Private Sub rbAppend_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbAppend.CheckedChanged
        rbOverwrite.Checked = False
        MsgBox("This option will append the 700.09 xls file to the existing table.  This will allow you to keep the existing election information so you can process multiple elections at the same time.  The historical layer will not be updated when using this option.")
    End Sub

    Private Sub rbOverwrite_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbOverwrite.CheckedChanged
        rbAppend.Checked = False
        MsgBox("This option will overwrite all existing data in the existing tables, delete the existing consolidations, and replace the precinct_present and precinct_historical layers.  You should backup the existing data in the Consolidation_Editor feature dataset before uploading the xls files.")
    End Sub

    Private Sub ElectionLayersWindow_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        ' New Strucure for Ini settings and new PropertySet for connection
        myAppStructure = New ElectionAppStruct

        Try
            ' Read ini and set all parameters in structure and class level
            Dim iRtn As Integer = myAppStructure.PopulateFromIni("ElectionApp.ini")
        Catch ex As Exception
            Dim strError As String
            Const ERR_PROC As String = "Problem load initialization file."
            strError = "ErrProcedure=" & ERR_PROC & vbCrLf & "ErrNumber=" & Err.Number & vbCrLf & "ErrDescription=" & Err.Description
            UpdateStatus("Initialization failed.")
            MsgBox(strError)
            Me.Close()
        End Try

        Try
            If myAppStructure.sTestMode.ToUpper = "NO" Then
                Dim ldapauth As LdapAuthentication = New LdapAuthentication
                If (Not ldapauth.IsAuthenticatedUser) Then
                    MsgBox("You are not in the ROV Supervisor group so are not authorized to run this application.")
                    Me.Close()
                End If
            End If

        Catch ex As Exception
            Dim strError As String
            Const ERR_PROC As String = "New()"
            strError = "ErrProcedure=" & ERR_PROC & vbCrLf & "ErrNumber=" & Err.Number & vbCrLf & "ErrDescription=" & Err.Description
            UpdateStatus("Initialization failed.")
            MsgBox(strError)
            Me.Close()
        End Try

    End Sub

    Private Sub ElectionLayersWindow_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Dispose()
    End Sub

    Public Function DeleteTable(ByVal sTblName As String) As Integer

        ' Intialize the Geoprocessor 
        Dim GP As Geoprocessor = New Geoprocessor
        Dim sRegInString As String

        ' Set the OverwriteOutput setting to True
        GP.OverwriteOutput = True

        ' New Register table class
        Dim pDelTbl As ESRI.ArcGIS.DataManagementTools.Delete = New ESRI.ArcGIS.DataManagementTools.Delete()

        ' create the dissolve strings
        sRegInString = myAppStructure.sSDEString & "\" & myAppStructure.sDSPrefix & sTblName

        'Set register attributes
        pDelTbl.in_data = sRegInString

        ' dissolve
        Try
            GP.Execute(pDelTbl, Nothing)
        Catch ex As Exception
            MsgBox(ex.Message & vbCr & "Unable to delete " & sRegInString)
        End Try

        GP = Nothing

        ' Success
        Return 0

    End Function

    ' Add field logic for Present feature class in SDE
    Public Function RegisterTableWithDB(ByVal sTableName As String) As Integer

        ' Intialize the Geoprocessor 
        Dim GP As Geoprocessor = New Geoprocessor
        Dim sRegInString As String

        ' Set the OverwriteOutput setting to True
        GP.OverwriteOutput = True

        ' New Register table class
        Dim pRegTbl As ESRI.ArcGIS.DataManagementTools.RegisterWithGeodatabase = New ESRI.ArcGIS.DataManagementTools.RegisterWithGeodatabase()

        ' create the dissolve strings
        sRegInString = myAppStructure.sSDEString & "\" & myAppStructure.sDSPrefix & sTableName

        'Set register attributes
        pRegTbl.in_dataset = sRegInString

        ' dissolve
        Try
            GP.Execute(pRegTbl, Nothing)
        Catch ex As Exception
            MsgBox(ex.Message & vbCr & "Unable to register the " & sRegInString & " with the geodatabase.  It may already be registered.")
        End Try

        GP = Nothing

        ' Success
        Return 0

    End Function
End Class
