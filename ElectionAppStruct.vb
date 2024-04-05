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
'//   FILE NAME		    : ElectionAppStruct.vb
'//   Namespace		    : ElectionSupervisor
'//   CLASS NAME(S)	    : ElectionAppStruct
'//   ORIGINATOR		: Keith Palmer (WESTON SOLUTIONS INC.)
'//   DATE OF ORIGIN	: 02/07/08
'//   Current Version #	: 1.0
'//
'//   Purpose           : Structure object to hold Ini variables & Election ID, Name.
'//
'///////////////////////////////////////////////////////////////////////////////////

Imports ESRI.ArcGIS.esriSystem

Public Structure ElectionAppStruct

    ' connection data types
    Dim sServer, sInstance, sDatabase As String
    Dim sUser, sPassword, sVersion As String
    Dim pConnPropSet As IPropertySet
    Dim sConnString As String

    ' DB element data types
    Dim sDBPrefix As String
    Dim sDSPrefix As String

    Dim sSrcDataset As String
    Dim sOutputDataset As String
    Dim sSrcFeatPrec As String
    Dim sSrcDimmHist As String
    Dim sSrcConsol As String

    Dim sTrgtHist As String
    Dim sTrgtPres As String
    Dim sTrgtConsol As String
    Dim sSrcDatasetCons As String
    Dim sSrcFeatPrecCons As String
    Dim sOutputDatasetCons As String

    Dim sPollLocationLayer As String

    Dim sSDEString As String

    Dim sSupervisorGrp As String
    Dim sTestMode As String

    ' election data types, not related to INI file
    Dim iElectionID As Integer
    Dim sElectionName As String

    Public Function PopulateFromIni(ByVal sIniFileName As String) As Integer

        Dim iFile As IniFile
        Dim sCurDir, sIniPath As String

        ' New propertset at this time, will hold sde connection below
        pConnPropSet = New ESRI.ArcGIS.esriSystem.PropertySetClass


        Dim sINIArgPath As String = ""

        Try
            sINIArgPath = My.Application.CommandLineArgs(0)
        Catch ex As Exception

        End Try

        If sINIArgPath <> "" Then
            sIniPath = sINIArgPath
        Else
            ' Construct Ini path and get file
            sCurDir = System.Environment.CurrentDirectory
            sIniPath = sCurDir & "\" & sIniFileName
        End If

        iFile = New IniFile(sIniPath)

        ' Get values for Connection String for SQL
        sConnString = iFile.GetString("SQLconnection", "ConnString", "")

        ' Get values for Connection Values for SDE
        sServer = iFile.GetString("SDEconnection", "Data Source", "")
        sInstance = iFile.GetString("SDEconnection", "Instance", "")
        sDatabase = iFile.GetString("SDEconnection", "Database", "")
        sUser = iFile.GetString("SDEconnection", "User", "")
        sPassword = iFile.GetString("SDEconnection", "Password", "")
        'sVersion = iFile.GetString("SDEconnection", "Version", "")

        ' Set SDE property set
        With pConnPropSet
            .SetProperty("SERVER", sServer)
            .SetProperty("INSTANCE", sInstance)
            .SetProperty("DATABASE", sDatabase)
            .SetProperty("USER", sUser)
            .SetProperty("PASSWORD", sPassword)
            '.SetProperty("VERSION", sVersion)
        End With

        ' Set DB elements - General
        sDBPrefix = iFile.GetString("DBElements_General", "DBPrefix", "")
        sDSPrefix = iFile.GetString("DBElements_General", "DSPrefix", "")
        sSupervisorGrp = iFile.GetString("DBElements_General", "SupervisorGrp", "")
        sTestMode = iFile.GetString("DBElements_General", "TestLogin", "NO")

        ' Set DB elements - Export
        sSrcDataset = iFile.GetString("DBElements_Export", "SrcDataset", "")
        sOutputDataset = iFile.GetString("DBElements_Export", "OutputDataset", "")
        sSrcFeatPrec = iFile.GetString("DBElements_Export", "SrcFeatPrec", "")
        sSrcDimmHist = iFile.GetString("DBElements_Export", "SrcDimmHist", "")
        sSrcConsol = iFile.GetString("DBElements_Export", "SrcConsol", "")
        'sSrcDimmPres = iFile.GetString("DBElements_Export", "SrcDimmPres", "")
        sTrgtHist = iFile.GetString("DBElements_Export", "TrgtHist", "")
        sTrgtPres = iFile.GetString("DBElements_Export", "TrgtPres", "")
        sTrgtConsol = iFile.GetString("DBElements_Export", "TrgtConsol", "")
        sSDEString = iFile.GetString("DBElements_Export", "SDEString", "")

        ' Set Polling location data        
        sPollLocationLayer = iFile.GetString("Election_Polling_Locations", "PollLocationLayer", "")

        ' Success
        Return 0

    End Function
End Structure
