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
'//   FILE NAME		    : ExportSDEFeatureClass.vb
'//   Namespace		    : ElectionSupervisor
'//   CLASS NAME(S)	    : ExportSDEFeatureClass
'//   ORIGINATOR		: Keith Palmer (WESTON SOLUTIONS INC.)
'//   DATE OF ORIGIN	: 02/07/08
'//   Current Version #	: 1.0
'//
'//   Purpose           : Class logic for Export of a feature class to SDE.  Also
'//                       has the delete SDE feature class function.
'//
'///////////////////////////////////////////////////////////////////////////////////

Imports System.Text
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.DataManagementTools

Public Class ExportSDEFeatureClass

    Private myAppStructure As ElectionAppStruct

    Public Sub New()

        ' New Strucure for Ini settings and new PropertySet for connection
        myAppStructure = New ElectionAppStruct
        Dim iRtn as Integer = myAppStructure.PopulateFromIni("ElectionApp.ini")
        ' Nothing for now
    End Sub


    Public Function ExportTo(ByVal sPropSet As PropertySet, ByVal sSDEInputFDS As String, ByVal sSrcFC As String, ByVal sSDEOutputFDS As String, ByVal sNewFC As String) As Integer

        ' Intialize the Geoprocessor 
        Dim GP As Geoprocessor = New Geoprocessor
        Dim sCopyInString As String
        Dim sCopyOutString As String

        ' Set the OverwriteOutput setting to True
        GP.OverwriteOutput = True


        sCopyInString = myAppStructure.sSDEString & "\" & sSDEInputFDS & "\" & sSrcFC
        sCopyOutString = myAppStructure.sSDEString & "\" & sSDEOutputFDS & "\" & sNewFC

        ' New Copy class
        'Dim pCopyClass As ESRI.ArcGIS.DataManagementTools.Copy = New ESRI.ArcGIS.DataManagementTools.Copy()
        Dim pCopyClass As ESRI.ArcGIS.DataManagementTools.CopyFeatures = New ESRI.ArcGIS.DataManagementTools.CopyFeatures()

        ' New SDE workspace factory and workspace, open with propset
        Dim pSdeFactWS As IWorkspaceFactory = New SdeWorkspaceFactoryClass
        Dim pSDEWorkspace As IWorkspace

        pSDEWorkspace = pSdeFactWS.Open(sPropSet, 0)
        If Not pSDEWorkspace.Exists Then
            Return -1
        Else
            'continue
        End If

        ' check for existing feature class
        Dim pSDEworkspace2 As IWorkspace2
        pSDEworkspace2 = pSDEWorkspace
        If pSDEworkspace2.NameExists(esriDatasetType.esriDTFeatureClass, sNewFC) Then
            MsgBox("Feature Class already exists.")
            Return -1
        Else
            'continue
        End If

        pCopyClass.in_features = sCopyInString
        pCopyClass.out_feature_class = sCopyOutString

        'pCopyClass.in_data = sCopyInString
        'pCopyClass.out_data = sCopyOutString

        ' dissolve
        Try
            GP.Execute(pCopyClass, Nothing)
            Dim i As Integer
            Dim sb As StringBuilder = New StringBuilder()
            For i = 0 To GP.MessageCount - 1
                sb.Append(GP.GetMessage(i) & vbLf)
            Next
            If InStr(UCase(sb.ToString), "FAILED") > 0 Then
                MsgBox(sb.ToString & vbCrLf & "Input: " & sCopyInString & vbCrLf & "Output: " & sCopyOutString)
                Return -1
            End If

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Input: " & sCopyInString & vbCrLf & "Output: " & sCopyOutString)
            Return -1
        End Try

        GP = Nothing

        ' Success
        Return 0

    End Function

    Public Function TruncateFrom(ByVal sPropSet As PropertySet, ByVal sSDEInputFDS As String, ByVal sSrcFC As String) As Integer

        ' Intialize the Geoprocessor 
        Dim GP As Geoprocessor = New Geoprocessor
        Dim sTruncateInString As String

        ' Set the OverwriteOutput setting to True
        GP.OverwriteOutput = True

        sTruncateInString = myAppStructure.sSDEString & "\" & sSDEInputFDS & "\" & sSrcFC

        ' New Copy class
        Dim pTruncateClass As ESRI.ArcGIS.DataManagementTools.TruncateTable = New ESRI.ArcGIS.DataManagementTools.TruncateTable()

        '' New SDE workspace factory and workspace, open with propset
        'Dim pSdeFactWS As IWorkspaceFactory = New SdeWorkspaceFactoryClass
        'Dim pSDEWorkspace As IWorkspace

        'pSDEWorkspace = pSdeFactWS.Open(sPropSet, 0)
        'If Not pSDEWorkspace.Exists Then
        '    Return -1
        'Else
        '    'continue
        'End If

        '' check for existing feature class
        'Dim pSDEworkspace2 As IWorkspace2
        'pSDEworkspace2 = pSDEWorkspace
        'If pSDEworkspace2.NameExists(esriDatasetType.esriDTFeatureClass, sSrcFC) Then

        'Else
        '    MsgBox(sSrcFC & " feature Class doesn't exists.")
        '    Return -1
        'End If


        pTruncateClass.in_table = sTruncateInString

        ' dissolve
        Try
            GP.Execute(pTruncateClass, Nothing)
            Dim i As Integer
            Dim sb As StringBuilder = New StringBuilder()
            For i = 0 To GP.MessageCount - 1
                sb.Append(GP.GetMessage(i) & vbLf)
            Next
            If InStr(UCase(sb.ToString), "FAILED") > 0 Then
                MsgBox(sb.ToString & vbCrLf & "Input: " & sTruncateInString & vbCrLf)
                Return -1
            End If

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Input: " & sTruncateInString & vbCrLf)
            Return -1
        End Try

        GP = Nothing

        ' Success
        Return 0

    End Function

    Public Function DeleteFrom(ByVal sPropSet As PropertySet, ByVal sSDEFDS As String, ByVal sDelFC As String) As Integer

        ' New SDE workspace factory and workspace, open with propset
        Dim pSdeFactWS As IWorkspaceFactory = New SdeWorkspaceFactoryClass
        Dim pSDEWorkspace As IWorkspace

        pSDEWorkspace = pSdeFactWS.Open(sPropSet, 0)
        If Not pSDEWorkspace.Exists Then
            pSDEWorkspace = Nothing
            pSdeFactWS = Nothing
            Return -1
        Else
            'continue
        End If

        ' New workspace name and set attributes
        Dim pSDEWorkspaceName As IWorkspaceName = New WorkspaceName
        pSDEWorkspaceName.ConnectionProperties = sPropSet
        pSDEWorkspaceName.WorkspaceFactoryProgID = "esriCore.SdeWorkspaceFactory.1"

        ' Create and assign name to SDE Dataset Name, assign workspace
        Dim pSDEFeatureDSName As IDatasetName
        pSDEFeatureDSName = New FeatureDatasetName
        pSDEFeatureDSName.Name = sSDEFDS
        pSDEFeatureDSName.WorkspaceName = pSDEWorkspaceName

        ' Create and assign name to Delete Feature Class Name, Dataset Name, Feature Class
        Dim pDelFCName As IFeatureClassName
        pDelFCName = New FeatureClassName
        pDelFCName.FeatureDatasetName = pSDEFeatureDSName

        Dim pDelDatasetName As IDatasetName
        pDelDatasetName = pDelFCName
        pDelDatasetName.Name = sDelFC

        ' New feature workspace manage
        Dim pFeatureWorkspaceManage As IFeatureWorkspaceManage
        pFeatureWorkspaceManage = pSDEWorkspace

        Dim pSDEworkspace2 As IWorkspace2
        pSDEworkspace2 = pSDEWorkspace
        Try
            ' check for existing feature class so can delete
            If pSDEworkspace2.NameExists(esriDatasetType.esriDTFeatureClass, sDelFC) Then
                pFeatureWorkspaceManage.DeleteByName(pDelFCName)
            Else
                'MsgBox("Feature Class does not exit for deletion.")
                Return 0  ' okay anyway
            End If
        Catch ex As Exception
            Return -1
        Finally
            ' Nothing the connections, workspaces and others
            pSdeFactWS = Nothing
            pSDEWorkspace = Nothing
            pSDEWorkspaceName = Nothing
            pSDEworkspace2 = Nothing
            pFeatureWorkspaceManage = Nothing
            pSDEFeatureDSName = Nothing
            pDelDatasetName = Nothing
            pDelFCName = Nothing
        End Try

        Return 0

    End Function

    Public Function IndexField(ByVal sPropSet As PropertySet, ByVal sSDEInputFDS As String, ByVal sSrcFC As String, ByVal sFieldName As String, ByVal sIdxName As String) As Integer

        ' Intialize the Geoprocessor 
        Dim GP As Geoprocessor = New Geoprocessor
        Dim sIndexInString As String

        ' Set the OverwriteOutput setting to True
        GP.OverwriteOutput = True

        sIndexInString = myAppStructure.sSDEString & "\" & sSDEInputFDS & "\" & sSrcFC

        ' New Copy class
        Dim pIndexClass As ESRI.ArcGIS.DataManagementTools.AddIndex = New ESRI.ArcGIS.DataManagementTools.AddIndex()

        ' New SDE workspace factory and workspace, open with propset
        Dim pSdeFactWS As IWorkspaceFactory = New SdeWorkspaceFactoryClass
        Dim pSDEWorkspace As IWorkspace

        pSDEWorkspace = pSdeFactWS.Open(sPropSet, 0)
        If Not pSDEWorkspace.Exists Then
            Return -1
        Else
            'continue
        End If

        ' check for existing feature class
        Dim pSDEworkspace2 As IWorkspace2
        pSDEworkspace2 = pSDEWorkspace
        If pSDEworkspace2.NameExists(esriDatasetType.esriDTFeatureClass, sSrcFC) Then

        Else
            MsgBox(sSrcFC & " feature Class doesn't exists.")
            Return -1
        End If

        pIndexClass.in_table = sIndexInString
        pIndexClass.fields = sFieldName
        pIndexClass.index_name = sIdxName

        ' dissolve
        Try
            GP.Execute(pIndexClass, Nothing)
            Dim i As Integer
            Dim sb As StringBuilder = New StringBuilder()
            For i = 0 To GP.MessageCount - 1
                sb.Append(GP.GetMessage(i) & vbLf)
            Next
            If InStr(UCase(sb.ToString), "FAILED") > 0 Then
                MsgBox(sb.ToString & vbCrLf & "Input: " & sIndexInString & vbCrLf)
                Return -1
            End If

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Input: " & sIndexInString & vbCrLf)
            Return -1
        End Try

        GP = Nothing
        ' Success
        Return 0

    End Function

End Class
