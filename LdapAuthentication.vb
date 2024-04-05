Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.DirectoryServices
Imports System.Windows.Forms

Public Class LdapAuthentication
    Private _path As String
    Private _filterAttribute As String
    Private isUserSupervisor As Boolean
    Private myAppStructure As ElectionAppStruct

    Public Function IsAuthenticatedUser() As Boolean

        _path = "LDAP://" & Environment.UserDomainName

        isUserSupervisor = False
        Dim rtnVal As String = "Invalid User"

        Dim domain As String = Environment.UserDomainName
        Dim username As String = Environment.UserName
        Dim domainAndUsername As String = domain.ToLower() & "\" & username.ToLower()


        Try
            Dim entry As DirectoryEntry = New DirectoryEntry(_path)

            'Bind to the native AdsObject to force authentication.			
            Dim obj As Object = entry.NativeObject

            Dim search As DirectorySearcher = New DirectorySearcher(entry)

            search.Filter = "(SAMAccountName=" + username + ")"
            search.PropertiesToLoad.Add("cn")
            Dim result As SearchResult = search.FindOne()
            If result.GetDirectoryEntry().Properties("cn").Value = "" Then
                Return "Invalid User"
            End If

            'Update the new path to the user in the directory.
            _path = result.Path
            _filterAttribute = DirectCast(result.Properties("cn")(0), String)
            Me.GetGroups()
            rtnVal = (isUserSupervisor)

        Catch ex As Exception
            Throw New Exception("Error authenticating user. " + ex.Message)
            Return False
        End Try

        Return rtnVal
    End Function

    Public Function GetGroups() As String

        Dim search As DirectorySearcher = New DirectorySearcher(_path)
        search.Filter = "(cn=" + _filterAttribute + ")"
        search.PropertiesToLoad.Add("memberOf")
        Dim groupName As String = ""

        Try
            Dim result As SearchResult = search.FindOne()
            Dim propertyCount As Integer = result.Properties("memberOf").Count
            Dim dn As String
            Dim equalsIndex, commaIndex As Integer

            For propertyCounter As Integer = 0 To propertyCount - 1
                dn = DirectCast(result.Properties("memberOf")(propertyCounter), String)
                equalsIndex = dn.IndexOf("=", 1)
                commaIndex = dn.IndexOf(",", 1)
                If equalsIndex = -1 Then
                    Return Nothing
                End If
                groupName = dn.Substring((equalsIndex + 1), (commaIndex - equalsIndex) - 1)
                If groupName.ToUpper = myAppStructure.sSupervisorGrp.ToUpper Then
                    isUserSupervisor = True
                End If

            Next

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Error obtaining group names. " + ex.Message)
        End Try

        Return groupName
    End Function

End Class

