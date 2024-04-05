Imports ESRI.ArcGIS.esriSystem

Namespace My

    Partial Friend Class MyApplication
        Private m_AOLicenseInitializer As LicenseInitializer = New ElectionSupervisor.LicenseInitializer()

        Private Sub MyApplication_Startup(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup
            ESRI.ArcGIS.RuntimeManager.Bind(ESRI.ArcGIS.ProductCode.Desktop)
            'ESRI License Initializer generated code.            
            m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeStandard, esriLicenseProductCode.esriLicenseProductCodeAdvanced}, _
            New esriLicenseExtensionCode() {})
        End Sub
    End Class


End Namespace

