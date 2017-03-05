Imports Microsoft.Office.Interop

Public Class CoCoCo_Invoicing

    Private Sub CoCoCo_Startup() Handles Me.Startup
    End Sub

    Private Sub CoCoCo_Shutdown() Handles Me.Shutdown

    End Sub
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Invoicing_ribbon()
    End Function

End Class
