Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        CreateRibbonExtensibilityObject()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As  _
        Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function
End Class
