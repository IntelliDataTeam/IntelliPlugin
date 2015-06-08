Imports System.Windows.Forms
Public Class loading_bar


    Private Sub stop_button_Click(sender As Object, e As EventArgs) Handles stop_button.Click
        If MsgBox("Do you want want to cancel?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Application.Exit()
        End If
    End Sub
End Class