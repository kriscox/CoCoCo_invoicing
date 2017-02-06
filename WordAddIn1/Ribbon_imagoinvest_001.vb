Imports Microsoft.Office.Tools.Ribbon

Public Class Imagoinvest

    Private Sub Imagoinvest_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Maak_provisie_Click(sender As Object, e As RibbonControlEventArgs) Handles Maak_provisie.Click
        Dim ProvisieForm As Erelonen_provisie_form

        ProvisieForm = New Erelonen_provisie_form
        ProvisieForm.Show()

    End Sub
End Class
