Public Class InputInvoice_Form
    Private Sub OGM_ok_Click(sender As Object, e As EventArgs) Handles OGM_ok.Click
        Dim OGMCode As String

        OGMCode = "+++"
        REM check if all fields are complete
        If Len(OGMcode1.Text) <> 3 Then
            MsgBox(Prompt:="Eerste deel moet uit 3 cijfers bestaan")
        ElseIf Len(OGMcode2.Text) <> 4 Then
            MsgBox(Prompt:="Tweede deel moet uit 4 cijfers bestaan")
        ElseIf Len(OGMcode3.Text) <> 5 Then
            MsgBox(Prompt:="Derde deel moet uit 5 vijfers bestaan")
        ElseIf Not check_omg(OGMcode1.Text, OGMcode2.Text, OGMcode3.Text) Then
            MsgBox(Prompt:="OMG code is foutief")
        Else
            Me.Tag = "OGM_OK"
            Me.Hide()
        End If
    End Sub

    Private Sub Dossier_ok_Click(sender As Object, e As EventArgs) Handles  Dossier_ok.Click
        REM check if all fields are complete
        If Len(Dossier_year.Text) <> 4 Then
            MsgBox(Prompt:="Eerste deel moet uit 4 cijfers bestaan")
        ElseIf Len(Dossier_nr.Text) <> 4 Then
            MsgBox(Prompt:="Tweede deel moet uit 4 cijfers bestaan")
        ElseIf Len(Dossier_nr2.Text) <> 1 Then
            MsgBox(Prompt:="Derde deel moet uit 1 cijfer bestaan")
        Else
            Me.Tag = "Dossier_OK"
            Me.Hide()
        End If
    End Sub

    Private Sub OGM_exit_Click(sender As Object, e As EventArgs) Handles OGM_exit.Click, Dossier_exit.Click
        Me.Tag = "EXIT"
        Me.Hide()
    End Sub

    Private Sub InputInvoice_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.AcceptButton = OGM_ok
        Me.CancelButton = OGM_exit
    End Sub

    Public Function check_omg(ByVal omg_code1 As String, ByVal omg_code2 As String, ByVal omg_code3 As String) As Boolean
        Dim rest, check As Double

        rest = CDbl(String.Concat(omg_code1, omg_code2, omg_code3)) \ 100
        rest = CDbl(rest) Mod 97

        REM get the last 2 chiffers
        check = CDbl(omg_code3) Mod 100

        check_omg = (rest = check)
    End Function

End Class