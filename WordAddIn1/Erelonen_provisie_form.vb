Public Class Erelonen_provisie_form
    Private Sub Erelonen_provisie_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Calculate_Texts()
    End Sub

    Private Sub TextBox_Leave(sender As Object, e As EventArgs) Handles Gerecht_input.Leave, Erelonen_input.Leave
        Dim Calculate As Boolean = True
        If 0 <> Erelonen_input.TextLength Then
            If Not IsNumeric(Erelonen_input.Text) Then
                MsgBox("Please only enter numbers")
                Erelonen_input.Focus()
                Calculate = False
            Else
                Erelonen_input.Text = Format(CDbl(Erelonen_input.Text), "€ 0.00")
            End If
        End If

        If 0 <> Gerecht_input.TextLength Then
            If Not IsNumeric(Gerecht_input.Text) Then
                MsgBox("Please only enter numbers")
                Gerecht_input.Focus()
                Calculate = False
            Else
                Gerecht_input.Text = Format(CDbl(Gerecht_input.Text), "€ 0.00")
            End If
        End If
        Calculate_Texts()
    End Sub

    Private Sub Calculate_Texts()
        If IsNumeric(Erelonen_input.Text) Then
            Erelonen_btw.Text = Format(CDbl(Erelonen_input.Text) * 0.21, "€ 0.00")
            Erelonen_totaal.Text = Format(CDbl(Erelonen_input.Text) + CDbl(Erelonen_btw.Text), "€ 0.00")
            If IsNumeric(Gerecht_input.Text) Then
                Totaal_ex_btw.Text = Format(CDbl(Erelonen_input.Text) + CDbl(Gerecht_input.Text), "€ 0.00")
            End If
            Totaal_btw.Text = Format(CDbl(Erelonen_btw.Text), "€ 0.00")
            Totaal_inc_btw.Text = Format(CDbl(Erelonen_totaal.Text) + CDbl(Gerecht_totaal.Text), "€ 0.00")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button_OK.Click
        On Error GoTo Fault
        If (Not IsNumeric(Gerecht_input.Text) Or Gerecht_input.Text.Length = 0 Or Gerecht_input.Text = 0) And
                (Not IsNumeric(Erelonen_input.Text) Or Erelonen_input.Text.Length = 0 Or Erelonen_input.Text = 0) Then
Fault:
            MsgBox("Gerechtskosten en Erelonen mogen niet beide nul zijn")
        Else
            Tag = "Ok"
            Hide()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button_Cancel.Click
        Throw New NotImplementedException()
    End Sub
End Class