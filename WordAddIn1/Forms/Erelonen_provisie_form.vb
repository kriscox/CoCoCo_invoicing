Imports System.Globalization
Imports System.Windows

Public Class Erelonen_provisie_form
    Dim culture As CultureInfo = CultureInfo.CurrentCulture
    Dim BTW As Double = 0.21

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Be sure all textboxes have the right currency formatting
        ValidateChildren()

    End Sub

    Private Sub Field_Validate(sender As Object, e As ComponentModel.CancelEventArgs) Handles Gerecht_input.Validating, Erelonen_input.Validating
        Dim faultSeparator As String
        Dim decimalSeparator As String = Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
        Dim Quantity As Double = Nothing
        Dim field As Forms.TextBox = CType(sender, Forms.TextBox)

        If decimalSeparator = "." Then faultSeparator = "," Else faultSeparator = "."

        If (0 = field.Text.Length) Then
            field.Text = 0
        ElseIf (field.Text.Contains(".") Xor field.Text.contains(",")) Then
            field.Text = field.Text.Replace(faultSeparator, decimalSeparator)
        End If

        If (Double.TryParse(field.Text, NumberStyles.Currency, culture, Quantity)) Then
            field.Text = FormatCurrency(Quantity)
        Else
            MsgBox("Waarde moet een getal zijn")
            e.Cancel = True
        End If

    End Sub



    Private Sub Field_Validated(sender As Object, e As EventArgs) Handles Gerecht_input.Validated, Erelonen_input.Validated

        Erelonen_btw.Text = FormatCurrency(CDbl(Erelonen_input.Text) * BTW)
        Erelonen_totaal.Text = FormatCurrency(CDbl(Erelonen_input.Text) + CDbl(Erelonen_btw.Text))
        Gerecht_totaal.Text = FormatCurrency(CDbl(Gerecht_input.Text))
        If IsNumeric(Gerecht_input.Text) Then
            Totaal_ex_btw.Text = FormatCurrency(CDbl(Erelonen_input.Text) + CDbl(Gerecht_input.Text))
        End If
        Totaal_btw.Text = FormatCurrency(CDbl(Erelonen_btw.Text))
        Totaal_inc_btw.Text = FormatCurrency(CDbl(Erelonen_totaal.Text) + CDbl(Gerecht_totaal.Text))
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

    Private Sub IC_CheckedChanged(sender As Object, e As EventArgs) Handles IC.CheckedChanged
        On Error GoTo Fault
        If IC.Checked Then
            Erelonen_btw.Visible = False
            BTW = 0
        Else
            Erelonen_btw.Visible = True
            BTW = 0.21
        End If
        Field_Validate(sender, e)
Fault:
        Throw New NotImplementedException("IC_checkbox exception in provisie form")
    End Sub
End Class