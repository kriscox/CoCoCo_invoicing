Imports System.Globalization
Imports System.Windows

Public Class Ereloon_Nota_form
    Dim culture As CultureInfo = CultureInfo.CurrentCulture

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Be sure all textboxes have the right currency formatting
        ValidateChildren()

    End Sub

    ''' <summary>
    ''' Vallidate if the entered values in the Textboxes are numeric or a currency depending on Tag "Currency"
    ''' </summary>
    ''' <param name="sender"> The field on the form which is vallidated</param>
    ''' <param name="e">Cancel argument</param>
    Private Sub Field_Validate(sender As Object, e As ComponentModel.CancelEventArgs) Handles Verplaatsingen_2_uren.Validating, Verplaatsingen_2_minuten.Validating, Verplaatsingen.Validating, Uitvoering.Validating, Rolzetting.Validating, Forfait.Validating, Fax.Validating, Derden.Validating, Dagvaarding.Validating, Dactylo.Validating, Consultaties_uren.Validating, Consultaties_minuten.Validating, Bijkomende_kosten.Validating, Betekening.Validating, Andere.Validating, Fotokopies.Validating
        Dim Quantity As Double = Nothing
        Dim field As Forms.TextBox = CType(sender, Forms.TextBox)

        If (0 = field.Text.Length) Then
            field.Text = 0
        End If

        If (field.Tag IsNot Nothing AndAlso field.Tag.ToString.Contains("Currency") AndAlso Double.TryParse(field.Text, NumberStyles.Currency, culture, Quantity)) Then
            field.Text = FormatCurrency(Quantity)

            'If field Is Not parsable
        ElseIf (field.Tag IsNot Nothing Or Not Double.TryParse(field.Text, Quantity)) Then
            MsgBox("Waarde moet een getal zijn")
            e.Cancel = True
        End If

    End Sub

    Private Sub Field_Validated(sender As Object, e As EventArgs) Handles Verplaatsingen_2_uren.Validated, Verplaatsingen_2_minuten.Validated, Verplaatsingen.Validated, Uitvoering.Validated, Rolzetting.Validated, Forfait.Validated, Fax.Validated, Derden.Validated, Dagvaarding.Validated, Dactylo.Validated, Consultaties_uren.Validated, Consultaties_minuten.Validated, Bijkomende_kosten.Validated, Betekening.Validated, Andere.Validated, Fotokopies.Validated
        Dim subtotal_excvat As Double
        Dim subtotal_NoVAT, subtotal As Double

        subtotal_excvat = CDbl(Dactylo.Text) + '* Kostenschema.Column(6) +
            CDbl(Fotokopies.Text) + '* Kostenschema.Column(5) +
            CDbl(Fax.Text) + '* Kostenschema.Column(4) +
            CDbl(Verplaatsingen.Text) + '* Kostenschema.Column(3) +
            CDbl(Bijkomende_kosten.Text) +
            (CDbl(Verplaatsingen_2_uren.Text) + CDbl(Verplaatsingen_2_minuten.Text) / 60) + '* Kostenschema.Column(2) +
            (CDbl(Consultaties_uren.Text) + CDbl(Consultaties_minuten.Text) / 60) + '* Kostenschema.Column(1) +
            CDbl(Forfait.Text)

        subtotal_NoVAT = CDbl(Rolzetting.Text) + CDbl(Dagvaarding.Text) + CDbl(Betekening.Text) +
            CDbl(Uitvoering.Text) + CDbl(Andere.Text)

        subtotal = subtotal_NoVAT + subtotal_excvat * 1.21

        Subtotaal.Text = FormatCurrency(subtotal)

        Totaal.Text = FormatCurrency(Subtotaal.Text - CDbl(Provisies.Text) - CDbl(Provisies2.Text) - CDbl(Derden.Text))
    End Sub
End Class