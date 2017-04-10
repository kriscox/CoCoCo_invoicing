Imports System.Globalization
Imports System.Windows

Public Class Derden_Gelden_Form
    Dim culture As CultureInfo = CultureInfo.CurrentCulture

    Private Sub Field_Validate(sender As Object, e As ComponentModel.CancelEventArgs) Handles Payment_amount.Validating
        Dim faultSeparator As String
        Dim decimalSeparator As String = Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
        Dim Quantity As Double = Nothing
        Dim field As Forms.TextBox = CType(sender, Forms.TextBox)

        If decimalSeparator = "." Then faultSeparator = "," Else faultSeparator = "."

        If (0 = field.Text.Length) Then
            field.Text = 0
        ElseIf (field.Text.Contains(".") Xor field.Text.Contains(",")) Then
            field.Text = field.Text.Replace(faultSeparator, decimalSeparator)
        End If

        If (field.Tag IsNot Nothing AndAlso field.Tag.ToString.Contains("Currency") AndAlso Double.TryParse(field.Text, NumberStyles.Currency, culture, Quantity)) Then
            field.Text = FormatCurrency(Quantity)

            'If field Is Not parsable
        ElseIf (field.Tag IsNot Nothing Or Not Double.TryParse(field.Text, Quantity)) Then
            MsgBox("Waarde moet een getal zijn")
            e.Cancel = True
        End If

    End Sub
End Class