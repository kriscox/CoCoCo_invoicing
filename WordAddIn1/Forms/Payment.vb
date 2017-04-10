Imports System.Globalization
Imports System.Windows

Public Class Payment
    Dim culture As CultureInfo = CultureInfo.CurrentCulture


    Private Sub Field_Validate(sender As Object, e As ComponentModel.CancelEventArgs) Handles MyBase.Validating
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

    Private Sub Payment_Cancel_Click(sender As Object, e As EventArgs) Handles payment_cancel.Click
        Me.Tag = "NOK"
        Me.Hide()
    End Sub

    Private Sub Payment_ok_Click(sender As Object, e As EventArgs) Handles Payment_ok.Click
        Dim Amount As String
        On Error GoTo ErrorHandler

        Me.Tag = "OK"
        REM check of euro's juist zijn
        Amount = Me.Payment_amount.Text
        REM replace . with ,
        If InStr(Amount, ".") Then
            Amount = Amount.Replace(oldChar:=".", newChar:=",")
        End If
        REM check for 2 decimals
        If (InStr(Amount, ",") = (Len(Amount) - 2)) And CDbl(Amount) > 0 Then
            Me.Payment_amount.Text = Amount
            Me.Hide()
        Else
            GoTo ErrorHandler
        End If
        Exit Sub

ErrorHandler:
        MsgBox(Prompt:="Geen getal", Buttons:=vbOKOnly)
        Me.Payment_amount.Text = ""

    End Sub

    Private Sub Payment_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.AcceptButton = Payment_ok
        Me.CancelButton = payment_cancel
    End Sub

End Class