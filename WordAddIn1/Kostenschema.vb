Public Class Kostenschema
    Property prestaties As Double
    Property wacht As Double
    Property verplaatsing As Double
    Property mail As Double
    Property fotokopie As Double
    Property dactylo As Double
    Property index As Integer

    Public Sub New(index As Integer)
        Me.index = index
        'read from excel

    End Sub
End Class
