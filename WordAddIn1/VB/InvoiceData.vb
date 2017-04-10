Imports Microsoft.Office.Interop.Excel
Public Class InvoiceData
    Property dossierNr As String
    Property dossierName As String
    Property aanspreektitel As String
    Property naam As String
    Property adres As String
    Property adres2 As String
    Property kostenSchema As Integer
    Property dactylo As Double
    Property fotokopies As Double
    Property fax As Double
    Property verplaatsing As Double
    Property bijkomende_kosten As Double
    Property forfait As Double
    Property erelonen As Double
    Property erelonen_uren As Integer
    Property erelonen_minuten As Integer
    Property wacht_kosten As Double
    Property wacht_uren As Integer
    Property wacht_minuten As Integer
    Property btw As Double
    Property rolzetting As Double
    Property dagvaarding As Double
    Property betekening As Double
    Property uitvoering As Double
    Property andere As Double
    Property derden As Double
    Property provisie_erelonen As Double
    Property provisie_gerechtskosten As Double
    Property Factuurnummer As Integer

    Public Sub New(ByRef ogm_Record As range)
        Dim SearchTable As ListObject = GlobalValues.GetWorkbook.Sheets("Ereloon Nota").ListObjects("Ereloon_Nota_Table")

        dactylo = ogm_Record.Cells(SearchTable.ListColumns("Dactylo").Index).Value
        fotokopies = ogm_Record.Cells(SearchTable.ListColumns("Fotokopies").Index).Value
        fax = ogm_Record.Cells(SearchTable.ListColumns("Fax").Index).Value
        verplaatsing = ogm_Record.Cells(SearchTable.ListColumns("Verplaatsing").Index).Value
        bijkomende_kosten = ogm_Record.Cells(SearchTable.ListColumns("Bijkomende_kosten").Index).Value
        forfait = ogm_Record.Cells(SearchTable.ListColumns("Forfait").Index).Value
        rolzetting = ogm_Record.Cells(SearchTable.ListColumns("Rolzetting").Index).Value
        dagvaarding = ogm_Record.Cells(SearchTable.ListColumns("Dagvaarding").Index).Value
        betekening = ogm_Record.Cells(SearchTable.ListColumns("Betekening").Index).Value
        uitvoering = ogm_Record.Cells(SearchTable.ListColumns("Uitvoering").Index).Value
        andere = ogm_Record.Cells(SearchTable.ListColumns("Andere").Index).Value

    End Sub

End Class
