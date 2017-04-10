Imports Microsoft.Office.Interop.Excel

Public Class Kostenschema
#Region "Variables"
    Property prestaties As Double
    Property wacht As Double
    Property verplaatsing As Double
    Property mail As Double
    Property fotokopie As Double
    Property dactylo As Double
    Property index As Integer
    Property VAT As Double
#End Region

    Public Sub New(index As Integer)
        index = index
        'read from excel
        Dim table As ListObject
        Dim row As ListRow
        Dim rownum As Integer = -1

        table = GlobalValues.GetWorkbook.Sheets("KostenSchemas").ListObjects("Kostenschema")

        'find rownr of asked kostenschema
        For Each row In table.ListRows
            If row.Range.Cells(1).Value2 = index Then
                rownum = row.Index
            End If
        Next

        If rownum = -1 Then
            MsgBox("Kostenschema niet meer gevonden")
            Return
        End If

        dactylo = table.ListColumns("Dactylo").DataBodyRange(rownum).Value
        fotokopie = table.ListColumns("fotokopie").DataBodyRange(rownum).Value
        mail = table.ListColumns("mail").DataBodyRange(rownum).Value
        prestaties = table.ListColumns("Prestatie").DataBodyRange(rownum).Value
        verplaatsing = table.ListColumns("verplaatsing").DataBodyRange(rownum).Value
        wacht = table.ListColumns("wacht").DataBodyRange(rownum).Value
        VAT = table.ListColumns("BTW").DataBodyRange(rownum).Value
    End Sub
End Class
