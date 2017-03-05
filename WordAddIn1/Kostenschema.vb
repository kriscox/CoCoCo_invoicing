Imports Microsoft.Office.Interop.Excel

Public Class Kostenschema
    Property prestaties As Double
    Property wacht As Double
    Property verplaatsing As Double
    Property mail As Double
    Property fotokopie As Double
    Property dactylo As Double
    Property index As Integer

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

        dactylo = table.ListColumns("Dactylo").DataBodyRange(rownum).Text
        fotokopie = table.ListColumns("fotokopie").DataBodyRange(rownum).Text
        mail = table.ListColumns("mail").DataBodyRange(rownum).Text
        prestaties = table.ListColumns("Prestatie").DataBodyRange(rownum).Text
        verplaatsing = table.ListColumns("verplaatsing").DataBodyRange(rownum).Text
        wacht = table.ListColumns("Wacht").DataBodyRange(rownum).Text
    End Sub
End Class
