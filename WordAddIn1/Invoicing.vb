Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

Public Class Invoicing
    Implements IDisposable

    Private globalvalues As GlobalValues = GlobalValues.GetInstance()
    Private ogm_Record As Excel.Range
    Private excel_Window As Excel.Window
    Private Workbook As Workbook = GlobalValues.GetWorkbook()
    Private ObjExcel As Excel.Application = GlobalValues.GetExcel()
    Dim dossierNr As String
    Dim inputInvoice As InputInvoice_Form
    Dim ereloonFct, gerechtskostenFct As Double
    Dim kostenSchema As Kostenschema
    Dim erelonen, Provisies As String
    Dim factuurData As InvoiceData
    Dim Factuurnummer As String
    Dim Total As Double
    Dim subtotal_Provisions, Prov_Erelonen, Prov_BTW, Prov_Gerecht As Double

    Public Property Workbook1 As Workbook
        Get
            Return Workbook
        End Get
        Set(value As Workbook)
            Workbook = value
        End Set
    End Property

    Public Sub startup()
        inputInvoice = New InputInvoice_Form()
        While True
            inputInvoice.OGMcode1.Text = ""
            inputInvoice.OGMcode2.Text = ""
            inputInvoice.OGMcode3.Text = ""
            inputInvoice.ShowDialog()

            Select Case inputInvoice.Tag

            REM Go to the excel
                Case Is = "TOWORKBOOK"

            REM loop to log a OGM code payment
                Case Is = "OGM_OK"
                    Dim ogm As String
                    ogm = "+++" & inputInvoice.OGMcode1.Text & "/" & inputInvoice.OGMcode2.Text & "/" & inputInvoice.OGMcode3.Text & "+++"
                    OGM_Payment(ogm)

            REM Exit loop
                Case Is = "OGM_EXIT"
                    Exit Sub
                Case Else
                    REM endless loop
            End Select
        End While
    End Sub

    Private Sub OGM_Payment(ByVal ogm As String)
        Dim Saldo, Amount, wages, rest_costs As Double
        Dim Dossier As String
        Dim Payment As Payment
        Dim ereloon As Boolean
        Dim rest As Double

        REM initialize global variables
        ereloonFct = 0
        gerechtskostenFct = 0
        erelonen = ""
        Provisies = ""

        REM Lookup ogm code
        If (Not ogm_lookup(ogm:=ogm, ereloon:=ereloon)) Then
            MsgBox(Prompt:="OGM niet gevonden of regel al afgesloten", Buttons:=vbCritical)
            GoTo Final
        End If

        REM read total Saldo
        If ereloon Then
            Saldo = Math.Round(ogm_Record.Cells(29).Value2 + 0.000001, 2)
        Else
            Saldo = Math.Round(ogm_Record.Cells(12).Value2 + 0.000001, 2)
        End If
        Dossier = ogm_Record.Cells(3).Value2

        REM *****
        REM request amount payed
        REM *****
        Payment = New Payment
        Payment.ogm_label.Text = ogm
        Payment.Dossier_label.Text = ogm_Record.Cells(3).Value2
        Payment.ShowDialog()

        If Payment.Tag <> "OK" Then
            Exit Sub
        End If

        Amount = CDbl(Payment.Payment_amount.Text)

        Factuurnummer = NextInvoiceNumber()

        REM *****
        REM Process payment
        REM *****
        If ereloon Then
            REM read DossierNr
            dossierNr = ogm_Record.Cells(3).Value2
            REM Saldo = Saldo - everything already provisioned
            Saldo = Saldo - getSaldo()
            If Saldo <= Amount Then

                If Saldo < Amount Then
                    rest = Amount - Saldo
                    MsgBox(Prompt:="Klant heeft teveel betaald. Restsaldo moet verwerkt worden, zit niet in de factuur.", Buttons:=vbExclamation)
                    Amount = Saldo
                End If

                REM Close erelonen entry
                Workbook1.Sheets("Ereloon Nota").Unprotect(Password:=GlobalValues.password)
                ogm_Record.Cells(32).Value2 = True
                ogm_Record.Cells(33).Value2 = Now
                ogm_Record.Cells(34).Value2 = Factuurnummer
                Workbook1.Sheets("Ereloon Nota").Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
                'TODO readkostenschem
                'kostenschema.readKostenSchema(kostenschemaID:=ogm_Record.Cells(9))
                erelonen = erelonen & " " & ogm_Record.Row
                Fee_invoice()
                Close_provisions()
            Else
                REM fill provisie for this amount
                rest = Fill_provisies(Amount, dossierNr)

                MsgBox(Prompt:="Klant heeft niet het volledige bedrag betaald, is geen eindfactuur.", Buttons:=vbExclamation)

                Provision_invoice()

            End If
            UpdateRecord()
        Else
            REM remove payed part
            Saldo = Saldo - CDbl(ogm_Record.Cells(15).Value2) - CDbl(ogm_Record.Cells(16).Value2) - CDbl(ogm_Record.Cells(17).Value2)

            REM everything payed close provisie
            If Saldo <= Amount Then
                Workbook1.Sheets("Provisies").Unprotect(Password:=GlobalValues.password)
                ereloonFct = ogm_Record.Cells(9).Value2 - ogm_Record.Cells(15).Value2
                gerechtskostenFct = ogm_Record.Cells(11).Value2 - ogm_Record.Cells(17).Value2
                ogm_Record.Cells(13).Value2 = True
                ogm_Record.Cells(15).Value2 = ogm_Record.Cells(9).Value2
                ogm_Record.Cells(16).Value2 = ogm_Record.Cells(10).Value2
                ogm_Record.Cells(17).Value2 = ogm_Record.Cells(11).Value2
                Workbook1.Sheets("Provisies").Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
                Provisies = "PE" & ogm_Record.Row
                'Fill rest on other open provisies
                rest = Amount - Saldo
                If rest > 0 Then
                    rest = Fill_provisies(rest, dossierNr)
                End If

            ElseIf Saldo > Amount Then
                ogm_Record.Cells(1, 13).text = False
                REM calculate remainder of costs
                rest_costs = ogm_Record.Cells(1, 11).text - ogm_Record.Cells(1, 17).text
                If Amount > rest_costs Then
                    Workbook1.Sheets("Provisies").Unprotect(Password:=GlobalValues.password)
                    ogm_Record.Cells(1, 17).text = ogm_Record.Cells(1, 11).text
                    gerechtskostenFct = rest_costs
                    Amount = Amount - rest_costs
                    REM devide the rest over wages and BTW
                    wages = Math.Round(Amount / 1.21, 2)
                    ogm_Record.Cells(1, 15).text = ogm_Record.Cells(1, 15).text + wages
                    ogm_Record.Cells(1, 16).text = ogm_Record.Cells(1, 16).text + Amount - wages
                    Workbook1.Sheets("Provisies").Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
                    ereloonFct = wages
                Else
                    Workbook1.Sheets("Provisies").Unprotect(Password:=GlobalValues.password)
                    ogm_Record.Cells(1, 17).text = ogm_Record.Cells(1, 17).text + Amount
                    Workbook1.Sheets("Provisies").Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
                    gerechtskostenFct = Amount
                End If
                Provisies = "PE" & ogm_Record.Row
            End If

            Provision_invoice()

        End If

Final:

    End Sub

    Private Sub Close_provisions()
        Dim sht As Worksheet
        Dim tbl As ListObject
        Dim row As Excel.Range

        sht = Workbook1.Sheets("Provisies")
        tbl = sht.ListObjects("Provisie_Table")

        On Error GoTo Final

        REM to place filter unprotect sheet
        sht.Unprotect(Password:=GlobalValues.password)

        REM remove the autofilter is necessairy
        With tbl
            .AutoFilter.ShowAllData()
            .Range.AutoFilter(Field:=3, Criteria1:=dossierNr)
            .Range.AutoFilter(Field:=1, Criteria1:="<" & DateValue(ogm_Record.Cells(1, 1)))
            For Each row In .DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows
                row.Cells(1, tbl.ListColumns("betaald").Index) = True
            Next
        End With
Final:
        If (sht.ProtectContents = False) Then
            sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
        End If

    End Sub

    Private Function ogm_lookup(ByVal ogm As String, ByRef ereloon As Boolean) As Boolean
        Dim searchSheet As Worksheet
        Dim SearchTable As ListObject
        Dim searchRange As Excel.Range
        Dim searchCount As Integer
        Dim Lst As Excel.ListRows
        Dim dossierNr As String

        On Error GoTo Final
        ogm_lookup = True
        ereloon = False
        dossierNr = ""
        REM *****
        REM first lookup in Provisie_table
        REM *****
        searchSheet = Workbook1.Sheets("Provisies")
        SearchTable = searchSheet.ListObjects("Provisie_Table")

        REM to place filter unprotect sheet
        Lst = SearchTable.ListRows
        searchSheet.Unprotect(Password:=GlobalValues.password)

        REM remove the autofilter is necessairy
        With SearchTable
            .AutoFilter.ShowAllData()
            .Range.AutoFilter(Field:=14, Criteria1:=ogm)
            searchCount = SearchTable.TotalsRowRange.Cells(1, SearchTable.ListColumns("dossiernr").Index).text
            If searchCount > 1 Then
                MsgBox(Prompt:="2 ogm codes gevonden, fout in excel")
                .AutoFilter.ShowAllData()
                searchSheet.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
                If Workbook1.Count > 1 Then
                    ObjExcel.Workbook.Close(SaveChanges:=True)
                Else
                    ObjExcel.Workbook.Save()
                    ObjExcel.Application.Quit()
                End If
            ElseIf searchCount = 1 Then
                If .DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows(1).Cells(13).Value2 = "Onwaar" Then
                    ogm_Record = .DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows(1)
                    GoTo Final
                Else
                    dossierNr = .DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows(1).Cells(3).Value2
                End If
            Else
                REM no row found cleaning
                .AutoFilter.ShowAllData()
                searchSheet.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
            End If
        End With

        REM *****
        REM Then lookup in Ereloon_nota
        REM *****
        searchSheet = Workbook1.Sheets("Ereloon Nota")
        SearchTable = searchSheet.ListObjects("Ereloon_Nota_Table")

        REM to place filter unprotect sheet
        Lst = SearchTable.ListRows
        searchSheet.Unprotect(Password:=GlobalValues.password)

        If (dossierNr = "") Then
            REM OGM not found in provisions
            REM remove the autofilter is necessairy
            With SearchTable
                .AutoFilter.ShowAllData()
                .Range.AutoFilter(Field:=31, Criteria1:=ogm)
                .Range.AutoFilter(Field:=32, Criteria1:="")
                searchCount = SearchTable.TotalsRowRange.Cells(1, SearchTable.ListColumns("dossiernr").Index).text
                If searchCount > 1 Then
                    MsgBox(Prompt:="2 ogm codes gevonden, fout in excel")
                    .AutoFilter.ShowAllData()
                    searchSheet.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
                    If ObjExcel.Workbooks.Count > 1 Then
                        Workbook1.Close(SaveChanges:=True)
                    Else
                        Workbook1.Save()
                        ObjExcel.Application.Quit()
                    End If
                ElseIf searchCount = 1 Then
                    ogm_Record = .DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows(1)
                    ereloon = True
                    GoTo Final
                Else
                    REM no row found
                    .AutoFilter.ShowAllData()
                    ogm_lookup = False
                End If
            End With
        Else
            REM find ogm for dossiernr
            With SearchTable
                .AutoFilter.ShowAllData()
                .Range.AutoFilter(Field:=3, Criteria1:=dossierNr)
                searchCount = SearchTable.TotalsRowRange.Cells(1, SearchTable.ListColumns("dossiernr").Index).text
                If searchCount <> 1 Then
                    MsgBox(Prompt:="Ogm van een afgesloten provisie, geen bijhorende ereloon nota gevonden")
                    GoTo Final
                Else
                    If MsgBox(Prompt:="OGM van een afgesloten provisie, ogm code van bijhorende ereloon nota is: " +
                    .DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows(1).Cells(31).Value2 + ". Mag ik hier  op boeken?",
                    Buttons:=vbYesNo + vbQuestion) = vbYes Then

                        ogm_Record = .DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows(1)
                        ereloon = True
                    Else
                        ogm_lookup = False
                    End If
                End If
            End With
        End If

Final:
        REM remove the autofilter
        SearchTable.AutoFilter.ShowAllData()

        If (searchSheet.ProtectContents = False) Then
            searchSheet.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
        End If

    End Function
    Private Sub UpdateRecord()
        Dim searchSheet As Worksheet
        Dim SearchTable As ListObject
        Dim searchRange As Excel.Range
        Dim searchCount As Integer
        Dim erelonen_payed, erelonen_vat_payed, provisions_payed As Double
        Dim Lst As Excel.ListRows

        On Error GoTo Final

        REM search in provisies
        searchSheet = Workbook1.Sheets("Provisies")
        SearchTable = searchSheet.ListObjects("Provisie_Table")

        REM to place filter unprotect sheet
        Lst = SearchTable.ListRows
        searchSheet.Unprotect(Password:=GlobalValues.password)

        REM remove the autofilter is necessairy
        With SearchTable
            .AutoFilter.ShowAllData()
            .Range.AutoFilter(Field:=3, Criteria1:=ogm_Record.Cells(3))
            REM get all payed Erelonen and provisions
            ogm_Record.Cells(27).Value2 = .TotalsRowRange.Cells(1, .ListColumns("Ereloon_betaald").Index).text * 1.21
            ogm_Record.Cells(28).Value2 = .TotalsRowRange.Cells(1, .ListColumns("gerechtskosten_betaald").Index).text

            REM calculate open saldo

            'calculate total
            ogm_Record.Cells(29).Value2 =
                ogm_Record.Cells(1, .ListColumns("Dactylo").Index).text * kostenSchema.dactylo +
                ogm_Record.Cells(1, .ListColumns("Fotokopies").Index).text * kostenSchema.fotokopie +
                ogm_Record.Cells(1, .ListColumns("Fax").Index).text * kostenSchema.dactylo +
                ogm_Record.Cells(1, .ListColumns("Verplaatsing").Index).text * kostenSchema.verplaatsing +
                ogm_Record.Cells(1, .ListColumns("Bijkomende_kosten").Index).text + _
 _
                ogm_Record.Cells(1, .ListColumns("Forfait").Index).text + _
 _
                (ogm_Record.Cells(1, .ListColumns("erelonen_uren").Index).text +
                  ogm_Record.Cells(1, .ListColumns("erelonen_minuten").Index).text / 60) * kostenSchema.prestaties + _
 _
                (ogm_Record.Cells(1, .ListColumns("wacht_uren").Index).text +
                  ogm_Record.Cells(1, .ListColumns("wacht_minuten").Index).text / 60) * kostenSchema.wacht + _
 _
                ogm_Record.Cells(1, .ListColumns("BTW").Index).text +
                ogm_Record.Cells(1, .ListColumns("Rolzetting").Index).text +
                ogm_Record.Cells(1, .ListColumns("Dagvaarding").Index).text +
                ogm_Record.Cells(1, .ListColumns("Betekening").Index).text +
                ogm_Record.Cells(1, .ListColumns("Uitvoering").Index).text +
                ogm_Record.Cells(1, .ListColumns("Andere").Index).text +
                ogm_Record.Cells(1, .ListColumns("Derden").Index).text - _
 _
                ogm_Record.Cells(1, .ListColumns("Provisies_erelonen").Index).text -
                ogm_Record.Cells(1, .ListColumns("Provisies_gerechtskosten").Index).text
        End With

Final:
        REM remove the autofilter
        SearchTable.AutoFilter.ShowAllData()

        If (searchSheet.ProtectContents = False) Then
            searchSheet.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
        End If
    End Sub

    Private Function getSaldo() As Double
        Dim sht As Worksheet
        Dim tbl As ListObject
        Dim Lst As ListRows
        Dim Provisies_Erelonen_VAT, Provisies_Erelonen_ExVAT As Double
        Dim Provisies_erelonen As Double
        Dim Provisies_GerechtsKosten As Double

        On Error GoTo ErrorHandler

        '------------------------------------------------------
        'read provisies
        '------------------
        sht = Workbook1.Sheets("Provisies")
        tbl = sht.ListObjects("Provisie_Table")
        Lst = tbl.ListRows
        sht.Unprotect(Password:=GlobalValues.password)

        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()
        tbl.Range.AutoFilter(Field:=3, Criteria1:=dossierNr)
        tbl.AutoFilter.ApplyFilter()
        Provisies_Erelonen_ExVAT = tbl.TotalsRowRange.Cells(1, tbl.ListColumns("Ereloon_betaald").Index).text
        Provisies_Erelonen_VAT = tbl.TotalsRowRange.Cells(tbl.ListColumns("BTW_betaald").Index).text
        Provisies_erelonen = Provisies_Erelonen_ExVAT + Provisies_Erelonen_VAT
        Provisies_GerechtsKosten = tbl.TotalsRowRange.Cells(tbl.ListColumns("gerechtskosten_betaald").Index).text

        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()

ErrorHandler:
        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
        getSaldo = Math.Round(Provisies_Erelonen_ExVAT + Provisies_Erelonen_VAT + Provisies_GerechtsKosten + 0.000001, 2)
    End Function

    Private Function Fill_provisies(ByVal Amount As Double, ByVal dossierNr As String) As Double

        On Error GoTo End_

        REM filter provisies
        Dim sht As Worksheet
        Dim tbl As ListObject
        Dim Lst As ListRows
        Dim rng As Excel.Range
        Dim row As Excel.Range
        Dim diff, gerechtskosten, ereloon As Double
        Dim idGerechtskostenToPay As Integer
        Dim idGerechtskostenPayed As Integer
        Dim idEreloonToPay As Integer
        Dim idEreloonPayed As Integer
        Dim idVATToPay As Integer
        Dim idVATPayed As Integer
        Dim idPayed As Integer
        Dim i As Integer

        REM get provisie table
        sht = Workbook1.Sheets("Provisies")
        tbl = sht.ListObjects("Provisie_Table")
        Lst = tbl.ListRows
        sht.Unprotect(Password:=GlobalValues.password)

        REM Get column id
        idGerechtskostenToPay = tbl.ListColumns("gerechtskosten").Index
        idGerechtskostenPayed = tbl.ListColumns("gerechtskosten_betaald").Index
        idEreloonToPay = tbl.ListColumns("Ereloon").Index
        idEreloonPayed = tbl.ListColumns("Ereloon_betaald").Index
        idVATToPay = tbl.ListColumns("BTW").Index
        idVATPayed = tbl.ListColumns("BTW_betaald").Index
        idPayed = tbl.ListColumns("betaald").Index

        REM remove the autofilter is necessairy and filter on dossierNr
        tbl.AutoFilter.ShowAllData()
        tbl.Range.AutoFilter(Field:=3, Criteria1:=dossierNr)
        tbl.AutoFilter.ApplyFilter()

        If (tbl.Range.SpecialCells(XlCellType.xlCellTypeVisible).Rows.Count <= 1) Then
            REM no lines exist jump to add_row
            GoTo Add_Row_
        End If

        REM for each row first fill up gerechtskosten
        rng = tbl.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible)
        For Each row In rng.Rows
            diff = row.Cells(1, idGerechtskostenToPay).text - row.Cells(1, idGerechtskostenPayed).text
            If diff > 0 Then
                REM if amount isn't sufficient to fill up, full up with amount and end function
                Provisies = Provisies & " P" & row.Row
                If diff > Amount Then
                    row.Cells(1, idGerechtskostenToPay).text = row.Cells(1, idGerechtskostenPayed).text + Amount
                    gerechtskostenFct = gerechtskostenFct + Amount
                    Amount = 0
                    Exit For
                Else
                    row.Cells(1, idGerechtskostenPayed).text = row.Cells(1, idGerechtskostenToPay).text
                    If row.Cells(1, idEreloonToPay).text = row.Cells(1, idEreloonPayed).text Then
                        row.Cells(1, idPayed).text = True
                    End If
                    gerechtskostenFct = gerechtskostenFct + diff
                    Amount = Amount - diff
                    If Amount = 0 Then
                        Exit For
                    End If
                End If
            End If
        Next

        REM goto end if no amount available
        If Amount = 0 Then
            GoTo End_
        End If

        REM for each row then fill up erelonen
        rng = tbl.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible)
        For Each row In rng.Rows
            diff = row.Cells(1, idEreloonToPay).text - row.Cells(1, idEreloonPayed).text
            If diff > 0 Then
                REM if amount isn't sufficient to fill up, full up with amount and end function
                REM diff is without 21% VAT
                Provisies = Provisies & " ES" & row.Row
                If diff * 1.21 > Amount Then
                    row.Cells(1, idEreloonPayed).text = row.Cells(1, idEreloonPayed).text + (Amount / 1.21)
                    row.Cells(1, idVATPayed).text = row.Cells(1, idVATPayed).text + (Amount / 1.21 * 0.21)
                    ereloonFct = ereloonFct + (Amount / 1.21)
                    Amount = 0
                    Exit For
                Else
                    row.Cells(1, idEreloonPayed).text = row.Cells(1, idEreloonToPay).text
                    row.Cells(1, idVATPayed).text = row.Cells(1, idVATToPay).text
                    row.Cells(1, idPayed).text = True
                    ereloonFct = ereloonFct + diff
                    Amount = Amount - diff * 1.21
                    If Amount <= 0 Then
                        Exit For
                    End If
                End If
            End If
        Next

        If Amount = 0 Then
            GoTo End_
        End If

Add_Row_:
        REM add a new provision for the rest
        With Lst.Add.Range
            .Cells(1) = Now
            For i = 2 To 8
                .Cells(i) = ogm_Record.Cells(i).text
            Next

            If ogm_Record.Columns.Count > 30 Then
                REM Eindnota

                tbl.AutoFilter.ShowAllData()
                tbl.Range.AutoFilter(Field:=3, Criteria1:=ogm_Record.Cells(3))
                tbl.AutoFilter.ApplyFilter()
                gerechtskosten = ogm_Record.Cells(21).Value2 + ogm_Record.Cells(22).Value2 + ogm_Record.Cells(23).Value2 _
                + ogm_Record.Cells(24).Value2 + ogm_Record.Cells(25).Value2 -
                tbl.TotalsRowRange.Cells(1, tbl.ListColumns("gerechtskosten_betaald").Index).text
            Else
                gerechtskosten = 0
            End If

            If Amount <= gerechtskosten Then
                gerechtskosten = Amount
                ereloon = 0
            Else
                ereloon = (Amount - gerechtskosten) / 1.21
            End If

            .Cells(tbl.ListColumns("Ereloon").Index).text = ereloon
            .Cells(tbl.ListColumns("BTW").Index).text = ereloon * 0.21
            .Cells(tbl.ListColumns("Ereloon_betaald").Index).text = ereloon
            .Cells(tbl.ListColumns("BTW_betaald").Index).text = ereloon * 0.21
            .Cells(tbl.ListColumns("totaal").Index).text = Amount
            .Cells(tbl.ListColumns("betaald").Index).text = True
            .Cells(tbl.ListColumns("gerechtskosten").Index).text = gerechtskosten
            .Cells(tbl.ListColumns("gerechtskosten_betaald").Index).text = gerechtskosten

            ereloonFct = ereloonFct + ereloon
            gerechtskostenFct = gerechtskostenFct + gerechtskosten

            .Cells(tbl.ListColumns("ogmnummer").Index).text = "+++ / / +++"

            Provisies = Provisies & " EP" & .Row
        End With
        Amount = 0

End_:
        On Error Resume Next

        Fill_provisies = Amount

        tbl.AutoFilter.ShowAllData()
        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)

    End Function

    Private Sub Fee_invoice()
        Dim Document As Document
        Dim objWord As Word.Application
        Dim wordWindow As Word.Window
        Dim table As Word.Table
        Dim row, bottomRow As Word.Row
        Dim subtotal_ExVAT, subtotal_NoVAT, subtotal As Double

        kostenSchema = New Kostenschema(index:=ogm_Record.Cells(9).Value2)
        factuurData = New InvoiceData(ogm_Record)
        'Create document
        objWord = CreateObject("Word.Application")
        Document = objWord.Documents.Add(Template:=GlobalValues.invoiceTemplate, Visible:=True)

        'Fill header
        AddHeader(Document:=Document, Factuurnummer:=Factuurnummer)

        'Fill table
        table = Document.Tables(2)
        table.Range.ParagraphFormat.KeepWithNext = True
        subtotal_ExVAT = AddWages(table:=table, kind:="ereloon")
        subtotal_ExVAT = subtotal_ExVAT + AddOfficeExpenses(table:=table)
        subtotal_NoVAT = AddLitigation(table:=table, kind:="ereloon")
        subtotal_NoVAT = subtotal_NoVAT + AddProvision(table:=table)
        subtotal_Provisions = AddPayedProvisions(table:=table)

        'Remove border of second row
        table.Rows(2).Borders(WdBorderType.wdBorderBottom).Visible = False

        ' Add Total
        bottomRow = table.Rows.Add

        With bottomRow
            .Cells.Borders(WdBorderType.wdBorderVertical).Visible = False
            .Cells.Borders(WdBorderType.wdBorderBottom).Visible = False
            .Cells.Borders(WdBorderType.wdBorderLeft).Visible = False
            .Cells.Borders(WdBorderType.wdBorderRight).Visible = False
            .Range.ParagraphFormat.KeepWithNext = True
        End With

        With table.Rows.Add
            .Cells(2).Merge(MergeTo:= .Cells(5))
            .Cells(2).Range.InsertAfter(Text:="Subtotaal excl Btw")
            .Cells(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
            .Cells(3).Range.InsertAfter(Text:=Format(Expression:=subtotal_ExVAT - Prov_Erelonen, Style:=GlobalValues.NumberFormat))
            .Range.ParagraphFormat.KeepWithNext = True
        End With

        With table.Rows.Add
            .Cells(2).Range.InsertAfter(Text:="Subtotaal Btw")
            .Cells(3).Range.InsertAfter(Text:=Format(Expression:=subtotal_ExVAT * 0.21 - Prov_BTW, Style:=GlobalValues.NumberFormat))
            .Range.ParagraphFormat.KeepWithNext = True
        End With

        With table.Rows.Add
            .Cells(2).Range.InsertAfter(Text:="Subtotaal derden en gerechtskosten")
            .Cells(3).Range.InsertAfter(Text:=Format(Expression:=subtotal_NoVAT - Prov_Gerecht, Style:=GlobalValues.NumberFormat))
            .Range.ParagraphFormat.KeepWithNext = True
        End With

        Total = subtotal_ExVAT * 1.21 + subtotal_NoVAT - subtotal_Provisions
        With table.Rows.Add
            .Cells(2).Range.InsertAfter(Text:="Totaal")
            .Cells(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
            .Cells(3).Range.InsertAfter(Text:=Format(Expression:=Total, Style:=GlobalValues.NumberFormat))
            .Cells(3).Borders(WdBorderType.wdBorderTop).Visible = True
            .Cells(2).Range.Font.Bold = True
            .Cells(3).Range.Font.Bold = True
            .Range.ParagraphFormat.KeepWithNext = True
        End With

        'add border under table
        Dim newTable As Table
        newTable = table.Split(bottomRow.Index + 1)
        bottomRow.Borders(WdBorderType.wdBorderTop).Visible = True
        bottomRow.Delete()
        table.Columns.SetWidth(ColumnWidth:=ObjExcel.CentimetersToPoints(2.5), RulerStyle:=WdRulerStyle.wdAdjustNone)
        table.Columns(1).SetWidth(ColumnWidth:=ObjExcel.CentimetersToPoints(5.25), RulerStyle:=WdRulerStyle.wdAdjustNone)
        table.Columns(2).SetWidth(ColumnWidth:=ObjExcel.CentimetersToPoints(1.5), RulerStyle:=WdRulerStyle.wdAdjustNone)
        table.Columns(6).SetWidth(ColumnWidth:=ObjExcel.CentimetersToPoints(3), RulerStyle:=WdRulerStyle.wdAdjustNone)

        'Log invoice
        logInvoice(provisie:=False)

        objWord.Visible = True
        Document.PrintPreview()
        MsgBox("Kijk de factuur na")

        Document.SaveAs2(FileName:=GlobalValues.InvoicePath + Factuurnummer)
        objWord.ActivePrinter = "Standaard"
        Document.PrintOut(Background:=True)
        objWord.ActivePrinter = "Standaard"
        On Error GoTo closeWord
        If objWord.Documents.Count > 1 Then
            Document.Close(SaveChanges:=True)
        Else
            objWord.Quit(SaveChanges:=True)
        End If
closeWord:
        objWord = Nothing

    End Sub

    Private Sub Provision_invoice()
        Dim Document As Document
        Dim objWord As Word.Application
        Dim wordWindow As Word.Window
        Dim table As Word.Table
        Dim row, bottomRow As Word.Row
        Dim subtotal_ExVAT, subtotal_NoVAT, subtotal As Double

        On Error GoTo Final

        'Create document
        objWord = CreateObject("Word.Application")
        Document = objWord.Documents.Add(Template:=GlobalValues.invoiceTemplate, Visible:=True)

        'Fill header
        AddHeader(Document:=Document, Factuurnummer:=Factuurnummer)

        'Fill table
        table = Document.Tables(2)
        table.Range.ParagraphFormat.KeepWithNext = True
        subtotal_ExVAT = AddWages(table:=table, kind:="provisie")
        subtotal_NoVAT = AddLitigation(table:=table, kind:="provisie")

        'Remove border of second row
        table.Rows(2).Borders(WdBorderType.wdBorderBottom).Visible = False

        ' Add Total
        bottomRow = table.Rows.Add

        REM insert if for not empty
        With bottomRow
            .Cells.Borders(WdBorderType.wdBorderVertical).Visible = False
            .Cells.Borders(WdBorderType.wdBorderBottom).Visible = False
            .Cells.Borders(WdBorderType.wdBorderLeft).Visible = False
            .Cells.Borders(WdBorderType.wdBorderRight).Visible = False
            .Range.ParagraphFormat.KeepWithNext = True
        End With

        If (subtotal_ExVAT <> 0) Then
            With table.Rows.Add
                .Cells(2).Merge(MergeTo:= .Cells(5))
                .Cells(2).Range.InsertAfter(Text:="Subtotaal excl Btw")
                .Cells(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                .Cells(3).Range.InsertAfter(Text:=Format(Expression:=subtotal_ExVAT, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
            End With

            With table.Rows.Add
                .Cells(2).Range.InsertAfter(Text:="Subtotaal Btw")
                .Cells(3).Range.InsertAfter(Text:=Format(Expression:=subtotal_ExVAT * 0.21, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
            End With
        End If

        If (subtotal_NoVAT <> 0) Then
            With table.Rows.Add
                .Cells(2).Range.InsertAfter(Text:="Subtotaal derden en gerechtskosten")
                .Cells(3).Range.InsertAfter(Text:=Format(Expression:=subtotal_NoVAT, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
            End With
        End If

        If (subtotal_NoVAT <> 0) Or (subtotal_ExVAT <> 0) Then
            With table.Rows.Add
                .Cells(2).Range.InsertAfter(Text:="Totaal")
                .Cells(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                .Cells(3).Range.InsertAfter(Text:=Format(Expression:=subtotal_ExVAT * 1.21 + subtotal_NoVAT, Style:=GlobalValues.NumberFormat))
                .Cells(3).Borders(WdBorderType.wdBorderTop).Visible = True
                .Cells(2).Range.Font.Bold = True
                .Cells(3).Range.Font.Bold = True
                .Range.ParagraphFormat.KeepWithNext = True
            End With
        End If

        'add border under table
        Dim newTable As Table
        newTable = table.Split(bottomRow.Index + 1)
        bottomRow.Borders(WdBorderType.wdBorderTop).Visible = True
        bottomRow.Delete()
        table.Columns.SetWidth(ColumnWidth:=ObjExcel.CentimetersToPoints(2.25), RulerStyle:=WdRulerStyle.wdAdjustNone)
        table.Columns(1).SetWidth(ColumnWidth:=ObjExcel.CentimetersToPoints(5.25), RulerStyle:=WdRulerStyle.wdAdjustNone)
        table.Columns(2).SetWidth(ColumnWidth:=ObjExcel.CentimetersToPoints(1.5), RulerStyle:=WdRulerStyle.wdAdjustNone)
        table.Columns(6).SetWidth(ColumnWidth:=ObjExcel.CentimetersToPoints(3), RulerStyle:=WdRulerStyle.wdAdjustNone)

        newTable.Columns.SetWidth(ColumnWidth:=ObjExcel.CentimetersToPoints(1.47), RulerStyle:=WdRulerStyle.wdAdjustNone)
        newTable.Columns(2).SetWidth(ColumnWidth:=ObjExcel.CentimetersToPoints(9.5), RulerStyle:=WdRulerStyle.wdAdjustNone)
        newTable.Columns(3).SetWidth(ColumnWidth:=ObjExcel.CentimetersToPoints(5.25), RulerStyle:=WdRulerStyle.wdAdjustNone)


Final:
        'Log invoice
        logInvoice(provisie:=True)

        objWord.Visible = True
        Document.PrintPreview()
        MsgBox("Kijk de factuur na")

        Document.SaveAs2(FileName:="i:\facturen\fa" + Factuurnummer)
        objWord.ActivePrinter = "Standaard"
        Document.PrintOut(Background:=True)
        objWord.ActivePrinter = "Standaard"
        On Error GoTo closeWord
        If objWord.Documents.Count > 1 Then
            Document.Close(SaveChanges:=True)
        Else
            objWord.Quit(SaveChanges:=True)
        End If
closeWord:
        objWord = Nothing
    End Sub

    REM checked


    REM checked
    Private Sub AddTitleRow(ByRef row As Row, ByVal title As String)
        With row
            .Cells(1).Range.InsertAfter(Text:=title)
            .Cells(1).Range.Font.Bold = True
            .Cells(1).Range.Font.Underline = WdUnderline.wdUnderlineSingle
            .Range.ParagraphFormat.KeepWithNext = True
        End With
    End Sub

    REM checked
    Private Sub AddSubtotal(ByRef table As Table, ByVal Total As Double)
        Dim totalrow As Row
        If Total <> 0 Then
            totalrow = table.Rows.Add
            totalrow.Cells(6).Range.InsertAfter(Format(Expression:=Total, Style:=GlobalValues.NumberFormat))

            'add empty row after total
            table.Rows.Add.Range.ParagraphFormat.KeepWithNext = False

            'format totalrow
            With totalrow
                .Cells(1).Range.InsertAfter("       subtotaal:")
                With .Cells(6).Borders(WdBorderType.wdBorderTop)
                    .Color = WdColor.wdColorBlack
                    .Visible = True
                End With
                .Range.ParagraphFormat.KeepWithNext = True
            End With
        End If
    End Sub

    REM Checked
    Private Sub AddHeader(ByRef Document As Document, ByVal Factuurnummer As String)
        On Error Resume Next

        'Fill header
        Document.CustomDocumentProperties("AdresBlok").Value = ogm_Record.Cells(1, 5).Value & " " & ogm_Record.Cells(1, 6).Value & vbCr &
                                                           ogm_Record.Cells(1, 7).Value & vbCr &
                                                           ogm_Record.Cells(1, 8).Value
        Document.CustomDocumentProperties("FactuurNummer").Value = Factuurnummer
        Document.CustomDocumentProperties("FactuurDatum").Value = Format(Expression:=Now, Style:="d mmmm yyyy")
        Document.CustomDocumentProperties("Vervaldatum").Value = Format(Expression:=DateAdd(Interval:=DateInterval.Month, Number:=1.0, DateValue:=Now), Style:="d mmmm yyyy")
        Document.CustomDocumentProperties("Dossier").Value = ogm_Record.Cells(1, 4).Value
        Document.CustomDocumentProperties("DossierNummer").Value = ogm_Record.Cells(1, 3).Value

        'Update fields
        Document.Fields.Update()
    End Sub

    REM Checked
    Private Function AddWages(ByRef table As Table, ByVal kind As String) As Double
        Dim subtotal As Double
        Dim titleRow As Row

        subtotal = 0

        'add title row
        titleRow = table.Rows.Add

        Select Case kind
            Case Is = "provisie"
                'add wages no details because provision
                If ereloonFct <> 0 Then
                    With table.Rows.Add
                        .Borders(WdBorderType.wdBorderTop).Visible = False
                        .Cells(1).Range.InsertAfter("- Provisie erelonen:")
                        .Cells(4).Range.InsertAfter(Format(Expression:=ereloonFct, Style:=GlobalValues.NumberFormat))
                        .Cells(5).Range.InsertAfter(Format(Expression:=ereloonFct * 0.21, Style:=GlobalValues.NumberFormat))
                        .Cells(6).Range.InsertAfter(Format(Expression:=ereloonFct * 1.21, Style:=GlobalValues.NumberFormat))
                        .Range.ParagraphFormat.KeepWithNext = True
                        subtotal = subtotal + ereloonFct
                    End With
                End If

            Case Is = "ereloon"
                'add wages
                Dim wages, pHours, pMinutes As Double
                pMinutes = CDbl(ogm_Record.Cells(1, 17).Value)
                pHours = CDbl(ogm_Record.Cells(1, 16).Value)
                wages = (pMinutes / 60 + pHours) * kostenSchema.prestaties

                If wages <> 0 Then
                    With table.Rows.Add
                        .Borders(WdBorderType.wdBorderTop).Visible = False
                        .Cells(1).Range.InsertAfter("- erelonen:")
                        .Cells(2).Range.InsertAfter(Format(Expression:=pHours, Style:="#0") &
                                                ":" & Format(Expression:=pMinutes, Style:="00"))
                        .Cells(3).Range.InsertAfter(Format(Expression:=kostenSchema.prestaties, Style:=GlobalValues.NumberFormat))
                        .Cells(4).Range.InsertAfter(Format(Expression:=wages, Style:=GlobalValues.NumberFormat))
                        .Cells(5).Range.InsertAfter(Format(Expression:=wages * 0.21, Style:=GlobalValues.NumberFormat))
                        .Cells(6).Range.InsertAfter(Format(Expression:=wages * 1.21, Style:=GlobalValues.NumberFormat))
                        .Range.ParagraphFormat.KeepWithNext = True
                        subtotal = subtotal + wages
                    End With
                End If

                'add waiting cost
                Dim wait, wHours, wMinutes As Double
                wMinutes = CDbl(ogm_Record.Cells(1, 19).Value)
                wHours = CDbl(ogm_Record.Cells(1, 18).Value)
                wait = (wMinutes / 60 + wHours) * kostenSchema.wacht

                If wait <> 0 Then
                    With table.Rows.Add
                        .Borders(WdBorderType.wdBorderBottom).Visible = False
                        .Cells(1).Range.InsertAfter("- verplaatsingen/wachttijden:")
                        .Cells(2).Range.InsertAfter(Format(Expression:=wHours, Style:="#0") &
                                            ":" & Format(Expression:=wMinutes, Style:="00"))
                        .Cells(3).Range.InsertAfter(Format(Expression:=kostenSchema.wacht, Style:=GlobalValues.NumberFormat))
                        .Cells(4).Range.InsertAfter(Format(Expression:=wait, Style:=GlobalValues.NumberFormat))
                        .Cells(5).Range.InsertAfter(Format(Expression:=wait * 0.21, Style:=GlobalValues.NumberFormat))
                        .Cells(6).Range.InsertAfter(Format(Expression:=wait * 1.21, Style:=GlobalValues.NumberFormat))
                        .Range.ParagraphFormat.KeepWithNext = True
                    End With
                    subtotal = subtotal + wait
                End If

            Case Else

        End Select

        'Add subtotal
        AddSubtotal(table:=table, Total:=subtotal * 1.21)
        If subtotal <> 0 Then
            AddTitleRow(row:=titleRow, title:="Erelonen")
        End If

        AddWages = subtotal
    End Function

    Private Function AddOfficeExpenses(ByRef table As Table) As Double
        Dim titleRow As Row
        Dim subtotal As Double
        Dim searchSheet As Worksheet
        Dim SearchTable As ListObject



        subtotal = 0

        'add header row
        titleRow = table.Rows.Add

        searchSheet = Workbook1.Sheets("Ereloon Nota")
        SearchTable = searchSheet.ListObjects("Ereloon_Nota_Table")

        'Add Dactylo
        If factuurData.dactylo <> 0 Then
            With table.Rows.Add
                Dim dactylo As Double
                dactylo = factuurData.dactylo * kostenSchema.dactylo
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- briefwisseling / dactylo:")
                .Cells(2).Range.InsertAfter(Format(Expression:=factuurData.dactylo, Style:="#0"))
                .Cells(3).Range.InsertAfter(Format(Expression:=kostenSchema.dactylo, Style:=GlobalValues.NumberFormat))
                .Cells(4).Range.InsertAfter(Format(Expression:=dactylo, Style:=GlobalValues.NumberFormat))
                .Cells(5).Range.InsertAfter(Format(Expression:=dactylo * 0.21, Style:=GlobalValues.NumberFormat))
                .Cells(6).Range.InsertAfter(Format(Expression:=dactylo * 1.21, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal + dactylo
            End With
        End If

        'Add fotocopie
        If factuurData.fotokopies <> 0 Then
            With table.Rows.Add
                Dim fotokopies As Double
                fotokopies = factuurData.fotokopies * kostenSchema.fotokopie
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- fotokopie:")
                .Cells(2).Range.InsertAfter(Format(Expression:=factuurData.fotokopies, Style:="#0"))
                .Cells(3).Range.InsertAfter(Format(Expression:=kostenSchema.fotokopie, Style:=GlobalValues.NumberFormat))
                .Cells(4).Range.InsertAfter(Format(Expression:=fotokopies, Style:=GlobalValues.NumberFormat))
                .Cells(5).Range.InsertAfter(Format(Expression:=fotokopies * 0.21, Style:=GlobalValues.NumberFormat))
                .Cells(6).Range.InsertAfter(Format(Expression:=fotokopies * 1.21, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal + fotokopies
            End With
        End If

        'Add Fax or incomming e-mail
        If factuurData.fax <> 0 Then
            With table.Rows.Add
                Dim fax As Double
                fax = factuurData.fax * kostenSchema.mail
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- inkomende fax of mail:")
                .Cells(2).Range.InsertAfter(Format(Expression:=factuurData.fax, Style:="#0"))
                .Cells(3).Range.InsertAfter(Format(Expression:=kostenSchema.mail, Style:=GlobalValues.NumberFormat))
                .Cells(4).Range.InsertAfter(Format(Expression:=fax, Style:=GlobalValues.NumberFormat))
                .Cells(5).Range.InsertAfter(Format(Expression:=fax * 0.21, Style:=GlobalValues.NumberFormat))
                .Cells(6).Range.InsertAfter(Format(Expression:=fax * 1.21, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal + fax
            End With
        End If

        'Add displacements
        If factuurData.verplaatsing <> 0 Then
            With table.Rows.Add
                Dim verplaatsing As Double
                verplaatsing = factuurData.verplaatsing * kostenSchema.verplaatsing
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- verplaatsingen (km):")
                .Cells(2).Range.InsertAfter(Format(Expression:=factuurData.verplaatsing, Style:="#0"))
                .Cells(3).Range.InsertAfter(Format(Expression:=kostenSchema.verplaatsing, Style:=GlobalValues.NumberFormat))
                .Cells(4).Range.InsertAfter(Format(Expression:=verplaatsing, Style:=GlobalValues.NumberFormat))
                .Cells(5).Range.InsertAfter(Format(Expression:=verplaatsing * 0.21, Style:=GlobalValues.NumberFormat))
                .Cells(6).Range.InsertAfter(Format(Expression:=verplaatsing * 1.21, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal + verplaatsing
            End With
        End If

        'add additional costs
        If factuurData.bijkomende_kosten <> 0 Then
            With table.Rows.Add
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- andere kostenen:")
                .Cells(4).Range.InsertAfter(Format(Expression:=factuurData.bijkomende_kosten, Style:=GlobalValues.NumberFormat))
                .Cells(5).Range.InsertAfter(Format(Expression:=factuurData.bijkomende_kosten * 0.21, Style:=GlobalValues.NumberFormat))
                .Cells(6).Range.InsertAfter(Format(Expression:=factuurData.bijkomende_kosten * 1.21, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal + factuurData.bijkomende_kosten
            End With
        End If

        'add forfait
        If factuurData.forfait <> 0 Then
            With table.Rows.Add
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- opstarten dossier:")
                .Cells(4).Range.InsertAfter(Format(Expression:=factuurData.forfait, Style:=GlobalValues.NumberFormat))
                .Cells(5).Range.InsertAfter(Format(Expression:=factuurData.forfait * 0.21, Style:=GlobalValues.NumberFormat))
                .Cells(6).Range.InsertAfter(Format(Expression:=factuurData.forfait * 1.21, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal + factuurData.forfait
            End With
        End If


        'Add subtotal
        AddSubtotal(table:=table, Total:=subtotal * 1.21)
        If subtotal <> 0 Then
            AddTitleRow(row:=titleRow, title:="Bureelkosten:")
        End If

        AddOfficeExpenses = subtotal
    End Function

    Private Function AddPayedProvisions(ByRef table As Table) As Double
        Dim titleRow As Row
        Dim subtotal As Double
        Dim sht As Worksheet
        Dim tbl As ListObject

        On Error GoTo Final

        sht = Workbook1.Sheets("Provisies")
        tbl = sht.ListObjects("Provisie_Table")

        sht.Unprotect(Password:=GlobalValues.password)
        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()
        tbl.Range.AutoFilter(Field:=3, Criteria1:=dossierNr)
        tbl.AutoFilter.ApplyFilter()
        Prov_Erelonen = tbl.TotalsRowRange.Cells(1, tbl.ListColumns("Ereloon_betaald").Index)
        Prov_Gerecht = tbl.TotalsRowRange.Cells(1, tbl.ListColumns("gerechtskosten_betaald").Index)
        Prov_BTW = Prov_Erelonen * 0.21
        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()

        subtotal = 0

        'add header row
        titleRow = table.Rows.Add

        'Add wages invoiced
        If Prov_Erelonen <> 0 Then
            With table.Rows.Add
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- al gefact erelonen:")
                .Cells(6).Range.InsertAfter(Format(Expression:=-Prov_Erelonen, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal + Prov_Erelonen
            End With

            'Add scheduling cost
            With table.Rows.Add
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- al gefact BTW:")
                .Cells(6).Range.InsertAfter(Format(Expression:=-Prov_BTW, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal + Prov_BTW
            End With
        End If

        'add summons
        If Prov_Gerecht <> 0 Then
            With table.Rows.Add
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- al gefact. provisies ")
                .Cells(6).Range.InsertAfter(Format(Expression:=-Prov_Gerecht, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal + Prov_Gerecht
            End With
        End If

        'Add subtotal
        AddSubtotal(table:=table, Total:=-subtotal)
        If subtotal <> 0 Then
            AddTitleRow(row:=titleRow, title:="Al gefactureerd:")
        End If

        AddPayedProvisions = subtotal

Final:
        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)

    End Function

    Private Function AddLitigation(ByRef table As Table, ByVal kind As String) As Double
        Dim titleRow As Row
        Dim subtotal As Double
        Dim sht As Worksheet
        Dim tbl As ListObject

        subtotal = 0

        'add header row
        titleRow = table.Rows.Add

        Select Case kind
        ' Add litigation costs
            Case Is = "provisie"
                If gerechtskostenFct <> 0 Then
                    With table.Rows.Add
                        .Borders(WdBorderType.wdBorderTop).Visible = False
                        .Cells(1).Range.InsertAfter("- gerechtskosten:")
                        .Cells(6).Range.InsertAfter(Format(Expression:=gerechtskostenFct, Style:=GlobalValues.NumberFormat))
                        .Range.ParagraphFormat.KeepWithNext = True
                        subtotal = subtotal + gerechtskostenFct
                    End With
                End If

            Case Is = "ereloon"

                sht = Workbook1.Sheets("Ereloon Nota")
                tbl = sht.ListObjects("Ereloon_Nota_Table")

                'Add scheduling cost
                If factuurData.rolzetting <> 0 Then
                    With table.Rows.Add
                        .Borders(WdBorderType.wdBorderTop).Visible = False
                        .Cells(1).Range.InsertAfter("- rolzetting:")
                        .Cells(6).Range.InsertAfter(Format(Expression:=factuurData.rolzetting, Style:=GlobalValues.NumberFormat))
                        .Range.ParagraphFormat.KeepWithNext = True
                        subtotal = subtotal + factuurData.rolzetting
                    End With
                End If

                'add summons
                If factuurData.dagvaarding <> 0 Then
                    With table.Rows.Add
                        .Borders(WdBorderType.wdBorderTop).Visible = False
                        .Cells(1).Range.InsertAfter("- dagvaardingen:")
                        .Cells(6).Range.InsertAfter(Format(Expression:=factuurData.dagvaarding, Style:=GlobalValues.NumberFormat))
                        .Range.ParagraphFormat.KeepWithNext = True
                        subtotal = subtotal + factuurData.dagvaarding
                    End With
                End If

                'add signification
                If factuurData.betekening <> 0 Then
                    With table.Rows.Add
                        .Borders(WdBorderType.wdBorderTop).Visible = False
                        .Cells(1).Range.InsertAfter("- betekeningen:")
                        .Cells(6).Range.InsertAfter(Format(Expression:=factuurData.betekening, Style:=GlobalValues.NumberFormat))
                        .Range.ParagraphFormat.KeepWithNext = True
                        subtotal = subtotal + factuurData.betekening
                    End With
                End If

                'add execution costs
                If factuurData.uitvoering <> 0 Then
                    With table.Rows.Add
                        .Borders(WdBorderType.wdBorderTop).Visible = False
                        .Cells(1).Range.InsertAfter("- uitvoering:")
                        .Cells(6).Range.InsertAfter(Format(Expression:=factuurData.uitvoering, Style:=GlobalValues.NumberFormat))
                        .Range.ParagraphFormat.KeepWithNext = True
                        subtotal = subtotal + factuurData.uitvoering
                    End With
                End If

                'add varia
                If factuurData.andere <> 0 Then
                    With table.Rows.Add
                        .Borders(WdBorderType.wdBorderTop).Visible = False
                        .Cells(1).Range.InsertAfter("- andere:")
                        .Cells(6).Range.InsertAfter(Format(Expression:=factuurData.andere, Style:=GlobalValues.NumberFormat))
                        .Range.ParagraphFormat.KeepWithNext = True
                        subtotal = subtotal + factuurData.andere
                    End With
                End If

            Case Else

        End Select

        'Add subtotal
        AddSubtotal(table:=table, Total:=subtotal)
        If subtotal <> 0 Then
            AddTitleRow(row:=titleRow, title:="Gerechts- en andere kosten:")
        End If

        AddLitigation = subtotal
    End Function

    Private Function AddProvision(ByRef table As Table) As Double
        Dim titleRow As Row
        Dim subtotal As Double

        subtotal = 0

        'add header row
        titleRow = table.Rows.Add

        'Add third-party funds
        If factuurData.derden <> 0 Then
            With table.Rows.Add
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- derdengelden:")
                .Cells(6).Range.InsertAfter(Format(Expression:=-factuurData.derden, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal - factuurData.derden
            End With
        End If

        'add provions wages
        If factuurData.provisie_erelonen <> 0 Then
            With table.Rows.Add
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- provisie erelonen:")
                .Cells(6).Range.InsertAfter(Format(Expression:=-factuurData.provisie_erelonen, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal - factuurData.provisie_erelonen
            End With
        End If

        'add provisons litigation
        If factuurData.provisie_gerechtskosten <> 0 Then
            With table.Rows.Add
                .Borders(WdBorderType.wdBorderTop).Visible = False
                .Cells(1).Range.InsertAfter("- provisie gerechtskosten:")
                .Cells(6).Range.InsertAfter(Format(Expression:=-factuurData.provisie_gerechtskosten, Style:=GlobalValues.NumberFormat))
                .Range.ParagraphFormat.KeepWithNext = True
                subtotal = subtotal - factuurData.provisie_gerechtskosten
            End With
        End If

        'Add subtotal
        AddSubtotal(table:=table, Total:=subtotal)
        If subtotal <> 0 Then
            AddTitleRow(row:=titleRow, title:="Provisies en derdengelden:")
        End If

        AddProvision = subtotal
    End Function

    Private Sub logInvoice(ByVal provisie As Boolean)
        Dim sht As Worksheet
        Dim tbl As ListObject
        Dim i As Integer

        On Error GoTo Final

        sht = Workbook1.Sheets("Facturen")
        tbl = sht.ListObjects("Invoice_table")

        sht.Unprotect(Password:=GlobalValues.password)

        With tbl.ListRows.Add.Range
            'Fill general information
            .Cells(1) = Now
            For i = 2 To 8
                .Cells(i) = ogm_Record.Cells(i).Value
            Next

            If provisie Then
                .Cells(27) = ereloonFct
                .Cells(28) = gerechtskostenFct
                .Cells(29) = ereloonFct * 1.21 + gerechtskostenFct
            Else
                For i = 9 To 26
                    .Cells(i) = ogm_Record.Cells(i).Value
                Next
                .Cells(29) = Total
            End If

            'insert reference lines
            .Cells(30) = Provisies
            .Cells(31) = erelonen
            .Cells(32) = Factuurnummer

        End With

Final:
        If (sht.ProtectContents = False) Then
            sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
        End If

    End Sub

    Private Function NextInvoiceNumber() As String
        Dim sht As Worksheet
        Dim Factuurnummer As Integer

        On Error GoTo Final

        sht = Workbook1.Sheets("Parameters")
        sht.Unprotect(Password:=GlobalValues.password)

        Factuurnummer = Workbook1.Names("FactuurNummer").RefersToRange.Cells(1).Value
        Workbook1.Names("FactuurNummer").RefersToRange.Cells(1) = Factuurnummer + 1

        NextInvoiceNumber = Format(Year(Now()), "0000") & Format(Factuurnummer, "00000")

Final:
        If (sht.ProtectContents = False) Then
            sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
        End If

    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                inputInvoice.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region


End Class
