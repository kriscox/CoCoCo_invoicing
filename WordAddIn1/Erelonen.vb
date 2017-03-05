Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

Public Class Erelonen
    REM versie van 20160702

    Private Structure str_Erelonen_bureelkosten
        Dim Kostenschema As Integer
        Dim Dactylo As Integer
        Dim Fotokopies As Integer
        Dim Fax As Integer
        Dim verplaatsing As Integer
        Dim Bijkomende_kosten As Double
        Dim Forfait As Double
        Dim erelonen_uren As Integer
        Dim erelonen_minuten As Integer
        Dim wacht_uren As Integer
        Dim wacht_minuten As Integer
    End Structure

    Private Structure str_gerechtskosten
        Dim Rolzetting As Double
        Dim Dagvaarding As Double
        Dim Betekening As Double
        Dim Uitvoering As Double
        Dim Andere As Double
    End Structure

    Dim KostenschemaList()() As String
    Dim index As Integer
    Dim dossierNr, dossierName, OGMCode As String
    Dim aanspreektitel, Name, Adres1, adres2 As String
    Dim Provisies_erelonen, Provisies_GerechtsKosten As Double
    Dim Derden, btw, Totaal, Subtotaal As Double
    Dim Amount_Schemas As Integer
    Dim ExWb As Workbook = GlobalValues.GetWorkbook()
    Dim gerechtskosten As str_gerechtskosten
    Dim Erelonen_bureelkosten As str_Erelonen_bureelkosten

    Private Function readFile() As Boolean
        Dim FileName As String
        Dim Result As Boolean
        Dim Line As String = ""
        Dim lineItems As String()

        On Error GoTo End_Routine

        Result = False
        FileName = Environ$("temp") & "\judaimp.csv"
        FileOpen(1, FileName, OpenMode.Input)

        If Not EOF(1) Then
            Input(1, Line)
            lineItems = Split(Line, ";")
            dossierNr = lineItems(0)
            dossierName = lineItems(1)
            aanspreektitel = lineItems(2)
            Name = lineItems(3)
            Adres1 = lineItems(4)
            adres2 = lineItems(5)
            Result = True
        End If

End_Routine:
        FileClose(1)
        readFile = Result

    End Function

    Private Function requestInputs() As Boolean

        'run Input_Form
        Dim Input_Form As New Ereloon_Nota_form
        Dim username As String

        username = Globals.CoCoCo_Invoicing.Application.UserInitials

        Totaal = 0
        With Input_Form
            '------------------------------------------------------
            'Put provisies's in place
            '------------------------
            .Provisies.Text = Format(Provisies_erelonen, "€## ##0.00")
            .Provisies2.Text = Format(Provisies_GerechtsKosten, "€## ##0.00")
            '------------------------------------------------------
            'Put kostenschema's in place
            '---------------------------
            .Kostenschema.DataSource = KostenschemaList
            '.Kostenschema. = UBound(KostenschemaList, 2)
            '.Kostenschema.ColumnWidths = "130;0;0;0;0;0;0"
            '.Kostenschema.Text = KostenschemaList(0)(0)
            '------------------------------------------------------
            'Show form
            '---------
            .Show()
            '------------------------------------------------------
        End With

        If (Input_Form.Tag <> "Cancelled") Then
            '--------------------------------------------------
            'Read Erelonen_bureelkosten Values
            '---------------------------------
            index = Input_Form.Kostenschema.SelectedValue
            With Erelonen_bureelkosten
                .Kostenschema = KostenschemaList(index)(7)

                .Dactylo = Input_Form.Dactylo.Text
                Subtotaal = Subtotaal + .Dactylo * KostenschemaList(index)(6)

                .Fotokopies = Input_Form.Fotokopies.Text
                Subtotaal = Subtotaal + .Fotokopies * KostenschemaList(index)(5)

                .Fax = Input_Form.Fax.Text
                Subtotaal = Subtotaal + .Fax * KostenschemaList(index)(4)

                .verplaatsing = Input_Form.Verplaatsingen.Text
                Subtotaal = Subtotaal + .verplaatsing * KostenschemaList(index)(3)

                .Bijkomende_kosten = Input_Form.Bijkomende_kosten.Text
                Subtotaal = Subtotaal + .Bijkomende_kosten

                .Forfait = Input_Form.Forfait.Text
                Subtotaal = Subtotaal + .Forfait

                .erelonen_uren = Input_Form.Consultaties_uren.Text
                .erelonen_minuten = Input_Form.Consultaties_minuten.Text
                Subtotaal = Subtotaal + (.erelonen_uren + .erelonen_minuten / 60) * KostenschemaList(index)(1)

                .wacht_uren = Input_Form.Verplaatsingen_2_uren.Text
                .wacht_minuten = Input_Form.Verplaatsingen_2_minuten.Text
                Subtotaal = Subtotaal + (.wacht_uren + .wacht_minuten / 60) * KostenschemaList(index)(2)

                btw = Subtotaal * 0.21
            End With
            '--------------------------------------------------
            'Read gerechtskosten Values
            '--------------------------
            With gerechtskosten
                Totaal = Subtotaal + btw
                .Betekening = Input_Form.Betekening.Text
                Totaal = Totaal + .Betekening
                .Dagvaarding = Input_Form.Dagvaarding.Text
                Totaal = Totaal + .Dagvaarding
                .Rolzetting = Input_Form.Rolzetting.Text
                Totaal = Totaal + .Rolzetting
                .Uitvoering = Input_Form.Uitvoering.Text
                Totaal = Totaal + .Uitvoering
                .Andere = Input_Form.Andere.Text
                Totaal = Totaal + .Andere
            End With
            Derden = Input_Form.Derden.Text
            requestInputs = True
        Else
            requestInputs = False
        End If

        Input_Form.Close()
        GoTo end_of_function

        On Error Resume Next
        Input_Form.Hide()
        Input_Form.Close()
        requestInputs = False

end_of_function:
    End Function

    Private Function ReadFromExcel() As Boolean
        Dim Lst As ListRows
        Dim sht As Worksheet
        Dim tbl As ListObject
        Dim rng As Excel.Range
        Dim rowRng As Excel.Range
        Dim row As Excel.Range
        Dim Number, Serial_Number As Double
        Dim CountDossier As Double
        Dim i As Integer = 0

        On Error GoTo ErrorHandler

        '------------------------------------------------------
        'read kostenschemas
        '------------------
        sht = ExWb.Sheets("Kostenschemas")
        tbl = sht.ListObjects("Kostenschema")
        sht.Unprotect(Password:=GlobalValues.password)

        'remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()
        tbl.Range.AutoFilter(Field:=9, Criteria1:=False)
        rowRng = tbl.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows
        'Read all filtered data
        ReDim KostenschemaList(rowRng.Count - 1)(7)
        Amount_Schemas = rowRng.Count
        On Error Resume Next
        For Each row In tbl.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows
            With row
                KostenschemaList(i)(0) = .Cells(2).Value
                KostenschemaList(i)(1) = .Cells(3).Value
                KostenschemaList(i)(2) = .Cells(4).Value
                KostenschemaList(i)(3) = .Cells(5).Value
                KostenschemaList(i)(4) = .Cells(6).Value
                KostenschemaList(i)(5) = .Cells(7).Value
                KostenschemaList(i)(6) = .Cells(8).Value
                KostenschemaList(i)(7) = .Cells(1).Value
                i = i + 1
            End With
        Next
        'Remove the autofilter
        tbl.AutoFilter.ShowAllData()
        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)


        '------------------------------------------------------
        'read provisies
        '------------------
        Dim Provisies_Erelonen_VAT, Provisies_Erelonen_ExVAT As Double
        sht = ExWb.Sheets("Provisies")
        tbl = sht.ListObjects("Provisie_Table")
        Lst = tbl.ListRows
        sht.Unprotect(Password:=GlobalValues.password)

        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()
        tbl.Range.AutoFilter(Field:=3, Criteria1:=dossierNr)
        tbl.AutoFilter.ApplyFilter()
        Provisies_Erelonen_ExVAT = sht.Evaluate("=Provisie_Table[[#Totals],[Ereloon_betaald]]")
        Provisies_Erelonen_VAT = sht.Evaluate("=Provisie_Table[[#Totals],[BTW_betaald]]")
        Provisies_erelonen = Provisies_Erelonen_ExVAT + Provisies_Erelonen_VAT
        Provisies_GerechtsKosten = sht.Evaluate("=Provisie_Table[[#Totals],[gerechtskosten_betaald]]")
        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()

        REM Tbl.Range.AutoFilter Field:=3, Criteria1:=DossierNr
        REM Tbl.Range.AutoFilter Field:=13, Criteria1:=false
        REM Tbl.AutoFilter.ApplyFilter
        REM Provisies1_Erelonen_ExVAT = Sht.Evaluate("=Provisie_Table[[#Totals],[Ereloon]]")
        REM Provisies1_Erelonen_VAT = Sht.Evaluate("=Provisie_Table[[#Totals],[BTW]]")
        REM Provisies1_erelonen = Provisies_Erelonen_ExVAT + Provisies_Erelonen_VAT
        REM Tbl.AutoFilter.ShowAllData

        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)

        ReadFromExcel = True
        Exit Function

ErrorHandler:
        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)
        ReadFromExcel = False
    End Function

    Private Function insert_text() As Boolean
        Dim text As String
        Dim tbl As Word.Table
        Dim row As Word.Row
        Dim rng As Word.Range
        Dim Subtotaal1 As Double
        Dim Selection As Word.Selection = Globals.CoCoCo_Invoicing.Application.Selection

        On Error GoTo ErrorHandler

        text = "Ik stelde de eindafrekening in dit dossier op. Het overzicht vindt u hieronder." + vbNewLine
        Selection.TypeText(text)

        '------------------------------------------------------
        'Bureel kosten en Erelonen
        '-------------------------
        rng = Selection.Range
        tbl = Globals.CoCoCo_Invoicing.Application.ActiveDocument.Tables.Add(Range:=rng, NumRows:=1, NumColumns:=7, DefaultTableBehavior:=WdDefaultTableBehavior.wdWord9TableBehavior)
        With tbl.Rows(1)
            .Cells(1).Range.Text = "Bureel kosten en Erelonen"
            .Borders(WdBorderType.wdBorderTop).Visible = False
            .Borders(WdBorderType.wdBorderLeft).Visible = False
            .Borders(WdBorderType.wdBorderRight).Visible = False
            .Borders(WdBorderType.wdBorderVertical).Visible = False
            .Cells(1).Borders(WdBorderType.wdBorderBottom).Visible = True
            .Cells(1).SetWidth(ColumnWidth:=Globals.CoCoCo_Invoicing.Application.CentimetersToPoints(7.5), RulerStyle:=WdRulerStyle.wdAdjustNone)
            .Cells(2).SetWidth(ColumnWidth:=Globals.CoCoCo_Invoicing.Application.CentimetersToPoints(1.3), RulerStyle:=WdRulerStyle.wdAdjustNone)
            .Cells(3).SetWidth(ColumnWidth:=Globals.CoCoCo_Invoicing.Application.CentimetersToPoints(1.3), RulerStyle:=WdRulerStyle.wdAdjustNone)
            .Cells(4).SetWidth(ColumnWidth:=Globals.CoCoCo_Invoicing.Application.CentimetersToPoints(0.5), RulerStyle:=WdRulerStyle.wdAdjustNone)
            .Cells(5).SetWidth(ColumnWidth:=Globals.CoCoCo_Invoicing.Application.CentimetersToPoints(2), RulerStyle:=WdRulerStyle.wdAdjustNone)
            .Cells(6).SetWidth(ColumnWidth:=Globals.CoCoCo_Invoicing.Application.CentimetersToPoints(0.5), RulerStyle:=WdRulerStyle.wdAdjustNone)
            .Cells(7).SetWidth(ColumnWidth:=Globals.CoCoCo_Invoicing.Application.CentimetersToPoints(3), RulerStyle:=WdRulerStyle.wdAdjustNone)
        End With

        With Erelonen_bureelkosten
            ' dactylo
            If (.Dactylo > 0) Then
                Insert_cost_amount_Row(tbl:=tbl, text:="- pagina's", Amount:= .Dactylo,
                UnitPrice:=KostenschemaList(index)(6), unit:="pag.")
            End If
            ' fotokopies
            If (.Fotokopies > 0) Then
                Insert_cost_amount_Row(tbl:=tbl, text:="- fotokopies", Amount:= .Fotokopies,
                UnitPrice:=KostenschemaList(index)(5), unit:="kop.")
            End If
            ' Inkomende mails/fax
            If (.Fax > 0) Then
                Insert_cost_amount_Row(tbl:=tbl, text:="- inkomen mails/fax", Amount:= .Fax,
                UnitPrice:=KostenschemaList(index)(4), unit:="")
            End If
            ' verplaatsingen
            If (.verplaatsing > 0) Then
                Insert_cost_amount_Row(tbl:=tbl, text:="- verplaatsingen", Amount:= .verplaatsing,
                UnitPrice:=KostenschemaList(index)(3), unit:="km.")
            End If
            ' andere kosten
            If (.Bijkomende_kosten > 0) Then
                Insert_cost_Row(tbl:=tbl, text:="- bijkomende kosten", cost:= .Bijkomende_kosten)
            End If
            ' Dossierkosten
            If (.Forfait > 0) Then
                Insert_cost_Row(tbl:=tbl, text:="- dossierkosten", cost:= .Forfait)
            End If
            ' prestaties
            If (.erelonen_uren + .erelonen_minuten > 0) Then
                Insert_cost_amount_Row(tbl:=tbl, text:="- prestaties", Amount:= .erelonen_uren + .erelonen_minuten / 60,
                UnitPrice:=KostenschemaList(index)(1), unit:="uren", hours:=True)
            End If
            ' verplaatsignen/wachtduur
            If (.wacht_uren + .wacht_minuten > 0) Then
                Insert_cost_amount_Row(tbl:=tbl, text:="- verplaatsingen/wachtuur", Amount:= .wacht_uren + .wacht_minuten / 60,
                UnitPrice:=KostenschemaList(index)(2), unit:="uren", hours:=True)
            End If
        End With

        tbl.Rows(1).Cells.Merge()
        tbl.Rows(1).Borders(WdBorderType.wdBorderBottom).Visible = True

        '------------------------------------------------------
        'VAT and gerechtskosten
        '----------------------

        'subtotaal
        If Subtotaal > 0 Then
            row = Insert_cost_Row(tbl:=tbl, text:="Subtotaal", cost:=Subtotaal)
            row.Range.Font.Bold = True
            row.Borders(WdBorderType.wdBorderTop).Visible = True
        End If
        'VAT
        If btw > 0 Then
            row = Insert_cost_Row(tbl:=tbl, text:=" - 21 %BTW", cost:=btw)
            row.Range.Font.Bold = False
        End If
        With gerechtskosten
            'Rolzetting
            If .Rolzetting > 0 Then
                row = Insert_cost_Row(tbl:=tbl, text:="gerechtskosten(Rolzetting)", cost:= .Rolzetting)
                row.Range.Font.Bold = False
            End If
            'Dagvaarding
            If .Dagvaarding > 0 Then
                row = Insert_cost_Row(tbl:=tbl, text:="gerechtskosten(Dagvaarding)", cost:= .Dagvaarding)
                row.Range.Font.Bold = False
            End If
            'Betekening
            If .Betekening > 0 Then
                row = Insert_cost_Row(tbl:=tbl, text:="gerechtskosten(Betekening)", cost:= .Betekening)
                row.Range.Font.Bold = False
            End If
            'Uitvoering
            If .Uitvoering > 0 Then
                row = Insert_cost_Row(tbl:=tbl, text:="gerechtskosten(Uitvoering)", cost:= .Uitvoering)
                row.Range.Font.Bold = False
            End If
            'Andere
            If .Andere > 0 Then
                row = Insert_cost_Row(tbl:=tbl, text:="gerechtskosten(Andere)", cost:= .Andere)
                row.Range.Font.Bold = False
            End If
            'Algemeen totaal
            If Totaal > 0 Then
                row = Insert_cost_Row(tbl:=tbl, text:="algemeen totaal", cost:=Totaal)
                row.Range.Bold = True
                row.Borders(WdBorderType.wdBorderTop).Visible = True
            End If
            Subtotaal1 = Totaal
            ' Provisies
            If Provisies_erelonen + Provisies_GerechtsKosten > 0 Then
                row = Insert_cost_Row(tbl:=tbl, text:="ontvangen provisies", cost:=-Provisies_erelonen - Provisies_GerechtsKosten)
                Subtotaal1 = Subtotaal1 - Provisies_erelonen - Provisies_GerechtsKosten
                row.Range.Font.Bold = False
            End If
        End With

        ' Derden
        If Derden > 0 Then
            row = Insert_cost_Row(tbl:=tbl, text:="ontvangen van derden", cost:=-Derden)
            Subtotaal1 = Subtotaal1 - Derden
            row.Range.Font.Bold = False
        End If
        'Saldo
        If Math.Round(Subtotaal1, 2) > 0 Then
            row = Insert_cost_Row(tbl:=tbl, text:="te betalen saldo", cost:=Subtotaal1)
            text = vbNewLine + "U kunt dit bedrag overmaken op rekeningnummer BE96 0012 4751 7505 met als mededeling: " +
                OGMCode + "."

        ElseIf Math.Round(Subtotaal1, 2) < 0 Then
            row = Insert_cost_Row(tbl:=tbl, text:="uit te keren saldo", cost:=Subtotaal1)
            row.Range.Font.ColorIndex = WdColor.wdColorDarkRed
            text = vbNewLine + "Dit bedrag zal overgemaakt worden op uw rekening binnen de 3 maanden"
        Else
            row = Insert_cost_Row(tbl:=tbl, text:="Totaal", cost:=Subtotaal1)
            text = vbNewLine
        End If

        row.Range.Bold = True
        row.Borders(WdBorderType.wdBorderTop).Visible = True

        tbl.Select()
        Selection.EndOf(Unit:=WdUnits.wdTable, Extend:=WdMovementType.wdMove)
        Selection.Move(Unit:=WdUnits.wdCharacter, Count:=1)
        Selection.TypeText(text)

        insert_text = True
        Exit Function
ErrorHandler:
        MsgBox("Text wrongly inserted")
        insert_text = False
    End Function

    Private Function Insert_cost_Row(ByRef tbl As Word.Table, ByVal text As String, ByVal cost As Double) As Word.Row
        Dim row As Word.Row
        row = tbl.Rows.Add
        With row
            .Cells(1).Range.Text = text
            .Cells(7).Range.Text = Format(cost, "€ ## ##0.00")
            .Cells(7).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
            .Borders.Enable = False
        End With
        Insert_cost_Row = row
    End Function

    Private Function Insert_cost_amount_Row(ByRef tbl As Word.Table, ByVal text As String, ByVal unit As String,
                                    ByVal Amount As Double, ByVal UnitPrice As Double,
                                    Optional ByVal hours As Boolean = False) As Word.Row
        Dim row As Word.Row
        row = tbl.Rows.Add
        With row
            .Cells(1).Range.Text = text
            If hours Then
                Dim hh, mm As Integer
                hh = Int(Amount)
                mm = (Amount - hh) * 60
                .Cells(2).Range.Text = hh & ":" & Format(mm, "00")
            Else
                .Cells(2).Range.Text = Amount
            End If
            .Cells(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
            .Cells(3).Range.Text = unit
            .Cells(3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
            .Cells(4).Range.Text = "x"
            .Cells(5).Range.Text = Format(UnitPrice, "€ ## ##0.00")
            .Cells(5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
            .Cells(6).Range.Text = "="
            .Cells(7).Range.Text = Format(Amount * UnitPrice, "€ ## ##0.00")
            .Cells(7).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
            .Borders.Enable = False
        End With
        Insert_cost_amount_Row = row
    End Function
    Private Function InsertInExcel() As Boolean
        Dim Lst As ListRows
        Dim sht As Worksheet
        Dim tbl As ListObject
        Dim rng As Excel.Range
        Dim row As Excel.Range
        Dim Number, Serial_Number As Integer
        Dim CountDossier As Double

        On Error GoTo ErrorHandler

        sht = ExWb.Sheets("Ereloon Nota")
        sht.Unprotect(Password:=GlobalValues.password)
        tbl = sht.ListObjects("Ereloon_Nota_Table")
        Lst = tbl.ListRows

        Lst.Add(Position:=1)
        On Error Resume Next
        With Lst(1)
            .Range(1).Value = Now
            .Range(2).Value = Globals.CoCoCo_Invoicing.Application.UserInitials
            .Range(3).Value = dossierNr
            .Range(4).Value = dossierName
            .Range(5).Value = aanspreektitel
            .Range(6).Value = Name
            .Range(7).Value = Adres1
            .Range(8).Value = adres2
            .Range(9).Value = Erelonen_bureelkosten.Kostenschema
            .Range(10).Value = Erelonen_bureelkosten.Dactylo
            .Range(11).Value = Erelonen_bureelkosten.Fotokopies
            .Range(12).Value = Erelonen_bureelkosten.Fax
            .Range(13).Value = Erelonen_bureelkosten.verplaatsing
            .Range(14).Value = Erelonen_bureelkosten.Bijkomende_kosten
            .Range(15).Value = Erelonen_bureelkosten.Forfait
            .Range(16).Value = Erelonen_bureelkosten.erelonen_uren
            .Range(17).Value = Erelonen_bureelkosten.erelonen_minuten
            .Range(18).Value = Erelonen_bureelkosten.wacht_uren
            .Range(19).Value = Erelonen_bureelkosten.wacht_minuten
            .Range(20).Value = btw
            .Range(21).Value = gerechtskosten.Rolzetting
            .Range(22).Value = gerechtskosten.Dagvaarding
            .Range(23).Value = gerechtskosten.Betekening
            .Range(24).Value = gerechtskosten.Uitvoering
            .Range(25).Value = gerechtskosten.Andere
            .Range(26).Value = Derden
            .Range(27).Value = Provisies_erelonen
            .Range(28).Value = Provisies_GerechtsKosten
            .Range(29).Value = Totaal
            .Range(30).Value = False
            .Range(31).Value = OGMCode
        End With

        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)

        InsertInExcel = True

        REM Sluit provisies af
        sht = ExWb.Sheets("Provisies")
        tbl = sht.ListObjects("Provisie_Table")
        Lst = tbl.ListRows
        sht.Unprotect(Password:=GlobalValues.password)

        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()
        tbl.Range.AutoFilter(Field:=3, Criteria1:=dossierNr)
        tbl.AutoFilter.ApplyFilter()

        rng = tbl.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible)
        For Each row In rng.Rows
            row.Cells(tbl.ListColumns("betaald").Index) = True
        Next row

        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()
        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)

        Exit Function

ErrorHandler:
        InsertInExcel = False
    End Function


    Public Function main() As Boolean
        Dim error_text As String

        If Not readFile() Then
            error_text = "CSV file not read"
            GoTo Exit_error
        ElseIf Not ReadFromExcel() Then
            error_text = "Error reading from Excel"
            GoTo Exit_error
        ElseIf Not requestInputs() Then
            error_text = "Error in Form"
            GoTo Exit_error
        Else
            OGMCode = GlobalValues.CoCoCo_Calculate_OGM(dossierNr)
            If Not InsertInExcel() Then
                error_text = "Error inserting in excel"
                GoTo Exit_error
            ElseIf Not insert_text() Then
                error_text = "Error inserting text"
                GoTo Exit_error
            End If
        End If

        main = True
        Exit Function

Exit_error:
        MsgBox(error_text)
        main = False
    End Function

End Class
