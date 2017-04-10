Imports Microsoft.Office.Interop.Excel

Public Class Provisie
#Region "Variables"
    Dim dossierNr, dossierName As String
    Dim aanspreektitel, Name As String
    Dim Adres1, adres2 As String
    Dim Prov1, btw, Prov2 As Double
    Dim erelonen, gerechtskosten, Totaal As Double
    Dim OGMCode As String
    Dim IC As Boolean = False
    Dim ExWb As Workbook = GlobalValues.GetWorkbook
#End Region

    Private Function ReadFile() As Boolean
        Dim FileName As String
        Dim Result As Boolean
        Dim Line As String = Nothing
        Dim LineItems As String()

        On Error GoTo End_Routine

        Result = False
        FileName = Environ$("temp") & "\judaimp.csv"
        FileOpen(FileNumber:=1, FileName:=FileName, Mode:=OpenMode.Input)

        If Not EOF(1) Then
            Line = LineInput(1)
            LineItems = Split(Line, ";")
            dossierNr = LineItems(0)
            dossierName = LineItems(1)
            aanspreektitel = LineItems(2)
            Name = LineItems(3)
            Adres1 = LineItems(4)
            adres2 = LineItems(5)
            Result = True
        End If

End_Routine:
        FileClose(1)

        ReadFile = Result

    End Function

    Private Function RequestAmounts() As Boolean

        'run Input_Form
        Dim Input_Form As New Erelonen_provisie_form
        Dim username As String = ""

        username = Globals.CoCoCo_Invoicing.Application.UserInitials

        Input_Form.Show()

        If (Input_Form.Tag <> "Cancelled") Then
            Prov1 = CDbl(Input_Form.Erelonen_input.Text)
            btw = CDbl(Input_Form.Totaal_btw.Text)
            IC = Input_Form.IC.Checked
            Prov2 = CDbl(Input_Form.Gerecht_input.Text)
            erelonen = CDbl(Input_Form.Erelonen_totaal.Text)
            gerechtskosten = CDbl(Input_Form.Gerecht_totaal.Text)
            Totaal = CDbl(Input_Form.Totaal_inc_btw.Text)
            RequestAmounts = True
        Else
            RequestAmounts = False
        End If
        Input_Form.Hide()
        GoTo end_of_function

        On Error Resume Next
        Input_Form.Hide()
        Input_Form.Close()
        RequestAmounts = False

end_of_function:
    End Function

    Private Function InsertInExcel() As Boolean
        Dim Lst As Microsoft.Office.Interop.Excel.ListRows
        Dim sht As Microsoft.Office.Interop.Excel.Worksheet
        Dim tbl As Microsoft.Office.Interop.Excel.ListObject
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim Number, Serial_Number As Integer
        Dim CountDossier As Double

        On Error GoTo ErrorHandler
        OGMCode = GlobalValues.CoCoCo_Calculate_OGM(dossierNr)

        sht = ExWb.Sheets("Provisies")
        sht.Unprotect(Password:=GlobalValues.password)
        tbl = sht.ListObjects("Provisie_Table")
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
            .Range(9).Value = Prov1
            .Range(10).Value = btw
            .Range(11).Value = Prov2
            .Range(12).Value = Prov1 + Prov2 + btw
            .Range(13).Value = False
            .Range(14).Value = OGMCode
            .Range(18).value = 0
            .Range(21).value = IC
        End With

        'Protect sheet
        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)

        InsertInExcel = True
        Exit Function

ErrorHandler:
        InsertInExcel = False
    End Function

    Private Function CloseExcel(Optional Write_On_Save As Boolean = True) As Boolean

        On Error GoTo ErrorHandler
        ' Close Excel bits
        ExWb.Save()
        ExWb.Close(SaveChanges:=Write_On_Save)
        ExWb = Nothing

        CloseExcel = True

        Exit Function

ErrorHandler:
        On Error Resume Next
        CloseExcel = False
        ExWb.Close(SaveChanges:=Write_On_Save)
        ExWb = Nothing

    End Function

    Private Sub Insert_text_provisie()
        Dim text As String
        Dim Selection As Word.Selection = Globals.CoCoCo_Invoicing.Application.Selection

        If (Prov1 = 0 And Prov2 > 0) Then

            text = "Mag ik u vragen om in dit dossier een provisie van" +
                Format(gerechtskosten, " € ## ##0.00 ") +
                "te betalen. Dit om mij toe te laten de gerechtsdeurwaarder te betalen."


            Selection.TypeText(text)
            REM Prov1 >0 and Prov2 >0
        ElseIf (Prov2 > 0) Then
            text = "Mag ik u vragen om in dit dossier een globale provisie te betalen van" +
                Format(Totaal, " € ## ##0.00") + "." +
                vbNewLine + "Dit bedrag is als volgt samengesteld: Provisie erelonen en bureelkosten" + Format(Prov1, " € ## ##0.00 ")
            If (Not IC) Then
                text += "vermeerderd met 21% btw of" + Format(btw, " € ## ##0.00 ")
            End If
            text += "en een provisie voor de gerechtskosten van" +
                Format(gerechtskosten, " € ## ##0.00") + "."

            Selection.TypeText(text)
            REM prov2 = 0
        Else
            text = "Mag ik u vragen om in dit dossier een globale provisie te betalen van" +
                Format(Totaal, " € ## ##0.00 ") +
                "samengesteld als volgt" + Format(Prov1, " € ## ##0.00 ") +
                "aan erelonen en bureelkosten "
            If (Not IC) Then
                text += " en" + Format(btw, " € ## ##0.00 ") + "aan BTW."
            Else
                text += "."
            End If

            Selection.TypeText(text)
        End If
        text = vbNewLine + "U kunt dit bedrag overmaken op rekeningnummer BE96 0012 4751 7505 met als mededeling: " +
           OGMCode + "."
        Selection.TypeText(text)

    End Sub

    Public Function Main() As Boolean
        Dim success As Boolean
        Dim error_text As String = ""
        Dim Write_On_Save As Boolean

        success = True
        Write_On_Save = True
        If Not ReadFile() Then
            error_text = "CSV file not read"
            GoTo Exit_error
        ElseIf Not RequestAmounts() Then
            error_text = "Error in Form"
            GoTo Exit_error
        ElseIf Not InsertInExcel() Then
            success = False
            error_text = "Error adding text to Excel"
            Write_On_Save = False
        End If

        If Not CloseExcel(Write_On_Save) Then
            success = False
            error_text = "Error adding text to Excel"
        End If

        If success Then
            Call Insert_text_provisie()
        Else
            GoTo Exit_error
        End If

        Main = success
        Exit Function

Exit_error:
        MsgBox(error_text)
        Main = success

    End Function

End Class
