Imports Microsoft.Office.Interop.Excel

Public Class DerdenGelden
#Region "Variables"
    Private dossierNr As String
    Private dossierName As String
    Private aanspreektitel As String
    Private Name As String
    Private Adres1 As String
    Private excelwb As Workbook = GlobalValues.GetWorkbook
    Private adres2 As String
    Private Totaal As Double = 0.00
    Private OGMCode As String
#End Region

    Private Function readFile() As Boolean
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

        readFile = Result

    End Function

    Private Function requestAmounts() As Boolean

        'run Input_Form
        Dim Input_Form As New Derden_Gelden_Form
        Dim username As String = ""

        username = Globals.CoCoCo_Invoicing.Application.UserInitials

        Input_Form.Show()

        If (Input_Form.Tag <> "Cancelled") Then
            Totaal = CDbl(Input_Form.Payment_amount.Text)
            requestAmounts = True
        Else
            requestAmounts = False
        End If
        Input_Form.Hide()
        GoTo end_of_function

        On Error Resume Next
        Input_Form.Hide()
        Input_Form.Close()
        requestAmounts = False

end_of_function:
    End Function

    Private Function InsertInExcel() As Boolean
        Dim Lst As ListRows
        Dim sht As Worksheet
        Dim tbl As ListObject
        Dim rng As Range
        Dim Number, Serial_Number As Integer
        Dim CountDossier As Double

        On Error GoTo ErrorHandler
        OGMCode = GlobalValues.CoCoCo_Calculate_OGM(dossierNr, True)

        sht = excelwb.Sheets("DerdenGelden")
        sht.Unprotect(Password:="mviw!wwGUp!zaX7A")
        tbl = sht.ListObjects("Derden_Gelden_Table")
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
            .Range(9).Value = Totaal
            .Range(10).Value = OGMCode
            .Range(12).Value = 0
        End With

        'Protect sheet
        sht.Protect(Password:="mviw!wwGUp!zaX7A", AllowSorting:=True, AllowFiltering:=True)

        InsertInExcel = True
        Exit Function

ErrorHandler:
        InsertInExcel = False
    End Function

    Private Sub Insert_text()
        Dim text As String
        Dim Selection As Word.Selection = Globals.CoCoCo_Invoicing.Application.Selection

        If (Totaal > 0) Then

            text = "Mag ik u vragen om in dit dossier een bedrag van" +
                Format(Totaal, " € ## ##0.00 ") +
                "te betalen. "


            Selection.TypeText(text)
            REM Prov1 >0 and Prov2 >0
            text = vbNewLine + "U kunt dit bedrag overmaken op rekeningnummer BE96 0012 4751 7505 met als mededeling: " +
               OGMCode + "."
            Selection.TypeText(text)
        End If

    End Sub

    Public Function main() As Boolean
        Dim success As Boolean
        Dim error_text As String = ""
        Dim Write_On_Save As Boolean

        success = True
        Write_On_Save = True
        If Not readFile() Then
            error_text = "CSV file not read"
            GoTo Exit_error
        ElseIf Not requestAmounts() Then
            error_text = "Error in Form"
            GoTo Exit_error
        ElseIf Not InsertInExcel() Then
            success = False
            error_text = "Error adding text to Excel"
            Write_On_Save = False
        End If

        If success Then
            Call Insert_text()
        Else
            GoTo Exit_error
        End If

        main = success
        Exit Function

Exit_error:
        MsgBox(error_text)
        main = success

    End Function
End Class
