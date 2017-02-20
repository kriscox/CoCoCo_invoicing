Imports Microsoft.Office.Interop

Public Class CoCoCo_Invoicing
    Dim Excel As Excel.Application
    Dim ExWb As Excel.Workbook
    Public Shared ReadOnly password As String = "mviw!wwGUp!zaX7A"

    Private Sub CoCoCo_Startup() Handles Me.Startup
    End Sub

    Private Sub CoCoCo_Shutdown() Handles Me.Shutdown

    End Sub
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Invoicing_ribbon()
    End Function

    Shared Function CoCoCo_Calculate_OGM(ByVal dossierNr As String, ByRef Exwb As Excel.Application) As String
        Dim sht As Excel.Worksheets
        Dim tbl As Excel.ListObject
        Dim Lst As Excel.ListRows
        Dim CountDossier As Integer

        On Error GoTo ErrorHandler
        '------------------------------------------------------
        'Calculate ogm code
        '----------------------
        sht = Exwb.Sheets("Ereloon Nota")
        tbl = sht.ListObjects("Ereloon_Nota_Table")
        Lst = tbl.ListRows
        sht.Unprotect(Password:=password)

        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()
        tbl.Range.AutoFilter(Field:=3, Criteria1:=dossierNr)
        tbl.AutoFilter.ApplyFilter()

        CountDossier = sht.Evaluate("=Ereloon_Nota_Table[[#Totals],[dossiernr]]")
        tbl.AutoFilter.ShowAllData()
        sht.Protect(Password:=password, AllowSorting:=True, AllowFiltering:=True)

        sht = Exwb.Sheets("Provisies")
        tbl = sht.ListObjects("Provisie_Table")
        Lst = tbl.ListRows
        sht.Unprotect(Password:=password)

        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()
        tbl.Range.AutoFilter(Field:=3, Criteria1:=dossierNr)
        tbl.AutoFilter.ApplyFilter()
        CountDossier = CountDossier + sht.Evaluate("=Provisie_Table[[#Totals],[dossiernr]]")
        tbl.AutoFilter.ShowAllData()
        sht.Protect(Password:=password, AllowSorting:=True, AllowFiltering:=True)

        Dim Serial_Number As Integer
        Dim List As String()
        Dim Volg_Number As Double
        Dim Volg_code As Double
        Dim Number As Double

        Serial_Number = CInt(CountDossier) + 1
        List = Split(dossierNr, "/")
        Volg_Number = Split(List(1), "-")(0) Mod 1000
        Volg_code = CDbl(Format(Volg_Number, "000") & Format(List(0), "0000")) Mod 97
        Number = CDbl(Format(Volg_code, "00") & Format(Serial_Number, "000"))

        Return "+++" & Format(Volg_Number, "000") & "/" & List(0) & "/" & Format(Serial_Number, "000") & Format(Number Mod 97, "00") & "+++"

        Exit Function
ErrorHandler:
        Return "+++//+++"
    End Function

End Class
