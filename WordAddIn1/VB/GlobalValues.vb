Imports Microsoft.Office.Interop.Excel

Public Class GlobalValues
    Implements IDisposable

    Private Shared myInstance As Object = Nothing
    Private Shared ExcelApp As Application = Nothing
    Private Shared ExcelWB As Workbook = Nothing
    Protected disposed As Boolean = False

    'Global variables
    Public Shared ReadOnly password As String = "mviw!wwGUp!zaX7A"
    Public Shared ReadOnly NumberFormat = "€ ## ##0.00;[RED]€ -## ##0.00;-"
    Public Shared ReadOnly invoiceTemplate = "C:\Users\krisc\OneDrive\Documents\01. CoCoCo\ImagoInvest\Factuur.dotx"
    Public Shared ReadOnly ExcelFileName = "C:\Users\krisc\OneDrive\Documents\01. CoCoCo\ImagoInvest\klantenboek.xlsx"
    Public Shared ReadOnly InvoicePath = "C:\Users\krisc\OneDrive\Documents\01. CoCoCo\FA"

    Private Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
    End Sub

    Public Shared Function GetInstance() As GlobalValues
        If myInstance Is Nothing Then
            myInstance = New GlobalValues()
        End If
        Return myInstance
    End Function

    Public Shared Function GetExcel() As Application
        If ExcelApp Is Nothing Then
            'Open excel
            ExcelApp = CreateObject("Excel.Application")
        End If
        GetExcel = ExcelApp
    End Function

    Public Shared Function GetWorkbook() As Workbook
        If ExcelWB Is Nothing Then
            Open_excel()
        End If
        GetWorkbook = ExcelWB
    End Function

    Public Shared Function CoCoCo_Calculate_OGM(ByVal dossierNr As String, Optional ByVal derdenGelden As Boolean = False) As String
        Dim tbl As ListObject
        Dim Lst As ListRows
        Dim sht As Worksheet
        Dim CountDossier As Integer

        On Error GoTo ErrorHandler
        '------------------------------------------------------
        'Calculate ogm code
        '----------------------
        If (Not derdenGelden) Then
            sht = ExcelWB.Sheets("Ereloon Nota")
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

            sht = ExcelWB.Sheets("Provisies")
            tbl = sht.ListObjects("Provisie_Table")
            Lst = tbl.ListRows
            sht.Unprotect(Password:=password)

            REM remove the autofilter is necessairy
            tbl.AutoFilter.ShowAllData()
            tbl.Range.AutoFilter(Field:=3, Criteria1:=dossierNr)
            tbl.AutoFilter.ApplyFilter()
            CountDossier = CountDossier + sht.Evaluate("=Provisie_Table[[#Totals],[dossiernr]]")
            If (CountDossier >= 299) Then
                Throw New IndexOutOfRangeException("Aantal OGM waarden is te groot (> 299)")
            End If
            tbl.AutoFilter.ShowAllData()
            sht.Protect(Password:=password, AllowSorting:=True, AllowFiltering:=True)
        Else
            sht = ExcelWB.Sheets("DerdenGelden")
            tbl = sht.ListObjects("Derden_Gelden_Table")
            Lst = tbl.ListRows
            sht.Unprotect(Password:=password)

            REM remove the autofilter is necessairy
            tbl.AutoFilter.ShowAllData()
            tbl.Range.AutoFilter(Field:=3, Criteria1:=dossierNr)
            tbl.AutoFilter.ApplyFilter()

            CountDossier = sht.Evaluate("=Derden_Gelden_Table[[#Totals],[dossiernr]]")
            CountDossier += 300 'Derden gelden OGM beginnen het derde deel met 3
            tbl.AutoFilter.ShowAllData()
            sht.Protect(Password:=password, AllowSorting:=True, AllowFiltering:=True)
        End If

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

    Private Shared Sub Open_excel()
        'If no ExcelApp exist, create a new one.
        Try
            If IsNothing(ExcelApp.Workbooks) Then
            End If
        Catch
            ExcelApp = Nothing
            GC.Collect()
            ExcelApp = CreateObject("Excel.Application")
        End Try
        ExcelWB = ExcelApp.Workbooks.Open(Filename:=ExcelFileName)

        If True = ExcelWB.ReadOnly Then
            ExcelWB.Close()
            Throw New Data.ReadOnlyException("!!The excel database is read-only!!")
        End If
    End Sub

    Private Shared Sub Close_excel()
        ExcelWB.Close(SaveChanges:=True)
        ExcelWB = Nothing
    End Sub

    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposed Then
            If disposing Then
                ' Insert code to free managed resources.  
                Close_excel()
                ExcelApp.Quit()
            End If
            ' Insert code to free unmanaged resources.  
        End If
        Me.disposed = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub

End Class
