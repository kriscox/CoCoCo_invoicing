Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Windows.Forms

Public Class PaymentOverview
    Implements IDisposable
#Region "Variables"
    REM version 20160830
    Dim ObjExcel As Excel.Application = GlobalValues.GetExcel()
    Dim ExWb As Excel.Workbook = GlobalValues.GetWorkbook()
    Dim colRecords As Collection
    Dim arrData(,), ExcelFileName As String
    Dim Input_Form As New ListForm
    Dim ListInvoices As Excel.Range
    Private Const ColumnCount = 7
#End Region

    Private Sub QuickSort(vArray As Object, inLow As Long, inHi As Long)

        Dim pivot As Object
        Dim tmpSwap As Object
        Dim tmpLow As Long
        Dim tmpHi As Long

        tmpLow = inLow
        tmpHi = inHi

        pivot = vArray((inLow + inHi) \ 2)

        While (tmpLow <= tmpHi)

            While (vArray(tmpLow) < pivot And tmpLow < inHi)
                tmpLow = tmpLow + 1
            End While


            While (pivot < vArray(tmpHi) And tmpHi > inLow)
                tmpHi = tmpHi - 1
            End While

            If (tmpLow <= tmpHi) Then
                tmpSwap = vArray(tmpLow)
                vArray(tmpLow) = vArray(tmpHi)
                vArray(tmpHi) = tmpSwap
                tmpLow = tmpLow + 1
                tmpHi = tmpHi - 1
            End If

        End While

        If (inLow < tmpHi) Then QuickSort(vArray, inLow, tmpHi)
        If (tmpLow < inHi) Then QuickSort(vArray, tmpLow, inHi)

    End Sub

    Private Sub RemoveDups(vArray() As String)
        Dim newArray() As String
        Dim i, Length As Integer

        Length = 0
        ReDim newArray(UBound(vArray))

        newArray(0) = vArray(0)

        For i = 1 To UBound(vArray)
            If (vArray(i) <> newArray(Length)) Then
                Length = Length + 1
                newArray(Length) = vArray(i)
            End If
        Next i

        ReDim Preserve newArray(Length)
        ReDim vArray(Length)
        vArray = newArray

    End Sub

    Private Function Get_Unpayed_Invoices(ByVal dossiernr As String) As Boolean

        Dim sht As Excel.Worksheet
        Dim tbl As Excel.ListObject
        Dim Lst As Excel.ListRows
        Dim tmpArray() As String
        Dim CountDossier As String
        Dim N, Dossier As Integer
        Dim Amount As Double

        On Error GoTo ErrorHandler

        sht = ExWb.Sheets("Ereloon Nota")
        tbl = sht.ListObjects("Ereloon_Nota_Table")
        Lst = tbl.ListRows
        sht.Unprotect(Password:=GlobalValues.password)

        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()
        tbl.Range.AutoFilter(Field:=32, Criteria1:="<>True")
        If (dossiernr IsNot Nothing) Then
            tbl.Range.AutoFilter(Field:=3, Criteria1:=dossiernr)
        End If
        tbl.AutoFilter.ApplyFilter()

        ListInvoices = tbl.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible)
        CountDossier = sht.Evaluate("=Ereloon_Nota_Table[[#Totals],[dossiernr]]").value - 1

        Dossier = tbl.ListColumns("dossiernr").Index

        ReDim tmpArray(CountDossier)
        For N = 0 To CountDossier
            tmpArray(N) = ListInvoices.Cells(N + 1, Dossier).value
        Next N

        tbl.AutoFilter.ShowAllData()
        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)

        sht = ExWb.Sheets("Provisies")
        tbl = sht.ListObjects("Provisie_Table")
        Lst = tbl.ListRows
        sht.Unprotect(Password:=GlobalValues.password)

        REM remove the autofilter is necessairy
        tbl.AutoFilter.ShowAllData()
        tbl.Range.AutoFilter(Field:=13, Criteria1:="<>True")
        If (dossiernr IsNot Nothing) Then
            tbl.Range.AutoFilter(Field:=3, Criteria1:=dossiernr)
        End If
        tbl.AutoFilter.ApplyFilter()

        ListInvoices = tbl.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible)

        Dossier = tbl.ListColumns("dossiernr").Index

        Amount = sht.Evaluate("=Provisie_Table[[#Totals],[dossiernr]]").value - 1
        ReDim Preserve tmpArray(CountDossier + Amount)
        For N = 0 To Amount
            tmpArray(CountDossier + N) = ListInvoices.Cells(N + 1, Dossier).value
        Next N

        tbl.AutoFilter.ShowAllData()
        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)

        Call QuickSort(tmpArray, 0, UBound(tmpArray))

        Call RemoveDups(tmpArray)

        ReDim Preserve arrData(UBound(tmpArray, 1), ColumnCount)

        sht = ExWb.Sheets("Parameters")

        For N = 0 To UBound(tmpArray) - 1
            sht.Cells(6, 5).value = tmpArray(N)
            arrData(N, 0) = sht.Cells(6, 7).value
            arrData(N, 1) = tmpArray(N)
            arrData(N, 2) = sht.Cells(6, 6).value
            If (sht.Cells(6, 9).value <> 0) Then
                arrData(N, 3) = FormatCurrency(sht.Cells(6, 9).value)
            Else
                arrData(N, 3) = FormatCurrency(sht.Cells(6, 8).value)
            End If
            arrData(N, 4) = FormatCurrency(Expression:=sht.Cells(6, 13).Value)
            If (sht.Cells(6, 14).value <> Nothing) Then
                arrData(N, 5) = Format(sht.Cells(6, 14).value, "Short Date")
            Else
                arrData(N, 5) = ""
            End If
            arrData(N, 6) = sht.Cells(6, 15).value
            If (sht.Cells(6, 16).value = Nothing) Then
                arrData(N, 7) = ""
            Else
                arrData(N, 7) = Format(sht.Cells(6, 16).value, "Short Date")
            End If
        Next N

        Get_Unpayed_Invoices = True
        Exit Function

ErrorHandler:
        Get_Unpayed_Invoices = False
        sht.Protect(Password:=GlobalValues.password, AllowSorting:=True, AllowFiltering:=True)

    End Function

    Private Function Show_ListForm(ByVal sel As Boolean) As Boolean

        Dim ColumnHeader As ListView.ColumnHeaderCollection
        Dim ListItems As ListView.ListViewItemCollection
        Dim Item As ListViewItem

        On Error GoTo ErrorHandler

        ColumnHeader = Input_Form.ListView1.Columns

        ColumnHeader.Clear()
        ColumnHeader.Add(text:="Klantnaam")
        ColumnHeader.Add(text:="Dossier")
        ColumnHeader.Add(text:="Omschrijving")
        ColumnHeader.Add(text:="Totaal bedrag")
        ColumnHeader.Add(text:="Openstaand bedrag")
        ColumnHeader.Add(text:="Datum laatste betaling")
        ColumnHeader.Add(text:="Laatste niveau aanmaning")
        ColumnHeader.Add(text:="Datum aanmaning")

        ListItems = Input_Form.ListView1.Items
        ListItems.Clear()
        For i = 0 To UBound(arrData, 1) - 1
            If arrData(i, 3) <> 0 Then
                Item = ListItems.Add(arrData(i, 0))
                For j = 1 To UBound(arrData, 2)
                    Item.SubItems.Add(arrData(i, j))
                Next j
            End If
        Next i

        With Input_Form.ListView1
            .AllowColumnReorder = True
            .GridLines = True
            .Enabled = True
            .View = System.Windows.Forms.View.Details

            .MultiSelect = True
            .CheckBoxes = True
        End With

        Input_Form.Show()

        Show_ListForm = True
        Exit Function

ErrorHandler:
        Show_ListForm = False
        Input_Form.Hide()
    End Function

    Public Function main(Optional ByVal dossiernr As String = Nothing) As Boolean
        Dim error_text As String

        If Not Get_Unpayed_Invoices(dossiernr) Then
            error_text = "Error getting invoices"
            GoTo Exit_error
        ElseIf Not Show_ListForm(dossiernr = Nothing) Then
            error_text = "Error showing Form"
            GoTo Exit_error
        End If

        main = True
        Exit Function

Exit_error:
        MsgBox(error_text)
        main = False

    End Function

    Public Function GetChoise() As Integer
        Return Input_Form.ListView1.SelectedIndices(0)
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                Input_Form.Dispose()
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
