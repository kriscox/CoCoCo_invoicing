Imports System.Collections
Imports System.Drawing
Imports System.Windows.Forms

Public Class ListForm

    Private Sub Print_Button_Click()
        Dim prn As New Printing.PrintDocument
        prn.PrinterSettings.PrinterName.ToString()
        AddHandler prn.PrintPage, AddressOf Me.PrintPageHandler
        prn.Print()
        RemoveHandler prn.PrintPage, AddressOf Me.PrintPageHandler

    End Sub

    Private Sub ListView1_ColumnClick(sender As Object, e As Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
        ' Set the ListViewItemSorter property to a new ListViewItemComparer
        ' object.
        If ListView1.ListViewItemSorter Is Nothing Then
            ListView1.ListViewItemSorter = New ListViewItemComparer(e.Column)
        Else
            Dim Sort As ListViewItemComparer = ListView1.ListViewItemSorter
            Sort.reverse(e.Column)
        End If
        ' Call the sort method to manually sort.
        ListView1.Sort()
    End Sub

    Private Sub PrintPageHandler(ByVal sender As Object, ByVal args As Printing.PrintPageEventArgs)

        Dim lvwItem As ListViewItem
        Dim lvwSubItem As ListViewItem.ListViewSubItem
        Dim xPos As Integer = 0
        Dim yPos As Integer = 0

        ' Counter for display purposes
        Dim listviewcount As Integer = 1

        ' Loop through our listview items
        For Each lvwItem In ListView1.Items
            xPos = 0

            ' Print the count
            ' Debug.Print(listviewcount)

            ' Print the subitems of this particular ListViewItem
            For Each lvwSubItem In lvwItem.SubItems
                xPos += 100
                yPos = 100 + (listviewcount * 15)
                args.Graphics.DrawString(lvwSubItem.Text(),
                    New Font("Arial", 10, FontStyle.Bold), Brushes.Black, xPos, yPos)
            Next

            ' Increment the count (for display purposes)
            listviewcount += 1
        Next

    End Sub
End Class

Class ListViewItemComparer
    Implements IComparer
    Private Ascending As Integer = 1
    Private col As Integer

    Public Sub New()
        col = 0
    End Sub

    Public Sub reverse(column As Integer)
        If col = column Then
            Ascending = Ascending * -1
        Else
            col = column
        End If

    End Sub

    Public Sub New(column As Integer)
        col = column
    End Sub

    Public Function Compare(x As Object, y As Object) As Integer _
                            Implements System.Collections.IComparer.Compare
        Dim returnVal As Integer = -1
        Dim val1 As String = CType(x, ListViewItem).SubItems(col).Text
        Dim val2 As String = CType(y, ListViewItem).SubItems(col).Text
        Dim dval1, dval2 As Double
        Dim date1, date2 As Date

        If Double.TryParse(val1, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture, dval1) And
            Double.TryParse(val2, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture, dval2) Then
            returnVal = dval1 - dval2
        ElseIf Date.TryParse(val1, date1) And Date.TryParse(val2, date2) Then
            returnVal = Date.Compare(date1, date2)
        Else
            returnVal = [String].Compare(val1, val2)
        End If

        Return Ascending * returnVal
    End Function
End Class