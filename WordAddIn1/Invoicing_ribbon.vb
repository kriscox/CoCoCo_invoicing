'TODO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Invoicing_ribbon()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

<Runtime.InteropServices.ComVisible(True)>
Public Class Invoicing_ribbon
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("WordAddIn1.Invoicing_ribbon.xml")
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Public Sub Provisie_nota_button(ByVal control As Office.IRibbonControl)
        Dim provisie As Provisie = New Provisie
        provisie.main()
    End Sub

    Public Sub Ereloon_nota_button(ByVal control As Office.IRibbonControl)
        Dim erelonen As Erelonen = New Erelonen
        erelonen.main()
    End Sub

    Public Sub Overzicht_button(ByVal control As Office.IRibbonControl)
        Dim PaymentOverview As PaymentOverview = New PaymentOverview
        PaymentOverview.main()
    End Sub
#End Region

    Public Sub Show_Provision_Form_Button(ByVal control As Office.IRibbonControl)
        Dim form As Erelonen_provisie_form = New Erelonen_provisie_form
        form.Show()
    End Sub

    Public Sub Show_Erelonen_Form_Button(ByVal control As Office.IRibbonControl)
        Dim erelonen_form As Ereloon_Nota_form = New Ereloon_Nota_form
        erelonen_form.Show()
    End Sub

    Public Sub Invoicing_Button(ByVal control As Office.IRibbonControl)
        Dim Invoicing As Invoicing = New Invoicing
        Invoicing.startup()
    End Sub

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
