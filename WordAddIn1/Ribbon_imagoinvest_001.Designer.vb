Partial Class Imagoinvest
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim RibbonDialogLauncherImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher = Me.Factory.CreateRibbonDialogLauncher
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Provisie = Me.Factory.CreateRibbonGroup
        Me.Maak_provisie = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Provisie.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Provisie)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Provisie
        '
        Me.Provisie.DialogLauncher = RibbonDialogLauncherImpl1
        Me.Provisie.Items.Add(Me.Maak_provisie)
        Me.Provisie.Label = "Provisie"
        Me.Provisie.Name = "Provisie"
        '
        'Maak_provisie
        '
        Me.Maak_provisie.Image = Global.WordAddIn1.My.Resources.Resources.Fases
        Me.Maak_provisie.Label = "Maak Provisie"
        Me.Maak_provisie.Name = "Maak_provisie"
        Me.Maak_provisie.ShowImage = True
        '
        'Imagoinvest
        '
        Me.Name = "Imagoinvest"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Provisie.ResumeLayout(False)
        Me.Provisie.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Provisie As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Maak_provisie As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Imagoinvest
        Get
            Return Me.GetRibbon(Of Imagoinvest)()
        End Get
    End Property
End Class
