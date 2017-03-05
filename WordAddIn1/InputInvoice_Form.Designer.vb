<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InputInvoice_Form
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.OGMinput = New System.Windows.Forms.TabPage()
        Me.OGM_exit = New System.Windows.Forms.Button()
        Me.OGM_ok = New System.Windows.Forms.Button()
        Me.OGMcode3 = New System.Windows.Forms.TextBox()
        Me.OGMcode2 = New System.Windows.Forms.TextBox()
        Me.OGMcode1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ShowExcel = New System.Windows.Forms.TabPage()
        Me.TabControl1.SuspendLayout()
        Me.OGMinput.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.OGMinput)
        Me.TabControl1.Controls.Add(Me.ShowExcel)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(395, 156)
        Me.TabControl1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
        Me.TabControl1.TabIndex = 0
        '
        'OGMinput
        '
        Me.OGMinput.BackColor = System.Drawing.SystemColors.Control
        Me.OGMinput.Controls.Add(Me.OGM_exit)
        Me.OGMinput.Controls.Add(Me.OGM_ok)
        Me.OGMinput.Controls.Add(Me.OGMcode3)
        Me.OGMinput.Controls.Add(Me.OGMcode2)
        Me.OGMinput.Controls.Add(Me.OGMcode1)
        Me.OGMinput.Controls.Add(Me.Label2)
        Me.OGMinput.Controls.Add(Me.Label4)
        Me.OGMinput.Controls.Add(Me.Label3)
        Me.OGMinput.Controls.Add(Me.Label1)
        Me.OGMinput.Location = New System.Drawing.Point(4, 22)
        Me.OGMinput.Name = "OGMinput"
        Me.OGMinput.Padding = New System.Windows.Forms.Padding(3)
        Me.OGMinput.Size = New System.Drawing.Size(387, 130)
        Me.OGMinput.TabIndex = 0
        Me.OGMinput.Text = "OGMinput"
        '
        'OGM_exit
        '
        Me.OGM_exit.Location = New System.Drawing.Point(264, 82)
        Me.OGM_exit.Name = "OGM_exit"
        Me.OGM_exit.Size = New System.Drawing.Size(82, 23)
        Me.OGM_exit.TabIndex = 4
        Me.OGM_exit.Text = "EXIT"
        Me.OGM_exit.UseVisualStyleBackColor = True
        '
        'OGM_ok
        '
        Me.OGM_ok.Location = New System.Drawing.Point(151, 82)
        Me.OGM_ok.Name = "OGM_ok"
        Me.OGM_ok.Size = New System.Drawing.Size(82, 23)
        Me.OGM_ok.TabIndex = 4
        Me.OGM_ok.Text = "OK"
        Me.OGM_ok.UseVisualStyleBackColor = True
        '
        'OGMcode3
        '
        Me.OGMcode3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OGMcode3.Location = New System.Drawing.Point(264, 31)
        Me.OGMcode3.MaxLength = 5
        Me.OGMcode3.Name = "OGMcode3"
        Me.OGMcode3.Size = New System.Drawing.Size(82, 20)
        Me.OGMcode3.TabIndex = 3
        '
        'OGMcode2
        '
        Me.OGMcode2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OGMcode2.Location = New System.Drawing.Point(151, 31)
        Me.OGMcode2.MaxLength = 4
        Me.OGMcode2.Name = "OGMcode2"
        Me.OGMcode2.Size = New System.Drawing.Size(82, 20)
        Me.OGMcode2.TabIndex = 2
        '
        'OGMcode1
        '
        Me.OGMcode1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OGMcode1.Location = New System.Drawing.Point(39, 32)
        Me.OGMcode1.MaxLength = 3
        Me.OGMcode1.Name = "OGMcode1"
        Me.OGMcode1.Size = New System.Drawing.Size(82, 20)
        Me.OGMcode1.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(353, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(26, 22)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "++"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(240, 31)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(18, 22)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "/"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(127, 31)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(18, 22)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "/"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(26, 22)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "++"
        '
        'ShowExcel
        '
        Me.ShowExcel.BackColor = System.Drawing.SystemColors.Control
        Me.ShowExcel.Location = New System.Drawing.Point(4, 22)
        Me.ShowExcel.Name = "ShowExcel"
        Me.ShowExcel.Padding = New System.Windows.Forms.Padding(3)
        Me.ShowExcel.Size = New System.Drawing.Size(387, 130)
        Me.ShowExcel.TabIndex = 1
        Me.ShowExcel.Text = "Excel"
        '
        'InputInvoice_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(419, 180)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "InputInvoice_Form"
        Me.Text = "Input betalingen"
        Me.TabControl1.ResumeLayout(False)
        Me.OGMinput.ResumeLayout(False)
        Me.OGMinput.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabControl1 As Windows.Forms.TabControl
    Friend WithEvents OGMinput As Windows.Forms.TabPage
    Friend WithEvents OGM_exit As Windows.Forms.Button
    Friend WithEvents OGM_ok As Windows.Forms.Button
    Friend WithEvents OGMcode3 As Windows.Forms.TextBox
    Friend WithEvents OGMcode2 As Windows.Forms.TextBox
    Friend WithEvents OGMcode1 As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents ShowExcel As Windows.Forms.TabPage
End Class
