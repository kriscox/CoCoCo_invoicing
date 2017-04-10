<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Derden_Gelden_Form
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
        Me.payment_cancel = New System.Windows.Forms.Button()
        Me.Payment_ok = New System.Windows.Forms.Button()
        Me.Payment_amount = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'payment_cancel
        '
        Me.payment_cancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.payment_cancel.Location = New System.Drawing.Point(279, 48)
        Me.payment_cancel.Name = "payment_cancel"
        Me.payment_cancel.Size = New System.Drawing.Size(83, 40)
        Me.payment_cancel.TabIndex = 7
        Me.payment_cancel.Text = "Annuleren"
        Me.payment_cancel.UseVisualStyleBackColor = True
        '
        'Payment_ok
        '
        Me.Payment_ok.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Payment_ok.Location = New System.Drawing.Point(187, 48)
        Me.Payment_ok.Name = "Payment_ok"
        Me.Payment_ok.Size = New System.Drawing.Size(83, 40)
        Me.Payment_ok.TabIndex = 6
        Me.Payment_ok.Text = "OK"
        Me.Payment_ok.UseVisualStyleBackColor = True
        '
        'Payment_amount
        '
        Me.Payment_amount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Payment_amount.Location = New System.Drawing.Point(81, 12)
        Me.Payment_amount.Name = "Payment_amount"
        Me.Payment_amount.Size = New System.Drawing.Size(280, 20)
        Me.Payment_amount.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 22)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Bedrag"
        '
        'Derden_Gelden_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(379, 109)
        Me.Controls.Add(Me.payment_cancel)
        Me.Controls.Add(Me.Payment_ok)
        Me.Controls.Add(Me.Payment_amount)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Derden_Gelden_Form"
        Me.Text = "Derden gelden"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents payment_cancel As Windows.Forms.Button
    Friend WithEvents Payment_ok As Windows.Forms.Button
    Friend WithEvents Payment_amount As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
End Class
