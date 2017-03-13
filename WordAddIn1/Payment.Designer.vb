<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Payment
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
        Me.ogm_label = New System.Windows.Forms.Label()
        Me.Dossier_label = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Payment_amount = New System.Windows.Forms.TextBox()
        Me.Payment_ok = New System.Windows.Forms.Button()
        Me.payment_cancel = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ogm_label
        '
        Me.ogm_label.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ogm_label.AutoSize = True
        Me.ogm_label.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ogm_label.Location = New System.Drawing.Point(88, 19)
        Me.ogm_label.Name = "ogm_label"
        Me.ogm_label.Size = New System.Drawing.Size(0, 22)
        Me.ogm_label.TabIndex = 0
        '
        'Dossier_label
        '
        Me.Dossier_label.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Dossier_label.AutoSize = True
        Me.Dossier_label.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dossier_label.Location = New System.Drawing.Point(88, 77)
        Me.Dossier_label.Name = "Dossier_label"
        Me.Dossier_label.Size = New System.Drawing.Size(0, 22)
        Me.Dossier_label.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(19, 135)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 22)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Bedrag"
        '
        'Payment_amount
        '
        Me.Payment_amount.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Payment_amount.Location = New System.Drawing.Point(88, 138)
        Me.Payment_amount.Name = "Payment_amount"
        Me.Payment_amount.Size = New System.Drawing.Size(269, 20)
        Me.Payment_amount.TabIndex = 1
        '
        'Payment_ok
        '
        Me.Payment_ok.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Payment_ok.Location = New System.Drawing.Point(128, 193)
        Me.Payment_ok.Name = "Payment_ok"
        Me.Payment_ok.Size = New System.Drawing.Size(75, 23)
        Me.Payment_ok.TabIndex = 2
        Me.Payment_ok.Text = "OK"
        Me.Payment_ok.UseVisualStyleBackColor = True
        '
        'payment_cancel
        '
        Me.payment_cancel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.payment_cancel.Location = New System.Drawing.Point(249, 193)
        Me.payment_cancel.Name = "payment_cancel"
        Me.payment_cancel.Size = New System.Drawing.Size(75, 23)
        Me.payment_cancel.TabIndex = 3
        Me.payment_cancel.Text = "Annuleren"
        Me.payment_cancel.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(43, 22)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "OGM"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 22)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "DossierNr"
        '
        'Payment
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(365, 241)
        Me.Controls.Add(Me.payment_cancel)
        Me.Controls.Add(Me.Payment_ok)
        Me.Controls.Add(Me.Payment_amount)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Dossier_label)
        Me.Controls.Add(Me.ogm_label)
        Me.Name = "Payment"
        Me.Text = "Betaling"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ogm_label As Windows.Forms.Label
    Friend WithEvents Dossier_label As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Payment_amount As Windows.Forms.TextBox
    Friend WithEvents Payment_ok As Windows.Forms.Button
    Friend WithEvents payment_cancel As Windows.Forms.Button
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
End Class
