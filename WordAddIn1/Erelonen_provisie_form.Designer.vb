<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Erelonen_provisie_form
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Erelonen_input = New System.Windows.Forms.TextBox()
        Me.Erelonen_btw = New System.Windows.Forms.Label()
        Me.Erelonen_totaal = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Gerecht_input = New System.Windows.Forms.TextBox()
        Me.Gerecht_totaal = New System.Windows.Forms.Label()
        Me.Label_BTW = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Totaal_btw = New System.Windows.Forms.Label()
        Me.Totaal_inc_btw = New System.Windows.Forms.Label()
        Me.Button_OK = New System.Windows.Forms.Button()
        Me.Button_Cancel = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Totaal_ex_btw = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(275, 22)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Provisie op erelonen en bureelkosten"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Erelonen_input
        '
        Me.Erelonen_input.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Erelonen_input.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Erelonen_input.Location = New System.Drawing.Point(301, 33)
        Me.Erelonen_input.Name = "Erelonen_input"
        Me.Erelonen_input.Size = New System.Drawing.Size(100, 19)
        Me.Erelonen_input.TabIndex = 1
        Me.Erelonen_input.Text = "0"
        Me.Erelonen_input.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Erelonen_btw
        '
        Me.Erelonen_btw.AutoSize = True
        Me.Erelonen_btw.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Erelonen_btw.Location = New System.Drawing.Point(478, 33)
        Me.Erelonen_btw.Name = "Erelonen_btw"
        Me.Erelonen_btw.Size = New System.Drawing.Size(53, 22)
        Me.Erelonen_btw.TabIndex = 2
        Me.Erelonen_btw.Text = "€ 0,00"
        Me.Erelonen_btw.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Erelonen_totaal
        '
        Me.Erelonen_totaal.AutoSize = True
        Me.Erelonen_totaal.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Erelonen_totaal.Location = New System.Drawing.Point(608, 33)
        Me.Erelonen_totaal.Name = "Erelonen_totaal"
        Me.Erelonen_totaal.Size = New System.Drawing.Size(53, 22)
        Me.Erelonen_totaal.TabIndex = 3
        Me.Erelonen_totaal.Text = "€ 0,00"
        Me.Erelonen_totaal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(12, 67)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(202, 22)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Provisie op gerechtskosten"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Gerecht_input
        '
        Me.Gerecht_input.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Gerecht_input.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Gerecht_input.Location = New System.Drawing.Point(301, 67)
        Me.Gerecht_input.Name = "Gerecht_input"
        Me.Gerecht_input.Size = New System.Drawing.Size(100, 19)
        Me.Gerecht_input.TabIndex = 2
        Me.Gerecht_input.Text = "0"
        Me.Gerecht_input.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Gerecht_totaal
        '
        Me.Gerecht_totaal.AutoSize = True
        Me.Gerecht_totaal.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Gerecht_totaal.Location = New System.Drawing.Point(608, 67)
        Me.Gerecht_totaal.Name = "Gerecht_totaal"
        Me.Gerecht_totaal.Size = New System.Drawing.Size(53, 22)
        Me.Gerecht_totaal.TabIndex = 3
        Me.Gerecht_totaal.Text = "€ 0,00"
        Me.Gerecht_totaal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label_BTW
        '
        Me.Label_BTW.AutoSize = True
        Me.Label_BTW.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_BTW.Location = New System.Drawing.Point(488, 11)
        Me.Label_BTW.Name = "Label_BTW"
        Me.Label_BTW.Size = New System.Drawing.Size(43, 22)
        Me.Label_BTW.TabIndex = 2
        Me.Label_BTW.Text = "BTW"
        Me.Label_BTW.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Location = New System.Drawing.Point(13, 104)
        Me.Label5.Margin = New System.Windows.Forms.Padding(0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(646, 2)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Label5"
        '
        'Totaal_btw
        '
        Me.Totaal_btw.AutoSize = True
        Me.Totaal_btw.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Totaal_btw.Location = New System.Drawing.Point(478, 111)
        Me.Totaal_btw.Name = "Totaal_btw"
        Me.Totaal_btw.Size = New System.Drawing.Size(53, 22)
        Me.Totaal_btw.TabIndex = 2
        Me.Totaal_btw.Text = "€ 0,00"
        Me.Totaal_btw.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Totaal_inc_btw
        '
        Me.Totaal_inc_btw.AutoSize = True
        Me.Totaal_inc_btw.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Totaal_inc_btw.Location = New System.Drawing.Point(608, 111)
        Me.Totaal_inc_btw.Name = "Totaal_inc_btw"
        Me.Totaal_inc_btw.Size = New System.Drawing.Size(53, 22)
        Me.Totaal_inc_btw.TabIndex = 3
        Me.Totaal_inc_btw.Text = "€ 0,00"
        Me.Totaal_inc_btw.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button_OK
        '
        Me.Button_OK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button_OK.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_OK.Location = New System.Drawing.Point(456, 140)
        Me.Button_OK.Name = "Button_OK"
        Me.Button_OK.Size = New System.Drawing.Size(75, 33)
        Me.Button_OK.TabIndex = 3
        Me.Button_OK.Text = "OK"
        Me.Button_OK.UseVisualStyleBackColor = True
        '
        'Button_Cancel
        '
        Me.Button_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button_Cancel.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Cancel.Location = New System.Drawing.Point(586, 140)
        Me.Button_Cancel.Name = "Button_Cancel"
        Me.Button_Cancel.Size = New System.Drawing.Size(75, 33)
        Me.Button_Cancel.TabIndex = 4
        Me.Button_Cancel.Text = "CANCEL"
        Me.Button_Cancel.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label9.Location = New System.Drawing.Point(607, 11)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(52, 22)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "Totaal"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(297, 11)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(59, 22)
        Me.Label10.TabIndex = 2
        Me.Label10.Text = "Bedrag"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Totaal_ex_btw
        '
        Me.Totaal_ex_btw.AutoSize = True
        Me.Totaal_ex_btw.Font = New System.Drawing.Font("Trebuchet MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Totaal_ex_btw.Location = New System.Drawing.Point(348, 111)
        Me.Totaal_ex_btw.Name = "Totaal_ex_btw"
        Me.Totaal_ex_btw.Size = New System.Drawing.Size(53, 22)
        Me.Totaal_ex_btw.TabIndex = 2
        Me.Totaal_ex_btw.Text = "€ 0,00"
        Me.Totaal_ex_btw.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Erelonen_provisie_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(673, 185)
        Me.Controls.Add(Me.Button_Cancel)
        Me.Controls.Add(Me.Button_OK)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Gerecht_totaal)
        Me.Controls.Add(Me.Totaal_inc_btw)
        Me.Controls.Add(Me.Erelonen_totaal)
        Me.Controls.Add(Me.Totaal_ex_btw)
        Me.Controls.Add(Me.Totaal_btw)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label_BTW)
        Me.Controls.Add(Me.Erelonen_btw)
        Me.Controls.Add(Me.Gerecht_input)
        Me.Controls.Add(Me.Erelonen_input)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Erelonen_provisie_form"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.RightToLeftLayout = True
        Me.Text = "Erelonen en gerechtskosten"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Erelonen_input As Windows.Forms.TextBox
    Friend WithEvents Erelonen_btw As Windows.Forms.Label
    Friend WithEvents Erelonen_totaal As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Gerecht_input As Windows.Forms.TextBox
    Friend WithEvents Gerecht_totaal As Windows.Forms.Label
    Friend WithEvents Label_BTW As Windows.Forms.Label
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Totaal_btw As Windows.Forms.Label
    Friend WithEvents Totaal_inc_btw As Windows.Forms.Label
    Friend WithEvents Button_OK As Windows.Forms.Button
    Friend WithEvents Button_Cancel As Windows.Forms.Button
    Friend WithEvents Label9 As Windows.Forms.Label
    Friend WithEvents Label10 As Windows.Forms.Label
    Friend WithEvents Totaal_ex_btw As Windows.Forms.Label
End Class
