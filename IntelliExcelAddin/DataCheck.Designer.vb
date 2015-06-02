<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DataCheck
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
        Me.ProdCat = New System.Windows.Forms.TextBox()
        Me.Mfg = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ok_button = New System.Windows.Forms.Button()
        Me.cancel_button = New System.Windows.Forms.Button()
        Me.Col = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(26, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ProdCat"
        '
        'ProdCat
        '
        Me.ProdCat.Location = New System.Drawing.Point(110, 30)
        Me.ProdCat.Name = "ProdCat"
        Me.ProdCat.Size = New System.Drawing.Size(100, 20)
        Me.ProdCat.TabIndex = 1
        '
        'Mfg
        '
        Me.Mfg.Location = New System.Drawing.Point(110, 68)
        Me.Mfg.Name = "Mfg"
        Me.Mfg.Size = New System.Drawing.Size(100, 20)
        Me.Mfg.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(26, 71)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(25, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Mfg"
        '
        'ok_button
        '
        Me.ok_button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ok_button.Location = New System.Drawing.Point(87, 153)
        Me.ok_button.Name = "ok_button"
        Me.ok_button.Size = New System.Drawing.Size(75, 23)
        Me.ok_button.TabIndex = 4
        Me.ok_button.Text = "OK"
        Me.ok_button.UseVisualStyleBackColor = True
        '
        'cancel_button
        '
        Me.cancel_button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancel_button.Location = New System.Drawing.Point(177, 153)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(75, 23)
        Me.cancel_button.TabIndex = 5
        Me.cancel_button.Text = "Cancel"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'Col
        '
        Me.Col.Location = New System.Drawing.Point(110, 105)
        Me.Col.Name = "Col"
        Me.Col.Size = New System.Drawing.Size(36, 20)
        Me.Col.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(26, 108)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Column"
        '
        'DataCheck
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(267, 187)
        Me.Controls.Add(Me.Col)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.ok_button)
        Me.Controls.Add(Me.Mfg)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ProdCat)
        Me.Controls.Add(Me.Label1)
        Me.Name = "DataCheck"
        Me.Text = "Check Data"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProdCat As System.Windows.Forms.TextBox
    Friend WithEvents Mfg As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ok_button As System.Windows.Forms.Button
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    Friend WithEvents Col As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
