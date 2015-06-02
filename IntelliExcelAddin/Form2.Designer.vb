<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class pop_form
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
        Me.headerCheckbox = New System.Windows.Forms.CheckBox()
        Me.cancel_button = New System.Windows.Forms.Button()
        Me.ok_button = New System.Windows.Forms.Button()
        Me.colNum = New System.Windows.Forms.NumericUpDown()
        Me.lastCol = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ExportCheckbox = New System.Windows.Forms.CheckBox()
        CType(Me.colNum, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'headerCheckbox
        '
        Me.headerCheckbox.AutoSize = True
        Me.headerCheckbox.Location = New System.Drawing.Point(15, 130)
        Me.headerCheckbox.Name = "headerCheckbox"
        Me.headerCheckbox.Size = New System.Drawing.Size(164, 17)
        Me.headerCheckbox.TabIndex = 3
        Me.headerCheckbox.Text = "Is the First Line the Headers?"
        Me.headerCheckbox.UseVisualStyleBackColor = True
        '
        'cancel_button
        '
        Me.cancel_button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancel_button.Location = New System.Drawing.Point(197, 192)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(75, 23)
        Me.cancel_button.TabIndex = 6
        Me.cancel_button.Text = "Cancel"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'ok_button
        '
        Me.ok_button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ok_button.Location = New System.Drawing.Point(97, 192)
        Me.ok_button.Name = "ok_button"
        Me.ok_button.Size = New System.Drawing.Size(75, 23)
        Me.ok_button.TabIndex = 5
        Me.ok_button.Text = "OK"
        Me.ok_button.UseVisualStyleBackColor = True
        '
        'colNum
        '
        Me.colNum.Location = New System.Drawing.Point(167, 66)
        Me.colNum.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.colNum.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.colNum.Name = "colNum"
        Me.colNum.Size = New System.Drawing.Size(56, 20)
        Me.colNum.TabIndex = 2
        Me.colNum.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'lastCol
        '
        Me.lastCol.Location = New System.Drawing.Point(166, 27)
        Me.lastCol.Name = "lastCol"
        Me.lastCol.Size = New System.Drawing.Size(56, 20)
        Me.lastCol.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(126, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Number of Input Columns"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(108, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Last Formula Cell (_2)"
        '
        'ExportCheckbox
        '
        Me.ExportCheckbox.AutoSize = True
        Me.ExportCheckbox.Location = New System.Drawing.Point(15, 153)
        Me.ExportCheckbox.Name = "ExportCheckbox"
        Me.ExportCheckbox.Size = New System.Drawing.Size(150, 17)
        Me.ExportCheckbox.TabIndex = 4
        Me.ExportCheckbox.Text = "Export results to CSV File?"
        Me.ExportCheckbox.UseVisualStyleBackColor = True
        '
        'pop_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(286, 224)
        Me.Controls.Add(Me.ExportCheckbox)
        Me.Controls.Add(Me.headerCheckbox)
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.ok_button)
        Me.Controls.Add(Me.colNum)
        Me.Controls.Add(Me.lastCol)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "pop_form"
        Me.Text = "Decision Time"
        CType(Me.colNum, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents headerCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    Friend WithEvents ok_button As System.Windows.Forms.Button
    Friend WithEvents colNum As System.Windows.Forms.NumericUpDown
    Friend WithEvents lastCol As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ExportCheckbox As System.Windows.Forms.CheckBox
End Class
