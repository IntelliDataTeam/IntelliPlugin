<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PopForm
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
        Me.lastCol = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.colNum = New System.Windows.Forms.NumericUpDown()
        Me.headerCheckbox = New System.Windows.Forms.CheckBox()
        Me.exportCheckbox = New System.Windows.Forms.CheckBox()
        Me.ok_button = New System.Windows.Forms.Button()
        Me.cancel_button = New System.Windows.Forms.Button()
        CType(Me.colNum, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(105, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Last Formula Column"
        '
        'lastCol
        '
        Me.lastCol.Location = New System.Drawing.Point(169, 29)
        Me.lastCol.Name = "lastCol"
        Me.lastCol.Size = New System.Drawing.Size(58, 20)
        Me.lastCol.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 83)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(121, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Number of Input Column"
        '
        'colNum
        '
        Me.colNum.Location = New System.Drawing.Point(169, 76)
        Me.colNum.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.colNum.Name = "colNum"
        Me.colNum.Size = New System.Drawing.Size(58, 20)
        Me.colNum.TabIndex = 2
        Me.colNum.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'headerCheckbox
        '
        Me.headerCheckbox.AutoSize = True
        Me.headerCheckbox.Location = New System.Drawing.Point(15, 125)
        Me.headerCheckbox.Name = "headerCheckbox"
        Me.headerCheckbox.Size = New System.Drawing.Size(161, 17)
        Me.headerCheckbox.TabIndex = 3
        Me.headerCheckbox.Text = "Is the First Row the Header?"
        Me.headerCheckbox.UseVisualStyleBackColor = True
        '
        'exportCheckbox
        '
        Me.exportCheckbox.AutoSize = True
        Me.exportCheckbox.Location = New System.Drawing.Point(15, 148)
        Me.exportCheckbox.Name = "exportCheckbox"
        Me.exportCheckbox.Size = New System.Drawing.Size(148, 17)
        Me.exportCheckbox.TabIndex = 4
        Me.exportCheckbox.Text = "Export to Results to CSV?"
        Me.exportCheckbox.UseVisualStyleBackColor = True
        '
        'ok_button
        '
        Me.ok_button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ok_button.Location = New System.Drawing.Point(99, 192)
        Me.ok_button.Name = "ok_button"
        Me.ok_button.Size = New System.Drawing.Size(75, 23)
        Me.ok_button.TabIndex = 5
        Me.ok_button.Text = "OK"
        Me.ok_button.UseVisualStyleBackColor = True
        '
        'cancel_button
        '
        Me.cancel_button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancel_button.Location = New System.Drawing.Point(200, 192)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(75, 23)
        Me.cancel_button.TabIndex = 6
        Me.cancel_button.Text = "Cancel"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'PopForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(287, 227)
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.ok_button)
        Me.Controls.Add(Me.exportCheckbox)
        Me.Controls.Add(Me.headerCheckbox)
        Me.Controls.Add(Me.colNum)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lastCol)
        Me.Controls.Add(Me.Label1)
        Me.Name = "PopForm"
        Me.Text = "Decision Time"
        CType(Me.colNum, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lastCol As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents colNum As System.Windows.Forms.NumericUpDown
    Friend WithEvents headerCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents exportCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents ok_button As System.Windows.Forms.Button
    Friend WithEvents cancel_button As System.Windows.Forms.Button
End Class
