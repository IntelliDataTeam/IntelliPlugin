<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ValForm
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lastCol = New System.Windows.Forms.TextBox()
        Me.colNum = New System.Windows.Forms.NumericUpDown()
        Me.ok_button = New System.Windows.Forms.Button()
        Me.cancel_button = New System.Windows.Forms.Button()
        Me.headerCheckbox = New System.Windows.Forms.CheckBox()
        Me.validColumn = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.exCol = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        CType(Me.colNum, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(30, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(108, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Last Formula Cell (_2)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(29, 108)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(126, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Number of Input Columns"
        '
        'lastCol
        '
        Me.lastCol.Location = New System.Drawing.Point(184, 34)
        Me.lastCol.Name = "lastCol"
        Me.lastCol.Size = New System.Drawing.Size(56, 20)
        Me.lastCol.TabIndex = 1
        '
        'colNum
        '
        Me.colNum.Location = New System.Drawing.Point(184, 101)
        Me.colNum.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.colNum.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.colNum.Name = "colNum"
        Me.colNum.Size = New System.Drawing.Size(56, 20)
        Me.colNum.TabIndex = 3
        Me.colNum.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'ok_button
        '
        Me.ok_button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ok_button.Location = New System.Drawing.Point(165, 223)
        Me.ok_button.Name = "ok_button"
        Me.ok_button.Size = New System.Drawing.Size(75, 23)
        Me.ok_button.TabIndex = 6
        Me.ok_button.Text = "OK"
        Me.ok_button.UseVisualStyleBackColor = True
        '
        'cancel_button
        '
        Me.cancel_button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancel_button.Location = New System.Drawing.Point(265, 223)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(75, 23)
        Me.cancel_button.TabIndex = 7
        Me.cancel_button.Text = "Cancel"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'headerCheckbox
        '
        Me.headerCheckbox.AutoSize = True
        Me.headerCheckbox.Location = New System.Drawing.Point(33, 182)
        Me.headerCheckbox.Name = "headerCheckbox"
        Me.headerCheckbox.Size = New System.Drawing.Size(164, 17)
        Me.headerCheckbox.TabIndex = 5
        Me.headerCheckbox.Text = "Is the First Line the Headers?"
        Me.headerCheckbox.UseVisualStyleBackColor = True
        '
        'validColumn
        '
        Me.validColumn.Location = New System.Drawing.Point(184, 138)
        Me.validColumn.Name = "validColumn"
        Me.validColumn.Size = New System.Drawing.Size(56, 20)
        Me.validColumn.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(30, 145)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(68, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Valid Column"
        '
        'exCol
        '
        Me.exCol.Location = New System.Drawing.Point(183, 70)
        Me.exCol.Name = "exCol"
        Me.exCol.Size = New System.Drawing.Size(56, 20)
        Me.exCol.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(29, 70)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(137, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Last Column to be Exported"
        '
        'ValForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(351, 269)
        Me.Controls.Add(Me.exCol)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.validColumn)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.headerCheckbox)
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.ok_button)
        Me.Controls.Add(Me.colNum)
        Me.Controls.Add(Me.lastCol)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "ValForm"
        Me.Text = "Decision Time"
        CType(Me.colNum, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lastCol As System.Windows.Forms.TextBox
    Friend WithEvents colNum As System.Windows.Forms.NumericUpDown
    Friend WithEvents ok_button As System.Windows.Forms.Button
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    Friend WithEvents headerCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents validColumn As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents exCol As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
