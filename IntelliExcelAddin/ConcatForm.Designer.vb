<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ConcatForm
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
        Me.InputRange = New System.Windows.Forms.TextBox()
        Me.OutputRange = New System.Windows.Forms.TextBox()
        Me.Delimiter = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.OK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(121, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Input Range (i.e. A1:B3)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 66)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(113, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Output Range (i.e. A1)"
        '
        'InputRange
        '
        Me.InputRange.Location = New System.Drawing.Point(166, 21)
        Me.InputRange.Name = "InputRange"
        Me.InputRange.Size = New System.Drawing.Size(46, 20)
        Me.InputRange.TabIndex = 1
        '
        'OutputRange
        '
        Me.OutputRange.Location = New System.Drawing.Point(166, 63)
        Me.OutputRange.Name = "OutputRange"
        Me.OutputRange.Size = New System.Drawing.Size(46, 20)
        Me.OutputRange.TabIndex = 2
        '
        'Delimiter
        '
        Me.Delimiter.Location = New System.Drawing.Point(166, 103)
        Me.Delimiter.Name = "Delimiter"
        Me.Delimiter.Size = New System.Drawing.Size(46, 20)
        Me.Delimiter.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 106)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(146, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Delimiter (New Line is default)"
        '
        'OK
        '
        Me.OK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OK.Location = New System.Drawing.Point(152, 156)
        Me.OK.Name = "OK"
        Me.OK.Size = New System.Drawing.Size(75, 23)
        Me.OK.TabIndex = 4
        Me.OK.Text = "OK"
        Me.OK.UseVisualStyleBackColor = True
        '
        'ConcatForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(239, 191)
        Me.Controls.Add(Me.OK)
        Me.Controls.Add(Me.Delimiter)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.OutputRange)
        Me.Controls.Add(Me.InputRange)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "ConcatForm"
        Me.Text = "Concat Cells"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents InputRange As System.Windows.Forms.TextBox
    Friend WithEvents OutputRange As System.Windows.Forms.TextBox
    Friend WithEvents Delimiter As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents OK As System.Windows.Forms.Button
End Class
