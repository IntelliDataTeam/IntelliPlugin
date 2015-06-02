<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Text2Column
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
        Me.TRange = New System.Windows.Forms.TextBox()
        Me.Delimiter = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DRange = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'TRange
        '
        Me.TRange.Location = New System.Drawing.Point(162, 17)
        Me.TRange.Name = "TRange"
        Me.TRange.Size = New System.Drawing.Size(50, 20)
        Me.TRange.TabIndex = 0
        '
        'Delimiter
        '
        Me.Delimiter.Location = New System.Drawing.Point(162, 93)
        Me.Delimiter.Name = "Delimiter"
        Me.Delimiter.Size = New System.Drawing.Size(50, 20)
        Me.Delimiter.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Text Range (ie A1:B2)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 100)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Delimiter"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.Highlight
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button1.Location = New System.Drawing.Point(137, 151)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Parsed"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(95, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Destination Range"
        '
        'DRange
        '
        Me.DRange.Location = New System.Drawing.Point(162, 57)
        Me.DRange.Name = "DRange"
        Me.DRange.Size = New System.Drawing.Size(50, 20)
        Me.DRange.TabIndex = 1
        '
        'Text2Column
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(224, 186)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.DRange)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Delimiter)
        Me.Controls.Add(Me.TRange)
        Me.Name = "Text2Column"
        Me.Text = "Parse Text To Columns"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TRange As System.Windows.Forms.TextBox
    Friend WithEvents Delimiter As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DRange As System.Windows.Forms.TextBox
End Class
