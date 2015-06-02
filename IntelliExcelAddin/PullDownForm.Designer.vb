<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PullDownForm
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
        Me.p_range = New System.Windows.Forms.TextBox()
        Me.p_limit = New System.Windows.Forms.TextBox()
        Me.ok_button = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(91, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Range of Formula"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(29, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Number of Rows"
        '
        'p_range
        '
        Me.p_range.Location = New System.Drawing.Point(149, 24)
        Me.p_range.Name = "p_range"
        Me.p_range.Size = New System.Drawing.Size(100, 20)
        Me.p_range.TabIndex = 1
        '
        'p_limit
        '
        Me.p_limit.Location = New System.Drawing.Point(149, 61)
        Me.p_limit.Name = "p_limit"
        Me.p_limit.Size = New System.Drawing.Size(100, 20)
        Me.p_limit.TabIndex = 2
        '
        'ok_button
        '
        Me.ok_button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ok_button.Location = New System.Drawing.Point(174, 108)
        Me.ok_button.Name = "ok_button"
        Me.ok_button.Size = New System.Drawing.Size(75, 23)
        Me.ok_button.TabIndex = 3
        Me.ok_button.Text = "OK"
        Me.ok_button.UseVisualStyleBackColor = True
        '
        'PullDownForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(271, 144)
        Me.Controls.Add(Me.ok_button)
        Me.Controls.Add(Me.p_limit)
        Me.Controls.Add(Me.p_range)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "PullDownForm"
        Me.Text = "Pull Down Formulas"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents p_range As System.Windows.Forms.TextBox
    Friend WithEvents p_limit As System.Windows.Forms.TextBox
    Friend WithEvents ok_button As System.Windows.Forms.Button
End Class
