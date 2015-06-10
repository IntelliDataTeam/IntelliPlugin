<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class loading_bar
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
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.stop_button = New System.Windows.Forms.Button()
        Me.total = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(72, 42)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(288, 25)
        Me.ProgressBar1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(169, 93)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Please Wait..."
        '
        'stop_button
        '
        Me.stop_button.DialogResult = System.Windows.Forms.DialogResult.Abort
        Me.stop_button.Location = New System.Drawing.Point(334, 110)
        Me.stop_button.Name = "stop_button"
        Me.stop_button.Size = New System.Drawing.Size(75, 23)
        Me.stop_button.TabIndex = 2
        Me.stop_button.Text = "Stop"
        Me.stop_button.UseVisualStyleBackColor = True
        '
        'total
        '
        Me.total.AutoSize = True
        Me.total.Location = New System.Drawing.Point(180, 70)
        Me.total.Name = "total"
        Me.total.Size = New System.Drawing.Size(0, 13)
        Me.total.TabIndex = 3
        '
        'loading_bar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(421, 145)
        Me.Controls.Add(Me.total)
        Me.Controls.Add(Me.stop_button)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Name = "loading_bar"
        Me.Text = "Exporting to CSV"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents stop_button As System.Windows.Forms.Button
    Friend WithEvents total As System.Windows.Forms.Label
End Class
