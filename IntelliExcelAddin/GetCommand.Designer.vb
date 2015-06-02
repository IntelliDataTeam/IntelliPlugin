<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GetCommand
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
        Me.mySelect = New System.Windows.Forms.TextBox()
        Me.myFrom = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.myWhere = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.myOrder = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ok_button = New System.Windows.Forms.Button()
        Me.cancel_button = New System.Windows.Forms.Button()
        Me.outCol = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(25, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "SELECT"
        '
        'mySelect
        '
        Me.mySelect.Location = New System.Drawing.Point(93, 20)
        Me.mySelect.Name = "mySelect"
        Me.mySelect.Size = New System.Drawing.Size(248, 20)
        Me.mySelect.TabIndex = 1
        '
        'myFrom
        '
        Me.myFrom.Location = New System.Drawing.Point(93, 56)
        Me.myFrom.Name = "myFrom"
        Me.myFrom.Size = New System.Drawing.Size(248, 20)
        Me.myFrom.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(25, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "FROM"
        '
        'myWhere
        '
        Me.myWhere.Location = New System.Drawing.Point(93, 92)
        Me.myWhere.Name = "myWhere"
        Me.myWhere.Size = New System.Drawing.Size(248, 20)
        Me.myWhere.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(25, 95)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "WHERE"
        '
        'myOrder
        '
        Me.myOrder.Location = New System.Drawing.Point(93, 129)
        Me.myOrder.Name = "myOrder"
        Me.myOrder.Size = New System.Drawing.Size(248, 20)
        Me.myOrder.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(25, 132)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "ORDER BY"
        '
        'ok_button
        '
        Me.ok_button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ok_button.Location = New System.Drawing.Point(153, 230)
        Me.ok_button.Name = "ok_button"
        Me.ok_button.Size = New System.Drawing.Size(75, 23)
        Me.ok_button.TabIndex = 6
        Me.ok_button.Text = "OK"
        Me.ok_button.UseVisualStyleBackColor = True
        '
        'cancel_button
        '
        Me.cancel_button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancel_button.Location = New System.Drawing.Point(266, 230)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(75, 23)
        Me.cancel_button.TabIndex = 7
        Me.cancel_button.Text = "Cancel"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'outCol
        '
        Me.outCol.Location = New System.Drawing.Point(117, 186)
        Me.outCol.Name = "outCol"
        Me.outCol.Size = New System.Drawing.Size(96, 20)
        Me.outCol.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(25, 189)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Output Column"
        '
        'GetCommand
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(372, 272)
        Me.Controls.Add(Me.outCol)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.ok_button)
        Me.Controls.Add(Me.myOrder)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.myWhere)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.myFrom)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.mySelect)
        Me.Controls.Add(Me.Label1)
        Me.Name = "GetCommand"
        Me.Text = "Get PN from DB"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents mySelect As System.Windows.Forms.TextBox
    Friend WithEvents myFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents myWhere As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents myOrder As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ok_button As System.Windows.Forms.Button
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    Friend WithEvents outCol As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
