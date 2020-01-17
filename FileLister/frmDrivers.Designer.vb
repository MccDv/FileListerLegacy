<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDrivers
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lstDrivers = New System.Windows.Forms.ListBox
        Me.cmdRemove = New System.Windows.Forms.Button
        Me.chkIncludeBin = New System.Windows.Forms.CheckBox
        Me.cmdDone = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lstDrivers
        '
        Me.lstDrivers.FormattingEnabled = True
        Me.lstDrivers.Location = New System.Drawing.Point(12, 12)
        Me.lstDrivers.Name = "lstDrivers"
        Me.lstDrivers.Size = New System.Drawing.Size(265, 173)
        Me.lstDrivers.TabIndex = 0
        '
        'cmdRemove
        '
        Me.cmdRemove.Enabled = False
        Me.cmdRemove.Location = New System.Drawing.Point(299, 28)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(105, 28)
        Me.cmdRemove.TabIndex = 1
        Me.cmdRemove.Text = "Uninstall Driver"
        Me.cmdRemove.UseVisualStyleBackColor = True
        '
        'chkIncludeBin
        '
        Me.chkIncludeBin.AutoSize = True
        Me.chkIncludeBin.Checked = True
        Me.chkIncludeBin.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkIncludeBin.Location = New System.Drawing.Point(307, 87)
        Me.chkIncludeBin.Name = "chkIncludeBin"
        Me.chkIncludeBin.Size = New System.Drawing.Size(97, 17)
        Me.chkIncludeBin.TabIndex = 2
        Me.chkIncludeBin.Text = "Delete Binaries"
        Me.chkIncludeBin.UseVisualStyleBackColor = True
        '
        'cmdDone
        '
        Me.cmdDone.Location = New System.Drawing.Point(319, 129)
        Me.cmdDone.Name = "cmdDone"
        Me.cmdDone.Size = New System.Drawing.Size(61, 32)
        Me.cmdDone.TabIndex = 3
        Me.cmdDone.Text = "E&xit"
        Me.cmdDone.UseVisualStyleBackColor = True
        '
        'frmDrivers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(424, 207)
        Me.Controls.Add(Me.cmdDone)
        Me.Controls.Add(Me.chkIncludeBin)
        Me.Controls.Add(Me.cmdRemove)
        Me.Controls.Add(Me.lstDrivers)
        Me.Name = "frmDrivers"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Uninstall Drivers"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lstDrivers As System.Windows.Forms.ListBox
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents chkIncludeBin As System.Windows.Forms.CheckBox
    Friend WithEvents cmdDone As System.Windows.Forms.Button
End Class
