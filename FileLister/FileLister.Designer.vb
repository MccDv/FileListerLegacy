<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FileLister
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FileLister))
        Me.cmdStart = New System.Windows.Forms.Button
        Me.lstKeys = New System.Windows.Forms.ListBox
        Me.cmdListFiles = New System.Windows.Forms.Button
        Me.lstValues = New System.Windows.Forms.ListBox
        Me.lstValNames = New System.Windows.Forms.ListBox
        Me.cmdSave = New System.Windows.Forms.Button
        Me.txtOutput = New System.Windows.Forms.TextBox
        Me.lblOutput = New System.Windows.Forms.Label
        Me.cmdClear = New System.Windows.Forms.Button
        Me.txtFileName = New System.Windows.Forms.TextBox
        Me.cmdDriverList = New System.Windows.Forms.Button
        Me.cmdListProps = New System.Windows.Forms.Button
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.cmdImage = New System.Windows.Forms.Button
        Me.cmdUnZip = New System.Windows.Forms.Button
        Me.cmdStartList = New System.Windows.Forms.Button
        Me.cmdListGac = New System.Windows.Forms.Button
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.DriversToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdStart
        '
        Me.cmdStart.Location = New System.Drawing.Point(12, 34)
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(69, 27)
        Me.cmdStart.TabIndex = 0
        Me.cmdStart.Text = "Start"
        Me.cmdStart.UseVisualStyleBackColor = True
        '
        'lstKeys
        '
        Me.lstKeys.FormattingEnabled = True
        Me.lstKeys.Location = New System.Drawing.Point(12, 73)
        Me.lstKeys.Name = "lstKeys"
        Me.lstKeys.Size = New System.Drawing.Size(165, 134)
        Me.lstKeys.TabIndex = 1
        '
        'cmdListFiles
        '
        Me.cmdListFiles.Enabled = False
        Me.cmdListFiles.Location = New System.Drawing.Point(692, 34)
        Me.cmdListFiles.Name = "cmdListFiles"
        Me.cmdListFiles.Size = New System.Drawing.Size(69, 27)
        Me.cmdListFiles.TabIndex = 2
        Me.cmdListFiles.Text = "List Files"
        Me.cmdListFiles.UseVisualStyleBackColor = True
        '
        'lstValues
        '
        Me.lstValues.FormattingEnabled = True
        Me.lstValues.Location = New System.Drawing.Point(183, 73)
        Me.lstValues.Name = "lstValues"
        Me.lstValues.Size = New System.Drawing.Size(168, 134)
        Me.lstValues.TabIndex = 3
        '
        'lstValNames
        '
        Me.lstValNames.FormattingEnabled = True
        Me.lstValNames.Location = New System.Drawing.Point(357, 73)
        Me.lstValNames.Name = "lstValNames"
        Me.lstValNames.Size = New System.Drawing.Size(411, 134)
        Me.lstValNames.TabIndex = 4
        '
        'cmdSave
        '
        Me.cmdSave.Enabled = False
        Me.cmdSave.Location = New System.Drawing.Point(692, 224)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(69, 27)
        Me.cmdSave.TabIndex = 5
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'txtOutput
        '
        Me.txtOutput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOutput.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOutput.Location = New System.Drawing.Point(12, 247)
        Me.txtOutput.MinimumSize = New System.Drawing.Size(4, 82)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOutput.Size = New System.Drawing.Size(674, 190)
        Me.txtOutput.TabIndex = 6
        '
        'lblOutput
        '
        Me.lblOutput.AutoSize = True
        Me.lblOutput.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOutput.Location = New System.Drawing.Point(12, 231)
        Me.lblOutput.Name = "lblOutput"
        Me.lblOutput.Size = New System.Drawing.Size(120, 13)
        Me.lblOutput.TabIndex = 7
        Me.lblOutput.Text = "Contents to Save to File"
        '
        'cmdClear
        '
        Me.cmdClear.Enabled = False
        Me.cmdClear.Location = New System.Drawing.Point(692, 302)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(69, 27)
        Me.cmdClear.TabIndex = 8
        Me.cmdClear.Text = "Clear List"
        Me.cmdClear.UseVisualStyleBackColor = True
        '
        'txtFileName
        '
        Me.txtFileName.BackColor = System.Drawing.SystemColors.Control
        Me.txtFileName.Location = New System.Drawing.Point(154, 224)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(532, 20)
        Me.txtFileName.TabIndex = 9
        Me.txtFileName.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmdDriverList
        '
        Me.cmdDriverList.Enabled = False
        Me.cmdDriverList.Location = New System.Drawing.Point(407, 34)
        Me.cmdDriverList.Name = "cmdDriverList"
        Me.cmdDriverList.Size = New System.Drawing.Size(69, 27)
        Me.cmdDriverList.TabIndex = 11
        Me.cmdDriverList.Text = "Driver List"
        Me.cmdDriverList.UseVisualStyleBackColor = True
        '
        'cmdListProps
        '
        Me.cmdListProps.Enabled = False
        Me.cmdListProps.Location = New System.Drawing.Point(617, 34)
        Me.cmdListProps.Name = "cmdListProps"
        Me.cmdListProps.Size = New System.Drawing.Size(69, 27)
        Me.cmdListProps.TabIndex = 12
        Me.cmdListProps.Text = "List Props"
        Me.cmdListProps.UseVisualStyleBackColor = True
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        '
        'cmdImage
        '
        Me.cmdImage.Location = New System.Drawing.Point(108, 34)
        Me.cmdImage.Name = "cmdImage"
        Me.cmdImage.Size = New System.Drawing.Size(69, 27)
        Me.cmdImage.TabIndex = 13
        Me.cmdImage.Text = "Image CD"
        Me.cmdImage.UseVisualStyleBackColor = True
        '
        'cmdUnZip
        '
        Me.cmdUnZip.Enabled = False
        Me.cmdUnZip.Location = New System.Drawing.Point(204, 34)
        Me.cmdUnZip.Name = "cmdUnZip"
        Me.cmdUnZip.Size = New System.Drawing.Size(76, 27)
        Me.cmdUnZip.TabIndex = 14
        Me.cmdUnZip.Text = "Unzip Files"
        Me.cmdUnZip.UseVisualStyleBackColor = True
        '
        'cmdStartList
        '
        Me.cmdStartList.Enabled = False
        Me.cmdStartList.Location = New System.Drawing.Point(482, 34)
        Me.cmdStartList.Name = "cmdStartList"
        Me.cmdStartList.Size = New System.Drawing.Size(76, 27)
        Me.cmdStartList.TabIndex = 15
        Me.cmdStartList.Text = "Menu Items"
        Me.cmdStartList.UseVisualStyleBackColor = True
        '
        'cmdListGac
        '
        Me.cmdListGac.Enabled = False
        Me.cmdListGac.Location = New System.Drawing.Point(564, 34)
        Me.cmdListGac.Name = "cmdListGac"
        Me.cmdListGac.Size = New System.Drawing.Size(47, 27)
        Me.cmdListGac.TabIndex = 16
        Me.cmdListGac.Text = "GAC"
        Me.cmdListGac.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DriversToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(780, 24)
        Me.MenuStrip1.TabIndex = 17
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'DriversToolStripMenuItem
        '
        Me.DriversToolStripMenuItem.Enabled = False
        Me.DriversToolStripMenuItem.Name = "DriversToolStripMenuItem"
        Me.DriversToolStripMenuItem.Size = New System.Drawing.Size(53, 20)
        Me.DriversToolStripMenuItem.Text = "Drivers"
        '
        'FileLister
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ClientSize = New System.Drawing.Size(780, 452)
        Me.Controls.Add(Me.cmdListGac)
        Me.Controls.Add(Me.cmdStartList)
        Me.Controls.Add(Me.cmdUnZip)
        Me.Controls.Add(Me.cmdImage)
        Me.Controls.Add(Me.cmdListProps)
        Me.Controls.Add(Me.cmdDriverList)
        Me.Controls.Add(Me.txtFileName)
        Me.Controls.Add(Me.cmdClear)
        Me.Controls.Add(Me.lblOutput)
        Me.Controls.Add(Me.txtOutput)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.lstValNames)
        Me.Controls.Add(Me.lstValues)
        Me.Controls.Add(Me.cmdListFiles)
        Me.Controls.Add(Me.lstKeys)
        Me.Controls.Add(Me.cmdStart)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FileLister"
        Me.Text = "File Lister for .Net"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdStart As System.Windows.Forms.Button
    Friend WithEvents lstKeys As System.Windows.Forms.ListBox
    Friend WithEvents cmdListFiles As System.Windows.Forms.Button
    Friend WithEvents lstValues As System.Windows.Forms.ListBox
    Friend WithEvents lstValNames As System.Windows.Forms.ListBox
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents txtOutput As System.Windows.Forms.TextBox
    Friend WithEvents lblOutput As System.Windows.Forms.Label
    Friend WithEvents cmdClear As System.Windows.Forms.Button
    Friend WithEvents txtFileName As System.Windows.Forms.TextBox
    Friend WithEvents cmdDriverList As System.Windows.Forms.Button
    Friend WithEvents cmdListProps As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents cmdImage As System.Windows.Forms.Button
    Friend WithEvents cmdUnZip As System.Windows.Forms.Button
    Friend WithEvents cmdStartList As System.Windows.Forms.Button
    Friend WithEvents cmdListGac As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents DriversToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
