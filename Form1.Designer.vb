<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnJCLJOBFilename = New System.Windows.Forms.Button()
        Me.txtJCLJOBFilename = New System.Windows.Forms.TextBox()
        Me.btnSourceFolder = New System.Windows.Forms.Button()
        Me.txtSourceFolderName = New System.Windows.Forms.TextBox()
        Me.txtJCLProclibFoldername = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.btnOutputFolder = New System.Windows.Forms.Button()
        Me.txtOutputFoldername = New System.Windows.Forms.TextBox()
        Me.btnADDILite = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDelimiter = New System.Windows.Forms.TextBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnJCLProclibFolder = New System.Windows.Forms.Button()
        Me.lblCopybookMessage = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'btnJCLJOBFilename
        '
        Me.btnJCLJOBFilename.Location = New System.Drawing.Point(13, 13)
        Me.btnJCLJOBFilename.Name = "btnJCLJOBFilename"
        Me.btnJCLJOBFilename.Size = New System.Drawing.Size(247, 40)
        Me.btnJCLJOBFilename.TabIndex = 0
        Me.btnJCLJOBFilename.Text = "JCL JOB Filename:"
        Me.btnJCLJOBFilename.UseVisualStyleBackColor = True
        '
        'txtJCLJOBFilename
        '
        Me.txtJCLJOBFilename.Location = New System.Drawing.Point(267, 20)
        Me.txtJCLJOBFilename.Name = "txtJCLJOBFilename"
        Me.txtJCLJOBFilename.Size = New System.Drawing.Size(824, 26)
        Me.txtJCLJOBFilename.TabIndex = 1
        '
        'btnSourceFolder
        '
        Me.btnSourceFolder.Location = New System.Drawing.Point(12, 106)
        Me.btnSourceFolder.Name = "btnSourceFolder"
        Me.btnSourceFolder.Size = New System.Drawing.Size(248, 40)
        Me.btnSourceFolder.TabIndex = 6
        Me.btnSourceFolder.Text = "Source Folder:"
        Me.btnSourceFolder.UseVisualStyleBackColor = True
        '
        'txtSourceFolderName
        '
        Me.txtSourceFolderName.Location = New System.Drawing.Point(266, 113)
        Me.txtSourceFolderName.Name = "txtSourceFolderName"
        Me.txtSourceFolderName.Size = New System.Drawing.Size(823, 26)
        Me.txtSourceFolderName.TabIndex = 7
        '
        'txtJCLProclibFoldername
        '
        Me.txtJCLProclibFoldername.Location = New System.Drawing.Point(267, 67)
        Me.txtJCLProclibFoldername.Name = "txtJCLProclibFoldername"
        Me.txtJCLProclibFoldername.Size = New System.Drawing.Size(824, 26)
        Me.txtJCLProclibFoldername.TabIndex = 5
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnOutputFolder
        '
        Me.btnOutputFolder.Location = New System.Drawing.Point(11, 343)
        Me.btnOutputFolder.Name = "btnOutputFolder"
        Me.btnOutputFolder.Size = New System.Drawing.Size(248, 40)
        Me.btnOutputFolder.TabIndex = 10
        Me.btnOutputFolder.Text = "Output Folder:"
        Me.btnOutputFolder.UseVisualStyleBackColor = True
        '
        'txtOutputFoldername
        '
        Me.txtOutputFoldername.Location = New System.Drawing.Point(267, 350)
        Me.txtOutputFoldername.Name = "txtOutputFoldername"
        Me.txtOutputFoldername.Size = New System.Drawing.Size(823, 26)
        Me.txtOutputFoldername.TabIndex = 11
        '
        'btnADDILite
        '
        Me.btnADDILite.Location = New System.Drawing.Point(13, 433)
        Me.btnADDILite.Name = "btnADDILite"
        Me.btnADDILite.Size = New System.Drawing.Size(96, 37)
        Me.btnADDILite.TabIndex = 13
        Me.btnADDILite.Text = "ADDILite"
        Me.btnADDILite.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(115, 441)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 20)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Delimiter:"
        '
        'txtDelimiter
        '
        Me.txtDelimiter.Location = New System.Drawing.Point(196, 438)
        Me.txtDelimiter.Name = "txtDelimiter"
        Me.txtDelimiter.Size = New System.Drawing.Size(31, 26)
        Me.txtDelimiter.TabIndex = 12
        Me.txtDelimiter.Text = "|"
        Me.txtDelimiter.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(1003, 433)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(86, 37)
        Me.btnClose.TabIndex = 14
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnJCLProclibFolder
        '
        Me.btnJCLProclibFolder.Location = New System.Drawing.Point(12, 60)
        Me.btnJCLProclibFolder.Name = "btnJCLProclibFolder"
        Me.btnJCLProclibFolder.Size = New System.Drawing.Size(248, 40)
        Me.btnJCLProclibFolder.TabIndex = 4
        Me.btnJCLProclibFolder.Text = "JCL Proclib Folder:"
        Me.btnJCLProclibFolder.UseVisualStyleBackColor = True
        '
        'lblCopybookMessage
        '
        Me.lblCopybookMessage.AutoSize = True
        Me.lblCopybookMessage.Location = New System.Drawing.Point(25, 477)
        Me.lblCopybookMessage.Name = "lblCopybookMessage"
        Me.lblCopybookMessage.Size = New System.Drawing.Size(57, 20)
        Me.lblCopybookMessage.TabIndex = 15
        Me.lblCopybookMessage.Text = "Label2"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(13, 507)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(1087, 23)
        Me.ProgressBar1.TabIndex = 16
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1122, 542)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.lblCopybookMessage)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.txtDelimiter)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnADDILite)
        Me.Controls.Add(Me.txtOutputFoldername)
        Me.Controls.Add(Me.btnOutputFolder)
        Me.Controls.Add(Me.txtJCLProclibFoldername)
        Me.Controls.Add(Me.btnJCLProclibFolder)
        Me.Controls.Add(Me.txtSourceFolderName)
        Me.Controls.Add(Me.btnSourceFolder)
        Me.Controls.Add(Me.txtJCLJOBFilename)
        Me.Controls.Add(Me.btnJCLJOBFilename)
        Me.Name = "Form1"
        Me.Text = "ADDILite"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnJCLJOBFilename As Button
    Friend WithEvents txtJCLJOBFilename As TextBox
    Friend WithEvents btnSourceFolder As Button
    Friend WithEvents txtSourceFolderName As TextBox
    Friend WithEvents txtJCLProclibFoldername As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents btnOutputFolder As Button
    Friend WithEvents txtOutputFoldername As TextBox
    Friend WithEvents btnADDILite As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtDelimiter As TextBox
    Friend WithEvents btnClose As Button
    Friend WithEvents btnJCLProclibFolder As Button
    Friend WithEvents lblCopybookMessage As Label
    Friend WithEvents ProgressBar1 As ProgressBar
End Class
