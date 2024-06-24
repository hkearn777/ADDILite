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
        Me.txtJCLJOBFolderName = New System.Windows.Forms.TextBox()
        Me.btnSourceFolder = New System.Windows.Forms.Button()
        Me.txtSourceFolderName = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.btnOutputFolder = New System.Windows.Forms.Button()
        Me.txtOutputFoldername = New System.Windows.Forms.TextBox()
        Me.btnADDILite = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDelimiter = New System.Windows.Forms.TextBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.lblCopybookMessage = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.cbLogStmt = New System.Windows.Forms.CheckBox()
        Me.lblProcessingJob = New System.Windows.Forms.Label()
        Me.lblProcessingSource = New System.Windows.Forms.Label()
        Me.lblProcessingWorksheet = New System.Windows.Forms.Label()
        Me.cbScanModeOnly = New System.Windows.Forms.CheckBox()
        Me.btnDataGatheringForm = New System.Windows.Forms.Button()
        Me.txtDataGatheringForm = New System.Windows.Forms.TextBox()
        Me.btnTelonFolder = New System.Windows.Forms.Button()
        Me.txtTelonFoldername = New System.Windows.Forms.TextBox()
        Me.btnScreenMapsFolder = New System.Windows.Forms.Button()
        Me.txtScreenMapsFolderName = New System.Windows.Forms.TextBox()
        Me.cbJOBS = New System.Windows.Forms.CheckBox()
        Me.cbJobComments = New System.Windows.Forms.CheckBox()
        Me.cbPrograms = New System.Windows.Forms.CheckBox()
        Me.cbFiles = New System.Windows.Forms.CheckBox()
        Me.cbRecords = New System.Windows.Forms.CheckBox()
        Me.cbFields = New System.Windows.Forms.CheckBox()
        Me.cbComments = New System.Windows.Forms.CheckBox()
        Me.cbexecSQL = New System.Windows.Forms.CheckBox()
        Me.cbexecCICS = New System.Windows.Forms.CheckBox()
        Me.cbIMS = New System.Windows.Forms.CheckBox()
        Me.cbCalls = New System.Windows.Forms.CheckBox()
        Me.cbScreenMaps = New System.Windows.Forms.CheckBox()
        Me.cbLibraries = New System.Windows.Forms.CheckBox()
        Me.cbBusinessRules = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'btnJCLJOBFilename
        '
        Me.btnJCLJOBFilename.Location = New System.Drawing.Point(13, 66)
        Me.btnJCLJOBFilename.Name = "btnJCLJOBFilename"
        Me.btnJCLJOBFilename.Size = New System.Drawing.Size(247, 40)
        Me.btnJCLJOBFilename.TabIndex = 2
        Me.btnJCLJOBFilename.Text = "JCL JOB Folder:"
        Me.btnJCLJOBFilename.UseVisualStyleBackColor = True
        '
        'txtJCLJOBFolderName
        '
        Me.txtJCLJOBFolderName.Location = New System.Drawing.Point(267, 73)
        Me.txtJCLJOBFolderName.Name = "txtJCLJOBFolderName"
        Me.txtJCLJOBFolderName.Size = New System.Drawing.Size(993, 26)
        Me.txtJCLJOBFolderName.TabIndex = 3
        '
        'btnSourceFolder
        '
        Me.btnSourceFolder.Location = New System.Drawing.Point(13, 114)
        Me.btnSourceFolder.Name = "btnSourceFolder"
        Me.btnSourceFolder.Size = New System.Drawing.Size(248, 40)
        Me.btnSourceFolder.TabIndex = 4
        Me.btnSourceFolder.Text = "Source Folder:"
        Me.btnSourceFolder.UseVisualStyleBackColor = True
        '
        'txtSourceFolderName
        '
        Me.txtSourceFolderName.Location = New System.Drawing.Point(267, 121)
        Me.txtSourceFolderName.Name = "txtSourceFolderName"
        Me.txtSourceFolderName.Size = New System.Drawing.Size(994, 26)
        Me.txtSourceFolderName.TabIndex = 5
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnOutputFolder
        '
        Me.btnOutputFolder.Location = New System.Drawing.Point(15, 262)
        Me.btnOutputFolder.Name = "btnOutputFolder"
        Me.btnOutputFolder.Size = New System.Drawing.Size(248, 40)
        Me.btnOutputFolder.TabIndex = 6
        Me.btnOutputFolder.Text = "Output Folder:"
        Me.btnOutputFolder.UseVisualStyleBackColor = True
        '
        'txtOutputFoldername
        '
        Me.txtOutputFoldername.Location = New System.Drawing.Point(267, 269)
        Me.txtOutputFoldername.Name = "txtOutputFoldername"
        Me.txtOutputFoldername.Size = New System.Drawing.Size(994, 26)
        Me.txtOutputFoldername.TabIndex = 7
        '
        'btnADDILite
        '
        Me.btnADDILite.Location = New System.Drawing.Point(1055, 363)
        Me.btnADDILite.Name = "btnADDILite"
        Me.btnADDILite.Size = New System.Drawing.Size(96, 53)
        Me.btnADDILite.TabIndex = 8
        Me.btnADDILite.Text = "ADDILite"
        Me.btnADDILite.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(357, 381)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 20)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Delimiter:"
        '
        'txtDelimiter
        '
        Me.txtDelimiter.Location = New System.Drawing.Point(438, 378)
        Me.txtDelimiter.Name = "txtDelimiter"
        Me.txtDelimiter.Size = New System.Drawing.Size(31, 26)
        Me.txtDelimiter.TabIndex = 12
        Me.txtDelimiter.Text = "|"
        Me.txtDelimiter.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(1172, 363)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(88, 51)
        Me.btnClose.TabIndex = 13
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'lblCopybookMessage
        '
        Me.lblCopybookMessage.AutoSize = True
        Me.lblCopybookMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCopybookMessage.Location = New System.Drawing.Point(16, 419)
        Me.lblCopybookMessage.Name = "lblCopybookMessage"
        Me.lblCopybookMessage.Size = New System.Drawing.Size(836, 20)
        Me.lblCopybookMessage.TabIndex = 15
        Me.lblCopybookMessage.Text = "Select Data Gathering Form, JCL JOB, Source, Telon, and Output buttons then click" &
    " ADDILite button to process files."
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 607)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(1248, 30)
        Me.ProgressBar1.TabIndex = 16
        '
        'cbLogStmt
        '
        Me.cbLogStmt.AutoSize = True
        Me.cbLogStmt.Location = New System.Drawing.Point(184, 379)
        Me.cbLogStmt.Name = "cbLogStmt"
        Me.cbLogStmt.Size = New System.Drawing.Size(141, 24)
        Me.cbLogStmt.TabIndex = 10
        Me.cbLogStmt.Text = "Log Stmt Array"
        Me.cbLogStmt.UseVisualStyleBackColor = True
        '
        'lblProcessingJob
        '
        Me.lblProcessingJob.AutoSize = True
        Me.lblProcessingJob.Location = New System.Drawing.Point(38, 488)
        Me.lblProcessingJob.Name = "lblProcessingJob"
        Me.lblProcessingJob.Size = New System.Drawing.Size(121, 20)
        Me.lblProcessingJob.TabIndex = 19
        Me.lblProcessingJob.Text = "Processing Job:"
        '
        'lblProcessingSource
        '
        Me.lblProcessingSource.AutoSize = True
        Me.lblProcessingSource.Location = New System.Drawing.Point(56, 519)
        Me.lblProcessingSource.Name = "lblProcessingSource"
        Me.lblProcessingSource.Size = New System.Drawing.Size(146, 20)
        Me.lblProcessingSource.TabIndex = 20
        Me.lblProcessingSource.Text = "Processing Source:"
        '
        'lblProcessingWorksheet
        '
        Me.lblProcessingWorksheet.AutoSize = True
        Me.lblProcessingWorksheet.Location = New System.Drawing.Point(75, 552)
        Me.lblProcessingWorksheet.Name = "lblProcessingWorksheet"
        Me.lblProcessingWorksheet.Size = New System.Drawing.Size(176, 20)
        Me.lblProcessingWorksheet.TabIndex = 21
        Me.lblProcessingWorksheet.Text = "Processing Worksheet: "
        '
        'cbScanModeOnly
        '
        Me.cbScanModeOnly.AutoSize = True
        Me.cbScanModeOnly.Location = New System.Drawing.Point(13, 379)
        Me.cbScanModeOnly.Name = "cbScanModeOnly"
        Me.cbScanModeOnly.Size = New System.Drawing.Size(151, 24)
        Me.cbScanModeOnly.TabIndex = 9
        Me.cbScanModeOnly.Text = "Scan Mode Only"
        Me.cbScanModeOnly.UseVisualStyleBackColor = True
        '
        'btnDataGatheringForm
        '
        Me.btnDataGatheringForm.Location = New System.Drawing.Point(13, 13)
        Me.btnDataGatheringForm.Name = "btnDataGatheringForm"
        Me.btnDataGatheringForm.Size = New System.Drawing.Size(247, 42)
        Me.btnDataGatheringForm.TabIndex = 0
        Me.btnDataGatheringForm.Text = "Data Gathering Form"
        Me.btnDataGatheringForm.UseVisualStyleBackColor = True
        '
        'txtDataGatheringForm
        '
        Me.txtDataGatheringForm.Location = New System.Drawing.Point(267, 21)
        Me.txtDataGatheringForm.Name = "txtDataGatheringForm"
        Me.txtDataGatheringForm.Size = New System.Drawing.Size(993, 26)
        Me.txtDataGatheringForm.TabIndex = 1
        '
        'btnTelonFolder
        '
        Me.btnTelonFolder.Location = New System.Drawing.Point(14, 163)
        Me.btnTelonFolder.Name = "btnTelonFolder"
        Me.btnTelonFolder.Size = New System.Drawing.Size(248, 40)
        Me.btnTelonFolder.TabIndex = 24
        Me.btnTelonFolder.Text = "Telon Members Folder"
        Me.btnTelonFolder.UseVisualStyleBackColor = True
        '
        'txtTelonFoldername
        '
        Me.txtTelonFoldername.Location = New System.Drawing.Point(267, 170)
        Me.txtTelonFoldername.Name = "txtTelonFoldername"
        Me.txtTelonFoldername.Size = New System.Drawing.Size(994, 26)
        Me.txtTelonFoldername.TabIndex = 25
        '
        'btnScreenMapsFolder
        '
        Me.btnScreenMapsFolder.Location = New System.Drawing.Point(16, 213)
        Me.btnScreenMapsFolder.Name = "btnScreenMapsFolder"
        Me.btnScreenMapsFolder.Size = New System.Drawing.Size(248, 40)
        Me.btnScreenMapsFolder.TabIndex = 26
        Me.btnScreenMapsFolder.Text = "Screen Maps Folder"
        Me.btnScreenMapsFolder.UseVisualStyleBackColor = True
        '
        'txtScreenMapsFolderName
        '
        Me.txtScreenMapsFolderName.Location = New System.Drawing.Point(270, 220)
        Me.txtScreenMapsFolderName.Name = "txtScreenMapsFolderName"
        Me.txtScreenMapsFolderName.Size = New System.Drawing.Size(994, 26)
        Me.txtScreenMapsFolderName.TabIndex = 27
        '
        'cbJOBS
        '
        Me.cbJOBS.AutoSize = True
        Me.cbJOBS.Checked = True
        Me.cbJOBS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbJOBS.Location = New System.Drawing.Point(18, 318)
        Me.cbJOBS.Name = "cbJOBS"
        Me.cbJOBS.Size = New System.Drawing.Size(77, 24)
        Me.cbJOBS.TabIndex = 30
        Me.cbJOBS.Text = "JOBS"
        Me.cbJOBS.UseVisualStyleBackColor = True
        '
        'cbJobComments
        '
        Me.cbJobComments.AutoSize = True
        Me.cbJobComments.Checked = True
        Me.cbJobComments.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbJobComments.Location = New System.Drawing.Point(101, 318)
        Me.cbJobComments.Name = "cbJobComments"
        Me.cbJobComments.Size = New System.Drawing.Size(147, 24)
        Me.cbJobComments.TabIndex = 31
        Me.cbJobComments.Text = "JOB Comments"
        Me.cbJobComments.UseVisualStyleBackColor = True
        '
        'cbPrograms
        '
        Me.cbPrograms.AutoSize = True
        Me.cbPrograms.Checked = True
        Me.cbPrograms.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbPrograms.Location = New System.Drawing.Point(254, 318)
        Me.cbPrograms.Name = "cbPrograms"
        Me.cbPrograms.Size = New System.Drawing.Size(103, 24)
        Me.cbPrograms.TabIndex = 32
        Me.cbPrograms.Text = "Programs"
        Me.cbPrograms.UseVisualStyleBackColor = True
        '
        'cbFiles
        '
        Me.cbFiles.AutoSize = True
        Me.cbFiles.Checked = True
        Me.cbFiles.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbFiles.Location = New System.Drawing.Point(363, 318)
        Me.cbFiles.Name = "cbFiles"
        Me.cbFiles.Size = New System.Drawing.Size(68, 24)
        Me.cbFiles.TabIndex = 33
        Me.cbFiles.Text = "Files"
        Me.cbFiles.UseVisualStyleBackColor = True
        '
        'cbRecords
        '
        Me.cbRecords.AutoSize = True
        Me.cbRecords.Checked = True
        Me.cbRecords.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbRecords.Location = New System.Drawing.Point(437, 318)
        Me.cbRecords.Name = "cbRecords"
        Me.cbRecords.Size = New System.Drawing.Size(95, 24)
        Me.cbRecords.TabIndex = 34
        Me.cbRecords.Text = "Records"
        Me.cbRecords.UseVisualStyleBackColor = True
        '
        'cbFields
        '
        Me.cbFields.AutoSize = True
        Me.cbFields.Checked = True
        Me.cbFields.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbFields.Location = New System.Drawing.Point(538, 318)
        Me.cbFields.Name = "cbFields"
        Me.cbFields.Size = New System.Drawing.Size(77, 24)
        Me.cbFields.TabIndex = 35
        Me.cbFields.Text = "Fields"
        Me.cbFields.UseVisualStyleBackColor = True
        '
        'cbComments
        '
        Me.cbComments.AutoSize = True
        Me.cbComments.Checked = True
        Me.cbComments.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbComments.Location = New System.Drawing.Point(621, 317)
        Me.cbComments.Name = "cbComments"
        Me.cbComments.Size = New System.Drawing.Size(112, 24)
        Me.cbComments.TabIndex = 36
        Me.cbComments.Text = "Comments"
        Me.cbComments.UseVisualStyleBackColor = True
        '
        'cbexecSQL
        '
        Me.cbexecSQL.AutoSize = True
        Me.cbexecSQL.Checked = True
        Me.cbexecSQL.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbexecSQL.Location = New System.Drawing.Point(739, 317)
        Me.cbexecSQL.Name = "cbexecSQL"
        Me.cbexecSQL.Size = New System.Drawing.Size(100, 24)
        Me.cbexecSQL.TabIndex = 37
        Me.cbexecSQL.Text = "execSQL"
        Me.cbexecSQL.UseVisualStyleBackColor = True
        '
        'cbexecCICS
        '
        Me.cbexecCICS.AutoSize = True
        Me.cbexecCICS.Checked = True
        Me.cbexecCICS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbexecCICS.Location = New System.Drawing.Point(845, 318)
        Me.cbexecCICS.Name = "cbexecCICS"
        Me.cbexecCICS.Size = New System.Drawing.Size(106, 24)
        Me.cbexecCICS.TabIndex = 38
        Me.cbexecCICS.Text = "execCICS"
        Me.cbexecCICS.UseVisualStyleBackColor = True
        '
        'cbIMS
        '
        Me.cbIMS.AutoSize = True
        Me.cbIMS.Checked = True
        Me.cbIMS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbIMS.Location = New System.Drawing.Point(957, 318)
        Me.cbIMS.Name = "cbIMS"
        Me.cbIMS.Size = New System.Drawing.Size(64, 24)
        Me.cbIMS.TabIndex = 39
        Me.cbIMS.Text = "IMS"
        Me.cbIMS.UseVisualStyleBackColor = True
        '
        'cbCalls
        '
        Me.cbCalls.AutoSize = True
        Me.cbCalls.Checked = True
        Me.cbCalls.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbCalls.Location = New System.Drawing.Point(1027, 317)
        Me.cbCalls.Name = "cbCalls"
        Me.cbCalls.Size = New System.Drawing.Size(69, 24)
        Me.cbCalls.TabIndex = 40
        Me.cbCalls.Text = "Calls"
        Me.cbCalls.UseVisualStyleBackColor = True
        '
        'cbScreenMaps
        '
        Me.cbScreenMaps.AutoSize = True
        Me.cbScreenMaps.Checked = True
        Me.cbScreenMaps.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbScreenMaps.Location = New System.Drawing.Point(16, 348)
        Me.cbScreenMaps.Name = "cbScreenMaps"
        Me.cbScreenMaps.Size = New System.Drawing.Size(129, 24)
        Me.cbScreenMaps.TabIndex = 41
        Me.cbScreenMaps.Text = "Screen Maps"
        Me.cbScreenMaps.UseVisualStyleBackColor = True
        '
        'cbLibraries
        '
        Me.cbLibraries.AutoSize = True
        Me.cbLibraries.Checked = True
        Me.cbLibraries.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbLibraries.Location = New System.Drawing.Point(151, 348)
        Me.cbLibraries.Name = "cbLibraries"
        Me.cbLibraries.Size = New System.Drawing.Size(95, 24)
        Me.cbLibraries.TabIndex = 42
        Me.cbLibraries.Text = "Libraries"
        Me.cbLibraries.UseVisualStyleBackColor = True
        '
        'cbBusinessRules
        '
        Me.cbBusinessRules.AutoSize = True
        Me.cbBusinessRules.Checked = True
        Me.cbBusinessRules.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbBusinessRules.Location = New System.Drawing.Point(252, 348)
        Me.cbBusinessRules.Name = "cbBusinessRules"
        Me.cbBusinessRules.Size = New System.Drawing.Size(145, 24)
        Me.cbBusinessRules.TabIndex = 43
        Me.cbBusinessRules.Text = "Business Rules"
        Me.cbBusinessRules.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1279, 652)
        Me.Controls.Add(Me.cbBusinessRules)
        Me.Controls.Add(Me.cbLibraries)
        Me.Controls.Add(Me.cbScreenMaps)
        Me.Controls.Add(Me.cbCalls)
        Me.Controls.Add(Me.cbIMS)
        Me.Controls.Add(Me.cbexecCICS)
        Me.Controls.Add(Me.cbexecSQL)
        Me.Controls.Add(Me.cbComments)
        Me.Controls.Add(Me.cbFields)
        Me.Controls.Add(Me.cbRecords)
        Me.Controls.Add(Me.cbFiles)
        Me.Controls.Add(Me.cbPrograms)
        Me.Controls.Add(Me.cbJobComments)
        Me.Controls.Add(Me.cbJOBS)
        Me.Controls.Add(Me.txtScreenMapsFolderName)
        Me.Controls.Add(Me.btnScreenMapsFolder)
        Me.Controls.Add(Me.txtTelonFoldername)
        Me.Controls.Add(Me.btnTelonFolder)
        Me.Controls.Add(Me.txtDataGatheringForm)
        Me.Controls.Add(Me.btnDataGatheringForm)
        Me.Controls.Add(Me.cbScanModeOnly)
        Me.Controls.Add(Me.lblProcessingWorksheet)
        Me.Controls.Add(Me.lblProcessingSource)
        Me.Controls.Add(Me.lblProcessingJob)
        Me.Controls.Add(Me.cbLogStmt)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.lblCopybookMessage)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.txtDelimiter)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnADDILite)
        Me.Controls.Add(Me.txtOutputFoldername)
        Me.Controls.Add(Me.btnOutputFolder)
        Me.Controls.Add(Me.txtSourceFolderName)
        Me.Controls.Add(Me.btnSourceFolder)
        Me.Controls.Add(Me.txtJCLJOBFolderName)
        Me.Controls.Add(Me.btnJCLJOBFilename)
        Me.Name = "Form1"
        Me.Text = "ADDILite"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnJCLJOBFilename As Button
    Friend WithEvents txtJCLJOBFolderName As TextBox
    Friend WithEvents btnSourceFolder As Button
    Friend WithEvents txtSourceFolderName As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents btnOutputFolder As Button
    Friend WithEvents txtOutputFoldername As TextBox
    Friend WithEvents btnADDILite As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtDelimiter As TextBox
    Friend WithEvents btnClose As Button
    Friend WithEvents lblCopybookMessage As Label
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents cbLogStmt As CheckBox
    Friend WithEvents lblProcessingJob As Label
    Friend WithEvents lblProcessingSource As Label
    Friend WithEvents lblProcessingWorksheet As Label
    Friend WithEvents cbScanModeOnly As CheckBox
    Friend WithEvents btnDataGatheringForm As Button
    Friend WithEvents txtDataGatheringForm As TextBox
    Friend WithEvents btnTelonFolder As Button
    Friend WithEvents txtTelonFoldername As TextBox
    Friend WithEvents btnScreenMapsFolder As Button
    Friend WithEvents txtScreenMapsFolderName As TextBox
    Friend WithEvents cbJOBS As CheckBox
    Friend WithEvents cbJobComments As CheckBox
    Friend WithEvents cbPrograms As CheckBox
    Friend WithEvents cbFiles As CheckBox
    Friend WithEvents cbRecords As CheckBox
    Friend WithEvents cbFields As CheckBox
    Friend WithEvents cbComments As CheckBox
    Friend WithEvents cbexecSQL As CheckBox
    Friend WithEvents cbexecCICS As CheckBox
    Friend WithEvents cbIMS As CheckBox
    Friend WithEvents cbCalls As CheckBox
    Friend WithEvents cbScreenMaps As CheckBox
    Friend WithEvents cbLibraries As CheckBox
    Friend WithEvents cbBusinessRules As CheckBox
End Class
