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
        Me.txtJCLJOBFolder = New System.Windows.Forms.TextBox()
        Me.txtCobolFolder = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.txtOutputFolder = New System.Windows.Forms.TextBox()
        Me.btnADDILite = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDelimiter = New System.Windows.Forms.TextBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.lblStatusMessage = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.cbLogStmt = New System.Windows.Forms.CheckBox()
        Me.lblProcessingJob = New System.Windows.Forms.Label()
        Me.lblProcessingSource = New System.Windows.Forms.Label()
        Me.lblProcessingWorksheet = New System.Windows.Forms.Label()
        Me.cbScanModeOnly = New System.Windows.Forms.CheckBox()
        Me.btnDataGatheringForm = New System.Windows.Forms.Button()
        Me.txtDataGatheringForm = New System.Windows.Forms.TextBox()
        Me.txtTelonFolder = New System.Windows.Forms.TextBox()
        Me.txtScreenMapsFolder = New System.Windows.Forms.TextBox()
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
        Me.cbDataCom = New System.Windows.Forms.CheckBox()
        Me.btnSandbox = New System.Windows.Forms.Button()
        Me.lblInitDirectory = New System.Windows.Forms.Label()
        Me.txtProcFolder = New System.Windows.Forms.TextBox()
        Me.btnAppFolder = New System.Windows.Forms.Button()
        Me.txtAppFolder = New System.Windows.Forms.TextBox()
        Me.lblJOBFolder = New System.Windows.Forms.Label()
        Me.lblPROCFolder = New System.Windows.Forms.Label()
        Me.lblCOBOLFolder = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblCopybooksFolder = New System.Windows.Forms.Label()
        Me.txtCopybookFolder = New System.Windows.Forms.TextBox()
        Me.lblEasytrieveFolder = New System.Windows.Forms.Label()
        Me.txtEasytrieveFolder = New System.Windows.Forms.TextBox()
        Me.lblDECLGenFolder = New System.Windows.Forms.Label()
        Me.txtDECLGenFolder = New System.Windows.Forms.TextBox()
        Me.lblASMFolder = New System.Windows.Forms.Label()
        Me.txtASMFolder = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblTelonFolder = New System.Windows.Forms.Label()
        Me.lblScreensFolder = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblOutputFolder = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtPUMLFolder = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtExpandedFolder = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtBusinessRulesFolder = New System.Windows.Forms.TextBox()
        Me.cbInstream = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'txtJCLJOBFolder
        '
        Me.txtJCLJOBFolder.Location = New System.Drawing.Point(169, 182)
        Me.txtJCLJOBFolder.Name = "txtJCLJOBFolder"
        Me.txtJCLJOBFolder.Size = New System.Drawing.Size(156, 26)
        Me.txtJCLJOBFolder.TabIndex = 3
        '
        'txtCobolFolder
        '
        Me.txtCobolFolder.Location = New System.Drawing.Point(169, 242)
        Me.txtCobolFolder.Name = "txtCobolFolder"
        Me.txtCobolFolder.Size = New System.Drawing.Size(156, 26)
        Me.txtCobolFolder.TabIndex = 5
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'txtOutputFolder
        '
        Me.txtOutputFolder.Location = New System.Drawing.Point(169, 411)
        Me.txtOutputFolder.Name = "txtOutputFolder"
        Me.txtOutputFolder.Size = New System.Drawing.Size(156, 26)
        Me.txtOutputFolder.TabIndex = 7
        '
        'btnADDILite
        '
        Me.btnADDILite.Enabled = False
        Me.btnADDILite.Location = New System.Drawing.Point(1061, 574)
        Me.btnADDILite.Name = "btnADDILite"
        Me.btnADDILite.Size = New System.Drawing.Size(96, 53)
        Me.btnADDILite.TabIndex = 8
        Me.btnADDILite.Text = "ADDILite"
        Me.btnADDILite.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(1103, 513)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 20)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Delimiter:"
        '
        'txtDelimiter
        '
        Me.txtDelimiter.Location = New System.Drawing.Point(1184, 510)
        Me.txtDelimiter.Name = "txtDelimiter"
        Me.txtDelimiter.Size = New System.Drawing.Size(31, 26)
        Me.txtDelimiter.TabIndex = 12
        Me.txtDelimiter.Text = "|"
        Me.txtDelimiter.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(1174, 574)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(88, 51)
        Me.btnClose.TabIndex = 13
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'lblStatusMessage
        '
        Me.lblStatusMessage.AutoSize = True
        Me.lblStatusMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatusMessage.Location = New System.Drawing.Point(14, 607)
        Me.lblStatusMessage.Name = "lblStatusMessage"
        Me.lblStatusMessage.Size = New System.Drawing.Size(836, 20)
        Me.lblStatusMessage.TabIndex = 15
        Me.lblStatusMessage.Text = "Select Data Gathering Form, JCL JOB, Source, Telon, and Output buttons then click" &
    " ADDILite button to process files."
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(14, 642)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(1248, 30)
        Me.ProgressBar1.TabIndex = 16
        '
        'cbLogStmt
        '
        Me.cbLogStmt.AutoSize = True
        Me.cbLogStmt.Location = New System.Drawing.Point(930, 511)
        Me.cbLogStmt.Name = "cbLogStmt"
        Me.cbLogStmt.Size = New System.Drawing.Size(141, 24)
        Me.cbLogStmt.TabIndex = 10
        Me.cbLogStmt.Text = "Log Stmt Array"
        Me.cbLogStmt.UseVisualStyleBackColor = True
        '
        'lblProcessingJob
        '
        Me.lblProcessingJob.AutoSize = True
        Me.lblProcessingJob.Location = New System.Drawing.Point(17, 546)
        Me.lblProcessingJob.Name = "lblProcessingJob"
        Me.lblProcessingJob.Size = New System.Drawing.Size(121, 20)
        Me.lblProcessingJob.TabIndex = 19
        Me.lblProcessingJob.Text = "Processing Job:"
        '
        'lblProcessingSource
        '
        Me.lblProcessingSource.AutoSize = True
        Me.lblProcessingSource.Location = New System.Drawing.Point(371, 546)
        Me.lblProcessingSource.Name = "lblProcessingSource"
        Me.lblProcessingSource.Size = New System.Drawing.Size(146, 20)
        Me.lblProcessingSource.TabIndex = 20
        Me.lblProcessingSource.Text = "Processing Source:"
        '
        'lblProcessingWorksheet
        '
        Me.lblProcessingWorksheet.AutoSize = True
        Me.lblProcessingWorksheet.Location = New System.Drawing.Point(755, 546)
        Me.lblProcessingWorksheet.Name = "lblProcessingWorksheet"
        Me.lblProcessingWorksheet.Size = New System.Drawing.Size(176, 20)
        Me.lblProcessingWorksheet.TabIndex = 21
        Me.lblProcessingWorksheet.Text = "Processing Worksheet: "
        '
        'cbScanModeOnly
        '
        Me.cbScanModeOnly.AutoSize = True
        Me.cbScanModeOnly.Location = New System.Drawing.Point(759, 511)
        Me.cbScanModeOnly.Name = "cbScanModeOnly"
        Me.cbScanModeOnly.Size = New System.Drawing.Size(151, 24)
        Me.cbScanModeOnly.TabIndex = 9
        Me.cbScanModeOnly.Text = "Scan Mode Only"
        Me.cbScanModeOnly.UseVisualStyleBackColor = True
        '
        'btnDataGatheringForm
        '
        Me.btnDataGatheringForm.Enabled = False
        Me.btnDataGatheringForm.Location = New System.Drawing.Point(13, 94)
        Me.btnDataGatheringForm.Name = "btnDataGatheringForm"
        Me.btnDataGatheringForm.Size = New System.Drawing.Size(247, 42)
        Me.btnDataGatheringForm.TabIndex = 0
        Me.btnDataGatheringForm.Text = "Data Gathering Form Filename:"
        Me.btnDataGatheringForm.UseVisualStyleBackColor = True
        '
        'txtDataGatheringForm
        '
        Me.txtDataGatheringForm.Location = New System.Drawing.Point(269, 102)
        Me.txtDataGatheringForm.Name = "txtDataGatheringForm"
        Me.txtDataGatheringForm.Size = New System.Drawing.Size(346, 26)
        Me.txtDataGatheringForm.TabIndex = 1
        '
        'txtTelonFolder
        '
        Me.txtTelonFolder.Location = New System.Drawing.Point(169, 347)
        Me.txtTelonFolder.Name = "txtTelonFolder"
        Me.txtTelonFolder.Size = New System.Drawing.Size(156, 26)
        Me.txtTelonFolder.TabIndex = 25
        '
        'txtScreenMapsFolder
        '
        Me.txtScreenMapsFolder.Location = New System.Drawing.Point(496, 347)
        Me.txtScreenMapsFolder.Name = "txtScreenMapsFolder"
        Me.txtScreenMapsFolder.Size = New System.Drawing.Size(137, 26)
        Me.txtScreenMapsFolder.TabIndex = 27
        '
        'cbJOBS
        '
        Me.cbJOBS.AutoSize = True
        Me.cbJOBS.Checked = True
        Me.cbJOBS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbJOBS.Location = New System.Drawing.Point(30, 479)
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
        Me.cbJobComments.Location = New System.Drawing.Point(113, 479)
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
        Me.cbPrograms.Location = New System.Drawing.Point(266, 479)
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
        Me.cbFiles.Location = New System.Drawing.Point(375, 479)
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
        Me.cbRecords.Location = New System.Drawing.Point(449, 479)
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
        Me.cbFields.Location = New System.Drawing.Point(550, 479)
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
        Me.cbComments.Location = New System.Drawing.Point(633, 478)
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
        Me.cbexecSQL.Location = New System.Drawing.Point(751, 478)
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
        Me.cbexecCICS.Location = New System.Drawing.Point(857, 479)
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
        Me.cbIMS.Location = New System.Drawing.Point(969, 479)
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
        Me.cbCalls.Location = New System.Drawing.Point(1148, 478)
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
        Me.cbScreenMaps.Location = New System.Drawing.Point(28, 509)
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
        Me.cbLibraries.Location = New System.Drawing.Point(163, 509)
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
        Me.cbBusinessRules.Location = New System.Drawing.Point(264, 509)
        Me.cbBusinessRules.Name = "cbBusinessRules"
        Me.cbBusinessRules.Size = New System.Drawing.Size(145, 24)
        Me.cbBusinessRules.TabIndex = 43
        Me.cbBusinessRules.Text = "Business Rules"
        Me.cbBusinessRules.UseVisualStyleBackColor = True
        '
        'cbDataCom
        '
        Me.cbDataCom.AutoSize = True
        Me.cbDataCom.Checked = True
        Me.cbDataCom.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbDataCom.Location = New System.Drawing.Point(1039, 479)
        Me.cbDataCom.Name = "cbDataCom"
        Me.cbDataCom.Size = New System.Drawing.Size(103, 24)
        Me.cbDataCom.TabIndex = 44
        Me.cbDataCom.Text = "DataCom"
        Me.cbDataCom.UseVisualStyleBackColor = True
        '
        'btnSandbox
        '
        Me.btnSandbox.Location = New System.Drawing.Point(12, 16)
        Me.btnSandbox.Name = "btnSandbox"
        Me.btnSandbox.Size = New System.Drawing.Size(98, 36)
        Me.btnSandbox.TabIndex = 45
        Me.btnSandbox.Text = "Sandbox"
        Me.btnSandbox.UseVisualStyleBackColor = True
        '
        'lblInitDirectory
        '
        Me.lblInitDirectory.AutoSize = True
        Me.lblInitDirectory.Location = New System.Drawing.Point(116, 24)
        Me.lblInitDirectory.Name = "lblInitDirectory"
        Me.lblInitDirectory.Size = New System.Drawing.Size(94, 20)
        Me.lblInitDirectory.TabIndex = 46
        Me.lblInitDirectory.Text = "InitDirectory"
        '
        'txtProcFolder
        '
        Me.txtProcFolder.Location = New System.Drawing.Point(494, 182)
        Me.txtProcFolder.Name = "txtProcFolder"
        Me.txtProcFolder.Size = New System.Drawing.Size(139, 26)
        Me.txtProcFolder.TabIndex = 48
        '
        'btnAppFolder
        '
        Me.btnAppFolder.Location = New System.Drawing.Point(12, 56)
        Me.btnAppFolder.Name = "btnAppFolder"
        Me.btnAppFolder.Size = New System.Drawing.Size(166, 32)
        Me.btnAppFolder.TabIndex = 49
        Me.btnAppFolder.Text = "Application Folder:"
        Me.btnAppFolder.UseVisualStyleBackColor = True
        '
        'txtAppFolder
        '
        Me.txtAppFolder.Location = New System.Drawing.Point(184, 59)
        Me.txtAppFolder.Name = "txtAppFolder"
        Me.txtAppFolder.Size = New System.Drawing.Size(214, 26)
        Me.txtAppFolder.TabIndex = 50
        '
        'lblJOBFolder
        '
        Me.lblJOBFolder.AutoSize = True
        Me.lblJOBFolder.Location = New System.Drawing.Point(34, 185)
        Me.lblJOBFolder.Name = "lblJOBFolder"
        Me.lblJOBFolder.Size = New System.Drawing.Size(78, 20)
        Me.lblJOBFolder.TabIndex = 51
        Me.lblJOBFolder.Text = "JOBS (0):"
        '
        'lblPROCFolder
        '
        Me.lblPROCFolder.AutoSize = True
        Me.lblPROCFolder.Location = New System.Drawing.Point(345, 185)
        Me.lblPROCFolder.Name = "lblPROCFolder"
        Me.lblPROCFolder.Size = New System.Drawing.Size(92, 20)
        Me.lblPROCFolder.TabIndex = 52
        Me.lblPROCFolder.Text = "PROCS (0):"
        '
        'lblCOBOLFolder
        '
        Me.lblCOBOLFolder.AutoSize = True
        Me.lblCOBOLFolder.Location = New System.Drawing.Point(34, 248)
        Me.lblCOBOLFolder.Name = "lblCOBOLFolder"
        Me.lblCOBOLFolder.Size = New System.Drawing.Size(91, 20)
        Me.lblCOBOLFolder.TabIndex = 53
        Me.lblCOBOLFolder.Text = "COBOL (0):"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label2.Location = New System.Drawing.Point(16, 154)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(122, 20)
        Me.Label2.TabIndex = 54
        Me.Label2.Text = "JCL FOLDERS:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label3.Location = New System.Drawing.Point(20, 218)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(163, 20)
        Me.Label3.TabIndex = 55
        Me.Label3.Text = "SOURCE FOLDERS:"
        '
        'lblCopybooksFolder
        '
        Me.lblCopybooksFolder.AutoSize = True
        Me.lblCopybooksFolder.Location = New System.Drawing.Point(345, 245)
        Me.lblCopybooksFolder.Name = "lblCopybooksFolder"
        Me.lblCopybooksFolder.Size = New System.Drawing.Size(115, 20)
        Me.lblCopybooksFolder.TabIndex = 56
        Me.lblCopybooksFolder.Text = "Copybooks (0):"
        '
        'txtCopybookFolder
        '
        Me.txtCopybookFolder.Location = New System.Drawing.Point(496, 242)
        Me.txtCopybookFolder.Name = "txtCopybookFolder"
        Me.txtCopybookFolder.Size = New System.Drawing.Size(137, 26)
        Me.txtCopybookFolder.TabIndex = 57
        '
        'lblEasytrieveFolder
        '
        Me.lblEasytrieveFolder.AutoSize = True
        Me.lblEasytrieveFolder.Location = New System.Drawing.Point(901, 245)
        Me.lblEasytrieveFolder.Name = "lblEasytrieveFolder"
        Me.lblEasytrieveFolder.Size = New System.Drawing.Size(109, 20)
        Me.lblEasytrieveFolder.TabIndex = 58
        Me.lblEasytrieveFolder.Text = "Easytrieve (0):"
        '
        'txtEasytrieveFolder
        '
        Me.txtEasytrieveFolder.Location = New System.Drawing.Point(1032, 242)
        Me.txtEasytrieveFolder.Name = "txtEasytrieveFolder"
        Me.txtEasytrieveFolder.Size = New System.Drawing.Size(183, 26)
        Me.txtEasytrieveFolder.TabIndex = 59
        '
        'lblDECLGenFolder
        '
        Me.lblDECLGenFolder.AutoSize = True
        Me.lblDECLGenFolder.Location = New System.Drawing.Point(645, 245)
        Me.lblDECLGenFolder.Name = "lblDECLGenFolder"
        Me.lblDECLGenFolder.Size = New System.Drawing.Size(110, 20)
        Me.lblDECLGenFolder.TabIndex = 60
        Me.lblDECLGenFolder.Text = "DECLGen (0):"
        '
        'txtDECLGenFolder
        '
        Me.txtDECLGenFolder.Location = New System.Drawing.Point(775, 242)
        Me.txtDECLGenFolder.Name = "txtDECLGenFolder"
        Me.txtDECLGenFolder.Size = New System.Drawing.Size(114, 26)
        Me.txtDECLGenFolder.TabIndex = 61
        '
        'lblASMFolder
        '
        Me.lblASMFolder.AutoSize = True
        Me.lblASMFolder.Location = New System.Drawing.Point(34, 284)
        Me.lblASMFolder.Name = "lblASMFolder"
        Me.lblASMFolder.Size = New System.Drawing.Size(111, 20)
        Me.lblASMFolder.TabIndex = 62
        Me.lblASMFolder.Text = "Assembler (0):"
        '
        'txtASMFolder
        '
        Me.txtASMFolder.Location = New System.Drawing.Point(169, 281)
        Me.txtASMFolder.Name = "txtASMFolder"
        Me.txtASMFolder.Size = New System.Drawing.Size(156, 26)
        Me.txtASMFolder.TabIndex = 63
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label4.Location = New System.Drawing.Point(20, 320)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(153, 20)
        Me.Label4.TabIndex = 64
        Me.Label4.Text = "ONLINE FOLDERS:"
        '
        'lblTelonFolder
        '
        Me.lblTelonFolder.AutoSize = True
        Me.lblTelonFolder.Location = New System.Drawing.Point(34, 350)
        Me.lblTelonFolder.Name = "lblTelonFolder"
        Me.lblTelonFolder.Size = New System.Drawing.Size(75, 20)
        Me.lblTelonFolder.TabIndex = 65
        Me.lblTelonFolder.Text = "Telon (0):"
        '
        'lblScreensFolder
        '
        Me.lblScreensFolder.AutoSize = True
        Me.lblScreensFolder.Location = New System.Drawing.Point(345, 350)
        Me.lblScreensFolder.Name = "lblScreensFolder"
        Me.lblScreensFolder.Size = New System.Drawing.Size(95, 20)
        Me.lblScreensFolder.TabIndex = 66
        Me.lblScreensFolder.Text = "Screens (0):"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label6.Location = New System.Drawing.Point(24, 385)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(158, 20)
        Me.Label6.TabIndex = 67
        Me.Label6.Text = "OUTPUT FOLDERS:"
        '
        'lblOutputFolder
        '
        Me.lblOutputFolder.AutoSize = True
        Me.lblOutputFolder.Location = New System.Drawing.Point(38, 414)
        Me.lblOutputFolder.Name = "lblOutputFolder"
        Me.lblOutputFolder.Size = New System.Drawing.Size(62, 20)
        Me.lblOutputFolder.TabIndex = 68
        Me.lblOutputFolder.Text = "Output:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(349, 414)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(57, 20)
        Me.Label8.TabIndex = 69
        Me.Label8.Text = "PUML:"
        '
        'txtPUMLFolder
        '
        Me.txtPUMLFolder.Location = New System.Drawing.Point(496, 411)
        Me.txtPUMLFolder.Name = "txtPUMLFolder"
        Me.txtPUMLFolder.Size = New System.Drawing.Size(137, 26)
        Me.txtPUMLFolder.TabIndex = 70
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(654, 414)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(85, 20)
        Me.Label9.TabIndex = 71
        Me.Label9.Text = "Expanded:"
        '
        'txtExpandedFolder
        '
        Me.txtExpandedFolder.Location = New System.Drawing.Point(775, 408)
        Me.txtExpandedFolder.Name = "txtExpandedFolder"
        Me.txtExpandedFolder.Size = New System.Drawing.Size(114, 26)
        Me.txtExpandedFolder.TabIndex = 72
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label10.Location = New System.Drawing.Point(14, 447)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(122, 20)
        Me.Label10.TabIndex = 73
        Me.Label10.Text = "RUN OPTIONS:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(901, 411)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(123, 20)
        Me.Label5.TabIndex = 74
        Me.Label5.Text = "Business Rules:"
        '
        'txtBusinessRulesFolder
        '
        Me.txtBusinessRulesFolder.Location = New System.Drawing.Point(1030, 408)
        Me.txtBusinessRulesFolder.Name = "txtBusinessRulesFolder"
        Me.txtBusinessRulesFolder.Size = New System.Drawing.Size(185, 26)
        Me.txtBusinessRulesFolder.TabIndex = 75
        '
        'cbInstream
        '
        Me.cbInstream.AutoSize = True
        Me.cbInstream.Checked = True
        Me.cbInstream.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbInstream.Location = New System.Drawing.Point(449, 509)
        Me.cbInstream.Name = "cbInstream"
        Me.cbInstream.Size = New System.Drawing.Size(98, 24)
        Me.cbInstream.TabIndex = 76
        Me.cbInstream.Text = "Instream"
        Me.cbInstream.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1279, 683)
        Me.Controls.Add(Me.cbInstream)
        Me.Controls.Add(Me.txtBusinessRulesFolder)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.txtExpandedFolder)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txtPUMLFolder)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.lblOutputFolder)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lblScreensFolder)
        Me.Controls.Add(Me.lblTelonFolder)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtASMFolder)
        Me.Controls.Add(Me.lblASMFolder)
        Me.Controls.Add(Me.txtDECLGenFolder)
        Me.Controls.Add(Me.lblDECLGenFolder)
        Me.Controls.Add(Me.txtEasytrieveFolder)
        Me.Controls.Add(Me.lblEasytrieveFolder)
        Me.Controls.Add(Me.txtCopybookFolder)
        Me.Controls.Add(Me.lblCopybooksFolder)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblCOBOLFolder)
        Me.Controls.Add(Me.lblPROCFolder)
        Me.Controls.Add(Me.lblJOBFolder)
        Me.Controls.Add(Me.txtAppFolder)
        Me.Controls.Add(Me.btnAppFolder)
        Me.Controls.Add(Me.txtProcFolder)
        Me.Controls.Add(Me.lblInitDirectory)
        Me.Controls.Add(Me.btnSandbox)
        Me.Controls.Add(Me.cbDataCom)
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
        Me.Controls.Add(Me.txtScreenMapsFolder)
        Me.Controls.Add(Me.txtTelonFolder)
        Me.Controls.Add(Me.txtDataGatheringForm)
        Me.Controls.Add(Me.btnDataGatheringForm)
        Me.Controls.Add(Me.cbScanModeOnly)
        Me.Controls.Add(Me.lblProcessingWorksheet)
        Me.Controls.Add(Me.lblProcessingSource)
        Me.Controls.Add(Me.lblProcessingJob)
        Me.Controls.Add(Me.cbLogStmt)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.lblStatusMessage)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.txtDelimiter)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnADDILite)
        Me.Controls.Add(Me.txtOutputFolder)
        Me.Controls.Add(Me.txtCobolFolder)
        Me.Controls.Add(Me.txtJCLJOBFolder)
        Me.Name = "Form1"
        Me.Text = "ADDILite"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtJCLJOBFolder As TextBox
    Friend WithEvents txtCobolFolder As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents txtOutputFolder As TextBox
    Friend WithEvents btnADDILite As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtDelimiter As TextBox
    Friend WithEvents btnClose As Button
    Friend WithEvents lblStatusMessage As Label
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents cbLogStmt As CheckBox
    Friend WithEvents lblProcessingJob As Label
    Friend WithEvents lblProcessingSource As Label
    Friend WithEvents lblProcessingWorksheet As Label
    Friend WithEvents cbScanModeOnly As CheckBox
    Friend WithEvents btnDataGatheringForm As Button
    Friend WithEvents txtDataGatheringForm As TextBox
    Friend WithEvents txtTelonFolder As TextBox
    Friend WithEvents txtScreenMapsFolder As TextBox
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
    Friend WithEvents cbDataCom As CheckBox
    Friend WithEvents btnSandbox As Button
    Friend WithEvents lblInitDirectory As Label
    Friend WithEvents txtProcFolder As TextBox
    Friend WithEvents btnAppFolder As Button
    Friend WithEvents txtAppFolder As TextBox
    Friend WithEvents lblJOBFolder As Label
    Friend WithEvents lblPROCFolder As Label
    Friend WithEvents lblCOBOLFolder As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents lblCopybooksFolder As Label
    Friend WithEvents txtCopybookFolder As TextBox
    Friend WithEvents lblEasytrieveFolder As Label
    Friend WithEvents txtEasytrieveFolder As TextBox
    Friend WithEvents lblDECLGenFolder As Label
    Friend WithEvents txtDECLGenFolder As TextBox
    Friend WithEvents lblASMFolder As Label
    Friend WithEvents txtASMFolder As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents lblTelonFolder As Label
    Friend WithEvents lblScreensFolder As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents lblOutputFolder As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents txtPUMLFolder As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents txtExpandedFolder As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents txtBusinessRulesFolder As TextBox
    Friend WithEvents cbInstream As CheckBox
End Class
