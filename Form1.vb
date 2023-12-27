﻿Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic.Logging
'Imports System.Reflection.Emit
'Imports System.Runtime.Remoting.Metadata.W3cXsd2001
'Imports System.Security.Cryptography

Public Class Form1
  ' ADDILite will read an IBM JCL syntax file and break apart its
  '  parts and pieces. Those parts are JOB, Executables, and Datasets.
  ' This will also create a Panvalet compatible extraction commands
  '  to pull out the Executables (programs) from their Library into
  '  our own library.
  ' This will also search for COBOL Sources to create the Data Details.
  ' '
  ' Latest change is to create a Plantuml compatible file to create flowcharts.
  ' Inputs:
  ' - JCL source text (*.jcl)
  ' - Proclib for PROCS
  ' - Source folder for COBOL
  ' - Source Includes for COBOL copybooks
  ' Outputs:
  ' - xlxs with all the Details.
  ' - PlantUml for creating flowchart
  '
  '***Be sure to change ProgramVersion when making changes!!!
  Dim ProgramVersion As String = "v1.6"
  'Change-History.
  ' 2023/12/05 v1.6 hk Handle IMS programs in the PROGRAMS tab.
  ' 2023/11/13 v1.5 hk While looking for comments ensure a valid COBOL division statement.
  '                   - Support all paragraphs in Identification division as Comments (except program-id)
  '                   - change 'unknown jcl control' message to put details in 3rd column
  '                   - Add start and end positions on the Fields tab and support Redefines
  ' 2023/10/30 v1.4 hk Provide for missing JOB Step Name in JCL
  '                   -Add to list of utilities
  '                   -Set SourceType to have something ie. NotFound or Unknown but not empty
  '                   -Add ORDER BY to stop scan for FROM Table-name(s) for ExecSQLs
  ' 2023/10/28 v1.3 hk Add Datagathering Form on first tab of spreadsheet
  '                   -fixed Easytrieve comments length
  ' 2023/10/25 v1.2 hk Added EXECSQL feature
  '                   -fix GetOpenMode and GetOpenModeSQL list
  ' 2023/09/25 v1.1 hk New program
  '
  Const JOBCARD As String = " JOB "
  Const PROCCARD As String = " PROC "
  Const PENDCARD As String = " PEND "
  Const EXECCARD As String = " EXEC "
  Const EXECCARDNOLABEL As String = "EXEC "
  Const DDCARD As String = " DD "
  Const DDCARDNOLABEL As String = "DD "
  Const SETCARD As String = " SET "
  Const SETCARDNOLABEL As String = "SET "
  Const OUTPUTCARD As String = " OUTPUT "
  Const IFCARD As String = "IF "
  Const ENDIFCARD As String = "ENDIF"
  Const QUOTE As Char = Chr(34)       'double-quote
  Const ESCAPENEWLINE As String = "\n"


  ' Arrays to hold the DB2 Declare to Member names
  ' these two array will share the same index
  Dim DB2Declares As New List(Of String)
  Dim MembersNames As New List(Of String)

  Dim ListOfDataGathering As New List(Of String)
  Dim NumberOfJobsToProcess As Integer = 0
  Dim ListOfJobs As New List(Of String)

  ' JCL
  Dim DirectoryName As String = ""
  Dim FileNameOnly As String = ""
  Dim tempFileName As String = ""
  Dim tempCobFileName As String = ""
  Dim tempEZTFileName As String = ""
  Dim Delimiter As String = ""
  Dim jControl As String = ""
  Dim jLabel As String = ""
  Dim jParameters As String = ""
  Dim procName As String = ""
  Dim jobName As String = ""
  Dim jobClass As String = ""
  Dim msgClass As String = ""
  Dim prevPgmName As String = ""
  Dim prevStepName As String = ""
  Dim prevDDName As String = ""
  Dim execName As String = ""
  Dim pgmName As String = ""
  Dim DDName As String = ""
  Dim stepName As String = ""
  Dim InstreamProc As String = ""

  Dim ddConcatSeq As Integer = 0
  Dim ddSequence As Integer = 0
  Dim jobSequence As Integer = 0
  Dim procSequence As Integer = 0
  Dim execSequence As Integer = 0

  Dim SummaryRow As Integer = 0
  Dim ProgramsRow As Integer = 0
  Dim RecordsRow As Integer = 0
  Dim FieldsRow As Integer = 0
  Dim CommentsRow As Integer = 0
  Dim EXECSQLRow As Integer = 0

  Dim jclStmt As New List(Of String)
  Dim ListOfExecs As New List(Of String)        'array holding the executable programs

  Dim swIPFile As StreamWriter = Nothing        'Instream proc file, temporary
  Dim swDDFile As StreamWriter = Nothing
  Dim swPumlFile As StreamWriter = Nothing
  Dim LogFile As StreamWriter = Nothing
  Dim LogStmtFile As StreamWriter = Nothing
  Dim swBRFile As StreamWriter = Nothing
  Dim swCallPgmsFile As StreamWriter = Nothing

  ' load the Excel References
  Dim objExcel As New Microsoft.Office.Interop.Excel.Application
  ' Data Gathering Form
  Dim dgfWorkbook As Microsoft.Office.Interop.Excel.Workbook
  Dim dgfWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  ' Model 
  Dim workbook As Microsoft.Office.Interop.Excel.Workbook
  Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim SummaryWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim RecordsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim FieldsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim CommentsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim EXECSQLWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim rngSummaryName As Microsoft.Office.Interop.Excel.Range
  Dim rngRecordName As Microsoft.Office.Interop.Excel.Range
  Dim rngRecordsName As Microsoft.Office.Interop.Excel.Range
  Dim rngFieldsName As Microsoft.Office.Interop.Excel.Range
  Dim rngComments As Microsoft.Office.Interop.Excel.Range
  Dim rngEXECSQL As Microsoft.Office.Interop.Excel.Range
  Dim DefaultFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault

  ' COBOL fields
  Dim SourceType As String = ""
  Dim CalledMember As String = ""
  Dim SrcStmt As New List(Of String)
  Dim cWord As New List(Of String)
  Dim lWord As New List(Of String)                    'Word Level value for IF syncs with cWord
  Dim ListOfFiles As New List(Of String)              'array to hold File & DB2 Table names
  Dim ListOfRecordNames As New List(Of String)          'array to hold read/written records
  Dim ListOfRecords As New List(Of String)              'array to hold read/written records
  Dim ListOfFields As New List(Of String)             'array to hold fields for each record
  Dim ListOfReadIntoRecords As New List(Of String)    'array to hold Read Into Records
  Dim ListOfWriteFromRecords As New List(Of String)   'array to hold Write from records
  Dim ListOfComments As New List(Of String)           'array to hold comments from source (cobol & easytrieve)
  Dim ListOfParagraphs As New List(Of String)         'array to hold COBOL paragraph names
  Dim ListOfCallPgms As New List(Of String)           'array to hold Call programs (sub routines)
  Dim ListOfEXECSQL As New List(Of String)            'array to hold EXEC SQL statments (cobol & easytrieve)
  Dim IFLevelIndex As New List(Of Integer)            'where in cWord the 'IF' is located
  Dim VerbNames As New List(Of String)
  Dim VerbCount As New List(Of Integer)
  Dim ProgramAuthor As String = ""
  Dim ProgramWritten As String = ""
  Dim IndentLevel As Integer = -1                  'how deep the if has gone
  Dim BRLevel As Integer = -1                       'how deep the Business Rule has gone
  Dim FirstWhenStatement As Boolean = False
  Dim WithinReadStatement As Boolean = False
  Dim WithinReadConditionStatement As Boolean = False
  'Dim WithinPerformWithEndPerformStatement As Boolean = False
  Dim WithinPerformCnt As Integer = 0
  Dim WithinIF As Boolean = False
  Dim pgmSeq As Integer = 0
  Dim pumlFile As StreamWriter = Nothing          'File holding the Plantuml commands


  Public Structure ProgramInfo
    Public ProgramId As String
    Public IdentificationDivision As Integer
    Public EnvironmentDivision As Integer
    Public DataDivision As Integer
    Public ProcedureDivision As Integer
    Public EndProgram As Integer
    Public Sub New(ByVal _ProgramId As String,
                   ByVal _IdentificationDivision As Integer,
                   ByVal _EnvironmentDivision As Integer,
                   ByVal _DataDivsision As Integer,
                   ByVal _ProcedureDivision As Integer,
                   ByVal _EndProgram As Integer)
      ProgramId = _ProgramId
      IdentificationDivision = _IdentificationDivision
      EnvironmentDivision = _EnvironmentDivision
      DataDivision = _DataDivsision
      ProcedureDivision = _ProcedureDivision
      EndProgram = _EndProgram
    End Sub
  End Structure
  Dim listOfPrograms As New List(Of ProgramInfo)
  Dim pgm As ProgramInfo = Nothing


  Public Structure CalledProgramInfo
    Public Name As String
    Public Count As Integer
    Public CalledFrom As String
    Public Sub New(ByVal _name As String,
                   ByVal _count As Short,
                   ByVal _calledFrom As String)
      Name = _name
      Count = _count
      CalledFrom = _calledFrom
    End Sub
  End Structure
  Dim list_CalledPrograms As New List(Of CalledProgramInfo)
  Dim List_Fields As New List(Of fieldInfo)
  Dim List_Usage As New List(Of String)({"BINARY", "COMP", "COMP-1", "COMP-2", "COMP-3", "COMP-4", "COMP-5", "COMPUTATIONAL", "COMPUTATIONAL-1", "COMPUTATIONAL-2", "COMPUTATIONAL-3", "COMPUTATIONAL-4", "COMPUTATIONAL-5", "DISPLAY", "DISPLAY-1", "INDEX", "NATIONAL", "PACKED-DECIMAL", "POINTER", "PROCEDURE-POINTER", "FUNCTION-POINTER"})

  Private Sub btnDataGatheringForm_Click(sender As Object, e As EventArgs) Handles btnDataGatheringForm.Click
    ' Open file dialog
    Dim ofd_DataGatheringForm As New OpenFileDialog With {
      .InitialDirectory = "C:\Users\906074897\Documents\All Projects\State of Illinois\Sandbox",
      .Filter = "Spreadsheet|*.xlsx",
      .Title = "Open the Data Gathering Form"
    }
    If ofd_DataGatheringForm.ShowDialog() = DialogResult.OK Then
      txtDataGatheringForm.Text = ofd_DataGatheringForm.FileName
    Else
      Exit Sub
    End If
  End Sub
  Private Sub btnJCLJOBFilename_Click(sender As Object, e As EventArgs) Handles btnJCLJOBFilename.Click
    ' grab the dgf's directory
    Dim myFileInfo As System.IO.FileInfo
    myFileInfo = My.Computer.FileSystem.GetFileInfo(txtDataGatheringForm.Text)
    Dim folderPath As String = myFileInfo.DirectoryName

    ' browse for and select folder name
    Dim bfd_JobLibFolder As New FolderBrowserDialog With {
      .Description = "Enter JCL JOBs folder name",
      .SelectedPath = folderPath
    }
    If bfd_JobLibFolder.ShowDialog() = DialogResult.OK Then
      txtJCLJOBFolderName.Text = bfd_JobLibFolder.SelectedPath
    Else
      Exit Sub
    End If

    NumberOfJobsToProcess = My.Computer.FileSystem.GetFiles(txtJCLJOBFolderName.Text).Count
    lblJobFileCount.Text = "JCL Job files found:" & Str(NumberOfJobsToProcess)
    For Each foundFile As String In My.Computer.FileSystem.GetFiles(txtJCLJOBFolderName.Text)
      ListOfJobs.Add(foundFile)
    Next
  End Sub

  'Private Sub btnProcLibFolder_Click(sender As Object, e As EventArgs) Handles btnProcLibFolder.Click
  '  ' browse for and select folder name
  '  Dim bfd_ProcLibFolder As New FolderBrowserDialog With {
  '    .Description = "Enter ProcLib folder name",
  '    .SelectedPath = txtJCLJOBFolderName.Text
  '  }
  '  If bfd_ProcLibFolder.ShowDialog() = DialogResult.OK Then
  '    txtProcLibFolderName.Text = bfd_ProcLibFolder.SelectedPath
  '  End If
  'End Sub

  Private Sub btnSourceFolder_Click(sender As Object, e As EventArgs) Handles btnSourceFolder.Click
    ' browse for and select folder name
    Dim bfd_SourceFolder As New FolderBrowserDialog With {
      .Description = "Enter Source folder name",
      .SelectedPath = txtJCLJOBFolderName.Text
    }
    If bfd_SourceFolder.ShowDialog() = DialogResult.OK Then
      txtSourceFolderName.Text = bfd_SourceFolder.SelectedPath
    End If
  End Sub

  Private Sub btnOutputFolder_Click(sender As Object, e As EventArgs) Handles btnOutputFolder.Click
    ' browse for and select folder name
    Dim bfd_OutputFolder As New FolderBrowserDialog With {
      .Description = "Enter OUTPUT folder name",
      .SelectedPath = txtJCLJOBFolderName.Text
    }
    If bfd_OutputFolder.ShowDialog() = DialogResult.OK Then
      txtOutputFoldername.Text = bfd_OutputFolder.SelectedPath
      DirectoryName = txtOutputFoldername.Text
      tempFileName = DirectoryName & "\" & FileNameOnly & "_expandedJCL.txt"
      tempCobFileName = DirectoryName & "\" & FileNameOnly & "_expandedCOB.txt"
      tempEZTFileName = DirectoryName & "\" & FileNameOnly & "_expandedEZT.txt"
    End If
  End Sub
  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub

  Private Sub txtDelimiter_TextChanged(sender As Object, e As EventArgs) Handles txtDelimiter.TextChanged
    Delimiter = txtDelimiter.Text
  End Sub

  Private Sub btnADDILite_Click(sender As Object, e As EventArgs) Handles btnADDILite.Click
    DirectoryName = Path.GetDirectoryName(txtJCLJOBFolderName.Text)

    Delimiter = txtDelimiter.Text
    lblCopybookMessage.Text = ""

    ' ready the progress bar
    ProgressBar1.Minimum = 0
    ProgressBar1.Maximum = NumberOfJobsToProcess + 2
    ProgressBar1.Step = 1
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True

    Me.Cursor = Cursors.WaitCursor

    Dim logFileName As String = txtOutputFoldername.Text & "\ADDILite_log.txt"
    LogFile = My.Computer.FileSystem.OpenTextFileWriter(logFileName, False)
    LogFile.WriteLine(Date.Now & ",Program Starts," & Me.Text)
    LogFile.WriteLine(Date.Now & ",Data Gathering Form," & txtDataGatheringForm.Text)
    LogFile.WriteLine(Date.Now & ",JOB Folder," & txtJCLJOBFolderName.Text)
    '**LogFile.WriteLine(Date.Now & ",JCL Proclib Folder," & txtJCLProclibFoldername.Text)
    LogFile.WriteLine(Date.Now & ",Source Folder," & txtSourceFolderName.Text)
    LogFile.WriteLine(Date.Now & ",Output Folder," & txtOutputFoldername.Text)
    LogFile.WriteLine(Date.Now & ",Delimiter," & txtDelimiter.Text)
    LogFile.WriteLine(Date.Now & ",ScanModeOnly," & cbScanModeOnly.Checked)

    'validations
    If Not FileNamesAreValid() Then
      LogFile.WriteLine(Date.Now & ",File Names are not Valid,")
      Me.Cursor = Cursors.Default
      Exit Sub
    End If

    objExcel.Visible = False

    ' Load the Data Gathering Form spreadsheet into the ListofDataGatheringForm array
    dgfWorkbook = objExcel.Workbooks.Open(txtDataGatheringForm.Text, True)
    SummaryRow = 1
    dgfWorksheet = dgfWorkbook.Sheets.Item(1)
    dgfWorksheet.Select(1)
    For SummaryRow = 1 To 50
      Dim row As Integer = LTrim(Str(SummaryRow))
      If Val(dgfWorksheet.Cells.Range("A" & row).Value2) > 0 Then
        ListOfDataGathering.Add(dgfWorksheet.Cells.Range("B" & row).Value2 &
                              Delimiter &
                              dgfWorksheet.Cells.Range("C" & row).Value)
      End If
    Next
    SummaryRow = 0
    dgfWorkbook.Close()


    ' Prepare for CallPgms file which holds all the Called Programs within the sources
    '  this file is processed as the last "JOB"
    ' Remove previous CallPgms.jcl file
    Dim CallPgmsFileName = txtJCLJOBFolderName.Text & "\CallPgms.jcl"
    If File.Exists(CallPgmsFileName) Then
      LogFile.WriteLine(Date.Now & ",Previous CallPgms.jcl file deleted," & CallPgmsFileName)
      Try
        File.Delete(CallPgmsFileName)
      Catch ex As Exception
        LogFile.WriteLine(Date.Now & ",Error deleting CallPgms.jcl file," & ex.Message)
        lblCopybookMessage.Text = "Error deleting CallPgms.jcl file:" & ex.Message
        Exit Sub
      End Try
    End If

    'build a cross-reference table of DB2 Tablenames with source members
    For Each foundFile As String In My.Computer.FileSystem.GetFiles(txtSourceFolderName.Text)
      Dim memberLines As String() = File.ReadAllLines(foundFile)
      For index = 0 To memberLines.Count - 1
        If memberLines(index).IndexOf(" EXEC SQL DECLARE ") > -1 Then
          Dim srcWords As New List(Of String)
          Call GetSourceWords(memberLines(index), srcWords)
          MembersNames.Add(Path.GetFileName(foundFile))
          'need to remove any schema name (before the period)
          Dim tableParameters As String() = srcWords(3).Split(".")
          If tableParameters.Count = 2 Then
            DB2Declares.Add(tableParameters(1))
          Else
            DB2Declares.Add(tableParameters(0))
          End If
          Exit For
        End If
      Next index
    Next

    ProgressBar1.PerformStep()
    ProgressBar1.Show()



    Dim ProgramsFileName = txtOutputFoldername.Text & "\ADDILite.xlsx"
    If File.Exists(ProgramsFileName) Then
      LogFile.WriteLine(Date.Now & ",Previous Model file deleted," & ProgramsFileName)
      Try
        File.Delete(ProgramsFileName)
      Catch ex As Exception
        LogFile.WriteLine(Date.Now & ",Error deleting Model," & ex.Message)
        lblCopybookMessage.Text = "Error Deleting Model:" & ex.Message
        ProgressBar1.Visible = False
        Exit Sub
      End Try
    End If

    ' Create the Summary tab (aka Datagathering form details)
    CreateSummary()


    ' Process All the jobs in the JCL Folder.
    '  An addtional job could be created if should there be call subroutines
    Dim Jobcount As Integer = 0
    For Each JobFile In ListOfJobs
      Jobcount += 1
      jobSequence += 1
      lblProcessingJob.Text = "Processing Job #" & Jobcount & ": " & JobFile
      LogFile.WriteLine(Date.Now & ",Processing Job," & Path.GetFileNameWithoutExtension(JobFile))
      FileNameOnly = Path.GetFileNameWithoutExtension(JobFile)
      ProcessJOBFile(JobFile)
      ProcessSourceFiles()
      ProgressBar1.PerformStep()
      ProgressBar1.Show()
      Call InitializeProgramVariables()
    Next

    ProcessCallPgms(CallPgmsFileName)

    CreateEXECSQLWorksheet()


    ' Save Application Model Spreadsheet
    If cbScanModeOnly.Checked Then
      objExcel.DisplayAlerts = False
      objExcel.Quit()
    Else
      ' Format, Save and close Excel
      lblCopybookMessage.Text = "Saving Spreadsheet"
      Call FormatWorksheets()
      workbook.SaveAs(ProgramsFileName, DefaultFormat)
      workbook.Close()
      objExcel.Quit()
    End If

    ProgressBar1.PerformStep()
    ProgressBar1.Show()

    LogFile.WriteLine(Date.Now & ",Program Ends,")
    LogFile.Close()
    Me.Cursor = Cursors.Default
    lblCopybookMessage.Text = "Process Complete"
    MessageBox.Show("Process Complete")
  End Sub

  Sub ProcessCallPgms(ByRef CallPgmsFileName As String)
    If ListOfCallPgms.Count = 0 Then
      Exit Sub
    End If

    ' create the CallPgms.jcl file
    swCallPgmsFile = New StreamWriter(CallPgmsFileName, False)
    Dim pgmCnt As Integer = 0
    swCallPgmsFile.WriteLine("//CALLPGMS JOB 'SUBROUTINES CALLED'")
    For Each callpgm In ListOfCallPgms
      pgmCnt += 1
      Dim execs As String() = callpgm.Split(Delimiter)
      swCallPgmsFile.WriteLine("//PGM" & LTrim(Str(pgmCnt)) & " EXEC PGM=" & execs(0).Replace(Delimiter, ""))
      swCallPgmsFile.WriteLine("//STEPLIB DD DSN=" & execs(2) & ",DISP=SHR")
    Next
    swCallPgmsFile.Close()

    'process the CallPgms file
    If File.Exists(CallPgmsFileName) Then
      lblProcessingJob.Text = "Processing Job CallPgms" & ": " & CallPgmsFileName
      LogFile.WriteLine(Date.Now & ",Processing Job," & CallPgmsFileName)
      FileNameOnly = Path.GetFileNameWithoutExtension(CallPgmsFileName)
      ProcessJOBFile(CallPgmsFileName)
      ProcessSourceFiles()
      ProgressBar1.PerformStep()
      ProgressBar1.Show()
      Call InitializeProgramVariables()
    Else
      LogFile.WriteLine(Date.Now & ",Call Pgms File not found?," & CallPgmsFileName)
    End If

  End Sub
  Function FileNamesAreValid() As Boolean
    FileNamesAreValid = False
    Select Case True
      Case Not File.Exists(txtDataGatheringForm.Text)
        LogFile.WriteLine(Date.Now & ",Data Gathering File spreadsheet not found," & txtDataGatheringForm.Text)
      Case txtJCLJOBFolderName.TextLength = 0
        LogFile.WriteLine(Date.Now & ",JCL JOB Folder name required,")
      Case Not IsValidFileNameOrPath(txtJCLJOBFolderName.Text)
        LogFile.WriteLine(Date.Now & ",JCL JOB Folder name has invalid characters,")

      Case txtOutputFoldername.TextLength = 0
        LogFile.WriteLine(Date.Now & ",OutFolder name required,")
      Case Not IsValidFileNameOrPath(txtOutputFoldername.Text)
        LogFile.WriteLine(Date.Now & ",OutFolder has invalid characters,")
      Case Not IsValidFileNameOrPath(txtSourceFolderName.Text)
        LogFile.WriteLine(Date.Now & ",Source folder name has invalid characters,")
        '      Case Not IsValidFileNameOrPath(txtJCLProclibFoldername.Text)
        '        LogFile.WriteLine(Date.Now & ",Proclib folder name has invalid characters,")
      Case Else
        FileNamesAreValid = True
    End Select
  End Function
  Function IsValidFileNameOrPath(ByVal name As String) As Boolean
    If name Is Nothing Then
      Return False
    End If

    For Each badChar As Char In System.IO.Path.GetInvalidPathChars
      If InStr(name, badChar) > 0 Then
        Return False
      End If
    Next

    Return True
  End Function

  ' * Subroutines
  Sub ProcessJOBFile(JobFile As String)

    'Load the Jobfile to the jclStmt List
    If Not File.Exists(JobFile) Then
      LogFile.WriteLine(Date.Now & ",Job file Not found!?," & JobFile)
      Exit Sub
    End If
    Dim jclRecordsCount As Integer = LoadJCLStatementsToArray(JobFile)
    If jclRecordsCount = 0 Then
      MessageBox.Show("No JCL records read from file:" & JobFile)
      Exit Sub
    End If

    If jclStmt.Count = 0 Then
      MessageBox.Show("No JCL statements loaded from File:" & JobFile)
    End If

    ' log file, optioned
    If cbLogStmt.Checked Then
      Call LogStmtArray(FileNameOnly, jclStmt)
    End If

    If WriteOutput() = -1 Then
      MessageBox.Show("Error while building output. See log file")
    End If

  End Sub
  Function LoadJCLStatementsToArray(JobFile As String) As Integer '
    ' Clear out the JCL Statement array
    jclStmt.Clear()

    ' Remove the temporary work file
    Try
      If My.Computer.FileSystem.FileExists(tempFileName) Then
        My.Computer.FileSystem.DeleteFile(tempFileName)
      End If
    Catch ex As Exception
      MessageBox.Show("Removal of Temp file error:" & ex.Message)
      LoadJCLStatementsToArray = -1
      Exit Function
    End Try

    Dim text1 As String
    Dim jStatement As String = ""
    Dim statement As String = ""
    Dim execIndex As Integer = 0
    Dim commaIndex As Integer = 0
    Dim continuation As Boolean = False
    LoadJCLStatementsToArray = 0
    Dim debugCount As Integer = -1

    ' load the job file to an array
    Dim JobLines As String() = File.ReadAllLines(JobFile)
    Dim JCLLines As List(Of String) = Nothing

    '' Include any called in PROC (exec proc=<proc> or exec <proc>) 
    ''   load from the Proclib folder
    'For Each JobLine In JobLines
    '  If txtProcLibFolderName.TextLength = 0 Then
    '    JCLLines.Add(JobLine)
    '    Continue For
    '  End If
    '  execIndex = JobLine.IndexOf(" EXEC ")
    '  If execIndex = -1 Then
    '    JCLLines.Add(JobLine)
    '    Continue For
    '  End If
    '  If JobLine.IndexOf(" PGM=") > -1 Then
    '    JCLLines.Add(JobLine)
    '    Continue For
    '  End If
    '  ' must reference a PROC 
    '  procName = GetParm(JobLine, "PROC=").Trim
    '  If procName.Length = 0 Then
    '    execIndex += 6
    '    Dim parmValues As String() = JobLine.Substring(execIndex).Split(",")
    '    procName = parmValues(0).Trim
    '  End If
    '  ' with a procname, now look in the proclib folder
    '  If File.Exists(txtProcLibFolderName.Text & procName) Then

    '  End If
    'Next

    '
    '' Write any Instream PROC(s) to a Proclib and drop it and drop empty lines
    '' Not handling Procs within Procs
    ''
    'Dim NumberOfInstreamProcsFound As Integer = 0
    'Dim swTemp As StreamWriter = Nothing

    'Dim ipName As String = ""
    ''
    '' Create a Proc file from the Instream proc, if any
    '' And drop the instream proc lines 
    '' And remove any columns 73-80 values
    ''
    'swTemp = New StreamWriter(tempFileName, False)
    'For index As Integer = 0 To JCLLines.Count - 1
    '  Dim jline As String = JCLLines(index) & Space(72)
    '  jline = jline.Substring(0, 72).Trim
    '  If jline.Substring(0, 2) = "/*" Then
    '    swTemp.WriteLine(jline)
    '    Continue For
    '  End If
    '  If jline.Substring(0, 3) = "//*" Then
    '    swTemp.WriteLine(jline)
    '    Continue For
    '  End If
    '  Dim procLocation As Integer = jline.IndexOf(" PROC ")
    '  If procLocation = -1 Then
    '    swTemp.WriteLine(jline)
    '    Continue For
    '  End If
    '  ipName = Trim(jline.Substring(2, procLocation - 3 + 1))
    '  Dim ipFileName = txtJCLProclibFoldername.Text & "\" & ipName & ".jcl"
    '  LogFile.WriteLine(Date.Now & ",Instream Proc Created," & ipFileName)
    '  swIPFile = New StreamWriter(ipFileName, False)
    '  Dim PendFound As Boolean = False
    '  For ipIndex As Integer = index To JCLLines.Count - 1
    '    Dim pline As String = JCLLines(ipIndex) & Space(72)
    '    pline = pline.Substring(0, 72).Trim
    '    swIPFile.WriteLine(pline)
    '    If pline.IndexOf(" PEND ") > -1 Then
    '      index = ipIndex
    '      PendFound = True
    '      Exit For
    '    End If
    '    If pline.EndsWith(" PEND") Then
    '      index = ipIndex
    '      PendFound = True
    '      Exit For
    '    End If
    '  Next
    '  If Not PendFound Then
    '    swIPFile.WriteLine("// PEND")
    '    index = JCLLines.Count - 1
    '  End If
    '  swIPFile.Close()
    'Next
    'swTemp.Close()

    '' Now include all PROCs from the ProcLib and place after the EXEC PROC statement
    ''
    'Dim jclWords As New List(Of String)
    'JCLLines = File.ReadAllLines(tempFileName)
    'swTemp = New StreamWriter(tempFileName, False)
    'For index As Integer = 0 To JCLLines.Count - 1
    '  If JCLLines(index).Substring(0, 2) = "/*" Then
    '    swTemp.WriteLine(JCLLines(index))
    '    Continue For
    '  End If
    '  If JCLLines(index).Substring(0, 3) = "//*" Then
    '    swTemp.WriteLine(JCLLines(index))
    '    Continue For
    '  End If
    '  Call GetJCLWords(JCLLines(index), jclWords)
    '  If jclWords.Count >= 2 Then
    '    jControl = jclWords(1)
    '  End If
    '  If jControl <> "EXEC" Then
    '    swTemp.WriteLine(JCLLines(index))
    '    Continue For
    '  End If
    '  procName = GetParm(JCLLines(index), "PROC=")
    '  If procName.Length = 0 Then
    '    procName = GetParm(JCLLines(index), "PROC")
    '  End If
    '  If procName.Length = 0 Then
    '    If jclWords.Count >= 3 Then
    '      If Not jclWords(2).StartsWith("PGM=") Then
    '        procName = jclWords(2)
    '      End If
    '    End If
    '  End If
    '  ' write all lines for this EXEC statement
    '  For contIndex As Integer = index To JCLLines.Count - 1
    '    swTemp.WriteLine(JCLLines(contIndex))
    '    If JCLLines(contIndex).Substring(0, 3) = "//*" Then
    '      Continue For
    '    End If
    '    If Microsoft.VisualBasic.
    '        Right(Trim(JCLLines(contIndex).PadRight(80).Substring(0, 70)), 1) <> "," Then
    '      index = contIndex
    '      Exit For
    '    End If
    '  Next

    '  If procName.Length = 0 Then
    '    Continue For
    '  End If
    '  ' if it is an "EXEC PROC=" copy the PROC file into here
    '  ' if it is an "EXEC <name>" this is also a PROC File to be copied
    '  Dim ProcFileName As String = txtJCLProclibFoldername.Text & "\" & procName & ".jcl"
    '  LogFile.WriteLine(Date.Now & ",Processing PROC source," & ProcFileName)
    '  Dim procLines As String() = File.ReadAllLines(ProcFileName)
    '  For procIndex = 0 To procLines.Count - 1
    '    Dim jline As String = "++" & Mid(procLines(procIndex), 3) & Space(72)
    '    jline = jline.Substring(0, 72).Trim
    '    swTemp.WriteLine(jline)
    '  Next
    'Next
    'swTemp.Close()
    ''
    '' Load JCL lines to a JCL statements array. 
    '' Basically dealing with continuations and removing comments
    ''
    'JCLLines = File.ReadAllLines(tempFileName)



    ' Load JCL Lines to a JCL Statement Array
    '   Remove continuations by combining to 1 line
    '   and remove comments
    '   and get rid of the slashes
    '   all that should be stored is the Label, Control, and Parameters

    For index As Integer = 0 To JobLines.Count - 1
      LoadJCLStatementsToArray += 1
      debugCount += 1
      text1 = JobLines(index).Replace(vbTab, Space(1))
      ' drop comments
      If Mid(text1, 1, 3) = "//*" Or Mid(text1, 1, 3) = "++*" Then
        Continue For
      End If
      ' drop data (of an DD * statement) or not a JCL statement
      If Mid(text1, 1, 2) = "//" Or Mid(text1, 1, 2) = "++" Then
      Else
        Continue For
      End If
      ' drop JES commands
      If Mid(text1, 1, 14) = "//SEND OUTPUT " Then
        Continue For
      End If
      If Mid(text1, 1, 9) = "/*JOBPARM" Then
        Continue For
      End If
      ' remove '+' in column 72 (which used to mean continuation?)
      If Len(text1) >= 72 Then
        If Mid(text1, 72, 1) = "+" Then
          Mid(text1, 72, 1) = " "
        End If
      End If
      ' format only the good stuff out of the line (no slashes, no comments)
      text1 = Trim(Microsoft.VisualBasic.Left(Mid(text1, 3) + Space(70), 70))
      ' determine if there will be a continuation
      text1 &= Space(1)
      continuation = JCLContinued(text1)
      ' Build the JCL statement
      jStatement &= text1
      ' if NOT continuing building of the JCL statement then add it to the List
      If continuation = False Then
        If jStatement.Trim.Length > 0 Then
          jclStmt.Add(jStatement)
        End If
        jStatement = ""
      End If
    Next


  End Function
  Function JCLContinued(ByRef text As String) As Boolean
    ' determine if there will be a continuation by looking for a comma + space on the line not within quotes
    Dim withinquote As Boolean = False
    JCLContinued = False
    For x = 1 To text.Length - 1
      If x > 72 Then
        Exit Function
      End If
      If Mid(text, x, 1) = "'" Then
        withinquote = Not withinquote
        Continue For
      End If
      If withinquote Then
        Continue For
      End If
      If Mid(text, x, 2) = ", " Then
        text = Mid(text, 1, x)                        'remove anything to the right of continuation comma+space
        JCLContinued = True
        Exit Function
      End If
    Next
  End Function
  Sub LogStmtArray(ByRef theFileName As String, theStmtArray As List(Of String))
    ' write the stmt array to a file for debugging purposes
    Dim logStmtCount As Integer = -1
    Dim logStmtFileName As String = txtOutputFoldername.Text & "\Debug\" & theFileName & "_logStmt.txt"
    Try
      LogStmtFile = My.Computer.FileSystem.OpenTextFileWriter(logStmtFileName, False)
    Catch ex As Exception
      LogFile.WriteLine(Date.Now & ",Error creating Log stmt file," & ex.Message)
      Exit Sub
    End Try
    For Each statement In theStmtArray
      logStmtCount += 1
      LogStmtFile.WriteLine(LTrim(Str(logStmtCount)) & ":" & statement)
    Next
    LogStmtFile.Close()

  End Sub
  Sub GetJCLWords(ByVal statement As String, ByRef jclWords As List(Of String))
    ' takes the input string and breaks into words surrouned by blanks
    ' and drops extra blanks
    jclWords.Clear()
    'statement = " DISPLAY '*** CRCALCX REC READ        = ' WS-REC-READ.   "
    statement = statement.Trim
    Dim WithinQuotes As Boolean = False
    Dim word As String = ""
    Dim aByte As String = ""
    For x As Integer = 0 To statement.Length - 1
      aByte = statement.Substring(x, 1)
      If aByte = "'" Then
        WithinQuotes = Not WithinQuotes
      End If
      If aByte = " " Then
        If WithinQuotes Then
          word &= aByte
        Else
          If word.Trim.Length > 0 Then
            jclWords.Add(word.ToUpper)
            word = ""
          End If
        End If
      Else
        word &= aByte
      End If
    Next
    If word.EndsWith(".") Then
      word = word.Remove(word.Length - 1)
      jclWords.Add(word.ToUpper)
      word = ""
    End If
    If word.Length > 0 Then
      jclWords.Add(word)
    End If
  End Sub
  Function GetParm(ByRef SearchWithinThis As String, ByRef SearchForThis As String) As String
    GetParm = ""
    Dim FirstCharacter As Integer = SearchWithinThis.IndexOf(SearchForThis)
    If FirstCharacter = -1 Then
      Exit Function
    End If
    Dim ParenCount As Integer = 0
    Dim ParmValue As String = ""
    Dim ByteValue As String = ""
    Dim WithinQuote As Boolean = False
    ' determine the value of the Parm looking for ending "," or ")"
    For x As Integer = FirstCharacter + SearchForThis.Length To SearchWithinThis.Length - 1
      ByteValue = SearchWithinThis.Substring(x, 1)
      Select Case ByteValue
        Case "'"
          If WithinQuote Then
            WithinQuote = False
          Else
            WithinQuote = True
          End If
        Case "("
          ParenCount += 1
          ParmValue &= ByteValue
        Case ")"
          ParenCount -= 1
          If ParenCount = 0 Then
            ParmValue &= ByteValue
            Exit For
          End If
          ParmValue &= ByteValue
        Case ","
          If WithinQuote Then
            ParmValue &= ByteValue
            Continue For
          End If
          If ParenCount = 0 Then
            Exit For
          End If
          ParmValue &= ByteValue
        Case Else
          ParmValue &= ByteValue
      End Select
    Next
    GetParm = ParmValue
  End Function
  Function GetParmPGM(SearchWithinThis As String) As String
    'look for and return value of "PGM=value"
    GetParmPGM = ""
    'find "PGM="
    Dim SearchForThis As String = "PGM="
    Dim FirstCharacter As Integer = SearchWithinThis.IndexOf(SearchForThis)
    If FirstCharacter = -1 Then
      Exit Function
    End If
    'find the value of "PGM=" (find the comma if there is one)
    Dim SecondHalf As String = Mid(SearchWithinThis, FirstCharacter + Len(SearchForThis) + 1)
    Dim CommaCharacter As Integer = SecondHalf.IndexOf(",")
    Select Case CommaCharacter
      Case -1
        GetParmPGM = SecondHalf
      Case 0                  'means a comma only with no value, no way
        Exit Function
      Case Else
        GetParmPGM = Microsoft.VisualBasic.Left(SecondHalf, CommaCharacter)
    End Select
  End Function

  Function GetFirstParm(parameter As String) As String
    ' Extract the first parameter from the jParameters, could have a comma or EOL
    GetFirstParm = ""
    Dim commaLocation As Integer = parameter.IndexOf(",")
    Select Case commaLocation
      Case -1
        GetFirstParm = RTrim(parameter)
      Case Else
        GetFirstParm = Microsoft.VisualBasic.Left(parameter, commaLocation)
    End Select
  End Function

  Function WriteOutput() As Integer
    ' Write the output JOB, EXEC, DD, Panvalet script, and PUML files.
    ' return of -1 means an error
    ' return of 0 means all is okay

    Dim DDFileName = txtOutputFoldername.Text & "/" & FileNameOnly & "_DD.csv"
    Try
      swDDFile = My.Computer.FileSystem.OpenTextFileWriter(DDFileName, False)
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error opening " & DDFileName)
      WriteOutput = -1
      Exit Function
    End Try

    WriteOutput = 0

    ' Write the details to the files
    jobName = ""
    jobClass = ""
    msgClass = ""

    For Each statement As String In jclStmt
      Call GetLabelControlParms(statement, jLabel, jControl, jParameters)
      If Len(jControl) = 0 Then
        'MessageBox.Show("JCL control not found:" & statement)
        LogFile.WriteLine(Date.Now & ",JCL control not found,'" & statement & ": " & FileNameOnly & "'")
        'WriteOutput = -1
        Continue For
      End If

      Select Case jControl
        Case "JOB"
          Call ProcessJOB()
        Case "PROC"
          procName = jLabel
        Case "PEND"
        Case "EXEC"
          Call ProcessEXEC()
        Case "DD"
          Call ProcessDD()          'this writes the _dd.csv record
        Case "SET"
          Continue For
        Case "OUTPUT"
          Continue For
        Case "IF"
          Continue For
        Case "ENDIF"
          Continue For
        Case "JCLLIB"
          Continue For

        Case Else
          If jLabel = "IF" And jControl = "RC" Then
            Continue For
          End If
          If jParameters = "PEND" Then
            Continue For
          End If
          If jParameters = "PROC" Then
            Continue For
          End If
          LogFile.WriteLine(Date.Now & ",Unknown JCL Control Value," & statement.Replace(",", ";") & " file:" & FileNameOnly)
          Continue For
      End Select

    Next
    ' close the files

    swDDFile.Close()

    Call CreatePuml()

    Call CreatePrograms()

  End Function

  Sub GetLabelControlParms(statement As String,
                           ByRef jLabel As String,
                           ByRef jControl As String,
                           ByRef jParameters As String)
    'This will split out the three basic components of a JCL Statement
    'Each statement is expected to have either JOB, PROC, EXEC, DD, PEND, SET, OUTPUT, IF, ENDIF, etc.
    'Enforcement of JCL syntax is not done here except that there must be a space between
    ' Label, Control and Parmeters. There may not be a label, so adjustments are made
    Dim jLabelPrev As String = jLabel
    jControl = ""
    jParameters = ""

    Dim jclWords As New List(Of String)
    Call GetSourceWords(statement, jclWords)
    Select Case jclWords.Count
      Case 0

      Case 1
        jLabel = ""
        jControl = jclWords(0)
        jParameters = ""

      Case 2
        If jclWords(1) = "PROC" Then
          jLabel = jclWords(0)
          jControl = jclWords(1)
          jParameters = ""
          Exit Select
        End If
        If jclWords(0) = "DD" Then
          jLabel = jLabelPrev
        Else
          jLabel = ""
        End If
        jControl = jclWords(0)
        jParameters = jclWords(1)

      Case >= 3
        jLabel = jclWords(0)
        jControl = jclWords(1)
        jParameters = jclWords(2)

      Case Else
        MessageBox.Show("GetLabelControlParameters unknown jclWords count:" & statement)

    End Select

    'For sPos As Integer = 1 To Len(statement)
    '  Select Case True
    '    Case Mid(statement, sPos, Len(JOBCARD)) = JOBCARD
    '      jLabel = RTrim(Mid(statement, 1, sPos - 1))
    '      jControl = Mid(statement, sPos + 1, 3)
    '      jParameters = Trim(Mid(statement, sPos + 5))
    '      Exit For

    '    Case Mid(statement, sPos, Len(PROCCARD)) = PROCCARD
    '      jLabel = RTrim(Mid(statement, 1, sPos - 1))
    '      jControl = Mid(statement, sPos + 1, 4)
    '      jParameters = Trim(Mid(statement, sPos + 6))
    '      Exit For
    '    Case Mid(statement, sPos, 5) = " PROC"
    '      jLabel = RTrim(Mid(statement, 1, sPos - 1))
    '      jControl = "PROC"
    '      jParameters = Trim(Mid(statement, sPos + 5))
    '      Exit For

    '    Case statement = "PEND"
    '      jLabel = ""
    '      jControl = "PEND"
    '      jParameters = ""
    '      Exit For
    '    Case Mid(statement, sPos, Len(PENDCARD)) = PENDCARD
    '      jLabel = RTrim(Mid(statement, 1, sPos - 1))
    '      jControl = Mid(statement, sPos + 1, 4)
    '      jParameters = Trim(Mid(statement, sPos + 6))
    '      Exit For
    '    Case statement.EndsWith("PEND ")
    '      jLabel = ""
    '      jControl = "PEND"
    '      jParameters = Trim(Mid(statement, 6))
    '      Exit For
    '    Case statement.EndsWith(" PEND")
    '      jLabel = ""
    '      jControl = "PEND"
    '      jParameters = ""
    '      Exit For

    '    Case Mid(statement, sPos, Len(EXECCARD)) = EXECCARD
    '      jLabel = RTrim(Mid(statement, 1, sPos - 1))
    '      jControl = Mid(statement, sPos + 1, 4)
    '      jParameters = Trim(Mid(statement, sPos + 6))
    '      Exit For

    '    Case Mid(statement, sPos, Len(EXECCARDNOLABEL)) = EXECCARDNOLABEL
    '      jLabel = ""
    '      jControl = Mid(statement, 1, 4)
    '      jParameters = Trim(Mid(statement, 6))
    '      Exit For

    '    Case Mid(statement, sPos, Len(DDCARD)) = DDCARD
    '      jLabel = RTrim(Mid(statement, 1, sPos - 1))
    '      jControl = Mid(statement, sPos + 1, 2)
    '      jParameters = Trim(Mid(statement, sPos + 4))
    '      Exit For

    '    Case Mid(statement, sPos, Len(SETCARD)) = SETCARD
    '      jLabel = RTrim(Mid(statement, 1, sPos - 1))
    '      jControl = Mid(statement, sPos + 1, 3)
    '      jParameters = Trim(Mid(statement, sPos + 5))
    '      Exit For
    '    Case Mid(statement, 1, Len(DDCARDNOLABEL)) = DDCARDNOLABEL    'concat DD
    '      jLabel = jLabelPrev
    '      jControl = Mid(statement, 1, 2)
    '      jParameters = Trim(Mid(statement, 4))
    '      Exit For
    '    Case Mid(statement, 1, Len(SETCARDNOLABEL)) = SETCARDNOLABEL
    '      jLabel = ""
    '      jControl = Mid(statement, 1, 3)
    '      jParameters = Trim(Mid(statement, 5))
    '      Exit For
    '    Case Mid(statement, sPos, Len(OUTPUTCARD)) = OUTPUTCARD
    '      jLabel = RTrim(Mid(statement, 1, sPos - 1))
    '      jControl = Mid(statement, sPos + 1, 6)
    '      jParameters = Trim(Mid(statement, sPos + 8))
    '      Exit For
    '    Case Mid(statement, sPos, Len(IFCARD)) = IFCARD
    '      jLabel = ""
    '      jControl = Mid(statement, sPos, 2)
    '      jParameters = Trim(Mid(statement, sPos + 3))
    '      Exit For
    '    Case Mid(statement, sPos, 1) = "("
    '      jLabel = ""
    '      jControl = "IF"
    '      jParameters = Trim(Mid(statement, 1))
    '      Exit For
    '    Case Mid(statement, sPos, Len(ENDIFCARD)) = ENDIFCARD
    '      jLabel = ""
    '      jControl = Trim(Mid(statement, sPos, 5))
    '      jParameters = ""
    '      Exit For

    '  End Select
    'Next
  End Sub
  Sub ProcessJOB()
    'jobSequence += 1
    procSequence = 0
    execSequence = 0
    ddSequence = 0
    jobName = jLabel
    msgClass = GetParm(jParameters, "MSGCLASS=")
    jobClass = GetParm(jParameters, "CLASS=")
    'swJobFile.WriteLine(jLabel & txtDelimiter.Text &
    '    LTrim(Str(jobSequence)) & txtDelimiter.Text &
    '    jobClass & Delimiter &
    '    msgClass & Delimiter &
    '    RTrim(jParameters))
    InstreamProc = ""
  End Sub
  Sub ProcessEXEC()
    ' The "EXEC" control is for either PROC or a PGM
    ' For PROC it could be "EXEC <procname>" or "EXEC PROC=<procname"
    ' For PGM it is "EXEC PGM=<pgmname>"

    ' If by this time we haven't gotten a JOBname then set the JOB details
    If jobName.Length = 0 Then
      jobName = FileNameOnly
      jobClass = "?"
      msgClass = "?"
    End If
    '
    stepName = jLabel

    ddSequence = 0
    execName = ""
    pgmName = ""

    pgmName = Trim(GetParmPGM(jParameters))

    ' Is this a PROC statement
    If pgmName.Length = 0 Then
      procSequence += 1
      procName = GetParm(jParameters, "PROC=")
      If procName.Length = 0 Then
        procName = GetFirstParm(jParameters)
      End If
      pgmName = ""
      Exit Sub
    End If

    ' Is this an IMS program? If so, we need to get the real program name.
    If pgmName <> "DFSRRC00" Then
      execSequence += 1
      SourceType = GetSourceType(pgmName)
      execName = pgmName
      Exit Sub
    End If

    ' get the real program name from the IMS PARM phrase, 2nd value
    ' i.e., PARM='DLI,P2BPCSD1,P2BPCSD1'
    execName = pgmName
    Dim tempstr As String = GetParm(jParameters, "PARM=")
    If tempstr.Length = 0 Then
      pgmName = "IMS Unknown"
      SourceType = "Unknown"
      Exit Sub
    End If

    'tempstr = 'DLI,P2BPCSD1,P2BPCSD1'
    Dim temparray As String() = tempstr.Split(",")
    pgmName = temparray(1)
    SourceType = GetSourceType(pgmName)

  End Sub

  Sub ProcessDD()
    ' Process the DD statement

    Dim db2 As String = ""
    If jLabel = "CAFIN" Then
      db2 = "DB2"
    End If
    Dim dsn As String = GetParm(jParameters, "DSN=")
    Dim reportID As String = ""
    Dim sysout As String = GetParm(jParameters, "SYSOUT=")
    Select Case sysout.Length
      Case 0

      Case 1
        Select Case jLabel
          Case "SYSOUT", "SYSPRINT", "SYSUDUMP"
            If sysout = "*" Then
              sysout = "SYSOUT=" & msgClass
            Else
              sysout = "SYSOUT=" & sysout
            End If
        End Select
        If sysout = "*" Then
          sysout = "SYSOUT=" & msgClass
        End If
        If dsn.Length = 0 Then
          dsn = sysout
        End If

      Case Else
        Dim sysoutParms As String() = sysout.Replace("(", "").Replace(")", "").Split(",")
        sysout = "SYSOUT=" & sysoutParms(0)
        reportID = sysoutParms(1)
    End Select


    Dim unit As String = GetParm(jParameters, "UNIT=")
    If dsn.Length = 0 And unit.Length > 0 Then
      dsn = "WORKSPACE"
    End If

    ' figure out the file's dispositions for start,end,abend
    Dim disp As String = GetParm(jParameters, "DISP=")
    Dim fileDisp As String() = Nothing
    If disp.Length > 0 Then
      fileDisp = disp.Replace("(", "").Replace(")", "").Split(",")
    End If
    ' determine start disp
    Dim startDisp As String = DetermineStartDisp(fileDisp)
    Dim endDisp As String = DetermineEndDisp(fileDisp)
    If dsn = "WORKSPACE" Then
      endDisp = "DELETE"
    End If
    Dim abendDisp As String = DetermineAbendDisp(fileDisp)
    Dim ddName As String = jLabel
    If prevDDName = ddName And prevStepName = stepName And prevPgmName = pgmName Then
      ddConcatSeq += 1
    Else
      ddConcatSeq = 0
      ddSequence += 1
    End If

    ' Grab any DCB info from the JCL
    Dim dcbRecfm As String = ""
    Dim dcbLrecl As String = ""
    Dim dcbBlksize As String = ""
    Dim fileDCB As String() = Nothing
    Dim dcb As String = GetParm(jParameters, "DCB=")
    If dcb.Length > 0 Then
      fileDCB = dcb.Replace("(", "").Replace(")", "").Split(",")
      For dcbIndex = 0 To fileDCB.Count - 1
        Select Case True
          Case fileDCB(dcbIndex).IndexOf("RECFM=") > -1
            dcbRecfm = GetParm(fileDCB(dcbIndex), "RECFM=")
          Case fileDCB(dcbIndex).IndexOf("LRECL=") > -1
            dcbLrecl = GetParm(fileDCB(dcbIndex), "LRECL=")
          Case fileDCB(dcbIndex).IndexOf("BLKSIZE=") > -1
            dcbBlksize = GetParm(fileDCB(dcbIndex), "BLKSIZE=")
        End Select
      Next
    End If
    If dcb.Length = 0 Then
      dcbRecfm = GetParm(jParameters, "RECFM=")
    End If

    ' write the csv record 
    swDDFile.WriteLine(jobName & txtDelimiter.Text &
                       LTrim(Str(jobSequence)) & txtDelimiter.Text &
                       procName & txtDelimiter.Text &
                       LTrim(Str(procSequence)) & txtDelimiter.Text &
                       stepName & txtDelimiter.Text &
                       pgmName & txtDelimiter.Text &
                       LTrim(Str(execSequence)) & txtDelimiter.Text &
                       ddName & txtDelimiter.Text &
                       LTrim(Str(ddSequence)) & txtDelimiter.Text &
                       LTrim(Str(ddConcatSeq)) & Delimiter &
                       dsn & Delimiter &
                       startDisp & Delimiter &
                       endDisp & Delimiter &
                       abendDisp & Delimiter &
                       dcbRecfm & Delimiter &
                       dcbLrecl & Delimiter &
                       db2 & Delimiter &
                       reportID & Delimiter &
                       "" & Delimiter &
                       SourceType & Delimiter &
                       execName)
    prevDDName = ddName
    prevPgmName = pgmName
    prevStepName = stepName
  End Sub

  Sub CreatePuml()
    ' Open the output file PUML
    Dim PumlFileName = txtOutputFoldername.Text & "\" & FileNameOnly & ".puml"
    swPumlFile = My.Computer.FileSystem.OpenTextFileWriter(PumlFileName, False)

    ' Write the top of file
    swPumlFile.WriteLine("@startuml " & FileNameOnly)
    swPumlFile.WriteLine("header ADDILite(c), by IBM")
    swPumlFile.WriteLine("title Flowchart of JOB: " & FileNameOnly)

    ' Read the DD CSV file back in and load to array
    Dim FileName = txtOutputFoldername.Text & "/" & FileNameOnly & "_DD.csv"
    If Not File.Exists(FileName) Then
      Exit Sub
    End If
    Dim csvCnt As Integer = 0
    Dim csvFile As FileIO.TextFieldParser = New FileIO.TextFieldParser(FileName)
    Dim csvRecord As String()           ' all fields(columns) for a given record
    csvFile.TextFieldType = FileIO.FieldType.Delimited
    csvFile.Delimiters = New String() {"|"}
    csvFile.HasFieldsEnclosedInQuotes = True
    Dim ListOfSteps As New List(Of String)

    Do While Not csvFile.EndOfData
      csvRecord = csvFile.ReadFields
      csvCnt += 1
      jobName = csvRecord(0)
      jobSequence = Val(csvRecord(1))
      procName = csvRecord(2)
      procSequence = Val(csvRecord(3))
      stepName = csvRecord(4)
      pgmName = csvRecord(5)
      If pgmName.Length = 0 Then
        Continue Do
      End If
      execSequence = Val(csvRecord(6))
      Dim DDName As String = csvRecord(7)
      Dim DDSeq As String = csvRecord(8)
      ddConcatSeq = Val(csvRecord(9))
      Dim dsn As String = csvRecord(10)
      Dim dispStart As String = csvRecord(11)
      Dim dispEnd As String = csvRecord(12)
      Dim dispAbend As String = csvRecord(13)
      Dim dcbRecFM As String = csvRecord(14)
      Dim dcbLrecl As String = csvRecord(15)
      Dim db2 As String = csvRecord(16)
      Dim reportID As String = csvRecord(17)
      Dim reportDescription As String = csvRecord(18)
      SourceType = csvRecord(19)

      If stepName = "STEPLIB" And ddConcatSeq > 0 Then
        stepName = stepName & LTrim(Str(ddConcatSeq))
      End If

      Dim InOrOut As String = " <-left- "
      Select Case dispStart
        Case "INPUT"
          InOrOut = " <-left- "
        Case "OUTPUT"
          InOrOut = " -right-> "
      End Select

      If Val(DDSeq) = 1 And Val(ddConcatSeq) = 0 Then
        ListOfSteps.Add(stepName)
        swPumlFile.WriteLine()
        swPumlFile.WriteLine("node " & Chr(34) & pgmName & Chr(34) & " as " & stepName)
      End If


      Select Case DDName
        Case "STEPLIB"
        Case "SYSOUT"
        Case "SYSPRINT"
        Case "SYSUDUMP"
        Case Else
          If dsn.Length > 0 Then
            If ddConcatSeq > 0 Then
              DDName = DDName & LTrim(Str(ddConcatSeq))
            End If
            If dispEnd = "DELETE" Then
              dsn = "<s:red>" & dsn & "</s>"
            End If
            swPumlFile.WriteLine("file " & Chr(34) & dsn & Chr(34) & " as " & stepName & "." & DDName)
            swPumlFile.WriteLine(stepName & InOrOut & stepName & "." & DDName)
          End If
      End Select

    Loop
    ' write the final step connections
    swPumlFile.WriteLine("' STEP CONNECTIONS")
    For stepIndex = 0 To ListOfSteps.Count - 1
      If stepIndex = ListOfSteps.Count - 1 Then
        Exit For
      End If
      stepName = ListOfSteps(stepIndex)
      swPumlFile.WriteLine(ListOfSteps(stepIndex) &
                           " -[#blue,plain,thickness=16]-->" &
                           ListOfSteps(stepIndex + 1))
    Next
    swPumlFile.WriteLine("@enduml")
    swPumlFile.Close()

  End Sub
  Function DetermineStartDisp(ByRef fileDisp As String()) As String
    ' determine start disp
    DetermineStartDisp = ""
    If fileDisp Is Nothing Then
      DetermineStartDisp = "OUTPUT"
      Exit Function
    End If
    If fileDisp.Count >= 1 Then
      If fileDisp(0).Length = 0 Then
        DetermineStartDisp = "OUTPUT"
      Else
        Select Case fileDisp(0)
          Case "SHR"
            DetermineStartDisp = "INPUT"
          Case "MOD", "NEW"
            DetermineStartDisp = "OUTPUT"
          Case Else
            DetermineStartDisp = "INPUT"
        End Select
      End If
    Else
      DetermineStartDisp = "INPUT"
    End If

  End Function
  Function DetermineEndDisp(ByRef fileDisp As String()) As String
    DetermineEndDisp = ""
    If fileDisp Is Nothing Then
      DetermineEndDisp = "KEEP"
      Exit Function
    End If
    If fileDisp.Count >= 2 Then
      If fileDisp(1).Length = 0 Then
        DetermineEndDisp = "KEEP"
      Else
        Select Case fileDisp(1)
          Case "KEEP"
            DetermineEndDisp = "KEEP"
          Case "DELETE"
            DetermineEndDisp = "DELETE"
          Case "CATLG"
            DetermineEndDisp = "KEEP"
        End Select
      End If
    Else
      DetermineEndDisp = "KEEP"
    End If
  End Function
  Function DetermineAbendDisp(ByRef fileDisp As String()) As String
    DetermineAbendDisp = ""
    If fileDisp Is Nothing Then
      DetermineAbendDisp = "DELETE"
      Exit Function
    End If
    If fileDisp.Count >= 3 Then
      If fileDisp(2).Length = 0 Then
        DetermineAbendDisp = "KEEP"
      Else
        Select Case fileDisp(2)
          Case "DELETE"
            DetermineAbendDisp = "DELETE"
          Case "KEEP"
            DetermineAbendDisp = "KEEP"
          Case "CATLG"
            DetermineAbendDisp = "KEEP"
        End Select
      End If
    Else
      DetermineAbendDisp = "KEEP"
    End If
  End Function

  Sub CreateSummary()
    workbook = objExcel.Workbooks.Add
    SummaryWorksheet = workbook.Sheets.Item(1)
    SummaryWorksheet.Name = "Summary"
    SummaryRow = 0
    SummaryWorksheet.Range("A1").Value = "Mainframe Documentation Project" & vbNewLine &
                                         "Data Gathering Form" & vbNewLine &
                                         Path.GetFileNameWithoutExtension(txtDataGatheringForm.Text) & vbNewLine &
                                         "Model Created:" & Date.Now & vbNewLine &
                                         "ADDILite, Version:" & ProgramVersion
    SummaryWorksheet.Range("A2").Value = ""
    SummaryWorksheet.Range("B1").Value = ""
    SummaryWorksheet.Range("B2").Value = ""
    SummaryRow = 2
    For Each dgf In ListOfDataGathering
      SummaryRow += 1
      Dim row As Integer = LTrim(Str(SummaryRow))
      Dim dgfRow As String() = dgf.Split(Delimiter)
      SummaryWorksheet.Range("A" & row).Value = dgfRow(0)
      SummaryWorksheet.Range("B" & row).Value = dgfRow(1)
    Next


  End Sub
  Sub CreatePrograms()

    ' Build the worksheetsheet. Programs sheet is a list of all JCL Jobs with programs and DD details.
    lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly & " : Programs"
    If ProgramsRow = 0 Then
      worksheet = workbook.Sheets.Add(After:=workbook.Worksheets(workbook.Worksheets.Count))
      worksheet.Name = "Programs"
      ' Write the column headings row
      worksheet.Range("A1").Value = "JobName"
      worksheet.Range("B1").Value = "JobSeq"
      worksheet.Range("C1").Value = "ProcName"
      worksheet.Range("D1").Value = "ProcSeq"
      worksheet.Range("E1").Value = "StepName"
      worksheet.Range("F1").Value = "ExecName"
      worksheet.Range("G1").Value = "PgmName"
      worksheet.Range("H1").Value = "PgmSeq"
      worksheet.Range("I1").Value = "DD"
      worksheet.Range("J1").Value = "DDSeq"
      worksheet.Range("K1").Value = "DDConcatSeq"
      worksheet.Range("L1").Value = "DatasetName"
      worksheet.Range("M1").Value = "StartDisp"
      worksheet.Range("N1").Value = "EndDisp"
      worksheet.Range("O1").Value = "AbendDisp"
      worksheet.Range("P1").Value = "RecFM"
      worksheet.Range("Q1").Value = "LRECL"
      worksheet.Range("R1").Value = "DBMS"
      worksheet.Range("S1").Value = "ReportId"
      worksheet.Range("T1").Value = "ReportDescription"
      worksheet.Range("U1").Value = "SourceType"
      ProgramsRow = 1
    End If

    ' Write the data

    ' Read the DD CSV file back in and load to array
    Dim FileName = txtOutputFoldername.Text & "/" & FileNameOnly & "_DD.csv"
    If Not File.Exists(FileName) Then
      Exit Sub
    End If
    Dim csvCnt As Integer = 0
    Dim csvFile As FileIO.TextFieldParser = New FileIO.TextFieldParser(FileName)
    Dim csvRecord As String()           ' all fields(columns) for a given record
    csvFile.TextFieldType = FileIO.FieldType.Delimited
    csvFile.Delimiters = New String() {"|"}
    csvFile.HasFieldsEnclosedInQuotes = True
    Dim row As String = ""
    Dim cnt As Integer = 0
    Do While Not csvFile.EndOfData
      csvRecord = csvFile.ReadFields
      cnt += 1
      jobName = csvRecord(0)
      jobSequence = Val(csvRecord(1))
      procName = csvRecord(2)
      procSequence = Val(csvRecord(3))
      stepName = csvRecord(4)
      pgmName = csvRecord(5)
      execSequence = Val(csvRecord(6))
      DDName = csvRecord(7)
      ddSequence = csvRecord(8)
      ddConcatSeq = Val(csvRecord(9))
      Dim dsn = csvRecord(10)
      Dim startDisp As String = csvRecord(11)
      Dim endDisp As String = csvRecord(12)
      Dim abendDisp As String = csvRecord(13)
      Dim dcbRecFM As String = csvRecord(14)
      Dim dcbLrecl As String = csvRecord(15)
      Dim db2 As String = csvRecord(16)
      Dim reportID As String = csvRecord(17)
      Dim reportDescription As String = csvRecord(18)
      SourceType = csvRecord(19)
      execName = csvRecord(20)
      ProgramsRow += 1
      row = LTrim(Str(ProgramsRow))
      worksheet.Range("A" & row).Value = jobName
      worksheet.Range("B" & row).Value = LTrim(Str(jobSequence))
      worksheet.Range("C" & row).Value = procName
      worksheet.Range("D" & row).Value = LTrim(Str(procSequence))
      worksheet.Range("E" & row).Value = stepName
      worksheet.Range("F" & row).Value = execName
      worksheet.Range("G" & row).Value = pgmName
      worksheet.Range("H" & row).Value = LTrim(Str(execSequence))
      worksheet.Range("I" & row).Value = DDName
      worksheet.Range("J" & row).Value = LTrim(Str(ddSequence))
      worksheet.Range("K" & row).Value = LTrim(Str(ddConcatSeq))
      worksheet.Range("L" & row).Value = dsn
      worksheet.Range("M" & row).Value = startDisp
      worksheet.Range("N" & row).Value = endDisp
      worksheet.Range("O" & row).Value = abendDisp
      worksheet.Range("P" & row).Value = dcbRecFM
      worksheet.Range("Q" & row).Value = dcbLrecl
      worksheet.Range("R" & row).Value = db2
      worksheet.Range("S" & row).Value = reportID
      worksheet.Range("T" & row).Value = reportDescription
      worksheet.Range("U" & row).Value = SourceType
      ' load up a list of executable programs to analyze
      If ddSequence = 1 And ddConcatSeq = 0 Then
        Select Case pgmName
          Case "IEFBR14", "SORT", "IEBGENER", "IEBCOPY", "IDCAMS", "DSNUTILB",
               "SRCHPRNT", "CMSAUTO1", "PKZIP", "IKJEFT01", "A4204030", "OFORMAT", "ODATE", "DFSRRC00", "FTP",
               "ICETOOL", "FILEMGR", "IRXJCL", "INITFILE", "EZTPA00", "IERRCO00", "ABEND"
          Case Else
            If ListOfExecs.IndexOf(pgmName & Delimiter & SourceType) = -1 Then
              ListOfExecs.Add(pgmName & Delimiter & SourceType)
            End If
        End Select
      End If
      If cnt Mod 100 = 0 Then
        lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly & " : Programs = " & cnt
      End If
    Loop
    lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly & " : Programs Complete"

  End Sub
  Sub ProcessSourceFiles()
    Dim SourceRecordsCount As Integer = 0
    ' loop through the list of executables. Note we may be adding while processing (called members)
    For Each exec In ListOfExecs
      lblProcessingSource.Text = "Processing Source: " & exec
      Dim execs As String() = exec.Split(Delimiter)
      If execs.Count >= 2 Then
        exec = execs(0).Replace(Delimiter, "")
        SourceType = execs(1)
      End If
      If exec.Length = 0 Then
        LogFile.WriteLine(Date.Now & ",Source file name empty?," & FileNameOnly)
        Continue For
      End If

      'Load the infile to the stmt List

      Select Case SourceType
        Case "COBOL"
          SourceRecordsCount = LoadCobolStatementsToArray(exec)
        Case "Easytrieve"
          SourceRecordsCount = LoadEasytrieveStatementsToArray(exec)
        Case Else
          Continue For
      End Select
      If SourceRecordsCount = -1 Then
        Continue For
      End If
      If SourceRecordsCount = 0 Then
        LogFile.WriteLine(Date.Now & ",No Source Records Found," & exec)
        Continue For
      End If

      If SrcStmt.Count = 0 Then
        LogFile.WriteLine(Date.Now & ",No Source statements found," & exec)
        Continue For
      End If

      ' log file, optioned
      If cbLogStmt.Checked Then
        Call LogStmtArray(exec, SrcStmt)
      End If

      ' Analyze Source Statement array (SrcStmt) to get list of programs
      listOfPrograms.Clear()
      listOfPrograms = GetListOfPrograms()      'list of programs within the exec source

      ' Analyze Source Statement array (SrcStmt) to get list of paragraph names
      'ListOfParagraphs.Clear()
      Call GetListOfParagraphs()

      ' Analyze Source Statement array (SrcStmt) to get list of EXEC SQL statments
      'ListOfEXECSQL.Clear()
      Call GetListOfEXECSQL()

      If pgm.ProcedureDivision = -1 Then
        LogFile.WriteLine(Date.Now & ",Source is not complete," & exec)
        Continue For
      End If

      If cbScanModeOnly.Checked Then
        Continue For
      End If

      'write the output files/excel
      Select Case SourceType
        Case "COBOL"
          If WriteOutputCOBOL(exec) = -1 Then
            LogFile.WriteLine(Date.Now & ",Error while building COBOL output,")
          End If

        Case "Easytrieve"
          If WriteOutputEasytrieve(exec) = -1 Then
            LogFile.WriteLine(Date.Now & ",Error while building Easytrieve output," & exec)
          End If

      End Select

    Next
    If Not cbScanModeOnly.Checked Then
      Call CreateCommentsWorksheet()
    End If

    lblProcessingSource.Text = "Processing Sources: complete"
    lblProcessingWorksheet.Text = "Processing Worksheet: complete"

  End Sub
  Function LoadCobolStatementsToArray(ByRef CobolFile As String) As Integer
    '*---------------------------------------------------------
    ' Load COBOL lines to a Cobol statements array. 
    '*---------------------------------------------------------
    '
    'Assign the TempFileName for this particular cobolfile
    '
    tempCobFileName = txtOutputFoldername.Text & "\" & CobolFile & "_expandedCOB.txt"

    ' Remove the temporary work file
    Try
      If My.Computer.FileSystem.FileExists(tempCobFileName) Then
        My.Computer.FileSystem.DeleteFile(tempCobFileName)
      End If
    Catch ex As Exception
      LogFile.WriteLine(Date.Now & ",Removal of Temp hlk error," & ex.Message)
      LoadCobolStatementsToArray = -1
      Exit Function
    End Try

    SrcStmt.Clear()

    ' Include all the COPY members to the temporary file
    ' Drop Empty lines
    ' Only keeping: Indicator area, Area A, and Area B (cols 7-72)

    Dim swTemp As StreamWriter = Nothing
    Dim CopybookName As String = ""
    Dim IncludeDirective As String = ""
    Dim SQLDirective As String = ""
    Dim NumberOfCopysFound As Integer = 0
    Dim NumberOfScans As Integer = 0
    Dim SequenceNumberArea As String = ""
    Dim IndicatorArea As String = ""
    Dim AreaA As String = ""
    Dim AreaB As String = ""
    Dim AreaAandB As String = ""
    Dim CommentArea As String = ""
    Dim execSequenceNumberArea As String = ""
    Dim execIndicatorArea As String = ""
    Dim execAreaA As String = ""
    Dim execAreaB As String = ""
    Dim execAreaAandB As String = ""
    Dim execCommentArea As String = ""
    Dim combinedEXEC As String = ""
    Dim PeriodIsPresent As Boolean = False
    Dim cIndex As Integer = -1
    Dim debugCnt As Integer = -1
    Dim startIndex As Integer = -1
    Dim endIndex As Integer = -1
    Dim WithinDataDivision As Boolean = False
    Dim Division As String = ""

    ' Verify COBOL FILE exists

    Dim CobolFileName = txtSourceFolderName.Text & "\" & CobolFile
    If Not File.Exists(CobolFileName) Then
      LogFile.WriteLine(Date.Now & ",Source not found," & CobolFile)
      LoadCobolStatementsToArray = -1
      Exit Function
    End If
    LogFile.WriteLine(Date.Now & ",Processing Source," & CobolFile)

    ' Load the COBOL file into the working Array
    Dim CobolLines As String() = File.ReadAllLines(CobolFileName)

    ' Look for Program Written and Program Author
    ProgramAuthor = ""
    ProgramWritten = ""
    For index As Integer = 0 To CobolLines.Count - 1
      If ProgramAuthor.Length = 0 Then
        If CobolLines(index).Length >= 60 Then
          If CobolLines(index).PadRight(80).Substring(7, 7) = "AUTHOR." Then
            ProgramAuthor = CobolLines(index).PadRight(80).Substring(14, 40).Trim
          End If
        End If
      End If
      If ProgramWritten.Length = 0 Then
        If CobolLines(index).PadRight(80).Substring(7, 13) = "DATE-WRITTEN." Then
          ProgramWritten = CobolLines(index).PadRight(80).Substring(20, 40).Trim
        End If
      End If
      If ProgramAuthor.Length > 0 And ProgramWritten.Length > 0 Then
        Exit For
      End If
    Next

    ' Save the COMMENTS found in the program to the ListOfComments array
    Division = "PRE IDENT"
    For index As Integer = 0 To CobolLines.Count - 1
      ' determine which Division we are in
      Dim tempDiv As Integer = CobolLines(index).PadRight(80).IndexOf(" DIVISION")
      If tempDiv > -1 Then
        If CobolLines(index).Substring(6, 1) <> "*" Then
          If (tempDiv - 8 + 1) > 15 Then            'skip if too long to be a COBOL division value
            Continue For
          End If
          Division = CobolLines(index).Substring(7, tempDiv - 8 + 1)
          Continue For
        End If
      End If
      ' write this cobol line if a Comment AND we have Division value
      If CobolLines(index).Length >= 7 Then
        If CobolLines(index).Substring(6, 1) = "*" And Division.Length > 0 Then
          Call ProcessComment(index, CobolLines(index), Division, CobolFile)
          Continue For
        End If
        If CobolLines(index).Substring(6, 1) = "/" And Division.Length > 0 Then
          Mid(CobolLines(index), 7, 1) = "*"
          Call ProcessComment(index, CobolLines(index), Division, CobolFile)
          Continue For
        End If
      End If
      ' if Division is Identification (or ID) and it is not 'program-id' then the cobol line 
      '   is treated as comments and so will we.
      If Division.ToUpper.Trim = "IDENTIFICATION" Or Division.ToUpper.Trim = "ID" Then
        If Mid(CobolLines(index), 8, 11) <> "PROGRAM-ID." Then
          Call ProcessComment(index, CobolLines(index), Division, CobolFile)
          Continue For
        End If
      End If
    Next
    Division = ""

    ' Expand all copy/include members into a single file, 
    '   we also change empty lines to comment lines
    '   we also remove EJECT and SKIP compiler directives
    Do
      NumberOfScans += 1
      NumberOfCopysFound = 0
      debugCnt = 0
      swTemp = New StreamWriter(tempCobFileName, False)
      ' Process through the Cobol Lines
      For index As Integer = 0 To CobolLines.Count - 1
        debugCnt += 1
        ' make a whole blank/empty line a comment line
        If Len(Trim(CobolLines(index))) = 0 Then
          swTemp.WriteLine(Space(6) & "*")
          Continue For
        End If

        Call FillInAreas(CobolLines(index),
                         SequenceNumberArea, IndicatorArea, AreaA, AreaB, CommentArea)

        ' special adjustment for slash in column 7, must be a Telon artifact
        If IndicatorArea = "/" Then
          IndicatorArea = "*"
          Mid(CobolLines(index), 7, 1) = "*"
        End If

        ' Special adjustment for Telon/CICS/DB2 EXEC SQL lines
        If IndicatorArea = "*" And AreaA = "****" And (AreaB.Substring(0, 8) = "EXEC SQL") Then
          For y = index To CobolLines.Count - 1
            Mid(CobolLines(y), 7, 5) = Space(5)
            swTemp.WriteLine(CobolLines(y))
            If CobolLines(y).IndexOf("END-EXEC") > -1 Then
              index = y
              Exit For
            End If
          Next y
          Continue For
        End If

        ' write the comment line back out
        If IndicatorArea = "*" Then
          swTemp.WriteLine(CobolLines(index))
          Continue For
        End If

        ' keep the blank/empty line as a comment line, ignore the CommentArea
        Dim IndicatorAreadAreaAAreaB As String = IndicatorArea & AreaA & AreaB
        If IndicatorAreadAreaAAreaB.Trim.Length = 0 Then
          swTemp.WriteLine(Space(6) & "*")
          Continue For
        End If

        ' get the Compiler directive, if any
        AreaAandB = AreaA & AreaB

        ' remove EJECT And SKIP compiler directives
        If AreaAandB.Trim.Length >= 5 Then
          If AreaAandB.Trim.Substring(0, 5) = "EJECT" Then
            mid(CobolLines(index), 7, 1) = "*"
            swTemp.WriteLine(CobolLines(index))
            Continue For
          End If
        End If
        If AreaAandB.Trim.Length >= 4 Then
          If AreaAandB.Trim.Substring(0, 4) = "SKIP" Then
            mid(CobolLines(index), 7, 1) = "*"
            swTemp.WriteLine(CobolLines(index))
            Continue For
          End If
        End If

        IncludeDirective = AreaAandB.ToUpper
        Dim IDirective As String() = IncludeDirective.Trim.Split(New Char() {" "c})

        ' Checking for copy/include statement to process
        Dim CopyType As String = ""
        Select Case True
          Case IDirective(0) = "COPY"
            CopybookName = Trim(IDirective(1).Replace(".", " "))
            CopyType = IDirective(0)
          Case IDirective(0) = "++INCLUDE"
            CopybookName = Trim(IDirective(1).Replace(".", " "))
            CopyType = IDirective(0)
          Case IDirective(0) = "EXEC"
            ' need to "string together" this till END-EXEC 
            combinedEXEC = ""
            startIndex = index
            endIndex = -1
            For execIndex As Integer = index To CobolLines.Count - 1
              Call FillInAreas(CobolLines(execIndex),
                execSequenceNumberArea, execIndicatorArea, execAreaA, execAreaB, execCommentArea)
              If execIndicatorArea = "*" Then
                Continue For
              End If
              execAreaAandB = execAreaA & execAreaB
              combinedEXEC &= execAreaAandB.ToUpper
              If combinedEXEC.IndexOf("END-EXEC") > -1 Then
                combinedEXEC = DropDuplicateSpaces(combinedEXEC)
                endIndex = execIndex
                Exit For
              End If
            Next
            ' safety check
            If endIndex = -1 Then
              LogFile.WriteLine(Date.Now & ",Malformed SQL statement; missing END-EXEC," &
                              CobolLines(index) & " line#:" & index + 1)
              LoadCobolStatementsToArray = -1
              Exit Function
            End If
            ' check to see if this an SQL INCLUDE or some other SQL command
            Dim execDirective As String() = combinedEXEC.Trim.Split(New Char() {" "c})
            If execDirective(1) = "SQL" And execDirective(2) = "INCLUDE" Then
              CopyType = "SQL"
              CopybookName = execDirective(3)
              ' comment out these SQL INCLUDE statement(s)
              For execIndex As Integer = startIndex To endIndex
                Mid(CobolLines(execIndex), 7, 1) = "*"
                swTemp.WriteLine(CobolLines(execIndex))
              Next
            Else
              '  ' write out these non-INCLUDE SQL statements
              '  For execindex As Integer = startIndex To endIndex
              '    swTemp.WriteLine(CobolLines(execindex))
              '  Next
              swTemp.WriteLine(CobolLines(index))
              Continue For
            End If
            '            index = endIndex            'bypass already processed cobolLines
            '
          Case Else
            swTemp.WriteLine(CobolLines(index))
            Continue For
        End Select

        If Len(CopybookName) > 8 Then
          CopybookName = Mid(CopybookName, 1, 8)
        End If

        ' Expand copybooks/includes into the source
        NumberOfCopysFound += 1
        Dim CopybookFileName As String = txtSourceFolderName.Text &
                                         "\" & CopybookName
        swTemp.WriteLine(Space(6) & "*" & CopyType & " " & CopybookName & " Begin Include")
        LogFile.WriteLine(Date.Now & ",Including COBOL copybook," & CopybookName)
        Call IncludeCopyMember(CopybookFileName, swTemp)
        swTemp.WriteLine(Space(6) & "*" & CopyType & " " & CopybookName & " End Include")
      Next
      swTemp.Close()

      ' check we expanded any copybooks, if so we scan again for any copy/includes
      If NumberOfCopysFound > 0 Then                      'we found at least 1 COPY stmt
        CobolLines = File.ReadAllLines(tempCobFileName)   ' so load what we got so far
      End If

    Loop Until NumberOfCopysFound = 0

    '
    ' We should now deal with the compiler directive: REPLACE if there are any.
    ' Directives are before the PROGRAM-ID.
    '
    Dim cStatement As String = ""
    Dim statement As String = ""
    Dim procIndex As Integer = 0
    Dim continuation As Boolean = True
    LoadCobolStatementsToArray = 0

    ' Load the temp file to the array
    CobolLines = File.ReadAllLines(tempCobFileName)

    ' scan for REPLACE directive and then do a Global Search and Replace
    cIndex = -1
    For Each text1 As String In CobolLines
      cIndex += 1
      If Trim(text1).Length = 0 Then                        'ignore empty lines
        Continue For
      End If
      Call FillInAreas(text1, SequenceNumberArea, IndicatorArea, AreaA, AreaB, CommentArea)
      If IndicatorArea = "*" Then
        Continue For
      End If
      AreaAandB = AreaA & AreaB
      IncludeDirective = AreaAandB.ToUpper
      Dim tDirective As String() = IncludeDirective.Trim.Split(New Char() {" "c})
      If tDirective(0) = "REPLACE" Then
        Call ReplaceAll(AreaAandB, CobolLines, cIndex)
      End If
    Next

    ' Process the WHOLE/ALL the cobol lines now that copybooks are now embedded
    ' and replace is done.
    ' This is also where we concatenate the lines, as needed, into a single statement.
    '
    Dim hlkcounter As Integer = 0
    Division = ""
    SrcStmt.Clear()

    For Each text1 As String In CobolLines
      hlkcounter += 1
      LoadCobolStatementsToArray += 1
      text1 = text1.Replace(vbTab, Space(1))                'replace TAB(S) with single space!
      text1 = text1.Replace(vbNullChar, Space(1))           'replace nulls with space
      text1 = text1.Replace("�", Space(1))
      If Trim(text1).Length = 0 Then                        'drop empty lines
        Continue For
      End If

      Call FillInAreas(text1, SequenceNumberArea, IndicatorArea, AreaA, AreaB, CommentArea)

      If IndicatorArea = "*" Then                           'keep comments
        SrcStmt.Add(IndicatorArea & AreaA & AreaB)
        Continue For
      End If

      AreaAandB = AreaA & AreaB
      If AreaAandB.Trim.Length = 0 Then
        Continue For
      End If
      If Microsoft.VisualBasic.Right(RTrim(AreaAandB), 1) = "." Then
        PeriodIsPresent = True
      Else
        PeriodIsPresent = False
      End If

      IncludeDirective = DropDuplicateSpaces(AreaAandB.ToUpper)
      Dim tDirective As String() = IncludeDirective.Trim.Split(New Char() {" "c})
      cWord.Clear()
      For Each word In tDirective
        If word.Trim.Length > 0 Then
          cWord.Add(UCase(word))
        End If
      Next


      If cWord(0) = "CBL" Then                     'Drop CBL Compiler directive
        Continue For
      End If
      If DivisionFound(cWord, Division) Then    'Division could have been updated
        cStatement = ""
      End If

      If IndicatorArea = "-" Then
        cStatement &= AreaB
        Continue For
      End If

      ' concatenate till end-of-statement, which is a period.

      If PeriodIsPresent Then
        cStatement &= AreaAandB
        cStatement = Mid(cStatement, 1, 4) & DropDuplicateSpaces(Mid(cStatement, 5))
        SrcStmt.Add(cStatement)
        cStatement = ""
      Else
        cStatement &= AreaAandB
      End If

    Next

  End Function
  Function GetListOfPrograms() As List(Of ProgramInfo)
    ' Scan through the source looking for the programs.
    ' Each program could have multiple sub programs inline (especially COBOL).
    ' Also a program could call a sub program, which we will store out to a
    '  separate file (CallPgms.jcl) for later analysis.
    pgm.IdentificationDivision = -1
    pgm.ProcedureDivision = -1
    pgm.EnvironmentDivision = -1
    pgm.DataDivision = -1
    pgm.ProcedureDivision = -1
    pgm.EndProgram = -1
    pgm.ProgramId = ""

    Dim srcWords As New List(Of String)

    Select Case SourceType
      Case "COBOL"
        For stmtIndex As Integer = 0 To SrcStmt.Count - 1
          Select Case True
            Case SrcStmt(stmtIndex).Substring(0, 1) = "*"
              Continue For
            Case (SrcStmt(stmtIndex).IndexOf("IDENTIFICATION DIVISION.") > -1) Or
                (SrcStmt(stmtIndex).IndexOf("ID DIVISION.") > -1)
              If pgm.ProcedureDivision >= 1 Then
                pgm.EndProgram = stmtIndex - 1
                listOfPrograms.Add(pgm)
                pgm = Nothing
              End If
              pgm.IdentificationDivision = stmtIndex
              pgm.EnvironmentDivision = -1
              pgm.DataDivision = -1
              pgm.ProcedureDivision = -1
            Case SrcStmt(stmtIndex).IndexOf("ENVIRONMENT DIVISION.") > -1
              pgm.EnvironmentDivision = stmtIndex
            Case SrcStmt(stmtIndex).IndexOf("DATA DIVISION.") > -1
              pgm.DataDivision = stmtIndex
            Case SrcStmt(stmtIndex).IndexOf("PROCEDURE DIVISION") > -1
              pgm.ProcedureDivision = stmtIndex
            Case SrcStmt(stmtIndex).IndexOf("PROGRAM-ID.") > -1
              pgm.ProgramId = SrcStmt(stmtIndex).Substring(11).Replace(".", "").Trim
          End Select
          If pgm.ProcedureDivision > -1 Then
            If SrcStmt(stmtIndex).IndexOf(" CALL ") > -1 Then
              Call AddToListOfCallPgms(SrcStmt(stmtIndex), srcWords)
            End If
          End If
        Next
        If Not IsNothing(pgm) Then
          pgm.EndProgram = SrcStmt.Count - 1
          listOfPrograms.Add(pgm)
        End If

      Case "Easytrieve"
        pgm.EndProgram = SrcStmt.Count - 1
        pgm.IdentificationDivision = 0
        'the first one will be the "Division"
        For stmtIndex As Integer = 0 To SrcStmt.Count - 1
          Call GetSourceWords(SrcStmt(stmtIndex), srcWords)
          Select Case True
            Case SrcStmt(stmtIndex).IndexOf("PROGRAM-ID.") > -1
              pgm.ProgramId = SrcStmt(stmtIndex).Substring(13).Trim
            Case SrcStmt(stmtIndex).Substring(0, 1) = "*"
              Continue For
            Case srcWords(0) = "FILE" Or
                srcWords(0) = "SQL"
              If pgm.EnvironmentDivision = -1 Then
                pgm.EnvironmentDivision = stmtIndex
                pgm.DataDivision = stmtIndex
              End If
            Case srcWords(0) = "JOB" Or
                 srcWords(0) = "SORT"
              If pgm.ProcedureDivision = -1 Then
                pgm.ProcedureDivision = stmtIndex
              End If
          End Select
        Next
        listOfPrograms.Add(pgm)
    End Select

    GetListOfPrograms = listOfPrograms
  End Function
  Sub AddToListOfCallPgms(ByRef statement As String, ByRef srcWords As List(Of String))
    Dim CalledFileName As String = ""
    Dim CalledSourceType As String = ""
    CalledMember = ""
    Call GetSourceWords(statement, srcWords)
    For x As Integer = 0 To srcWords.Count - 1
      If srcWords(x) = "CALL" Then
        CalledMember = srcWords(x + 1)
        If Mid(CalledMember, 1, 1) <> "'" Then         'skip dynamic called routines
          LogFile.WriteLine(Date.Now & ",Called Member is not Static called so skipped," & CalledMember)
          Continue For
        End If
        CalledMember = CalledMember.Replace("'", "").Trim
        ' if a utility, ie ABEND do not add to list
        Select Case CalledMember
          Case "DSNHADDR", "DSNHADD2", "DSNHLI", "ADRABND", "ABEND", "DSNTIAR", "CBLTDLI",
               "DHSFINAL", "ADMAAB0", "ADMAABT", "BINCONV", "CNTYNAME", "ICOMMA2",
               "OCDATE", "OCDATE9", "OCNTRYDT", "ODATE", "OFORMAT", "OMOCNTRY",
               "R8460720", "PA8403AB", "PA1581BA"
            LogFile.WriteLine(Date.Now & ",Called Member is Utility so skipped," & CalledMember)
            Exit Sub
        End Select
        CalledFileName = txtSourceFolderName.Text & "\" & CalledMember
        If File.Exists(CalledFileName) Then
          CalledSourceType = GetSourceType(CalledFileName)
          If ListOfCallPgms.IndexOf(CalledMember & Delimiter & CalledSourceType & Delimiter & pgm.ProgramId) = -1 Then
            ListOfCallPgms.Add(CalledMember & Delimiter & CalledSourceType & Delimiter & pgm.ProgramId)
            LogFile.WriteLine(Date.Now & ",Called Member added to ListOfCallPgms," & CalledMember)
          End If
        Else
          LogFile.WriteLine(Date.Now & ",Called Member Source Not Found," & CalledMember)
        End If
      End If
    Next
  End Sub
  Sub GetListOfParagraphs()
    For Each pgm In listOfPrograms
      Select Case SourceType
        Case "COBOL"
          For stmtIndex As Integer = pgm.ProcedureDivision + 1 To SrcStmt.Count - 1
            If SrcStmt(stmtIndex).Substring(0, 1) = "*" Then
              Continue For
            End If
            If SrcStmt(stmtIndex).Length > 5 Then
              If SrcStmt(stmtIndex).Substring(0, 4) <> Space(4) Then
                Call GetSourceWords(SrcStmt(stmtIndex), cWord)
                ListOfParagraphs.Add(cWord(0))
              End If
            End If
          Next stmtIndex

        Case "Easytrieve"
      End Select
    Next pgm

  End Sub
  Sub GetListOfEXECSQL()
    Dim StartIndex As Integer = -1
    Dim EndIndex As Integer = -1
    Dim execCnt As Integer = 0
    Dim Statement As String = ""
    Dim ExecSQL As String = ""
    Dim Table As String = ""
    Dim Cursor As String = ""
    Dim x As Integer = 0
    Dim y As Integer = 0
    Dim z As Integer = 0
    Dim ListOfTables As New List(Of String)
    Dim JustTheTable As String = ""
    For Each pgm In listOfPrograms
      Select Case SourceType
        Case "COBOL"
          For stmtIndex As Integer = pgm.DataDivision + 1 To pgm.EndProgram
            If SrcStmt(stmtIndex).IndexOf("EXEC SQL") = -1 Then
              Continue For
            End If

            Call GetSourceWords(SrcStmt(stmtIndex), cWord)

            For x = 0 To cWord.Count - 1
              If cWord(x) = "END-EXEC" Then
                Continue For
              End If

              If (x + 2) < cWord.Count - 1 Then
                If cWord(x) = "EXEC" And cWord(x + 1) = "SQL" And cWord(x + 2) = "ROLLBACK" Then
                  ExecSQL = cWord(x + 2)
                  Table = ""
                  Cursor = ""
                  Statement = ""
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                  x += 2
                  Continue For
                End If
              End If

              If (x + 3) < cWord.Count - 1 Then
                If cWord(x) = "EXEC" And cWord(x + 1) = "SQL" And cWord(x + 2) = "OPEN" Then
                  ExecSQL = cWord(x + 2)
                  Table = ""
                  Cursor = cWord(x + 3)
                  Statement = ""
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                  x += 3
                  Continue For
                End If

                If cWord(x) = "EXEC" And cWord(x + 1) = "SQL" And cWord(x + 2) = "CLOSE" Then
                  ExecSQL = cWord(x + 2)
                  Table = ""
                  Cursor = cWord(x + 3)
                  Statement = ""
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                  x += 3
                  Continue For
                End If

                If cWord(x) = "EXEC" And cWord(x + 1) = "SQL" And cWord(x + 2) = "UPDATE" Then
                  ExecSQL = cWord(x + 2)
                  Table = cWord(x + 3)
                  Cursor = ""
                  Statement = ""
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                  x += 3
                  Continue For
                End If

                If cWord(x) = "EXEC" And cWord(x + 1) = "SQL" And cWord(x + 2) = "SELECT" Then
                  ExecSQL = cWord(x + 2)
                  Table = ""
                  Cursor = ""
                  ListOfTables.Clear()
                  Statement = ""
                  ' Build the statement
                  For y = x + 2 To cWord.Count - 1
                    If cWord(y) = "END-EXEC" Then
                      Exit For
                    End If
                    Statement &= cWord(y) & " "
                  Next y
                  ' Build the Table using FROM clause
                  For y = x + 2 To cWord.Count - 1
                    If cWord(y) = "END-EXEC" Then
                      Exit For
                    End If
                    If cWord(y) = "FROM" Then
                      Table = ""
                      For z = y + 1 To cWord.Count - 1
                        If cWord(z) = "WHERE" Or cWord(z) = "END-EXEC" Then
                          If ListOfTables.IndexOf(Table) = -1 Then
                            ListOfTables.Add(Table)
                            Table = ""
                            Exit For
                          End If
                        End If
                        If Not cWord(z).EndsWith(",") Then
                          If Table.Length = 0 Then
                            Table = cWord(z)
                          End If
                        Else
                          If Table.Length = 0 Then
                            Table = cWord(z)
                          End If
                          If ListOfTables.IndexOf(Table) = -1 Then
                            ListOfTables.Add(Table)
                            Table = ""
                          End If
                        End If
                      Next z
                      y = z + 1
                    End If
                  Next y
                  ' finalize Table
                  Table = ""
                  For Each TableEntry In ListOfTables
                    Dim SchemaOrTable As String() = TableEntry.Split(".")
                    If SchemaOrTable.Count > 1 Then
                      JustTheTable = SchemaOrTable(1)
                    Else
                      JustTheTable = SchemaOrTable(0)
                    End If
                    If Not JustTheTable.EndsWith(",") Then
                      Table &= JustTheTable.Trim & ","
                    Else
                      Table &= JustTheTable.Trim
                    End If
                  Next
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement.Trim)
                  x = y + 1
                  Continue For
                End If
              End If

              If (x + 4) < cWord.Count - 1 Then
                If cWord(x) = "*" And cWord(x + 1) = "EXEC" And cWord(x + 2) = "SQL" And cWord(x + 3) = "INCLUDE" Then
                  ExecSQL = cWord(x + 3)
                  Statement = cWord(x + 4)
                  Table = ""
                  Cursor = ""
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                  x += 4
                  Continue For
                End If

                If cWord(x) = "EXEC" And cWord(x + 1) = "SQL" And cWord(x + 2) = "DECLARE" And cWord(x + 4) = "TABLE" Then
                  ExecSQL = cWord(x + 2) & " " & cWord(x + 4)
                  Statement = ""
                  Table = cWord(x + 3)
                  Cursor = ""
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                  x += 4
                  Continue For
                End If

                If cWord(x) = "EXEC" And cWord(x + 1) = "SQL" And cWord(x + 2) = "INSERT" Then
                  ExecSQL = cWord(x + 2)
                  Table = cWord(x + 4)
                  Cursor = ""
                  Statement = ""
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                  x += 4
                  Continue For
                End If

                If cWord(x) = "EXEC" And cWord(x + 1) = "SQL" And cWord(x + 2) = "DELETE" Then
                  ExecSQL = cWord(x + 2)
                  Table = cWord(x + 4)
                  Cursor = ""
                  Statement = ""
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                  x += 4
                  Continue For
                End If
              End If

              If (x + 5) < cWord.Count - 1 Then
                If cWord(x) = "EXEC" And cWord(x + 1) = "SQL" And cWord(x + 2) = "FETCH" Then
                  ExecSQL = cWord(x + 2)
                  Table = ""
                  Cursor = cWord(x + 3)
                  Statement = ""
                  For y = x + 4 To cWord.Count - 1
                    If cWord(y) = "END-EXEC" Then
                      Exit For
                    End If
                    Statement &= cWord(y) & " "
                  Next y
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement.Trim)
                  x = y + 1
                  Continue For
                End If
              End If

              If (x + 6) < cWord.Count - 1 Then
                If cWord(x) = "EXEC" And cWord(x + 1) = "SQL" And cWord(x + 2) = "DECLARE" And cWord(x + 4) = "CURSOR" Then
                  ExecSQL = cWord(x + 2) & " " & cWord(x + 4)
                  Table = ""
                  Cursor = cWord(x + 3)
                  Statement = ""
                  ListOfTables.Clear()
                  ' Build the Statement
                  For y = x + 2 To cWord.Count - 1
                    If cWord(y) = "END-EXEC" Then
                      Exit For
                    End If
                    Statement &= cWord(y) & " "
                  Next y
                  ' Build the Table using the FROM clause
                  For y = x + 6 To cWord.Count - 1
                    If cWord(y) = "END-EXEC" Then
                      Exit For
                    End If
                    If cWord(y) = "FROM" Then
                      Table = ""
                      For z = y + 1 To cWord.Count - 1
                        Select Case cWord(z)
                          Case "WHERE", "INNER", "JOIN", "(", "SELECT", "END-EXEC", "ORDER"
                            If ListOfTables.IndexOf(Table) = -1 Then
                              ListOfTables.Add(Table)
                              y = z + 1
                              Exit For
                            End If
                        End Select
                        If Not cWord(z).EndsWith(",") Then
                          If Table.Length = 0 Then
                            Table = cWord(z)
                          End If
                        Else
                          If Table.Length = 0 Then
                            Table = cWord(z)
                          End If
                          Table &= ","
                          If ListOfTables.IndexOf(Table) = -1 Then
                            ListOfTables.Add(Table)
                          End If
                          Table = ""
                        End If
                      Next z
                    End If
                  Next y
                  ' finalize build of Table
                  Table = ""
                  For Each TableEntry In ListOfTables
                    Dim SchemaOrTable As String() = TableEntry.Split(".")
                    If SchemaOrTable.Count > 1 Then
                      JustTheTable = SchemaOrTable(1)
                    Else
                      JustTheTable = SchemaOrTable(0)
                    End If
                    If Not JustTheTable.EndsWith(",") Then
                      Table &= JustTheTable.Trim & ","
                    Else
                      Table &= JustTheTable.Trim
                    End If
                  Next
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement.Trim)
                  x = y + 1
                  Continue For
                End If
              End If
            Next x
          Next stmtIndex

        Case "Easytrieve"
          For stmtIndex As Integer = pgm.DataDivision + 1 To pgm.EndProgram
            Call GetSourceWords(SrcStmt(stmtIndex), cWord)
            x = 0

            If (cWord.Count - 1) >= 2 Then
              If cWord(0) = "SQL" And cWord(1) = "OPEN" Then
                ExecSQL = cWord(1)
                Table = cWord(2)
                Cursor = ""
                Statement = ""
                Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                Continue For
              End If
            End If

            If (cWord.Count - 1) >= 2 Then
              If cWord(0) = "SQL" And cWord(1) = "CLOSE" Then
                ExecSQL = cWord(1)
                Table = cWord(2)
                Cursor = ""
                Statement = ""
                Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                Continue For
              End If
            End If

            If (cWord.Count - 1) >= 3 Then
              If cWord(0) = "*SQL" And cWord(1) = "INCLUDE" Then
                ExecSQL = cWord(1)
                Statement = cWord(3)
                Table = ""
                Cursor = ""
                Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                Continue For
              End If
              If cWord(0) = "SQL" And cWord(1) = "INCLUDE" Then
                ExecSQL = cWord(1)
                Statement = cWord(3)
                Table = ""
                Cursor = ""
                Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                Continue For
              End If

              If cWord(0) = "SQL" And cWord(1) = "FETCH" Then
                ExecSQL = cWord(1)
                Table = cWord(2)
                Cursor = ""
                Statement = ""
                For y = 3 To cWord.Count - 1
                  Statement &= cWord(y).Trim & " "
                Next y
                Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement.Trim)
                Continue For
              End If

              If cWord(0) = "SQL" And cWord(1) = "SELECT" Then
                ExecSQL = cWord(1)
                Table = ""
                Cursor = ""
                ' find the FROM
                Statement = ""
                ListOfTables.Clear()
                ' build the Statement
                For y = 1 To cWord.Count - 1
                  If cWord(y) = "END-EXEC" Then
                    Exit For
                  End If
                  Statement &= cWord(y) & " "
                Next y
                ' build the Table using the FROM clause
                For y = 3 To cWord.Count - 1
                  If cWord(y) = "END-EXEC" Then
                    Exit For
                  End If
                  If cWord(y) = "FROM" Then
                    Table = ""
                    For z = y + 1 To cWord.Count - 1
                      Select Case cWord(z)
                        Case "WHERE", "INNER", "JOIN", "(", "SELECT", "END-EXEC"
                          If ListOfTables.IndexOf(Table) = -1 Then
                            ListOfTables.Add(Table)
                            Table = ""
                            Exit For
                          End If
                      End Select
                      If Not cWord(z).EndsWith(",") Then
                        If Table.Length = 0 Then
                          Table = cWord(z)
                        End If
                      Else
                        If Table.Length = 0 Then
                          Table = cWord(z)
                        End If
                        If ListOfTables.IndexOf(Table) = -1 Then
                          ListOfTables.Add(Table)
                          Table = ""
                        End If
                      End If
                    Next z
                    y = z + 1
                  End If
                Next y
                ' finalize build of Table
                Table = ""
                For Each TableEntry In ListOfTables
                  Dim SchemaOrTable As String() = TableEntry.Split(".")
                  If SchemaOrTable.Count > 1 Then
                    JustTheTable = SchemaOrTable(1)
                  Else
                    JustTheTable = SchemaOrTable(0)
                  End If
                  If Not JustTheTable.EndsWith(",") Then
                    Table &= JustTheTable.Trim & ","
                  Else
                    Table &= JustTheTable.Trim
                  End If
                Next
                Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement.Trim)
                Continue For
              End If
            End If

            If (cWord.Count - 1) >= 4 Then
              If cWord(0) = "EXEC" And cWord(1) = "SQL" And cWord(2) = "DECLARE" And cWord(4) = "TABLE" Then
                ExecSQL = cWord(2) & " " & cWord(4)
                Statement = ""
                Table = cWord(3)
                Cursor = ""
                Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                Continue For
              End If
            End If

            If (cWord.Count - 1) >= 6 Then
              If cWord(0) = "SQL" And cWord(1) = "DECLARE" And cWord(3) = "CURSOR" Then
                ExecSQL = cWord(1) & " " & cWord(3)
                Table = ""
                Cursor = cWord(2)
                Statement = ""
                For y = 3 To cWord.Count - 1
                  Statement &= cWord(y).Trim & " "
                  If cWord(y) = "FROM" Then
                    For z = y + 1 To cWord.Count - 1
                      If cWord(z) = "WHERE" Or cWord(z) = "END-EXEC" Then
                        If ListOfTables.IndexOf(Table) = -1 Then
                          ListOfTables.Add(Table)
                          Exit For
                        End If
                      End If
                      If Not cWord(z).EndsWith(",") Then
                        If Table.Length = 0 Then
                          Table = cWord(z)
                        End If
                      Else
                        If Table.Length = 0 Then
                          Table = cWord(z)
                        End If
                        If ListOfTables.IndexOf(Table) = -1 Then
                          ListOfTables.Add(Table)
                          Table = ""
                        End If
                      End If
                    Next z
                  End If
                Next y
                Table = ""
                For Each TableEntry In ListOfTables
                  Dim SchemaOrTable As String() = TableEntry.Split(".")
                  If SchemaOrTable.Count > 1 Then
                    JustTheTable = SchemaOrTable(1)
                  Else
                    JustTheTable = SchemaOrTable(0)
                  End If
                  If Not JustTheTable.EndsWith(",") Then
                    Table &= JustTheTable.Trim & ","
                  Else
                    Table &= JustTheTable.Trim
                  End If
                Next
                Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement.Trim)
                Continue For
              End If
            End If

            If (cWord.Count - 1) >= 2 Then
              If cWord(0) = "SQL" And cWord(1) = "OPEN" Then
                ExecSQL = cWord(1)
                Table = cWord(2)
                Cursor = ""
                Statement = ""
                Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement.Trim)
                Continue For
              End If
            End If
          Next stmtIndex

      End Select
    Next pgm

  End Sub
  Sub AddToListOfEXECSQL(ByRef Execcnt As Integer,
                         ByRef execSql As String,
                         ByRef Table As String,
                         ByRef Cursor As String,
                         ByVal Statement As String)
    ' Need to remove any Delimiters within the fields
    Statement = Statement.Replace(Delimiter, "&")
    Execcnt += 1
    ListOfEXECSQL.Add(FileNameOnly & Delimiter &
                      pgm.ProgramId & Delimiter &
                      execSql & Delimiter &
                      LTrim(Str(Execcnt)) & Delimiter &
                      Table & Delimiter &
                      Cursor & Delimiter &
                      Statement)
  End Sub
  Function GetSourceType(ByRef FileName As String) As String
    ' Identify if this file is COBOL or Easytrieve or Utility
    GetSourceType = ""
    If FileName.Length = 0 Then
      LogFile.WriteLine(Date.Now & ",Filename for GetSourcetype is empty," & FileNameOnly)
      Exit Function
    End If
    Select Case FileName
      Case "IEFBR14", "IDCAMS", "SORT", "ICETOOL", "IEBGENER", "PKZIP", "IKJEFT01", "DFSRRC00",
           "FTP", "FILEMGR"
        GetSourceType = "UTILITY"
        Exit Function
      Case "SRCHPRNT", "CMSAUTO1", "DSNUTILB", "OFORMAT", "ODATE", "A4204030", "IRXJCL", "INITFILE", "EZTPA00",
           "IERRCO00", "ABEND"
        GetSourceType = "UTILITY"
        Exit Function
    End Select
    Dim SourceFileName As String = txtSourceFolderName.Text & "\" & FileName
    If Not File.Exists(SourceFileName) Then
      LogFile.WriteLine(Date.Now & ",Source File Not found," & FileName)
      GetSourceType = "NotFound"
      Exit Function
    End If
    Dim CobolLines As String() = File.ReadAllLines(SourceFileName)
    For index As Integer = 0 To CobolLines.Count - 1
      If Len(Trim(CobolLines(index))) = 0 Then
        Continue For
      End If
      If (CobolLines(index).ToUpper.IndexOf("IDENTIFICATION DIVISION.") > -1) Or
        (CobolLines(index).ToUpper.IndexOf("ID DIVISION.") > -1) Then
        GetSourceType = "COBOL"
        Exit Function
      End If
      If CobolLines(index).Length >= 6 Then
        Select Case CobolLines(index).ToUpper.Substring(0, 4)
          Case "PARM", "FILE", "SORT", "JOB "
            GetSourceType = "Easytrieve"
            Exit Function
        End Select
      End If
    Next
    LogFile.WriteLine(Date.Now & ",Unknown Type of Source File," & FileName)
    GetSourceType = "Unknown"

  End Function
  Sub FillInAreas(ByVal CobolLine As String,
                  ByRef SequenceNumberArea As String,
                  ByRef IndicatorArea As String,
                  ByRef AreaA As String,
                  ByRef AreaB As String,
                  ByRef CommentArea As String)
    ' break the line into COBOL format areas
    ' Ensure line is 80 characters in length
    Dim Line As String = CobolLine.PadRight(80)
    ' extract out the COBOL areas (remember on substring startindex is zero-based!)
    SequenceNumberArea = Line.Substring(0, 6)   'cols 1-6
    IndicatorArea = Line.Substring(6, 1)        'cols 7
    AreaA = Line.Substring(7, 4)                'cols 8-11
    AreaB = Line.Substring(11, 61)              'cols 12-72
    CommentArea = Line.Substring(72, 8)         'cols 73-80
  End Sub
  Function DropDuplicateSpaces(ByVal text As String) As String
    DropDuplicateSpaces = Regex.Replace(text, " +", " ")
  End Function
  Sub IncludeCopyMember(ByVal CopyMember As String,
                        ByRef swTemp As StreamWriter)
    If File.Exists(CopyMember) = False Then
      swTemp.Write(Space(6) & "*Member not found:")
      swTemp.WriteLine(Path.GetFileNameWithoutExtension(CopyMember))
      lblCopybookMessage.Text = "Copy Member not found:" & CopyMember
      LogFile.WriteLine(Date.Now & ",Copy Member not found," & Path.GetFileNameWithoutExtension(CopyMember))
      Exit Sub
    End If
    'Dim debugName As String = Path.GetFileNameWithoutExtension(CopyMember)
    Dim debugCnt As Integer = 0
    Dim IncludeLines As String() = File.ReadAllLines(CopyMember)
    ' append copymember to temp file and drop blank lines
    For Each line As String In IncludeLines
      debugCnt += 1
      If Len(Trim(line)) > 0 Then
        swTemp.WriteLine(line)
      End If
    Next
  End Sub
  Sub ReplaceAll(ByRef cStatement As String, ByRef CobolLines As String(), ByRef cIndex As Integer)
    ' for Compiler directive 'Replace'. Do the substitutions here.
    Dim tLine As String = RTrim(cStatement)
    If Microsoft.VisualBasic.Right(tLine, 1).Equals(".") Then        'remove the period
      tLine = tLine.Remove(tLine.Length - 1, 1)
    End If
    Dim tWord = tLine.Trim.Split(New Char() {" "c})
    Dim SearchFor As String = tWord(1).Replace("=", " ").Trim
    Dim ReplaceWith As String = tWord(3).Replace("=", " ").Trim
    ' loop through the CobolLines array replacing all <SearchFor> with <ReplaceWith>
    For index As Integer = cIndex + 1 To CobolLines.Count - 1
      If CobolLines(index).IndexOf(SearchFor) > -1 Then
        CobolLines(index) = CobolLines(index).Replace(SearchFor, ReplaceWith)
      End If
    Next
  End Sub
  Function DivisionFound(ByRef cWord As List(Of String), ByRef Division As String) As Boolean
    ' Caution! This not only returns true/false but ALSO updates the Division value
    '  if it encountered a division statement.
    DivisionFound = False
    If cWord.Count < 2 Then
      Exit Function
    End If
    If cWord(1).IndexOf("DIVISION") > -1 Then
      Select Case cWord(0)
        Case "IDENTIFICATION", "ID",
             "ENVIRONMENT",
             "DATA",
             "PROCEDURE"
          Division = cWord(0)
          DivisionFound = True
      End Select
    End If
  End Function
  Function WriteOutputEasytrieve(ByRef exec As String) As Integer
    WriteOutputEasytrieve = 0
    ' Create a Plantuml file, step by step, based on the Procedure division.
    Call CreatePumlEasytrieve(exec)

    ' Create a Records/Fields spreadsheet
    Call CreateRecordFile()

    ' Call CreateComponentsFile()


  End Function
  Function WriteOutputCOBOL(ByRef exec As String) As Integer
    ' Write the output Pgm, Data, Procedure, copy files.
    ' return of -1 means an error
    ' return of 0 means all is okay

    WriteOutputCOBOL = 0

    ' Create a Plantuml file, step by step, based on the Procedure division.
    Call CreatePumlCOBOL(exec)

    ' Create a Records/Fields spreadsheet
    Call CreateRecordFile()

    ' Create the Business Rules
    Call CreateBRCOBOL(exec)

    ' Call CreateComponentsFile()

  End Function
  Function LoadEasytrieveStatementsToArray(ByRef exec As String) As Integer
    '*---------------------------------------------------------
    ' Load Easytrieve lines to a statements array. 
    '*---------------------------------------------------------
    '
    'Assign the Temporay File Name for this particular Easytrieve file
    '
    tempEZTFileName = txtOutputFoldername.Text & "\" & exec & "_expandedEZT.txt"

    ' Remove the temporary work file
    Try
      If My.Computer.FileSystem.FileExists(tempEZTFileName) Then
        My.Computer.FileSystem.DeleteFile(tempEZTFileName)
      End If
    Catch ex As Exception
      LogFile.WriteLine(Date.Now & ",Removal of Temp EZT error," & ex.Message)
      LoadEasytrieveStatementsToArray = -1
      Exit Function
    End Try

    Dim FileName As String = txtSourceFolderName.Text & "\" & exec
    If Not File.Exists(FileName) Then
      LogFile.WriteLine(Date.Now & ",Source File Not Found," & exec)
      LoadEasytrieveStatementsToArray = -1
      Exit Function
    Else
      LogFile.WriteLine(Date.Now & ",Processing Source," & exec)
    End If

    ' put all the lines into the array
    Dim EztLinesLoaded As String() = File.ReadAllLines(FileName)

    ' Load the COMMENTS found in the program to the ListOfComments array
    Dim Division As String = "IDENTIFICATION"
    For index As Integer = 0 To EztLinesLoaded.Count - 1
      ' determine which "Division" we are in
      ' Easytrieve anything before "FILE" or "SQL" is IDENTIFICATION
      '   anything before "JOB" or "SORT" is DATA
      '   anything after "JOB" or "SORT" is PROCEDURE
      If EztLinesLoaded(index).Length >= 5 Then
        If EztLinesLoaded(index).Substring(0, 5) = "FILE " Then
          Division = "DATA"
        End If
        If EztLinesLoaded(index).Substring(0, 4) = "SQL " Then
          Division = "DATA"
        End If
        If EztLinesLoaded(index).Substring(0, 4) = "JOB " Then
          Division = "PROCEDURE"
        End If
        If EztLinesLoaded(index).Substring(0, 5) = "SORT " Then
          Division = "PROCEDURE"
        End If
      End If
      ' write line if this is a Comment and we have Division value
      If EztLinesLoaded(index).Length >= 1 Then
        If EztLinesLoaded(index).Substring(0, 1) = "*" And Division.Length > 0 Then
          Call ProcessComment(index, EztLinesLoaded(index), Division, exec)
        End If
      End If
    Next
    Division = ""


    Dim statement As String = ""
    Dim newLine As String = ""
    Dim swTemp As StreamWriter = Nothing
    swTemp = New StreamWriter(tempEZTFileName, False)
    swTemp.WriteLine("*PROGRAM-ID. " & exec)
    Dim reccnt As Integer = 0

    ' process the eztlinesloaded array
    '  we will drop empty/blank lines, trim off leading spaces,
    '  and combine continued lines
    '
    For index As Integer = 0 To EztLinesLoaded.Length - 1
      ' remove columns 73-80 and trim off extra spaces
      EztLinesLoaded(index) = Trim(Microsoft.VisualBasic.Left(EztLinesLoaded(index) & Space(72), 72))
      If Trim(EztLinesLoaded(index)).Length = 0 Then
        Continue For
      End If
      If Mid(EztLinesLoaded(index), 1, 1) = "*" Then
        swTemp.WriteLine(EztLinesLoaded(index))
        Continue For
      End If

      ' combine continued lines
      statement = Trim(Microsoft.VisualBasic.Left(EztLinesLoaded(index) & Space(72), 72))
      If statement.EndsWith(" +") = False Then
        swTemp.WriteLine(statement.Trim)
        reccnt += 1
        Continue For
      End If

      For continuedIndex As Integer = index + 1 To EztLinesLoaded.Length - 1
        statement &= Trim(Microsoft.VisualBasic.Left(EztLinesLoaded(continuedIndex) & Space(72), 72))
        If statement.EndsWith(" +") <> True Then
          swTemp.WriteLine(statement.Replace(" +", " ").Trim)
          reccnt += 1
          index = continuedIndex
          Exit For
        End If
      Next continuedIndex
    Next index
    swTemp.Close()

    ' Resolve the includes/copybooks
    ' TODO: convert to easytrieve format? '*here

    EztLinesLoaded = File.ReadAllLines(tempEZTFileName)
    swTemp = New StreamWriter(tempEZTFileName, False)
    Dim CBFileName As String = ""
    Dim srcWords As New List(Of String)
    Dim db2TableName As String = ""
    For index As Integer = 0 To EztLinesLoaded.Length - 1
      statement = EztLinesLoaded(index)
      Call GetSourceWords(statement, srcWords)
      If srcWords.Count >= 4 Then
        If srcWords(0) = "SQL" And srcWords(1) = "INCLUDE" And srcWords(2) = "FROM" Then
          Dim tableParameters As String() = srcWords(3).Split(".")
          If tableParameters.Count = 2 Then
            db2TableName = tableParameters(1)
          Else
            db2TableName = tableParameters(0)
          End If
          'search DB2 Declares table to get member name(include file name)
          CBFileName = ""
          Dim db2Index As Integer = 0
          Dim db2Found As Boolean = False
          For db2Index = 0 To DB2Declares.Count - 1
            If DB2Declares(db2Index) = db2TableName Then
              db2Found = True
              Exit For
            End If
          Next
          ' Include the member
          If db2Found Then
            swTemp.WriteLine("*%COPYBOOK SQL " & MembersNames(db2Index))
            swTemp.WriteLine("*" & statement & " Begin Include")
            CBFileName = txtSourceFolderName.Text & "\" & MembersNames(db2Index)
            Call IncludeCopyMember(CBFileName, swTemp)
          Else
            swTemp.WriteLine("*Crossref to DB2Declares not found in SOURCES folder")
            LogFile.WriteLine(Date.Now & ",Crossref to DB2Declares not found in SOURCES folder," &
                              db2TableName)
          End If
          swTemp.WriteLine("*" & statement & " End Include")
          Continue For
        End If
      End If
      If statement.Length > 0 Then
        If statement.Substring(0, 1) = "%" Then
          swTemp.WriteLine("*%COPYBOOK FILE " & statement.Substring(1).Trim)
          swTemp.WriteLine("*" & statement & " Begin Include")
          CBFileName = txtSourceFolderName.Text & "\" & statement.Substring(1)
          Call IncludeCopyMember(CBFileName, swTemp)
          swTemp.WriteLine("*" & statement & " End Include")
          Continue For
        End If
      End If
      swTemp.WriteLine(statement)
    Next
    swTemp.Close()

    ' load the Source Statement Array with final content
    SrcStmt.Clear()

    reccnt = 0
    EztLinesLoaded = File.ReadAllLines(tempEZTFileName)
    For index = 0 To EztLinesLoaded.Count - 1
      SrcStmt.Add(EztLinesLoaded(index).Trim)
      reccnt += 1
    Next

    LoadEasytrieveStatementsToArray = reccnt
  End Function
  Sub CreateRecordFile()
    ' This creates the tabs Records & Fields for the xlxs file which will
    ' hold all things about DATA.
    ' This routine supports both COBOL and Easytrieve syntax
    ' 
    Dim srcWords As New List(Of String)
    Dim statement As String = ""
    Dim FDFileName As String = ""
    Dim FileNameFields As String()
    Dim FileName As String = ""
    Dim FileNameType As String = ""
    Dim FileNameDD As String = ""
    Dim FileNameIndex As Integer = 0
    Dim RecordNameFields As String()
    Dim recordName As String = ""
    Dim recordNameIndex As Integer = 0
    Dim recordNameLevel As String = ""
    Dim recordNameOpenMode As String = ""
    Dim recordNameRecFM As String = ""
    Dim recordNameMinLrecl As String = ""
    Dim recordNameMaxLrecl As String = ""
    Dim recordNameOrganization As String = ""
    Dim CopybookName As String = ""
    Dim recordLength As Integer = 0
    Dim verb As String = ""
    ListOfFiles.Clear()
    ListOfRecordNames.Clear()
    ListOfRecords.Clear()
    List_Fields.Clear()
    ListOfFields.Clear()
    '
    '
    ' Process each program module in this source file
    '
    For Each pgm In listOfPrograms
      pgmName = pgm.ProgramId

      ListOfFiles = GetListOfFiles()
      'now with list of files, go thru the list getting the recordname(s), copybook(s)
      For Each file In ListOfFiles
        FileNameFields = file.Split(Delimiter)
        FileName = FileNameFields(0)
        FileNameDD = FileNameFields(1)
        FileNameType = FileNameFields(2)
        FileNameIndex = Val(FileNameFields(3))
        Select Case FileNameType
          Case "FILE"
            ListOfRecordNames = GetListOfRecordNamesFILE(FileName, FileNameIndex)
          Case "SQL"
            ListOfRecordNames = GetListOfRecordNamesSQL(FileName, FileNameIndex)
        End Select
        For Each recname In ListOfRecordNames
          RecordNameFields = recname.Split(Delimiter)
          recordName = RecordNameFields(0)
          recordNameIndex = Val(RecordNameFields(1))
          recordNameLevel = RecordNameFields(2)
          recordNameOpenMode = RecordNameFields(3)
          recordNameRecFM = RecordNameFields(4)
          recordNameMinLrecl = RecordNameFields(5)
          recordNameMaxLrecl = RecordNameFields(6)
          recordNameOrganization = RecordNameFields(7)
          recordLength = GetRecordLength(recordNameIndex)
          CopybookName = FindCopybookName(recordNameIndex, recordName)
          'write the filename,recordname,recordname index
          ListOfRecords.Add(FileNameOnly & Delimiter &
                            pgmName & Delimiter &
                            FileName & Delimiter &
                            FileNameDD & Delimiter &
                            FileNameType & Delimiter &
                              recordName & Delimiter &
                              LTrim(Str(recordLength)) & Delimiter &
                              recordNameIndex & Delimiter &
                              recordNameLevel & Delimiter &
                              recordNameOpenMode & Delimiter &
                              recordNameRecFM & Delimiter &
                              recordNameMinLrecl & Delimiter &
                              recordNameMaxLrecl & Delimiter &
                              CopybookName & Delimiter &
                              recordNameOrganization)
          ' Write the Copybook
          Dim fields As New fieldInfo("", "", "", "", "DISPLAY", 0, 0, 0, -1, -1, "", -1, False)
          Dim FieldSeq As Integer = 0
          For Each fields In List_Fields
            If Val(fields.Level) = 0 Then
              Continue For
            End If
            Dim indent As String = Space(Val(fields.Level) - 1)
            If fields.Level = "01" Then
              FieldSeq = 0
            End If
            FieldSeq += 1
            ListOfFields.Add(FileNameOnly & Delimiter &
                                 pgmName & Delimiter &
                                 FileName & Delimiter &
                                 FileNameDD & Delimiter &
                                 FileNameType & Delimiter &
                                 recordName & Delimiter &
                                 CopybookName & Delimiter &
                                 FieldSeq & Delimiter &
                                  fields.Level & Delimiter &
                                  fields.FieldName & Delimiter &
                                  fields.Picture & Delimiter &
                                  fields.StartPos & Delimiter &
                                  fields.EndPos & Delimiter &
                                  fields.Length & Delimiter &
                                  fields.Redefines)
          Next fields
        Next recname
      Next file
    Next pgm

    '
    ' Create the Excel Records sheet.
    '
    lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly & " : Records = " & ListOfRecords.Count
    If RecordsRow = 0 Then
      RecordsWorksheet = workbook.Sheets.Add(After:=workbook.Worksheets(workbook.Worksheets.Count))
      RecordsWorksheet.Name = "Records"
      ' Write the Column Headers row
      RecordsWorksheet.Range("A1").Value = "Source"
      RecordsWorksheet.Range("B1").Value = "Program"
      RecordsWorksheet.Range("C1").Value = "File/Table"
      RecordsWorksheet.Range("D1").Value = "DD"
      RecordsWorksheet.Range("E1").Value = "Type"
      RecordsWorksheet.Range("F1").Value = "RecordName"
      RecordsWorksheet.Range("G1").Value = "Length"
      RecordsWorksheet.Range("H1").Value = "@Line"
      RecordsWorksheet.Range("I1").Value = "Level"
      RecordsWorksheet.Range("J1").Value = "Open"
      RecordsWorksheet.Range("K1").Value = "RecFM"
      RecordsWorksheet.Range("L1").Value = "FDMinLen"
      RecordsWorksheet.Range("M1").Value = "FDMaxLen"
      RecordsWorksheet.Range("N1").Value = "Copybook"
      RecordsWorksheet.Range("O1").Value = "FDOrg"
      RecordsRow = 1
    End If

    ' write the Records data

    Dim DelimText As String()
    Dim row As Integer = LTrim(Str(RecordsRow))
    Dim cnt As Integer = 0
    If ListOfRecords.Count > 0 Then
      For Each record In ListOfRecords
        cnt += 1
        RecordsRow += 1
        row = LTrim(Str(RecordsRow))
        DelimText = record.Split(Delimiter)
        If DelimText.Count >= 15 Then
          RecordsWorksheet.Range("A" & row).Value = DelimText(0)       'Source
          RecordsWorksheet.Range("B" & row).Value = DelimText(1)       'Program
          RecordsWorksheet.Range("C" & row).Value = DelimText(2)       'file/table
          RecordsWorksheet.Range("D" & row).Value = DelimText(3)       'DD
          RecordsWorksheet.Range("E" & row).Value = DelimText(4)       'Type
          RecordsWorksheet.Range("F" & row).Value = DelimText(5)       'RecordName
          RecordsWorksheet.Range("G" & row).Value = DelimText(6)       'Length
          RecordsWorksheet.Range("H" & row).Value = DelimText(7)       '@line
          RecordsWorksheet.Range("I" & row).Value = DelimText(8)       'Level
          RecordsWorksheet.Range("J" & row).Value = DelimText(9)       'Open Mode
          RecordsWorksheet.Range("K" & row).Value = DelimText(10)      'RecFM
          RecordsWorksheet.Range("L" & row).Value = DelimText(11)      'FDMinLen
          RecordsWorksheet.Range("M" & row).Value = DelimText(12)      'FDMaxLen
          RecordsWorksheet.Range("N" & row).Value = DelimText(13)      'Copybook
          RecordsWorksheet.Range("O" & row).Value = DelimText(14)      'FDOrg
        End If
        If cnt Mod 100 = 0 Then
          lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly &
            " : Records = " & ListOfRecords.Count &
            " # " & cnt
        End If
      Next

    End If
    lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly & " : Records Complete"
    '
    ' Create the Fields worksheet
    '
    lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly & " : Fields = " & ListOfFields.Count

    If FieldsRow = 0 Then
      FieldsWorksheet = workbook.Sheets.Add(After:=workbook.Worksheets(workbook.Worksheets.Count))
      FieldsWorksheet.Name = "Fields"
      ' Write the Column Headers row
      FieldsWorksheet.Range("A1").Value = "Source"
      FieldsWorksheet.Range("B1").Value = "Program"
      FieldsWorksheet.Range("C1").Value = "File/Table"
      FieldsWorksheet.Range("D1").Value = "DD"
      FieldsWorksheet.Range("E1").Value = "Type"
      FieldsWorksheet.Range("F1").Value = "RecordName"
      FieldsWorksheet.Range("G1").Value = "CopyBook"
      FieldsWorksheet.Range("H1").Value = "FieldSeq"
      FieldsWorksheet.Range("I1").Value = "Level"
      FieldsWorksheet.Range("J1").Value = "FieldName"
      FieldsWorksheet.Range("K1").Value = "Picture"
      FieldsWorksheet.Range("L1").Value = "Start"
      FieldsWorksheet.Range("M1").Value = "End"
      FieldsWorksheet.Range("N1").Value = "Length"
      FieldsWorksheet.Range("O1").Value = "Redefines"
      FieldsRow = 1
    End If
    '
    ' write the Fields data
    '
    row = LTrim(Str(FieldsRow))
    cnt = 0
    If ListOfFields.Count > 0 Then
      For Each FieldRow In ListOfFields
        cnt += 1
        FieldsRow += 1
        DelimText = FieldRow.Split(Delimiter)
        row = LTrim(Str(FieldsRow))
        If DelimText.Count >= 15 Then
          FieldsWorksheet.Range("A" & row).Value = DelimText(0)       'Source
          FieldsWorksheet.Range("B" & row).Value = DelimText(1)       'Program
          FieldsWorksheet.Range("C" & row).Value = DelimText(2)       'file/table
          FieldsWorksheet.Range("D" & row).Value = DelimText(3)       'DD
          FieldsWorksheet.Range("E" & row).Value = DelimText(4)       'Type
          FieldsWorksheet.Range("F" & row).Value = DelimText(5)       'RecordName
          FieldsWorksheet.Range("G" & row).Value = DelimText(6)       'Copybook
          FieldsWorksheet.Range("H" & row).Value = DelimText(7)       'FieldSeq
          FieldsWorksheet.Range("I" & row).Value = DelimText(8)       'Level
          FieldsWorksheet.Range("J" & row).Value = DelimText(9)       'Fieldname
          FieldsWorksheet.Range("K" & row).Value = DelimText(10)      'Picture
          FieldsWorksheet.Range("L" & row).Value = DelimText(11)      'Start
          FieldsWorksheet.Range("M" & row).Value = DelimText(12)      'End
          FieldsWorksheet.Range("N" & row).Value = DelimText(13)      'Length
          FieldsWorksheet.Range("O" & row).Value = DelimText(14)      'Redefines fieldnames
        End If
        If cnt Mod 100 = 0 Then
          lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly &
            " : Fields = " & ListOfFields.Count &
            " # " & cnt
        End If
      Next
    End If
    lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly & " : Fields Complete"

  End Sub
  Sub FormatWorksheets()
    If SummaryRow > 0 Then
      rngSummaryName = SummaryWorksheet.Range("A1:A1")
      rngSummaryName.Font.Bold = True
      rngSummaryName.Font.Size = 16
      rngSummaryName = SummaryWorksheet.Range("A1:B" & LTrim(Str(SummaryRow)))
      rngSummaryName.Columns.AutoFit()
      rngSummaryName.Rows.AutoFit()
    End If

    ' Format the Programs sheet - first row bold the columns
    If ProgramsRow > 1 Then
      Dim row As Integer = LTrim(Str(ProgramsRow))
      ' Format the Sheet - first row bold the columns
      rngRecordName = worksheet.Range("A1:U1")
      rngRecordName.Font.Bold = True
      ' data area autofit all columns
      rngRecordName = worksheet.Range("A1:U" & row)
      'rngRecordName.AutoFilter()
      workbook.Worksheets("Programs").Range("A1").AutoFilter
      rngRecordName.Columns.AutoFit()
      ' ignore error flag that numbers being loaded into a text field
      objExcel.ErrorCheckingOptions.NumberAsText = False
      'worksheet.Select(1)
      'worksheet.Activate()
      'worksheet.FreezePanes(2, 1)
    End If

    ' Format the Records Sheet - first row bold the columns
    If RecordsRow > 0 Then
      Dim row As Integer = LTrim(Str(RecordsRow))
      rngRecordsName = RecordsWorksheet.Range("A1:O1")
      rngRecordsName.Font.Bold = True
      ' data area autofit all columns
      rngRecordsName = RecordsWorksheet.Range("A1:O" & row)
      workbook.Worksheets("Records").Range("A1").AutoFilter
      rngRecordsName.Columns.AutoFit()
      ' ignore error flag that numbers being loaded into a text field
      objExcel.ErrorCheckingOptions.NumberAsText = False
      'RecordsWorksheet.Select(2)
      'RecordsWorksheet.Activate()
      'RecordsWorksheet.FreezePanes(2, 1)
    End If

    ' Format the Fields worksheet - first row bold the columns
    If FieldsRow > 0 Then
      Dim row As Integer = LTrim(Str(FieldsRow))
      rngFieldsName = FieldsWorksheet.Range("A1:O1")
      rngFieldsName.Font.Bold = True
      ' data area autofit all columns
      rngFieldsName = FieldsWorksheet.Range("A1:O" & row)
      workbook.Worksheets("Fields").Range("A1").AutoFilter
      rngFieldsName.Columns.AutoFit()
      ' ignore error flag that numbers being loaded into a text field
      objExcel.ErrorCheckingOptions.NumberAsText = False
      'FieldsWorksheet.Select(3)
      'FieldsWorksheet.Activate()
      'FieldsWorksheet.FreezePanes(2, 1)
    End If

    If CommentsRow > 0 Then
      Dim row As Integer = LTrim(Str(CommentsRow))
      ' Format the Sheet - first row bold the columns
      rngComments = CommentsWorksheet.Range("A1:F1")
      rngComments.Font.Bold = True
      ' data area autofit all columns
      rngComments = CommentsWorksheet.Range("A1:F" & row)
      workbook.Worksheets("Comments").Range("A1").AutoFilter
      rngComments.Columns.AutoFit()
      rngComments.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
      ' ignore error flag that numbers being loaded into a text field
      objExcel.ErrorCheckingOptions.NumberAsText = False
      'CommentsWorksheet.Select(4)
      'CommentsWorksheet.Activate()
      'CommentsWorksheet.FreezePanes(2, 1)
    End If

    If EXECSQLRow > 0 Then
      Dim row As Integer = LTrim(Str(EXECSQLRow))
      ' Format the Sheet - first row bold the columns
      rngEXECSQL = EXECSQLWorksheet.Range("A1:G1")
      rngEXECSQL.Font.Bold = True
      ' data area autofit all columns
      rngEXECSQL = EXECSQLWorksheet.Range("A1:G" & row)
      workbook.Worksheets("ExecSQL").Range("A1").AutoFilter
      rngEXECSQL.Columns.AutoFit()
      rngEXECSQL.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
      ' ignore error flag that numbers being loaded into a text field
      objExcel.ErrorCheckingOptions.NumberAsText = False
      'CommentsWorksheet.Select(4)
      'CommentsWorksheet.Activate()
      'CommentsWorksheet.FreezePanes(2, 1)
    End If


    SummaryWorksheet.Select(1)
    SummaryWorksheet.Activate()

  End Sub
  Sub CreateCommentsWorksheet()
    '* Create the Comments worksheet from the listofcomments array
    lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly & " : Comments = " & ListOfComments.Count

    Dim cnt As Integer = 0
    Dim row As Integer = 0
    If CommentsRow = 0 Then
      CommentsWorksheet = workbook.Sheets.Add(After:=workbook.Worksheets(workbook.Worksheets.Count))
      CommentsWorksheet.Name = "Comments"
      ' Write the Column Headers row
      CommentsWorksheet.Range("A1").Value = "Source"
      CommentsWorksheet.Range("B1").Value = "Program"
      CommentsWorksheet.Range("C1").Value = "Type"
      CommentsWorksheet.Range("D1").Value = "Division"
      CommentsWorksheet.Range("E1").Value = "Line#"
      CommentsWorksheet.Range("F1").Value = "Comment"
      CommentsRow = 1
    End If
    ' load comments to spreadsheet. Merge sequential lines numbers into one row/cell
    Dim prevLineNum As Integer = -1
    Dim currLineNum As Integer = 0
    Dim prevProgram As String = ""
    Dim currProgram As String = ""
    For Each comment In ListOfComments
      cnt += 1
      Dim commentColumns As String() = comment.Split(Delimiter)
      currLineNum = Val(commentColumns(4))
      currProgram = commentColumns(1)
      If currLineNum - 1 <> prevLineNum Or currProgram <> prevProgram Then
        CommentsRow += 1
        row = LTrim(Str(CommentsRow))
        CommentsWorksheet.Range("A" & row).Value = commentColumns(0)       'Source
        CommentsWorksheet.Range("B" & row).Value = commentColumns(1)       'Program
        CommentsWorksheet.Range("C" & row).Value = commentColumns(2)       'TYPE
        CommentsWorksheet.Range("D" & row).Value = commentColumns(3)       'division
        CommentsWorksheet.Range("E" & row).Value = commentColumns(4)       'Line#
        CommentsWorksheet.Range("F" & row).Value = commentColumns(5)       'Comment
      Else
        CommentsWorksheet.Range("F" & row).Value &= vbNewLine & commentColumns(5)
      End If
      prevLineNum = currLineNum
      prevProgram = currProgram
      If cnt Mod 100 = 0 Then
        lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly &
          " : Comments = " & ListOfComments.Count &
          " # " & cnt
      End If
    Next
    lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly & " : Comments Complete"
  End Sub
  Sub CreateEXECSQLWorksheet()
    '* Create the ExecSQL worksheet from the listofexecsql array
    lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly & " : ExecSql = " & ListOfEXECSQL.Count

    Dim cnt As Integer = 0
    Dim row As Integer = 0
    Dim Statement As String = ""
    Dim Table As String = ""
    If EXECSQLRow = 0 Then
      EXECSQLWorksheet = workbook.Sheets.Add(After:=workbook.Worksheets(workbook.Worksheets.Count))
      EXECSQLWorksheet.Name = "ExecSQL"
      ' Write the Column Headers row
      EXECSQLWorksheet.Range("A1").Value = "Source"
      EXECSQLWorksheet.Range("B1").Value = "Program"
      EXECSQLWorksheet.Range("C1").Value = "EXECSQL"
      EXECSQLWorksheet.Range("D1").Value = "Seq"
      EXECSQLWorksheet.Range("E1").Value = "Table"
      EXECSQLWorksheet.Range("F1").Value = "Cursor"
      EXECSQLWorksheet.Range("G1").Value = "Statement"
      EXECSQLRow = 1
    End If
    ' load EXECSQL to spreadsheet.
    For Each execsql In ListOfEXECSQL
      cnt += 1
      Dim ExecSqlColumns As String() = execsql.Split(Delimiter)
      EXECSQLRow += 1
      row = LTrim(Str(EXECSQLRow))
      EXECSQLWorksheet.Range("A" & row).Value = ExecSqlColumns(0)       'Source
      EXECSQLWorksheet.Range("B" & row).Value = ExecSqlColumns(1)       'Program
      EXECSQLWorksheet.Range("C" & row).Value = ExecSqlColumns(2)       'ExecSql
      EXECSQLWorksheet.Range("D" & row).Value = ExecSqlColumns(3)       'seq
      Table = ExecSqlColumns(4).Replace(",", vbNewLine).Trim
      EXECSQLWorksheet.Range("E" & row).Value = Table
      EXECSQLWorksheet.Range("F" & row).Value = ExecSqlColumns(5)       'Cursor
      EXECSQLWorksheet.Range("G" & row).Value = AddNewLineAboutEveryNthCharacters(ExecSqlColumns(6), vbNewLine, 60) 'Statement

      If cnt Mod 100 = 0 Then
        lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly &
          " : ExecSQL = " & ListOfEXECSQL.Count &
          " # " & cnt
      End If
    Next
    lblProcessingWorksheet.Text = "Processing Worksheet: " & FileNameOnly & " : ExecSql Complete"
  End Sub
  Sub CreatePumlCOBOL(ByRef exec As String)
    ' create the flowchart (puml) file for COBOL

    Dim EndCondIndex As Integer = -1
    Dim StartCondIndex As Integer = -1
    Dim ParagraphStarted As Boolean = False
    Dim condStatement As String = ""
    Dim condStatementCR As String = ""
    Dim imperativeStatement As String = ""
    Dim imperativeStatementCR As String = ""
    Dim statement As String = ""
    Dim vwordIndex As Integer = -1
    WithinReadConditionStatement = False
    WithinReadStatement = False
    'WithinPerformWithEndPerformStatement = False
    Dim WithinQuotes As Boolean = False
    Dim IfCnt As Integer = 0


    ' Open the output file Puml 
    Dim PumlFileName = txtOutputFoldername.Text & "\" & exec & ".puml"

    ' Open and write at least one time. Not worrying (try/catch) about subsequent writes
    Try
      pumlFile = My.Computer.FileSystem.OpenTextFileWriter(PumlFileName, False)
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error opening PumlFile COBOL")
      Exit Sub
    End Try

    ' Write the top of file
    pumlFile.WriteLine("@startuml " & exec)
    pumlFile.WriteLine("header ADDILite(c), by IBM")
    pumlFile.WriteLine("title Flowchart of COBOL Program: " & exec &
                       "\nProgram Author: " & ProgramAuthor &
                       "\nDate written: " & ProgramWritten)

    For Each pgm In listOfPrograms
      pgmName = pgm.ProgramId

      For index As Integer = pgm.ProcedureDivision + 1 To pgm.EndProgram
        ' Paragraph names
        If SrcStmt(index).Substring(0, 4) <> Space(4) Then
          If SrcStmt(index).Substring(0, 1) = "*" Then
            Continue For
          End If
          Call ProcessPumlParagraph(ParagraphStarted, SrcStmt(index))
          Continue For
        End If
        WithinQuotes = False
        WithinPerformCnt = 0
        'WithinPerformWithEndPerformStatement = False

        ' break the statement in words
        Call GetSourceWords(SrcStmt(index).Trim, cWord)

        ' Analyze this statement for Logic Levels
        'Call AnalyzeLevelsCobol(cWord, lWord)

        ' Process every VERB word in this statement 
        ' Every verb should be a plum object created.

        IndentLevel = 1
        IFLevelIndex.Clear()
        'If index = 504 Then
        '  MessageBox.Show("pause")
        'End If
        For wordIndex = 0 To cWord.Count - 1
          Select Case cWord(wordIndex)
            Case "IF"
              IfCnt += 1
              IFLevelIndex.Add(wordIndex)
              Call ProcessPumlIF(wordIndex)
            Case "ELSE"
              Call ProcessPumlELSE(wordIndex, IfCnt)
            Case "END-IF"
              IfCnt -= 1
              IndentLevel -= 1
              pumlFile.WriteLine(Indent() & "endif")
              If IfCnt = 0 Then
                WithinIF = False
              End If
            Case "EVALUATE"
              Call ProcessPumlEVALUATE(wordIndex)
            Case "WHEN"
              Call ProcessPumlWHEN(wordIndex)
            Case "END-EVALUATE"
              Call ProcessPumlENDEVALUATE(wordIndex)
            Case "PERFORM"
              Call ProcessPumlPERFORM(wordIndex)
            Case "END-PERFORM"
              Call ProcessPumlENDPERFORM()
            Case "COMPUTE"
              Call ProcessPumlCOMPUTE(wordIndex)
            Case "SEARCH"
              Call ProcessPumlSEARCH(wordIndex)
            Case "READ"
              Call ProcessPumlREAD(wordIndex)
            Case "AT", "END", "NOT"
              ProcessPumlReadCondition(wordIndex)
            Case "END-READ"
              ProcessPumlENDREAD(wordIndex)
            Case "GO"
              Call ProcessPumlGOTO(wordIndex)
              If WithinIF Then
                ' if next word is available, if NOT an end-if then write the end-if 
                '   if there is an ELSE just leave it alone
                '   otherwise just write the end-if
                If wordIndex + 1 > cWord.Count - 1 Then
                  Continue For
                End If
                If cWord(wordIndex + 1) = "ELSE" Then
                  Continue For
                End If
                If cWord(wordIndex + 1) <> "END-IF" Then
                  IndentLevel -= 1
                  pumlFile.WriteLine(Indent() & "endif")
                  IfCnt -= 1
                  If IfCnt = 0 Then
                    WithinIF = False
                  End If
                  Continue For
                End If
              End If
            Case "EXEC"
              ProcessPumlEXEC(wordIndex)
            Case "DISPLAY"
              ProcessPumlDisplay(wordIndex)

            Case Else
              Dim EndIndex As Integer = 0
              Dim MiscStatement As String = ""
              Call GetStatement(wordIndex, EndIndex, MiscStatement)
              pumlFile.WriteLine(Indent() & ":" & MiscStatement.Trim & ";")
              wordIndex = EndIndex
          End Select
        Next wordIndex
        If WithinReadStatement And WithinReadConditionStatement Then
          IndentLevel -= 1
          pumlFile.WriteLine(Indent() & "endif")
        End If
        If WithinIF Or IfCnt > 0 Then
          For x As Integer = 1 To IfCnt
            IndentLevel -= 1
            pumlFile.WriteLine(Indent() & "endif")
          Next
          IfCnt = 0
          WithinIF = False
        End If
        WithinReadConditionStatement = False
        WithinReadStatement = False
        WithinIF = False
        Do Until WithinPerformCnt = 0
          Call ProcessPumlENDPERFORM()
        Loop
      Next index

      If ParagraphStarted = True Then
        pumlFile.WriteLine("end")
        ParagraphStarted = False
      End If

    Next
    pumlFile.WriteLine("@enduml")

    pumlFile.Close()
  End Sub
  Sub AnalyzeLevelsCobol(ByRef cWord As List(Of String), ByRef lWord As List(Of String))
    ' Assign a logic level number to each word. Always start at 0 until there is an IF
    lWord.Clear()
    For x As Integer = 0 To cWord.Count - 1
      lWord.Add(0)
    Next
    Dim EndIndex As Integer = 0
    Dim level As Integer = 0
    For cIndex As Integer = 0 To cWord.Count - 1
      If cWord(cIndex) = "IF" Then
        EndIndex = IndexToNextVerb(cIndex)
        For x As Integer = cIndex To EndIndex
          lWord(x) = level
        Next
        level += 1
        cIndex = EndIndex
        Continue For
      End If
      If cWord(cIndex) = "ELSE" Then
        level -= 1
        lWord(cIndex) = level
        level += 1
        Continue For
      End If
      If cWord(cIndex) = "END-IF" Then
        level -= 1
        lWord(cIndex) = level
        Continue For
      End If
      lWord(cIndex) = level
    Next
  End Sub
  Sub CreatePumlEasytrieve(ByRef exec As String)
    ' create the flowchart (puml) file for Easytrieve
    Dim EndCondIndex As Integer = -1
    Dim StartCondIndex As Integer = -1
    Dim ParagraphStarted As Boolean = False
    Dim condStatement As String = ""
    Dim condStatementCR As String = ""
    Dim imperativeStatement As String = ""
    Dim imperativeStatementCR As String = ""
    Dim statement As String = ""
    Dim vwordIndex As Integer = -1
    Dim ifcnt As Integer = 0

    ' Open the output file Puml 
    Dim PumlFileName = txtOutputFoldername.Text & "\" & exec & ".puml"

    ' Open and write at least one time. Not worrying (try/catch) about subsequent writes
    Try
      pumlFile = My.Computer.FileSystem.OpenTextFileWriter(PumlFileName, False)
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error opening PumlFile for Easytrieve")
      Exit Sub
    End Try

    ' Write the top of file
    pumlFile.WriteLine("@startuml " & exec)
    pumlFile.WriteLine("header ADDILite(c), by IBM")
    pumlFile.WriteLine("title Flowchart of Easytrieve Program: " & exec)

    For Each pgm In listOfPrograms
      pgmName = pgm.ProgramId

      For index As Integer = pgm.ProcedureDivision To pgm.EndProgram
        ' skip comments
        If SrcStmt(index).Length >= 1 Then
          If SrcStmt(index).Substring(0, 1) = "*" Then
            Continue For
          End If
        End If

        ' break the statement in words
        Call GetSourceWords(SrcStmt(index).Trim, cWord)

        ' Process every VERB word in this statement 
        ' Every verb should/could be a plum object created.

        'IndentLevel = 1
        'IFLevelIndex.Clear()

        For wordIndex = 0 To cWord.Count - 1
          Select Case cWord(wordIndex)
            Case "JOB"
              Call ProcessPumlParagraphEasytrieve(ParagraphStarted, cWord(wordIndex))
            Case "INPUT"
              Call ProcessPumlInput(wordIndex)
            Case "SORT"
              Call ProcessPumlSortEasytrieve(wordIndex)
            Case "START"
              Call ProcessPumlStart(wordIndex)
            Case "FINISH"
              Call ProcessPumlStart(wordIndex)
            Case "IF"
              IFLevelIndex.Add(wordIndex)
              Call ProcessPumlIFEasytrieve(wordIndex)
            Case "ELSE"
              Call ProcessPumlELSE(wordIndex, ifcnt)
            Case "END-IF"
              IndentLevel -= 1
              pumlFile.WriteLine(Indent() & "endif")
            Case "CASE"
              Call ProcessPumlEVALUATE(wordIndex)
            Case "WHEN"
              Call ProcessPumlWHEN(wordIndex)
            Case "END-EVALUATE"
              Call ProcessPumlENDEVALUATE(wordIndex)
            'Case "PERFORM"
            '  Call ProcessPumlPERFORM(wordIndex)
            'Case "END-PERFORM"
            '  Call ProcessPumlENDPERFORM(wordIndex)
            Case "DO"
              Call ProcessPumlDO(wordIndex)
            Case "END-DO"
              Call ProcessPumlENDDO(wordIndex)
            Case "COMPUTE"
              Call ProcessPumlCOMPUTE(wordIndex)
            'Case "READ"
            '  Call ProcessPumlREAD(wordIndex)
            Case "GET"
              Call ProcessPumlGET(wordIndex)
            'Case "AT", "END", "NOT"
            '  ProcessPumlReadCondition(wordIndex)
            'Case "END-READ"
            '  ProcessPumlENDREAD(wordIndex)
            Case "GO"
              Call ProcessPumlGOTO(wordIndex)
            Case "EXEC"
              ProcessPumlEXEC(wordIndex)
            Case "STOP", "END-PROC"
              pumlFile.WriteLine("end")
              pumlFile.WriteLine("")
              ParagraphStarted = False
            Case "REPORT"
              Call ProcessPumlParagraphEasytrieve(ParagraphStarted, cWord(wordIndex + 1))
              wordIndex = cWord.Count - 1
            Case "SEQUENCE"
              wordIndex = cWord.Count - 1
            Case "CONTROL"
              wordIndex = cWord.Count - 1
            Case "TITLE"
              If cWord.Count >= 5 Then
                If cWord(1) = "1" Or cWord(1) = "2" Then
                  pumlFile.WriteLine(Indent() & ":" & cWord(4).Replace("'", "").Replace("*", "").Trim & ";")
                End If
              End If
              wordIndex = cWord.Count - 1
            Case "LINE"
              wordIndex = cWord.Count - 1
            Case Else
              If cWord(wordIndex).IndexOf(".") > -1 Then
                ' handle performed paragraph name
                If wordIndex < cWord.Count - 1 Then
                  Call ProcessPumlParagraphEasytrieve(ParagraphStarted, cWord(wordIndex))
                  wordIndex += 1
                End If
              Else
                Dim EndIndex As Integer = 0
                Dim MiscStatement As String = ""
                Call GetStatement(wordIndex, EndIndex, MiscStatement)
                pumlFile.WriteLine(Indent() & ":" & MiscStatement.Trim & ";")
                wordIndex = EndIndex
              End If
          End Select
        Next wordIndex
      Next index

      If ParagraphStarted = True Then
        pumlFile.WriteLine("end")
        ParagraphStarted = False
      End If

    Next
    pumlFile.WriteLine("@enduml")

    pumlFile.Close()
  End Sub

  Sub CreateBRCOBOL(ByRef exec As String)
    ' create the Business Rules file for COBOL

    Dim EndCondIndex As Integer = -1
    Dim StartCondIndex As Integer = -1
    Dim IFStarted As Boolean = False
    Dim ELSEStarted As Boolean = False
    Dim condStatement As String = ""
    Dim condStatementCR As String = ""
    Dim imperativeStatementTrue As String = ""
    Dim imperativeStatementFalse = ""
    Dim imperativeStatementCR As String = ""
    Dim statement As String = ""
    Dim IfStatement As String = ""
    Dim IFLevel As Integer = 0
    Dim vwordIndex As Integer = -1
    Dim RuleNumber As Integer = 0
    Dim wordIndex As Integer = -1

    ' Open the output file BR
    Dim BRFileName = txtOutputFoldername.Text & "\" & exec & "_BR.csv"

    ' Open and write at least one time. Not worrying (try/catch) about subsequent writes
    Try
      swBRFile = My.Computer.FileSystem.OpenTextFileWriter(BRFileName, False)
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error opening BRFile COBOL")
      Exit Sub
    End Try

    ' Write the top of file (row 1 header)
    swBRFile.WriteLine("Source" & Delimiter &
                       "Program" & Delimiter &
                       "Rule#" & Delimiter &
                       "Level" & Delimiter &
                       "Condition" & Delimiter &
                       "Then when true" & Delimiter &
                       "Else when false")

    For Each pgm In listOfPrograms
      pgmName = pgm.ProgramId

      For index As Integer = pgm.ProcedureDivision + 1 To pgm.EndProgram
        ' ignore comments
        If SrcStmt(index).Substring(0, 1) = "*" Then
          Continue For
        End If
        ' ignore paragraphs
        If SrcStmt(index).Length >= 1 Then
          If SrcStmt(index).Substring(0, 1) <> Space(1) Then
            Continue For
          End If
        End If

        ' break the statement in words
        Call GetSourceWords(SrcStmt(index).Trim, cWord)

        ' Process every VERB word in this statement 
        ' Every IF / Evaluate verb should be a Business Rule created.

        BRLevel = 1
        IFLevel = 0
        IFStarted = False
        ELSEStarted = False
        imperativeStatementTrue = ""
        imperativeStatementFalse = ""
        For wordIndex = 0 To cWord.Count - 1
          Select Case cWord(wordIndex)
            Case "IF"
              IFLevel += 1
              RuleNumber += 1
              BRLevel += 1
              IFStarted = True
              imperativeStatementTrue = ""
              Dim EndIndex As Integer = 0
              'to next cobol verb and next index
              Call GetStatement(wordIndex, EndIndex, IfStatement)
              wordIndex = EndIndex

            Case "ELSE"
              imperativeStatementFalse = ""
              ELSEStarted = True

            Case "END-IF"
              IFLevel -= 1
              IFStarted = False
              ELSEStarted = False
              swBRFile.WriteLine(FileNameOnly & Delimiter &
                                 pgm.ProgramId & Delimiter &
                                 LTrim(Str(index)) & Delimiter &
                                 LTrim(Str(wordIndex)) & Delimiter &
                                 IfStatement & Delimiter &
                                 imperativeStatementTrue & Delimiter &
                                 imperativeStatementFalse)
              imperativeStatementTrue = ""
              imperativeStatementFalse = ""
              BRLevel = 1

            Case Else
              If IFStarted Then
                Dim EndIndex As Integer = 0
                Dim MiscStatement As String = ""
                Call GetStatement(wordIndex, EndIndex, MiscStatement)
                If ELSEStarted Then
                  imperativeStatementFalse &= MiscStatement & " "
                Else
                  imperativeStatementTrue &= MiscStatement & " "
                End If
                wordIndex = EndIndex
              End If
          End Select
        Next wordIndex
        If IFStarted = True Then
          IFStarted = False
          ELSEStarted = False
          IFLevel = -1
          swBRFile.WriteLine(FileNameOnly & Delimiter &
                                 pgm.ProgramId & Delimiter &
                                 LTrim(Str(index)) & Delimiter &
                                 LTrim(Str(wordIndex)) & Delimiter &
                                 IfStatement & Delimiter &
                                 imperativeStatementTrue & Delimiter &
                                 imperativeStatementFalse)
        End If

      Next index


    Next pgm

    swBRFile.Close()
  End Sub

  Function GetListOfFiles() As List(Of String)
    ' Scan through the stmt array looking for all data "FILES"
    '   A "FILE" is something stated with either "SELECT" or "EXEC SQL DECLARE"
    '  Store also the DDName and indicate if FILE or SQL
    ' in format of: Filename,DDName,FILE,index
    '           or: Tablename,,SQL,index
    Dim statement As String = ""
    Dim FDFileName As String = ""
    Dim srcWords As New List(Of String)
    ListOfFiles.Clear()
    For stmtIndex As Integer = pgm.EnvironmentDivision + 1 To pgm.ProcedureDivision - 1
      statement = SrcStmt(stmtIndex)
      If statement.Length >= 1 Then
        If statement.Substring(0, 1) = "*" Then
          Continue For
        End If
      End If
      Call GetSourceWords(statement, srcWords)
      Select Case SourceType
        Case "COBOL"
          If srcWords(0) = "SELECT" Then
            Dim file_name_1 As String = ""
            If srcWords(1).Equals("OPTIONAL") Then
              FDFileName = srcWords(2)
            Else
              FDFileName = srcWords(1)
            End If
            For x As Integer = 0 To srcWords.Count - 1
              If srcWords(x) = "ASSIGN" Then
                If srcWords(x + 1) <> "TO" Then
                  DDName = srcWords(x + 1)
                Else
                  If x + 2 <= srcWords.Count - 1 Then
                    DDName = srcWords(x + 2)
                  End If
                End If
              End If
            Next
            ListOfFiles.Add(srcWords(1) & Delimiter &
                            DDName & Delimiter &
                            "FILE" & Delimiter &
                            LTrim(Str(stmtIndex)))
            Continue For
          End If
          If srcWords.Count >= 5 Then
            If srcWords(0) = "EXEC" And
              srcWords(1) = "SQL" And
              srcWords(2) = "DECLARE" And
              srcWords(4) = "TABLE" Then
              ListOfFiles.Add(srcWords(3) & Delimiter &
                            "" & Delimiter &
                            "SQL" & Delimiter &
                            LTrim(Str(stmtIndex)))
            End If
          End If
        Case "Easytrieve"
          If srcWords(0) = "FILE" Then
            FDFileName = srcWords(1)
            DDName = srcWords(1)
            ListOfFiles.Add(FDFileName & Delimiter &
                            DDName & Delimiter &
                            "FILE" & Delimiter &
                            LTrim(Str(stmtIndex)))
            Continue For
          End If
          If srcWords.Count >= 5 Then
            If srcWords(0) = "EXEC" And
              srcWords(1) = "SQL" And
              srcWords(2) = "DECLARE" And
              srcWords(4) = "TABLE" Then
              FDFileName = srcWords(3)
              Dim tableParameters As String() = FDFileName.Split(".")
              If tableParameters.Count = 2 Then
                FDFileName = tableParameters(1)
              Else
                FDFileName = tableParameters(0)
              End If
              ListOfFiles.Add(FDFileName & Delimiter &
                            "" & Delimiter &
                            "SQL" & Delimiter &
                            LTrim(Str(stmtIndex)))
            End If
          End If

      End Select
    Next
    GetListOfFiles = ListOfFiles
  End Function
  Function GetListOfRecordNamesFILE(ByRef filename As String, ByRef FileIndex As Integer) As List(Of String)
    ' Use the Data Division index to search Stmt array to get the FD record location,
    ' then the next stmt record is the "01-Level" with the File's record name.
    ' Could be multiple 01-levels(record names) for this file
    ' 0-Record Name,
    ' 1-index to record name
    ' 2-Level (FD or SD)
    ' 3-OpenMode (I,O)
    ' 4-recfm (F or V)
    ' 5-minlrecl
    ' 6-maxlrecl
    ' 7-organization
    Dim FDWords As New List(Of String)
    Dim ListOfRecordNames As New List(Of String)
    Dim FDDetails As String()
    Dim FDDetail As String = ""
    Dim FDDetaillevel As String = ""
    Dim FDDetailOpenMode As String = ""
    Dim FDDetailrecfm As String = ""
    Dim FDDetailminLrecl As String = ""
    Dim FDDetailmaxLrecl As String = ""
    Dim FDDetailorganization As String = ""
    Dim recname As String = ""
    Dim FDRecName As String = ""      'first FD 01-level
    For FDIndex As Integer = pgm.DataDivision To pgm.ProcedureDivision
      Call GetSourceWords(SrcStmt(FDIndex), FDWords)
      If FDWords.Count >= 1 Then
        If FDWords(0) = "FD" Or FDWords(0) = "SD" Or FDWords(0) = "FILE" Then
          If FDWords(1) = filename Then
            Call GetSourceWords(SrcStmt(FileIndex), cWord)
            FDDetail = GetFDDetails(cWord)
            FDDetails = FDDetail.Split(Delimiter)
            FDDetaillevel = FDDetails(4)
            FDDetailOpenMode = FDDetails(5)
            FDDetailrecfm = FDDetails(6)
            FDDetailminLrecl = FDDetails(7)
            FDDetailmaxLrecl = FDDetails(8)
            FDDetailorganization = FDDetails(10)
            ' loop for all Records (01-Levels)'s related to this filename
            For recIndex = FDIndex + 1 To pgm.ProcedureDivision
              Call GetSourceWords(SrcStmt(recIndex), FDWords)
              Select Case FDWords(0)
                Case "FD", "SD", "WORKING-STORAGE", "LOCAL-STORAGE", "LINKAGE"
                  Exit For
                Case "01", "FILE"
                  recname = FDWords(1).Replace(".", "").Trim
                  If FDRecName.Length = 0 Then
                    FDRecName = recname
                  End If
                  ListOfRecordNames.Add(recname & Delimiter &
                                     LTrim(Str(recIndex)) & Delimiter &
                                     FDDetaillevel & Delimiter &
                                     FDDetailOpenMode & Delimiter &
                                     FDDetailrecfm & Delimiter &
                                     FDDetailminLrecl & Delimiter &
                                     FDDetailmaxLrecl & Delimiter &
                                     FDDetailorganization)
              End Select
            Next
            Exit For
          End If
        End If
      End If
    Next
    Select Case SourceType
      Case "COBOL"
        ' now that I found the record name(s) for the FD, I now need to search
        ' the procedure area for (a) the file name and (b) FD record name(s).
        ' (a) File name will be for the READ (INTO) verb
        ' (b) FD Record name will be for WRITE verb
        ' I need to link all the record names in order to find any copybooks associated with 
        ' the record name(s)
        '
        ' looking for READ verb with file name to see if any INTO option
        '
        Dim recordLength As Integer = -1
        Dim recordNameIndex As Integer = -1
        ' check for READ verb
        Dim ReadDetails As New List(Of String)
        For ReadIndex As Integer = pgm.ProcedureDivision To pgm.EndProgram
          If SrcStmt(ReadIndex).Substring(0, 1) = "*" Then
            Continue For
          End If
          If SrcStmt(ReadIndex).IndexOf("READ " & filename) = -1 Then
            Continue For
          End If
          Call GetSourceWords(SrcStmt(ReadIndex), ReadDetails)
          For ReadLocation As Integer = 0 To ReadDetails.Count - 1
            If ReadDetails(ReadLocation) = "READ" Then
              If ReadDetails(ReadLocation + 1) = filename Then
                recname = FindReadRecordName(ReadLocation, ReadDetails)
                If recname.Length > 0 Then
                  If ListOfReadIntoRecords.IndexOf(recname) = -1 Then
                    recordNameIndex = FindWSRecordNameIndex(pgm.DataDivision, recname)
                    ' if 01 Recname not found, skip this record name
                    If recordNameIndex = -1 Then
                      Continue For
                    End If
                    recordLength = GetRecordLength(recordNameIndex)
                    ListOfReadIntoRecords.Add(recname)
                    ListOfRecordNames.Add(recname & Delimiter &
                                      LTrim(Str(recordNameIndex)) & Delimiter &
                                      "WS" & Delimiter &
                                      "INPUT" & Delimiter &
                                      "F" & Delimiter &
                                      LTrim(Str(recordLength)) & Delimiter &
                                      "0" & Delimiter &
                                      "SEQUENTIAL")
                    ReadLocation += 2
                  End If
                End If
              End If
            End If
          Next ReadLocation
        Next ReadIndex

        ' Check for WRITE verb, need to find the FD's Record name (ie Write FDrecname)
        Dim WriteDetails As New List(Of String)
        For WriteIndex As Integer = pgm.ProcedureDivision + 1 To pgm.EndProgram
          If SrcStmt(WriteIndex).Substring(0, 1) = "*" Then
            Continue For
          End If
          If SrcStmt(WriteIndex).IndexOf("WRITE " & FDRecName) = -1 Then
            Continue For
          End If
          Call GetSourceWords(SrcStmt(WriteIndex), WriteDetails)
          For WriteLocation As Integer = 0 To WriteDetails.Count - 1
            If WriteDetails(WriteLocation) = "WRITE" Then
              If WriteDetails(WriteLocation + 1) = FDRecName Then
                recname = FindWriteRecordName(WriteLocation, WriteDetails) '*here
                If recname.Length > 0 Then
                  If ListOfWriteFromRecords.IndexOf(recname) = -1 Then
                    recordNameIndex = FindWSRecordNameIndex(pgm.DataDivision, recname)
                    ' if 01 Recname not found, skip this record name
                    If recordNameIndex = -1 Then
                      Continue For
                    End If
                    recordLength = GetRecordLength(recordNameIndex)
                    ListOfWriteFromRecords.Add(recname)
                    ListOfRecordNames.Add(recname & Delimiter &
                                      LTrim(Str(recordNameIndex)) & Delimiter &
                                      "WS" & Delimiter &
                                      "OUTPUT" & Delimiter &
                                      "F" & Delimiter &
                                      LTrim(Str(recordLength)) & Delimiter &
                                      "0" & Delimiter &
                                      "SEQUENTIAL")
                    WriteLocation += 2
                  End If
                End If
              End If
            End If
          Next WriteLocation
        Next WriteIndex
      Case "Easytrieve"
        'Easytrieve does not have the concept of multiple records
    End Select

    GetListOfRecordNamesFILE = ListOfRecordNames
  End Function
  Function GetListOfRecordNamesSQL(ByRef filename As String, ByRef DeclareIndex As Integer) As List(Of String)
    ' starting at the DeclareIndex look for the record name (01 level)
    ' return: record name,index,level,recfm,minlrecl,maxlrecl,organization

    ' Presume that SQL only has one Record name (01 level)

    Dim SQLRecordNames As New List(Of String)
    Dim SelectWords As New List(Of String)

    Dim recnameOpenMode = GetOpenModeSQL(filename)
    For dataIndex As Integer = DeclareIndex + 1 To pgm.ProcedureDivision
      Call GetSourceWords(SrcStmt(dataIndex), SelectWords)
      If SelectWords(0) = "01" Then
        Dim recname As String = SelectWords(1)
        'getrecordlength(recname)
        SQLRecordNames.Add(recname & Delimiter &
                              LTrim(Str(dataIndex)) & Delimiter &
                              "SQL" & Delimiter &
                              recnameOpenMode & Delimiter &
                              "F" & Delimiter &
                              "" & Delimiter &
                              "" & Delimiter &
                              "RDBMS")
        Exit For
      End If
    Next
    GetListOfRecordNamesSQL = SQLRecordNames
  End Function
  Function GetRecordLength(ByRef StartOf01 As Integer) As Integer
    ' To get the record length we get the start and ending index of the 01-level
    ' Note! the List_fields (global) array is updated & returned here!
    GetRecordLength = 0
    If SourceType <> "COBOL" Then
      Exit Function
    End If
    Dim DataWords As New List(Of String)
    Dim EndOf01 As Integer = -1
    ' Now find the end of this 01-level
    For stmtindex As Integer = StartOf01 + 1 To pgm.ProcedureDivision
      Call GetSourceWords(SrcStmt(stmtindex), DataWords)
      Select Case DataWords(0)
        Case "FD", "01", "LINKAGE", "PROCEDURE", "FILE", "SQL", "JOB", "SORT"
          EndOf01 = stmtindex - 1
          Exit For
      End Select
    Next
    If EndOf01 = -1 Then
      MessageBox.Show("Not able to find End of 01, start of 01-Level@" & StartOf01 &
                      ", Pgm=" & pgm.ProgramId)
      GetRecordLength = 0
      Exit Function
    End If
    ' Now I have the start and end indexes of this 01-level

    ' To figure out the length of this 01-level, we got to figure out the length
    '  of all the fields.
    ' Load the fields to array so we can compute lengths, etc.
    Dim fieldWords As New List(Of String)
    'Dim fields As New RecordInfo

    Dim pic As Integer = -1
    Dim usageIndex As Integer = -1
    Dim occurMinIndex As Integer = -1
    Dim occurMaxIndex As Integer = -1
    Dim dependOnIndex As Integer = -1
    Dim withinRedefines As Boolean = False
    Dim redefinesLevel As String = "00"

    List_Fields.Clear()
    Dim currentParent As Integer = -1
    For fieldIndex As Integer = StartOf01 To EndOf01
      ' ignore any comments
      If SrcStmt(fieldIndex).Substring(0, 1) = "*" Then
        Continue For
      End If

      Call GetSourceWords(SrcStmt(fieldIndex), fieldWords)
      If fieldWords.Count >= 1 Then
        If Not IsNumeric(fieldWords(0)) Then
          Continue For
        End If
      End If
      Dim fields As New fieldInfo("", "", "", "", "DISPLAY", 0, 0, 0, -1, -1, "", -1, -1)

      ' level clause
      fields.Level = Microsoft.VisualBasic.Right("000" & fieldWords(0), 2)
      If fieldWords(0) = "01" Then
        fields.StartPos = 1
        fields.Parent = -1
      End If

      ' field name clause
      fields.FieldName = fieldWords(1)

      ' Redefines Clause
      If fieldWords.IndexOf("REDEFINES") > -1 Then
        fields.Redefines = fieldWords(3)
        'find the main definition field
        For searchIdx As Integer = List_Fields.Count - 1 To 0 Step -1
          Dim searchField = List_Fields(searchIdx)
          If searchField.FieldName = fields.Redefines Then
            fields.RedefField = searchIdx
            Exit For
          End If
        Next
        If fields.RedefField = -1 Then
          LogFile.WriteLine(Date.Now & ",Redefine parent not found!," & fields.FieldName & "[" & FileNameOnly & "]")
        End If
        withinRedefines = True
        redefinesLevel = fields.Level
      Else
        fields.Redefines = ""
        If withinRedefines And fields.Level > redefinesLevel Then
          fields.RedefField = 1
        Else
          fields.RedefField = 0
          withinRedefines = False
        End If
      End If

      ' picture clause
      pic = fieldWords.IndexOf("PIC")
      If pic = -1 Then
        pic = fieldWords.IndexOf("PICTURE")
      End If
      If pic = -1 Then
        fields.Picture = ""
      Else
        If fieldWords(pic + 1) = "IS" Then
          pic += 2
        Else
          pic += 1
        End If
        fields.Picture = fieldWords(pic)
      End If


      ' Usage clause (ie, COMP, COMP-3, DISPLAY, etc.
      fields.Usage = "DISPLAY"
      If pic > -1 And pic < (fieldWords.Count - 1) Then
        usageIndex = fieldWords.IndexOf("USAGE")
        If usageIndex > -1 Then
          If fieldWords(usageIndex + 1) = "IS" Then
            usageIndex += 2
          Else
            usageIndex += 1
          End If
          fields.Usage = fieldWords(usageIndex)
        Else
          If pic > -1 And pic <= (fieldWords.Count - 1) Then
            usageIndex = pic + 1
            If fieldWords(usageIndex) <> "VALUE" Then
              If List_Usage.IndexOf(fieldWords(usageIndex)) > -1 Then
                fields.Usage = fieldWords(usageIndex)
              End If
            End If
          End If
        End If
      End If
      ' Occurs clause
      occurMinIndex = fieldWords.IndexOf("OCCURS")
      If occurMinIndex > -1 Then
        fields.OccursMinimumTimes = fieldWords(occurMinIndex + 1)
        occurMaxIndex = fieldWords.IndexOf("TO")
        If occurMaxIndex > -1 Then
          fields.OccursMaximumTimes = fieldWords(occurMaxIndex + 1)
        End If
        If fields.OccursMaximumTimes >= 0 Then
          dependOnIndex = fieldWords.IndexOf("DEPENDING")
          If dependOnIndex > -1 Then
            If fieldWords(dependOnIndex + 1) = "ON" Then
              dependOnIndex += 1
            End If
            fields.DependingOn = fieldWords(dependOnIndex + 1)
          End If
        End If
      End If
      ' Assign the length
      If fields.Picture.Length > 0 Then
        fields.Length = DetermineDigits(fields.Picture)
        Select Case fields.Usage
          Case "COMP", "COMPUTATIONAL", "BINARY", "COMP-5", "COMPUTATIONAL-5"
            Select Case fields.Length
              Case 1 To 4
                fields.Length = 2
              Case 5 To 9
                fields.Length = 4
              Case 10 To 18
                fields.Length = 8
            End Select
          Case "COMP-3", "COMPUTATIONAL-3", "PACKED-DECIMAL"
            Dim rbytes As Integer = (fields.Length + 1) Mod 2
            fields.Length = (fields.Length + 1) \ 2
            If rbytes > 0 Then
              fields.Length += 1
            End If
        End Select
      End If
      ' Determine Parent. decrement an Index, check previous fields' LEVEL:
      ' if previous LEVEL value same as my LEVEL then
      '   its my sibling and thus same parent, assign my sibling's parent value to my PARENT value, exit loop
      ' if previous LEVEL value is less than my LEVEL then
      '   its my parent, so assign previous Index to my PARENT, exit loop

      For idx = List_Fields.Count - 1 To 0 Step -1
        Dim prevField As fieldInfo = List_Fields(idx)
        If prevField.Level = fields.Level Then
          fields.Parent = prevField.Parent
          Exit For
        End If
        If prevField.Level < fields.Level Then
          fields.Parent = idx
          Exit For
        End If
      Next
      ' Do not care about VALUE clause
      List_Fields.Add(fields)
    Next
    ' determine the length of the 01-level (Record)
    ' and Tally up all lengths 
    Dim totalRecordLength As Integer = 0
    Dim totalFieldLength As Integer = 0
    For Each fields In List_Fields
      If fields.Picture.Length = 0 Then
        Continue For
      End If
      'If fields.RedefField = 1 Then
      '  Continue For
      'End If
      'Dim idx = List_Fields.IndexOf(fields)
      If fields.OccursMinimumTimes > -1 Then
        totalRecordLength += fields.Length * fields.OccursMinimumTimes
        totalFieldLength = fields.Length * fields.OccursMinimumTimes
      Else
        totalRecordLength += fields.Length
        totalFieldLength = fields.Length
      End If
      ' now loop thru and add totalFieldLength to all its Parent Heritage
      Dim parIdx = fields.Parent
      Do Until parIdx = -1
        Dim parField = List_Fields(parIdx)
        parField.Length += fields.Length
        parIdx = parField.Parent
      Loop
    Next
    ' resolve the Start and End Positions
    Dim startPos As Integer = 1
    Dim endPos As Integer = -1
    Dim redefStartPos As Integer = -1
    Dim redefEndPos As Integer = -1
    For Each fields In List_Fields
      If fields.Picture.Length = 0 Then
        fields.StartPos = startPos
        fields.EndPos = fields.StartPos + fields.Length - 1
        If fields.Redefines.Length > 0 Then
          If fields.RedefField > -1 Then
            Dim mainField = List_Fields(fields.RedefField)
            redefStartPos = mainField.StartPos
            fields.StartPos = redefStartPos
            fields.EndPos = fields.StartPos + fields.Length - 1
          End If
        End If
        Continue For
      End If
      If fields.RedefField = 1 Then
        fields.StartPos = redefStartPos
        fields.EndPos = fields.StartPos + fields.Length - 1
        redefEndPos = fields.EndPos
        redefStartPos = redefEndPos + 1
      Else
        fields.StartPos = startPos
        fields.EndPos = fields.StartPos + fields.Length - 1
        endPos = fields.EndPos
        startPos = endPos + 1
      End If
    Next
    ' resolve group item lengths for its own fields
    'For Each fields In List_Fields
    '  If fields.Picture.Length > 0 Then
    '    For idx = List_Fields.IndexOf(fields) - 1 To 0 Step -1
    '      Dim searchFields As fieldInfo = List_Fields(idx)
    '      If searchFields.Picture.Length = 0 And
    '        searchFields.Level < fields.Level Then
    '        searchFields.Length += fields.Length
    '        Exit For
    '      End If
    '    Next
    '  End If
    'Next
    ' resolve group lengths that are still zero

    ' put final length on the 01-level. Messy, I know
    'For Each fields In List_Fields
    '  If fields.Level = 1 Then
    '    fields.Length = totalRecordLength
    '    Exit For
    '  End If
    'Next

    GetRecordLength = totalRecordLength
  End Function
  Function FindCopybookName(ByRef DataIndex As Integer, ByVal RecordName As String) As String
    ' Use the Data Division index to search Stmt array to get the Record  location,
    ' then look previous lines to see what the possible copybook name would be.

    Dim CopyWords As New List(Of String)
    'Dim RecordWords As New List(Of String)
    ' look upward to see if we find 'COPY/INCLUDE/SQL INCLUDE' statement
    ' here are some examples:
    '  COBOL
    '*COPY CRCALC.
    '    EXEC SQL INCLUDE SQLCA END-EXEC.
    '*INCLUDE++ PM044016  (from Panvalet)
    '* Easytrieve
    '*%COPYBOOK SQL CRPVBI
    '*%COPYBOOK FILE PM044025
    ' we can stop searching at a NON-commented line.
    FindCopybookName = ""
    For CopyIndex As Integer = DataIndex - 1 To pgm.DataDivision Step -1
      FindCopybookName = FindCopyOrInclude(CopyIndex, CopyWords)
      If FindCopybookName.Length > 0 Then
        Exit For
      End If
    Next
    If FindCopybookName = "EXIT FOR" Then
      FindCopybookName = ""
    End If
    If FindCopybookName.Length > 0 Then
      Exit Function
    End If

    ' if still no copybook found at this point then maybe the copy/include is
    ' placed after the record name; so search downward. ugh.
    FindCopybookName = ""
    For CopyIndex As Integer = DataIndex + 1 To pgm.ProcedureDivision
      FindCopybookName = FindCopyOrInclude(CopyIndex, CopyWords)
      If FindCopybookName.Length > 0 Then
        Exit For
      End If
    Next
    If FindCopybookName = "EXIT FOR" Then
      FindCopybookName = ""
    End If
    If FindCopybookName.Length = 0 Then
      FindCopybookName = "NONE"
    End If
  End Function
  Sub ProcessPumlParagraph(ByRef ParagraphStarted As Boolean, ByRef statement As String)
    If ParagraphStarted = True Then
      pumlFile.WriteLine("end")
      pumlFile.WriteLine("")
    End If
    pumlFile.WriteLine("start")
    pumlFile.WriteLine(":**" & Trim(statement.Replace(".", "")) & "**;")
    ParagraphStarted = True
  End Sub
  Sub ProcessPumlParagraphEasytrieve(ByRef ParagraphStarted As Boolean, ByRef statement As String)
    'If ParagraphStarted = True Then
    '  pumlFile.WriteLine("end")
    '  pumlFile.WriteLine("")
    'End If
    pumlFile.WriteLine("start")
    pumlFile.WriteLine(":**" & statement.Trim & "**;")
    ParagraphStarted = True
    IndentLevel = 1
    IFLevelIndex.Clear()
  End Sub
  Sub ProcessPumlSortEasytrieve(ByRef WordIndex As Integer)
    ' note that cWord is global
    Dim EndIndex As Integer = cWord.Count - 1
    Dim TogetherWords As String = StringTogetherWords(WordIndex, (cWord.Count - 1))
    Dim SortStatement As String = AddNewLineAboutEveryNthCharacters(TogetherWords, ESCAPENEWLINE, 30)

    pumlFile.WriteLine("start")
    pumlFile.WriteLine(":" & SortStatement.Trim & ";")
    WordIndex = cWord.Count - 1
  End Sub
  Sub ProcessPumlIF(ByRef WordIndex As Integer)
    ' find the 'IF' aka Conditional statement
    ' Indentlevel is global
    Dim EndIndex As Integer = 0
    Dim Statement As String = ""
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlFile.WriteLine(Indent() & "if (" & Statement.Trim & ") then (yes)")
    IndentLevel += 1
    WordIndex = EndIndex
    WithinIF = True
  End Sub
  Sub ProcessBRIF(ByRef WordIndex As Integer, ByRef IFStatement As String)
    ' find the 'IF' aka Conditional statement
    ' Indentlevel is global
    Dim EndIndex As Integer = 0
    Call GetStatement(WordIndex, EndIndex, IFStatement) 'to next cobol verb
    BRLevel += 1
    WordIndex = EndIndex
    IndentLevel += 1
  End Sub
  Sub ProcessPumlIFEasytrieve(ByRef WordIndex As Integer)
    ' find the 'IF' aka Conditional statement
    ' Indentlevel is global
    Dim EndIndex As Integer = cWord.Count - 1
    Dim Statement As String = ""
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlFile.WriteLine(Indent() & "if (" & Statement.Trim & ") then (yes)")
    IndentLevel += 1
    WordIndex = EndIndex
  End Sub

  Sub ProcessPumlELSE(ByRef WordIndex As Integer, ByRef ifCnt As Integer)
    ' Does current 'ELSE' belong to my 'IF' or another 'IF'???
    ' cWord is global
    ' IndentLevel is global
    'hoping/presuming a well-formed if/else/end-if

    'Look to see if this ELSE has an IF before I hit another ELSE...
    For x = WordIndex - 1 To 0 Step -1
      If cWord(x) = "ELSE" Then
        'malformed? so add an END-IF
        pumlFile.WriteLine("floating note left: Malformed")
        IndentLevel -= 1
        pumlFile.WriteLine(Indent() & "endif")
        ifCnt -= 1
        Exit For
      End If
      If cWord(x) = "IF" Then
        Exit For
      End If
      If cWord(x) = "END-IF" Then
        Exit For
      End If
    Next
    '
    IndentLevel -= 1
    pumlFile.WriteLine(Indent() & "else (no)")
    IndentLevel += 1

  End Sub
  Sub ProcessPumlEVALUATE(ByRef wordIndex As Integer)
    'TODO: need to fix embedded Evaluates
    ' find the end of 'EVALUATE' / 'CASE' statement which should be at the first 'WHEN' clause
    'cWord is global
    'IndentLevel is global
    Dim Statement As String = ""
    Dim EndIndex As Integer = wordIndex + 1
    For EndIndex = EndIndex To cWord.Count - 1
      If cWord(EndIndex) = "WHEN" Then
        Exit For
      End If
    Next
    EndIndex -= 1
    Call GetStatement(wordIndex, EndIndex, Statement)
    pumlFile.WriteLine(Indent() & ":" & Statement.Trim & ";")
    IndentLevel += 1
    FirstWhenStatement = True
    wordIndex = EndIndex
  End Sub
  Sub ProcessPumlWHEN(ByRef wordindex As Integer)
    'TODO: need to fix embedded WHENs
    ' find the end of 'EVALUATE' statement which should be at the first 'WHEN' clause
    'cWord is global
    'IndentLevel is global
    Dim Statement As String = ""
    Dim EndIndex As Integer = wordindex + 1
    For EndIndex = EndIndex To cWord.Count - 1
      If cWord(EndIndex) = "WHEN" Then
        Exit For
      End If
    Next
    EndIndex -= 1
    Call GetStatement(wordindex, EndIndex, Statement)
    If FirstWhenStatement = True Then
      FirstWhenStatement = False
      pumlFile.WriteLine(Indent() & "if (" & Statement.Trim & ") then (yes)")
      IndentLevel += 1
    Else
      IndentLevel -= 1
      pumlFile.WriteLine(Indent() & "elseif (" & Statement.Trim & ") then (yes)")
      IndentLevel += 1
    End If
    wordindex = EndIndex

  End Sub
  Sub ProcessPumlENDEVALUATE(ByRef wordindex As Integer)
    'TODO: Need to handle embedded end-evaluate
    FirstWhenStatement = False
    IndentLevel -= 1
    pumlFile.WriteLine(Indent() & "endif")
  End Sub
  Sub ProcessPumlCOMPUTE(ByRef WordIndex As Integer)
    ' find the end of 'COMPUTE' statement
    Dim EndIndex As Integer = 0
    Dim Statement As String = ""
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlFile.WriteLine(Indent() & ":" & Statement.Trim & ";")
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlGOTO(ByRef WordIndex As Integer)
    ' find the end of 'GO TO' statement
    Dim EndIndex As Integer = 0
    Dim Statement As String = ""
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlFile.WriteLine(Indent() & "#pink:" & Statement.Trim & ";")
    pumlFile.WriteLine(Indent() & "detach")
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlGET(ByRef WordIndex As Integer)
    ' Next word is file/record name
    Dim EndIndex As Integer = WordIndex + 1
    Dim Statement As String = ""
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlFile.WriteLine(Indent() & ":" & Statement.Trim & "/")
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlREAD(ByRef WordIndex As Integer)
    'TODO: need to fix embedded READ
    ' find the end of 'READ' statement which should be at either a verb, AT, END, NOT or END-READ
    'cWord is global
    'IndentLevel is global

    Dim TogetherWords As String = ""
    Dim ReadStatement As String = ""

    ' Format 1:Sequential Read
    Dim StartIndex = WordIndex
    Dim EndIndex As Integer = -1
    Dim NextVerb As Integer = -1

    NextVerb = IndexToNextVerb(WordIndex + 1)
    If NextVerb = -1 Then
      EndIndex = cWord.Count - 1
      TogetherWords = StringTogetherWords(WordIndex, EndIndex)
      ReadStatement = AddNewLineAboutEveryNthCharacters(TogetherWords, ESCAPENEWLINE, 30)
      pumlFile.WriteLine(Indent() & ":" & ReadStatement.Trim & "/")
      WordIndex = EndIndex
      Exit Sub
    End If

    For EndIndex = WordIndex + 1 To NextVerb
      Select Case cWord(EndIndex)
        Case "AT", "END", "NOT", "END-READ"
          WithinReadStatement = True
          Exit For
      End Select
    Next
    If EndIndex > NextVerb Then
      EndIndex = NextVerb - 1
    End If
    EndIndex -= 1
    If StartIndex < 0 Or EndIndex < 0 Then
      LogFile.WriteLine(Date.Now & ",Problem with indexes at processPumlRead," & StartIndex & "/" & EndIndex)
      WordIndex = cWord.Count - 1
      Exit Sub
    End If
    ' just to be sure indexes are right.
    If StartIndex < EndIndex Then
      TogetherWords = StringTogetherWords(StartIndex, EndIndex)
      ReadStatement = AddNewLineAboutEveryNthCharacters(TogetherWords, ESCAPENEWLINE, 30)
      pumlFile.WriteLine(Indent() & ":" & ReadStatement.Trim & "/")
      WordIndex = EndIndex
      IndentLevel += 1
    Else
      LogFile.WriteLine(Date.Now & ",Problem #2 with indexes at processPumlRead," & StartIndex & "/" & EndIndex)
      WordIndex = cWord.Count - 1
    End If


    ''Format 2:random retrieval
  End Sub
  Sub ProcessPumlInput(ByRef WordIndex As Integer)
    ' find the file of 'READ' statement which should next next word
    'cWord is global
    'IndentLevel is global

    ' Format 1:Sequential Read
    WithinReadStatement = True
    Dim StartIndex = WordIndex
    Dim EndIndex As Integer = StartIndex + 1
    Dim TogetherWords As String = StringTogetherWords(StartIndex, EndIndex)
    Dim ReadStatement As String = AddNewLineAboutEveryNthCharacters(TogetherWords, ESCAPENEWLINE, 30)
    pumlFile.WriteLine(Indent() & ":" & ReadStatement.Trim & "/")
    WordIndex = EndIndex
    IndentLevel += 1
  End Sub
  Sub ProcessPumlReadCondition(ByRef WordIndex As Integer)
    If WithinReadStatement = False Then
      Exit Sub
    End If
    If WordIndex + 3 > cWord.Count - 1 Then
      Exit Sub
    End If
    If WithinReadConditionStatement = True Then
      IndentLevel -= 1
      pumlFile.WriteLine(Indent() & "endif")
    End If
    Dim ReadCondition As String = ""
    Dim ReadConditionCount As Integer = 0

    For x As Integer = WordIndex To WordIndex + 3
      Select Case cWord(x)
        Case "AT", "END", "NOT"
          ReadCondition &= cWord(x) & " "
          ReadConditionCount += 1
      End Select
    Next
    WithinReadConditionStatement = True
    pumlFile.WriteLine(Indent() & "if (" & ReadCondition.Trim & "?) then (yes)")
    IndentLevel += 1
    WordIndex += ReadConditionCount - 1
  End Sub
  Sub ProcessPumlENDREAD(ByRef WordIndex As Integer)
    If WithinReadConditionStatement = True Then
      IndentLevel -= 1
      pumlFile.WriteLine(Indent() & "endif")
      IndentLevel -= 1
    End If
    'IndentLevel -= 1
    WithinReadConditionStatement = False
    WithinReadStatement = False
  End Sub
  Sub ProcessPumlSEARCH(ByRef WordIndex As Integer)
    ' for now just one big block for the search statement
    ' find end of statement or the END-SEARCH phrase
    Dim WordsTogether As String = ""
    Dim EndIndex As Integer = 0
    Dim Statement As String = ""
    For EndIndex = WordIndex + 1 To cWord.Count - 1
      If cWord(EndIndex) = "END-SEARCH" Then
        WordsTogether = StringTogetherWords(WordIndex, EndIndex)
        Statement = AddNewLineAboutEveryNthCharacters(WordsTogether, ESCAPENEWLINE, 30)
        pumlFile.WriteLine(Indent() & ":" & Statement.Trim & ";")
        WordIndex = EndIndex
        Exit Sub
      End If
    Next
    ' keyword END-SEARCH not found set to end of statement
    EndIndex = cWord.Count - 1
    WordsTogether = StringTogetherWords(WordIndex, EndIndex)
    Statement = AddNewLineAboutEveryNthCharacters(WordsTogether, ESCAPENEWLINE, 30)
    pumlFile.WriteLine(Indent() & ":" & Statement.Trim & ";")
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlPERFORM(ByRef WordIndex As Integer)
    ' Need to find the end of 'PERFORM' statement
    ' From the manual------
    ' The PERFORM statement is: 
    ' - An out-of-line PERFORM statement When procedure-name-1 is specified. 
    ' - An in-line PERFORM statement When procedure-name-1 is omitted.
    '
    ' An in-line PERFORM must be delimited by the END-PERFORM phrase. 
    '
    ' The in-line and out-of-line formats cannot be combined. For example, if procedure-name-1 is specified, imperative statements and the END-PERFORM 'phrase must not be specified. 
    '
    ' The PERFORM statement formats are: 
    ' - Basic PERFORM 
    ' - TIMES phrase PERFORM 
    ' - UNTIL phrase PERFORM 
    ' - VARYING phrase PERFORM
    '
    Dim EndIndex As Integer = 0
    Dim Statement As String = ""

    ' BASIC Perform
    ' If the next phrase/word is a procedure-name there is no end-perform
    '  even if there is a TIMES or VARYING phrase
    If ListOfParagraphs.IndexOf(cWord(WordIndex + 1)) > -1 Then
      Call GetStatement(WordIndex, EndIndex, Statement)
      pumlFile.WriteLine(Indent() & ":" & Statement.Trim & "|")
      WordIndex = EndIndex
      Exit Sub
    End If

    ' looking for Conditional (UNTIL) phrase
    For EndIndex = WordIndex + 1 To cWord.Count - 1
      If cWord(EndIndex) = "UNTIL" Then
        Call GetStatement(WordIndex, EndIndex, Statement)
        pumlFile.WriteLine(Indent() & "while (" & Statement.Trim & ") is (true)")
        IndentLevel += 1
        WordIndex = EndIndex
        'WithinPerformWithEndPerformStatement = True
        WithinPerformCnt += 1
        Exit Sub
      End If
    Next

    ' looking for TIMES phrase
    For EndIndex = WordIndex + 1 To cWord.Count - 1
      If cWord(EndIndex) = "TIMES" Then
        Call GetStatement(WordIndex, EndIndex, Statement)
        pumlFile.WriteLine(Indent() & "while (" & Statement.Trim & ") is (true)")
        IndentLevel += 1
        WordIndex = EndIndex
        'WithinPerformWithEndPerformStatement = True
        WithinPerformCnt += 1
        Exit Sub
      End If
    Next

    ' all other combinations of the PERFORM should have an END-PERFORM
    '  this is just a start of a series of commands, not really a PERFORM with a loop
    For EndIndex = WordIndex + 1 To cWord.Count - 1
      If cWord(EndIndex) = "END-PERFORM" Then
        Call GetStatement(WordIndex, EndIndex, Statement)
        pumlFile.WriteLine(Indent() & ":DO;")
        IndentLevel += 1
        WordIndex = EndIndex
        Exit Sub
      End If
    Next

    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlFile.WriteLine(Indent() & ":" & Statement.Trim & "|")
    WordIndex = EndIndex

  End Sub
  Sub ProcessPumlDO(ByRef WordIndex As Integer)
    ' find the end of 'DO' statement
    '
    Dim EndIndex As Integer = 0
    Dim Statement As String = ""
    Dim EndPerformFound As Boolean = False

    EndIndex = cWord.Count - 1
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlFile.WriteLine(Indent() & "while (" & Statement.Trim & ") is (true)")
    IndentLevel += 1
    WordIndex = EndIndex
  End Sub

  Sub ProcessPumlStart(ByRef WordIndex As Integer)
    ' The end of 'PERFORM' statement is the next word
    Dim EndIndex As Integer = WordIndex + 1
    Dim Statement As String = ""
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlFile.WriteLine(Indent() & ":" & Statement.Trim & "|")
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlENDPERFORM()
    IndentLevel -= 1
    If WithinPerformCnt > 0 Then
      pumlFile.WriteLine(Indent() & "endwhile (Complete)")
      'WithinPerformWithEndPerformStatement = False
      WithinPerformCnt -= 1
    Else
      pumlFile.WriteLine(Indent() & ":END DO;")
    End If
  End Sub
  Sub ProcessPumlENDDO(ByRef wordindex As Integer)
    IndentLevel -= 1
    pumlFile.WriteLine(Indent() & "endwhile (Complete)")
  End Sub
  Sub ProcessPumlEXEC(ByRef WordIndex As Integer)
    Dim EndIndex As Integer = 0
    Dim EXECStatement As String = ""
    Call GetStatement(WordIndex, EndIndex, EXECStatement)
    pumlFile.WriteLine(Indent() & ":" & EXECStatement.Trim & "}")
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlDisplay(ByRef WordIndex As Integer)
    ' find end of display being careful to handle quoted items as one
    Dim EndIndex As Integer = 0
    Dim WithinQuotes As Boolean = False
    For EndIndex = WordIndex + 1 To cWord.Count - 1
      If cWord(EndIndex).StartsWith("'") Or cWord(EndIndex).StartsWith(QUOTE) Then
        If WithinQuotes Then
          WithinQuotes = False
        Else
          WithinQuotes = True
        End If
        If cWord(EndIndex).Length > 1 Then
          If cWord(EndIndex).EndsWith("'") Or cWord(EndIndex).EndsWith(QUOTE) Then
            WithinQuotes = False
          End If
          If cWord(EndIndex).EndsWith("',") Or cWord(EndIndex).EndsWith(QUOTE & ",") Then
            WithinQuotes = False
          End If
        End If
        Continue For
      End If
      If cWord(EndIndex).EndsWith("'") Or cWord(EndIndex).EndsWith(QUOTE) Then
        WithinQuotes = False
        Continue For
      End If
      If cWord(EndIndex).EndsWith("',") Or cWord(EndIndex).EndsWith(QUOTE & ",") Then
        WithinQuotes = False
        Continue For
      End If
      If WithinQuotes = True Then
        Continue For
      End If
      'is next word a verb?
      If VerbNames.IndexOf(cWord(EndIndex)) > -1 Then
        EndIndex -= 1
        Exit For
      End If
    Next
    If EndIndex > cWord.Count - 1 Then
      EndIndex = cWord.Count - 1
    End If
    Dim WordsTogether As String = StringTogetherWords(WordIndex, EndIndex)
    Dim Statement As String = AddNewLineAboutEveryNthCharacters(WordsTogether, ESCAPENEWLINE, 30)
    pumlFile.WriteLine(Indent() & ":" & Statement.Trim & ";")
    WordIndex = EndIndex

  End Sub
  Function Indent() As String
    If IndentLevel > 0 Then
      Return Space(IndentLevel * 2)
    End If
    Indent = ""
  End Function
  Sub GetStatement(ByRef WordIndex As Integer, ByRef EndIndex As Integer, ByRef statement As String)
    ' get the whole COBOL statement of this verb by looking for the next verb
    'Dim StartIndex As Integer = WordIndex
    EndIndex = IndexToNextVerb(WordIndex)
    If EndIndex = -1 Then
      EndIndex = cWord.Count - 1
    End If
    Dim WordsTogether As String = StringTogetherWords(WordIndex, EndIndex)
    statement = AddNewLineAboutEveryNthCharacters(WordsTogether, ESCAPENEWLINE, 30)
  End Sub
  Function GetFDDetails(ByRef cWord As List(Of String)) As String
    ' Analyze a File's Definition
    ' COBOL
    ' For the given SELECT (provided in cWord).
    ' This also looks at the FD/SD statements.
    ' Easytrieve
    ' For the given FILE
    ' Returns a string separated with delimiter:
    '0-FileNameOnly
    '1-pgmName
    '2-pgmSeq
    '3-file_name_1
    '4-Level  (FD/SD)
    '5-OpenMode (Input/Output)
    '6-RecordingMode (Fixed/Variable) RECFM
    '7-RecordSizeMinimum
    '8-RecordSizeMaximum
    '9-assignment_name_1 (DD name)
    '10-organization (sequential/indexed)
    Dim file_name_1 As String = ""
    Dim level As String = ""
    Dim OpenMode As String = ""
    Dim RecordingMode As String = ""
    Dim RecordSizeMinimum As Integer = 0
    Dim RecordSizeMaximum As Integer = 0
    Dim assignment_name_1 As String = ""
    Dim organization As String = ""
    Dim fdWords As New List(Of String)

    Select Case SourceType
      Case "COBOL"
        ' Find the file-name-1 value presuming this value is after the SELECT and/or OPTIONAL
        If cWord(1).Equals("OPTIONAL") Then
          file_name_1 = cWord(2)
        Else
          file_name_1 = cWord(1)
        End If

        assignment_name_1 = ""
        Dim index As Integer = GetKeywordIndex("ASSIGN")
        Dim index2 As Integer = 0
        Dim index3 As Integer = 0

        If index > -1 Then
          If cWord(index + 1).Equals("TO") Then
            assignment_name_1 = cWord(index + 2)
          Else
            assignment_name_1 = cWord(index + 1)
          End If
          assignment_name_1 = assignment_name_1.Replace(".", "")
        End If

        ' need ORGANIZATION value SEQUENTIAL, INDEXED, RELATIVE, or LINE SEQUENTIAL
        ' If no 'ORGANIZATION' value the default is SEQUENTIAL
        organization = "SEQUENTIAL"
        index = GetKeywordIndex("SEQUENTIAL")
        If index > -1 Then
          Select Case True
            Case "ORGANIZATION".Equals(cWord(index - 1))
            Case "ORGANIZATION".Equals(cWord(index - 2)) Or "IS".Equals(cWord(index - 1))
              organization = "SEQUENTIAL"
          End Select
        End If
        index = GetKeywordIndex("INDEXED")
        If index > -1 Then
          organization = "INDEXED"
        End If
        index = GetKeywordIndex("RELATIVE")
        If index > -1 Then
          organization = "RELATIVE"
        End If
        index = GetKeywordIndex("LINE")
        If index > -1 Then
          index2 = GetKeywordIndex("SEQUENTIAL")
          If index2 > -1 Then
            organization = "LINE SEQUENTIAL"
          End If
        End If

        ' Locate the FD statement for this SELECT statement
        If LocateFDStatement(file_name_1, fdWords) = -1 Then
          MessageBox.Show("FD/SD statement not found:PGM:" & pgmName & ";FILE:" & file_name_1)
        End If

        level = fdWords(0)

        Dim fdword1 As String = ""
        Dim fdword2 As String = ""
        Dim fdword3 As String = ""
        RecordingMode = "V"
        index = fdWords.IndexOf("RECORDING")
        If index > -1 Then
          ' sloppy code, I know...
          If index + 1 <= fdWords.Count - 1 Then
            fdword1 = fdWords(index + 1)
          End If
          If index + 2 <= fdWords.Count - 1 Then
            fdword2 = fdWords(index + 2)
          End If
          If index + 3 <= fdWords.Count - 1 Then
            fdword3 = fdWords(index + 3)
          End If
          If fdword1 = "MODE" And fdword2 = "IS" Then
            RecordingMode = fdword3
          End If
          If fdword1 = "MODE" Or fdword1 = "IS" Then
            RecordingMode = fdword2
          End If
          RecordingMode = fdword1
        End If

        RecordSizeMinimum = 0
        ' search to find the RECORD CLAUSE
        index2 = -1
        index3 = -1
        For index = 0 To fdWords.Count - 1
          If fdWords(index).Equals("RECORD") Then
            Select Case True
              Case IsNumeric(fdWords(index + 1))
                index2 = index + 1
                If fdWords(index2 + 1) = "TO" Then
                  index3 = index2 + 2
                End If
                Exit For
              Case fdWords(index + 1).Equals("CONTAIN") Or
                   fdWords(index + 1).Equals("CONTAINS")
                If IsNumeric(fdWords(index + 2)) Then
                  index2 = index + 2
                End If
                If fdWords(index2 + 1) = "TO" Then
                  index3 = index2 + 2
                End If
                Exit For
              Case fdWords(index + 1).Equals("IS")
                If fdWords(index + 2).Equals("VARYING") Then
                  index2 = index + 3
                  index2 = GetIndexForRecordSize(index2, fdWords)
                  If fdWords(index2 + 1) = "TO" Then
                    index3 = index2 + 2
                  End If
                  Exit For
                End If
              Case fdWords(index + 1).Equals("VARYING")
                index2 = index + 2
                index2 = GetIndexForRecordSize(index2, fdWords)
                If fdWords(index2 + 1) = "TO" Then
                  index3 = index2 + 2
                End If
                Exit For
            End Select
          End If
        Next
        If index2 > -1 Then
          RecordSizeMinimum = Val(fdWords(index2))
        End If

        RecordSizeMaximum = RecordSizeMinimum
        If index3 > -1 Then
          RecordSizeMaximum = Val(fdWords(index3))
        End If

        OpenMode = GetOpenMode(pgmName, file_name_1)

      Case "Easytrieve"
        file_name_1 = cWord(1)
        level = "FILE"
        OpenMode = GetOpenMode(pgmName, file_name_1)
        RecordingMode = "F"
        RecordSizeMinimum = 0
        RecordSizeMaximum = 0
        assignment_name_1 = cWord(1)
        organization = "SEQUENTIAL"
    End Select

    GetFDDetails = (FileNameOnly & Delimiter &
                         pgmName & Delimiter &
                         LTrim(Str(pgmSeq)) & Delimiter &
                         file_name_1 & Delimiter &
                         level & Delimiter &
                         RTrim(OpenMode) & Delimiter &
                         RecordingMode & Delimiter &
                         LTrim(Str(RecordSizeMinimum)) & Delimiter &
                         LTrim(Str(RecordSizeMaximum)) & Delimiter &
                         assignment_name_1 & Delimiter &
                         organization)

  End Function
  Function GetIndexForRecordSize(ByVal index As Integer, ByRef fdWords As List(Of String)) As Integer
    GetIndexForRecordSize = index
    Select Case True
      Case IsNumeric(fdWords(index)) : Exit Select
      Case IsNumeric(fdWords(index + 1)) : GetIndexForRecordSize += 1
      Case IsNumeric(fdWords(index + 2)) : GetIndexForRecordSize += 2
      Case IsNumeric(fdWords(index + 3)) : GetIndexForRecordSize += 3
      Case Else
        MessageBox.Show("Unknown 'IS VARYING' syntax@" & pgmName & "FD:" & fdWords.ToString)
    End Select
  End Function
  Function LocateFDStatement(ByRef filename As String, ByRef fdWords As List(Of String)) As Integer
    ' returns TWO things:
    ' index to the stmt array where the FD is located
    '  -1 = filename not found!
    ' fdWords of the found FD parsed out for this filename
    '
    LocateFDStatement = -1          'Not found
    Dim statement As String = ""
    For index As Integer = 0 To SrcStmt.Count - 1
      statement = SrcStmt(index)
      If Len(statement) = 0 Then
        Continue For
      End If
      ' parse statement into fd words
      Dim tWord = statement.Replace(".", " ").Split(New Char() {" "c})
      fdWords.Clear()
      For Each word As String In tWord
        If word.Trim.Length > 0 Then        'dropping empty words
          fdWords.Add(word.ToUpper)
        End If
      Next
      If fdWords.Count >= 2 Then
        If (fdWords(0) = "FD" Or fdWords(0) = "SD") Then
          If fdWords(1) = filename Then
            LocateFDStatement = index
            Exit Function
          End If
        End If
      End If
    Next
  End Function
  Function GetOpenMode(ByRef pgmName As String, ByRef file_name_1 As String) As String
    ' Determine the OPEN mode of the file
    ' have to scan through the statement file looking for file_name_1 for 
    ' 'PROGRAM-ID=<pgmName> and OPEN ('INPUT' or 'OUTPUT' or 'I-O' or 'EXTEND')
    ' It could have all open modes.
    ' 
    GetOpenMode = ""
    Dim srcWords As New List(Of String)
    Dim ListOfOpenModes As New List(Of String)

    For Index As Integer = pgm.ProcedureDivision + 1 To pgm.EndProgram
      If SrcStmt(Index).Substring(0, 1) = "*" Then
        Continue For
      End If
      Call GetSourceWords(SrcStmt(Index), srcWords)
      ' search this statement line if it holds any reference to file_name_1
      ' if it does, see what open mode it has.
      For fnIndex = 0 To srcWords.Count - 1 Step 1
        fnIndex = srcWords.IndexOf(file_name_1, fnIndex)
        If fnIndex = -1 Then Exit For
        For x As Integer = fnIndex - 1 To 0 Step -1
          Select Case srcWords(x)
            Case "INPUT"
              AddToListOfOpenModes(ListOfOpenModes, "INPUT")
              Exit For
            Case "OUTPUT"
              AddToListOfOpenModes(ListOfOpenModes, "OUTPUT")
              Exit For
            Case "I-O"
              AddToListOfOpenModes(ListOfOpenModes, "I/O")
              Exit For
            Case "EXTEND"
              AddToListOfOpenModes(ListOfOpenModes, "EXTEND")
              Exit For
            ' these cases below indicate this was not an OPEN verb
            Case "READ"
              Exit For
            Case "CLOSE"
              Exit For
            Case "SORT"
              AddToListOfOpenModes(ListOfOpenModes, "SORT")
              Exit For
            Case "MERGE"
              AddToListOfOpenModes(ListOfOpenModes, "MERGE")
              Exit For
            Case "USING"
              AddToListOfOpenModes(ListOfOpenModes, "SORTIN")
              Exit For
            Case "GIVING"
              AddToListOfOpenModes(ListOfOpenModes, "SORTOUT")
              Exit For
            Case "OPEN"
              MessageBox.Show("Never found open mode!:" &
                              file_name_1 & ":" & SrcStmt(Index))
              Exit For
            Case "PUT"
              AddToListOfOpenModes(ListOfOpenModes, "WRITE")
          End Select
        Next x
      Next fnIndex
    Next Index
    Dim modes As String = ""
    For Each mode In ListOfOpenModes
      modes &= mode & " "
    Next
    GetOpenMode = modes.Trim()
  End Function
  Sub AddToListOfOpenModes(ByRef ListOfOpenModes As List(Of String), ByRef theMode As String)
    If ListOfOpenModes.IndexOf(theMode) = -1 Then
      ListOfOpenModes.Add(theMode)
    End If
  End Sub
  Function GetOpenModeSQL(ByRef filename As String) As String
    ' Determine the OPEN mode (SELECT, INSERT, UPDATE, etc) of the SQL filename.
    ' search each statement for an "EXEC SQL".
    '  then search if filename is referenced, if so get that SELECT/INSERT/UPDATE...
    '  then search for another "EXEC SQL" on same line.
    '
    ' It could have all open modes.
    ' 
    GetOpenModeSQL = ""
    Dim srcWords As New List(Of String)
    Dim ListOfOpenModes As New List(Of String)


    For Index As Integer = pgm.ProcedureDivision + 1 To pgm.EndProgram
      If SrcStmt(Index).Substring(0, 1) = "*" Then
        Continue For
      End If
      Call GetSourceWords(SrcStmt(Index), srcWords)
      If srcWords.Count <= 6 Then
        Continue For
      End If
      For cblIndex = 0 To srcWords.Count - 1
        If srcWords(cblIndex) = "EXEC" Then
          If srcWords(cblIndex + 1) = "SQL" Then
            Dim filenamefound As Boolean = False
            ' find any part of that filename (could have DB Qualifier on it)
            Dim tblIndex As Integer
            For tblIndex = cblIndex + 2 To srcWords.Count - 1
              If srcWords(tblIndex) = "END-EXEC" Then
                Exit For
              End If
              If InStr(srcWords(tblIndex), filename) > 0 Then
                filenamefound = True
                Exit For
              End If
            Next
            If filenamefound Then
              If ListOfOpenModes.IndexOf(srcWords(cblIndex + 2)) = -1 Then
                ListOfOpenModes.Add(srcWords(cblIndex + 2))
              End If
            End If
            cblIndex = tblIndex
          End If
        End If
      Next cblIndex
    Next Index
    ' Check if there is an SQL CURSOR reference
    For index = pgm.ProcedureDivision - 1 To pgm.DataDivision Step -1
      If SrcStmt(index).Substring(0, 1) = "*" Then
        Continue For
      End If
      Call GetSourceWords(SrcStmt(index), srcWords)
      If srcWords.Count < 5 Then
        Continue For
      End If
      If srcWords(0) = "EXEC" And
            srcWords(1) = "SQL" And
            srcWords(2) = "DECLARE" And
            srcWords(4) = "CURSOR" Then
        Dim filenamefound As Boolean = False
        Dim tblIndex As Integer
        For tblIndex = tblIndex + 5 To srcWords.Count - 1
          ' find any part of that filename (could have DB Qualifier on it)
          If srcWords(tblIndex) = "END-EXEC" Then
            Exit For
          End If
          If InStr(srcWords(tblIndex), filename) > 0 Then
            filenamefound = True
            Exit For
          End If
        Next
        If filenamefound Then
          If ListOfOpenModes.IndexOf(srcWords(4)) = -1 Then
            ListOfOpenModes.Add(srcWords(4))
          End If
        End If
      End If
    Next
    Dim modes As String = ""
    For Each mode In ListOfOpenModes
      modes &= mode & " "
    Next
    GetOpenModeSQL = modes.Trim()
  End Function
  Sub GetSourceWords(ByVal statement As String, ByRef srcWords As List(Of String))
    ' takes the stmt and breaks into words and drops blanks
    srcWords.Clear()
    'statement = " DISPLAY '*** CRCALCX REC READ        = ' WS-REC-READ.   "
    statement = statement.Trim
    Dim WithinQuotes As Boolean = False
    Dim word As String = ""
    Dim aByte As String = ""
    For x As Integer = 0 To statement.Length - 1
      aByte = statement.Substring(x, 1)
      If aByte = "'" Then
        WithinQuotes = Not WithinQuotes
      End If
      If aByte = " " Then
        If WithinQuotes Then
          word &= aByte
        Else
          If word.Trim.Length > 0 Then
            srcWords.Add(word.ToUpper)
            word = ""
          End If
        End If
      Else
        word &= aByte
      End If
    Next
    If word.EndsWith(".") Then
      word = word.Remove(word.Length - 1)
      srcWords.Add(word.ToUpper)
      word = ""
    End If
    If word.Length > 0 Then
      srcWords.Add(word)
    End If
  End Sub
  Function GetKeywordIndex(keyword As String) As Integer
    ' find the keyword, if any, in the list
    GetKeywordIndex = cWord.IndexOf(keyword)
  End Function
  Function FindReadRecordName(ByRef fnIndex As Integer, ByRef srcWords As List(Of String)) As String
    ' fnIndex points to the READ filename verb
    ' loop thru this read verb to find the "INTO" if there is any to find
    ' COBOL syntax: READ filename [RECORD] [INTO recordname]
    '               READ filename [NEXT] [RECORD]
    '               READ filename [PREVIOUS] [RECORD]
    For IntoIndex As Integer = fnIndex + 2 To srcWords.Count - 1
      Select Case srcWords(IntoIndex)
        Case "INTO"
          FindReadRecordName = srcWords(IntoIndex + 1)
          Exit Function
        Case "NEXT", "PREVIOUS", "RECORD"
          Continue For
        Case Else
          Exit For
      End Select
    Next
    FindReadRecordName = ""
  End Function
  Function FindWSRecordNameIndex(ByRef DataIndex As Integer, ByVal WSRecordName As String) As Integer
    ' Use the Data Division index to search Stmt array to get the WS record index/location,
    ' 
    Dim FDWords As New List(Of String)
    Dim RecordWords As New List(Of String)
    Dim FDIndex As Integer = -1
    FindWSRecordNameIndex = -1
    For FDIndex = DataIndex To pgm.ProcedureDivision
      If SrcStmt(FDIndex).Substring(0, 1) = "*" Then
        Continue For
      End If
      Call GetSourceWords(SrcStmt(FDIndex), FDWords)
      If FDWords.Count >= 2 Then
        If FDWords(0) = "01" And FDWords(1) = WSRecordName Then
          Exit For
        End If
      End If
    Next
    If FDIndex < pgm.ProcedureDivision Then
      FindWSRecordNameIndex = FDIndex
    End If
  End Function
  Function FindWriteRecordName(ByRef fnIndex As Integer, ByRef srcWords As List(Of String)) As String
    'check this write verb for a 'FROM' otherwise use the record name from the FD
    ' COBOL syntax: WRITE FDrecordname [FROM recordname] 
    If fnIndex + 1 >= srcWords.Count - 1 Then
      FindWriteRecordName = srcWords(fnIndex + 1)
      Exit Function
    End If
    If srcWords(fnIndex + 2) = "FROM" Then
      FindWriteRecordName = srcWords(fnIndex + 3)
    Else
      FindWriteRecordName = srcWords(fnIndex + 1)
    End If
  End Function
  Function DetermineDigits(ByRef pic As String) As Integer
    'ie pic 9(02) = 2 digits (99)
    'ie pic 99 = 2 digits (99)
    'ie pic 9(7)v99 = 9 digits (9999999v99)
    'ie pic x = 1 digit (X)
    Dim startRepeat As Integer = 0
    Dim picdigits As Integer = 0
    For picIndex As Integer = 0 To pic.Length - 1
      Select Case pic.Substring(picIndex, 1)
        Case "("
          startRepeat = picIndex + 1
        Case ")"
          picdigits += Val(pic.Substring(startRepeat, (picIndex - 1) - startRepeat + 1)) - 1
          startRepeat = 0
        Case "X", "9", "B", ".", ",", "-", "+"
          If startRepeat = 0 Then
            picdigits += 1
          End If
      End Select
    Next
    DetermineDigits = picdigits
  End Function
  Function FindCopyOrInclude(ByRef CopyIndex As Integer, ByRef CopyWords As List(Of String)) As String
    FindCopyOrInclude = ""
    Call GetSourceWords(SrcStmt(CopyIndex), CopyWords)
    If CopyWords.Count >= 2 Then
      If CopyWords(0) = "*COPY" Then
        FindCopyOrInclude = CopyWords(1)
        Exit Function
      End If
    End If
    If CopyWords.Count >= 5 Then
      If CopyWords(1) = "EXEC" And CopyWords(2) = "SQL" And CopyWords(3) = "INCLUDE" Then
        FindCopyOrInclude = CopyWords(4)
        Exit Function
      End If
    End If
    If CopyWords.Count >= 2 Then
      If CopyWords(0) = "*INCLUDE++" Then
        FindCopyOrInclude = CopyWords(1)
        Exit Function
      End If
      If CopyWords(0) = "*%COPYBOOK" Then
        FindCopyOrInclude = CopyWords(2)
      End If
    End If
    If IsNumeric(CopyWords(0)) Then
      FindCopyOrInclude = "EXIT FOR"
      Exit Function
    End If
    Select Case CopyWords(0)
      Case "FILE", "FD", "WORKING-STORAGE", "LOCAL-STORAGE", "LINKAGE"
        FindCopyOrInclude = "EXIT FOR"
        Exit Function
    End Select
  End Function
  Function StringTogetherWords(ByRef StartCondIndex As Integer, ByRef EndCondIndex As Integer) As String
    ' string together from startofconditionindex to endofconditionindex
    ' cWord is a global variable
    Dim wordsStrungTogether As String = ""
    For condIndex As Integer = StartCondIndex To EndCondIndex
      wordsStrungTogether &= cWord(condIndex) & " "
    Next
    StringTogetherWords = wordsStrungTogether.TrimEnd
  End Function
  Function AddNewLineAboutEveryNthCharacters(ByRef condStatement As String,
                                            ByRef theNewLine As String,
                                            ByVal Size As Integer) As String
    ' add "\n" or vbnewline (theNewLine) about every SIZE number of characters
    Dim condStatementCR As String = ""
    Dim bytesMoved As Integer = 0
    If condStatement.Length = 0 Then
      AddNewLineAboutEveryNthCharacters = ""
      Exit Function
    End If
    If condStatement.Length > Size Then
      For condIndex As Integer = 0 To condStatement.Length - 1
        If condStatement.Substring(condIndex, 1) = Space(1) And bytesMoved > (Size - 1) Then
          condStatementCR &= theNewLine
          bytesMoved = 0
        End If
        condStatementCR &= condStatement.Substring(condIndex, 1)
        bytesMoved += 1
      Next
    Else
      condStatementCR = condStatement
    End If
    AddNewLineAboutEveryNthCharacters = condStatementCR
  End Function
  Function IndexToNextVerb(ByRef StartCondIndex As Integer) As Integer
    ' cWord is a global variable
    ' VerNames is a global variable
    ' find ending index to next COBOL verb in cWord
    Dim EndCondIndex As Integer = -1
    Dim VerbIndex As Integer = -1
    For EndCondIndex = StartCondIndex + 1 To cWord.Count - 1
      If WithinReadStatement = True Then
        Select Case cWord(EndCondIndex)
          Case "AT", "END", "NOT"
            IndexToNextVerb = EndCondIndex - 1
            Exit Function
        End Select
      End If
      VerbIndex = VerbNames.IndexOf(cWord(EndCondIndex))
      If VerbIndex > -1 Then
        IndexToNextVerb = EndCondIndex - 1
        Exit Function
      End If
    Next
    ' there is not another verb in this statement
    IndexToNextVerb = -1
  End Function

  Sub ProcessComment(ByRef Index As Integer, ByRef statement As String, ByRef Division As String, ByRef CobolFile As String)
    Dim comment As String = ""
    Select Case SourceType
      Case "COBOL"
        comment = statement.
                  PadRight(80).
                  Replace(Delimiter, ":").
                  Replace("*", " ").
                  Substring(7, 65).
                  Trim
      Case "Easytrieve"
        comment = statement.
                  PadRight(80).
                  Replace(Delimiter, ":").
                  Replace("*", " ").
                  Substring(1, 71).
                  Trim
      Case Else

    End Select
    comment = StrConv(comment, vbProperCase)

    If comment.Length > 0 Then
      ListOfComments.Add(FileNameOnly & Delimiter &
                        CobolFile & Delimiter &
                        SourceType & Delimiter &
                        Division & Delimiter &
                        LTrim(Str(Index + 1)) & Delimiter &
                        Chr(34) & comment & Chr(34))
    End If
  End Sub


  Sub InitializeProgramVariables()
    lblCopybookMessage.Text = ""

    ' re Initialize all beginning Variables and tables.
    ' JCL
    FileNameOnly = ""
    tempFileName = ""
    tempCobFileName = ""
    tempEZTFileName = ""
    jControl = ""
    jLabel = ""
    jParameters = ""
    procName = ""
    jobName = ""
    jobClass = ""
    msgClass = ""
    prevPgmName = ""
    prevStepName = ""
    prevDDName = ""
    pgmName = ""
    DDName = ""
    stepName = ""
    InstreamProc = ""
    ddConcatSeq = 0
    ddSequence = 0
    jobSequence = 0
    procSequence = 0
    execSequence = 0
    jclStmt.Clear()
    ListOfExecs.Clear()
    ' COBOL fields
    SourceType = ""
    SrcStmt.Clear()
    cWord.Clear()
    ListOfFiles.Clear()
    ListOfRecordNames.Clear()
    ListOfRecords.Clear()
    ListOfFields.Clear()
    ListOfReadIntoRecords.Clear()
    ListOfWriteFromRecords.Clear()
    ListOfComments.Clear()
    IFLevelIndex.Clear()
    ProgramAuthor = ""
    ProgramWritten = ""
    IndentLevel = -1
    BRLevel = -1
    FirstWhenStatement = False
    WithinReadStatement = False
    WithinReadConditionStatement = False
    WithinIF = False
    pgmSeq = 0

  End Sub
  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Me.Text = "ADDILite " & ProgramVersion



    ' This area is the COBOL Verb array with counts. 
    ' **BE SURE TO KEEP VerbNames AND VerbCount ARRAYS IN SYNC!!!**
    ' Flow commands
    VerbNames.Add("GO")
    VerbNames.Add("ALTER")
    VerbNames.Add("CALL")
    VerbNames.Add("PERFORM")
    VerbNames.Add("EVALUATE")
    VerbNames.Add("WHEN")
    VerbNames.Add("CONTINUE")
    VerbNames.Add("IF")
    VerbNames.Add("ELSE")
    VerbNames.Add("GOBACK")
    VerbNames.Add("STOP")
    VerbNames.Add("CHAIN")
    ' I/O
    VerbNames.Add("OPEN")
    VerbNames.Add("READ")
    VerbNames.Add("WRITE")
    VerbNames.Add("REWRITE")
    VerbNames.Add("CLOSE")
    VerbNames.Add("EXEC")
    VerbNames.Add("COMMIT")
    VerbNames.Add("CANCEL")
    VerbNames.Add("DELETE")
    VerbNames.Add("MERGE")
    VerbNames.Add("SORT")
    VerbNames.Add("RETURN")
    VerbNames.Add("NEXT")
    ' Maths
    VerbNames.Add("COMPUTE")
    VerbNames.Add("ADD")
    VerbNames.Add("SUBTRACT")
    VerbNames.Add("MULTIPLY")
    VerbNames.Add("DIVIDE")
    ' Misc
    VerbNames.Add("MOVE")
    VerbNames.Add("DISABLE")
    VerbNames.Add("DISPLAY")
    VerbNames.Add("ENABLE")
    VerbNames.Add("END-READ")
    VerbNames.Add("END-EVALUATE")
    VerbNames.Add("END-IF")
    VerbNames.Add("END-INVOKE")
    VerbNames.Add("END-PERFORM")
    VerbNames.Add("END-SET")
    VerbNames.Add("ENTER")
    VerbNames.Add("ENTRY")
    VerbNames.Add("EXAMINE")
    VerbNames.Add("EXECUTE")
    VerbNames.Add("EXHIBIT")
    VerbNames.Add("EXIT")
    VerbNames.Add("GENERATE")
    VerbNames.Add("INITIALIZE")
    VerbNames.Add("INITIATE")
    VerbNames.Add("INSPECT")
    VerbNames.Add("INVOKE")
    VerbNames.Add("NOTE")
    VerbNames.Add("OTHERWISE")
    VerbNames.Add("READY")
    VerbNames.Add("RECEIVE")
    VerbNames.Add("RECOVER")
    VerbNames.Add("RELEASE")
    VerbNames.Add("RESET")
    VerbNames.Add("ROLLBACK")
    VerbNames.Add("SEARCH")
    VerbNames.Add("SEND")
    VerbNames.Add("SERVICE")
    VerbNames.Add("SET")
    VerbNames.Add("START")
    VerbNames.Add("STRING")
    VerbNames.Add("SUPPRESS")
    VerbNames.Add("TERMINATE")
    VerbNames.Add("TRANSFORM")
    VerbNames.Add("UNLOCK")
    VerbNames.Add("UNSTRING")

    ' Flow commands
    VerbCount.Add(0)    'GO
    VerbCount.Add(0)    'ALTER
    VerbCount.Add(0)    'CALL
    VerbCount.Add(0)    'PERFORM
    VerbCount.Add(0)    'EVALUATE
    VerbCount.Add(0)    'WHEN
    VerbCount.Add(0)    'CONTINUE
    VerbCount.Add(0)    'IF
    VerbCount.Add(0)    'ELSE
    VerbCount.Add(0)    'GOBACK
    VerbCount.Add(0)    'STOP
    VerbCount.Add(0)    'CHAIN
    ' I/O
    VerbCount.Add(0)    'OPEN
    VerbCount.Add(0)    'READ
    VerbCount.Add(0)    'WRITE
    VerbCount.Add(0)    'REWRITE
    VerbCount.Add(0)    'CLOSE
    VerbCount.Add(0)    'EXEC
    VerbCount.Add(0)    'COMMIT
    VerbCount.Add(0)    'CANCEL
    VerbCount.Add(0)    'DELETE
    VerbCount.Add(0)    'MERGE
    VerbCount.Add(0)    'SORT
    VerbCount.Add(0)    'RETURN
    VerbCount.Add(0)    'NEXT
    ' Maths
    VerbCount.Add(0)    'COMPUTE
    VerbCount.Add(0)    'ADD
    VerbCount.Add(0)    'SUBTRACT
    VerbCount.Add(0)    'MULTIPLY
    VerbCount.Add(0)    'DIVIDE
    ' Misc
    VerbCount.Add(0)    'MOVE
    VerbCount.Add(0)    'DISABLE
    VerbCount.Add(0)    'DISPLAY
    VerbCount.Add(0)    'ENABLE
    VerbCount.Add(0)    'END-READ
    VerbCount.Add(0)    'END-EVALUATE
    VerbCount.Add(0)    'END-IF
    VerbCount.Add(0)    'END-INVOKE
    VerbCount.Add(0)    'END-PERFORM
    VerbCount.Add(0)    'END-SET
    VerbCount.Add(0)    'ENTER
    VerbCount.Add(0)    'ENTRY
    VerbCount.Add(0)    'EXAMINE
    VerbCount.Add(0)    'EXECUTE
    VerbCount.Add(0)    'EXHIBIT
    VerbCount.Add(0)    'EXIT
    VerbCount.Add(0)    'GENERATE
    VerbCount.Add(0)    'INITIALIZE
    VerbCount.Add(0)    'INITIATE
    VerbCount.Add(0)    'INSPECT
    VerbCount.Add(0)    'INVOKE
    VerbCount.Add(0)    'NOTE
    VerbCount.Add(0)    'OTHERWISE
    VerbCount.Add(0)    'READY
    VerbCount.Add(0)    'RECEIVE
    VerbCount.Add(0)    'RECOVER
    VerbCount.Add(0)    'RELEASE
    VerbCount.Add(0)    'RESET
    VerbCount.Add(0)    'ROLLBACK
    VerbCount.Add(0)    'SEARCH
    VerbCount.Add(0)    'SEND
    VerbCount.Add(0)    'SERVICE
    VerbCount.Add(0)    'SET
    VerbCount.Add(0)    'START
    VerbCount.Add(0)    'STRING
    VerbCount.Add(0)    'SUPPRESS
    VerbCount.Add(0)    'TERMINATE
    VerbCount.Add(0)    'TRANSFORM
    VerbCount.Add(0)    'UNLOCK
    VerbCount.Add(0)    'UNSTRING

  End Sub


End Class