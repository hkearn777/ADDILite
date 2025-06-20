﻿Imports System.IO
'Imports System.Reflection
Imports System.Text.RegularExpressions
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ListView
'Imports Microsoft.Office
'Imports Microsoft.Office.Interop
Imports ClosedXML.Excel
'Imports DocumentFormat.OpenXml.Drawing
'Imports Microsoft.VisualBasic.Logging

Public Class Form1
  ' ADDILite will read an IBM JCL syntax file and break apart its
  '  parts and pieces. Those parts are JOB, Executables, and Datasets.

  ' This will analyze COBOL and Easytrieve Sources to create the Data Details.
  ' '
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
  Dim ProgramVersion As String = "v2.5.0"
  'Change-History.
  ' 2025/05/16 v2.5.0 hk Migrate Interop to ClosedXML for Excel processing
  ' 2025/05/16 v2.4.3 hk Create a builder CLASS for hyperlink functions
  ' 2025/05/16 v2.4.2 hk Fix Folder names on Summary Tab and hyperlinks
  ' 2025/05/15 v2.4.1 hk Remove Hyperlink for Utilities on Programs tab
  ' 2025/05/15 v2.4.0 hk Remove BusinessRulesFolder 
  ' 2025/05/14 v2.3.0 hk Remove Business Rules COBOL Module, see CreateCobolBusinessRules.sln
  ' 2025/05/14 v2.2.0 hk Remove Old Comments
  ' 2025/05/08 v2.1.0 hk Add JCL EXEC COND column to Programs tab
  ' 2025/03/02 v2.0.1 hk Set Environment variables ADDILite to the folder path
  ' 2025/02/27 v2     hk Add subfolders for sources (ie cbl, cob, cpy, etc.)
  '                      - fix CALLPGMS .jcl to have unique program names
  '                      - delete the #ADDI## files at start of program (at click)
  ' 2025/02/12 v1.8.9 hk add Status display messages
  '                      Comment out a COBOL line if Indictor column as '#"
  '                      For Data gathering summary show file extension
  '                      Include *.txt for Source code COBOL look up
  ' 2025/02/08 v1.8.8 hk move array to range for each tab
  ' 2025/02/04 v1.8.7 hk correct finding Record name starting at Working-storage
  ' 2025/01/27 v1.8.6 hk add Log entry for COBOLAlias.csv file
  ' 2025/01/22 v1.8.5 hk replace period in indicator area with asterisk
  ' 2025/01/19 v1.8.4 hk Update Cursor find
  ' 2025/01/14 v1.8.3 hk Comment out COBOL debug indicator (D)
  ' 2025/01/16 v1.8.2 hk Add new Tab Instream
  ' 2024/10/24 v1.8   hk count source lines and place on Programs tab
  '                      - fixed flowchart to max 45 character lines
  '                      - fixed COBOL continuation with blank lines
  '                      - Included statement for UPDATE, INSERT, DELETE execSQL
  '                      - New Folder structure for PUML, Flowcharts, Business Rules
  '                      - Move JOB Flow logic to standalone program: CreateJobFlowcharts
  '                      - Add COBOL program-id Alias filenames feature
  '                      - Fix Open Mode for SQL on Records tab
  '                      - Fix SQL Copybooks
  ' 2024/09/30 v1.7   hk Flowchart Links
  ' 2024/09/27 v1.6.7 hk fix drop empty '//' and '/*' JCL statements
  '                      - fix missing execname and pgmname when PROC is utility
  '                      - fix writing dynamic routines to CALLPGMS.jcl
  ' 2024/09/24 v1.6.6 hk fix BR value remove equal sign
  ' 2024/09/23 v1.6.5 hk fixed key/value parse, fixed MISSING PROC message
  ' 2024/09/20 v1.6.4 hk Reference PROCs in PROC folder instead of Sources
  '                      - remove ENDIF jcl 
  '                      - remove JCL INCLUDE statements
  '                      - remove PROC= on PROC names
  '                      - assign ++ to JCL lines that are part of PROC
  '                      - clear PROCNAME as needed
  '                      - uppercase callpgms name and Account to INTERNAL
  ' 2024/09/13 v1.6.3 hk Fix symbolic for program name
  ' 2024/09/13 v1.6.2 hk Fix Parsing of DD SYSOUT
  ' 2024/08/09 v1.6  hk Initial Directory settings on start up (especially new install)
  ' 2024/07/24 v1.5  hk Support CA-Datacom databases
  ' 2024/07/03 v1.4  hk Business Rules. Implement ExtractBR into this code.
  '                  - Paragraph to Paragraph diagram
  '                  - Code clean up
  ' 2024/04/15 v1.3  hk Create Batch ADDILite feature
  '                  - Optionally Create worksheet Tabs
  '                  - fixed CALLS tab content
  '                  - fixed execCICS tab content
  '                  - fixed continuation of COBOL lines
  ' 2024/04/13 v1.2.1 hk Revised Programs tab and new Files Tab
  '                    - Added Libraries tab.
  ' 2024/04/10 v1.2  hk Create Jobs Tab
  '                    -Create Libraries Tab
  ' 2024/03/20 v1.13 hk Create JCL Comments Tab
  '                    - Handle inline Easytrieve code via in-stream data sets (DD *)
  '                    - New Source Type: Assembler
  ' 2024/03/20 v1.13 hk Create JCL Comments Tab
  ' 2024/02/27 v1.12 hk Create ScreenMaps Tab
  '                    - Added Readonly option on excel SaveAs
  '                    - Set Freeze Frames on first rows of the worksheets
  ' 2024/02/09 v1.11 hk v1.11 Create CICS Tab
  '                    - Create MAPS tab
  '                    - Corrected various errors in Puml flowcharts
  '                    - Corrected File name validations
  '                    - Support PC COBOL syntax
  ' 2023/01/25 v1.10 hk Page Break on PUML file for COBOL
  ' 2023/12/28 v1.8 hk add Utility: DFSBBO00
  '                    - Support PROCLIB
  '                    - Add utility: IKJEFT1B
  '                    - Change caps for comments
  '                    - Add CALLS tab
  '                    - Clean up log entry for source file not found
  ' 2023/12/19 v1.7 hk For IMS programs, create a PSPNames list file. This is for later DBDName extract.
  '                    - Create IMS Tab, read DBDNames.txt and TELON files
  ' 2023/12/05 v1.6 hk Handle IMS programs in the PROGRAMS tab by adding ExecName Column.
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

  ' Initial directory TODO make this an environment setting.
  Dim InitDirectory As String = ""
  Dim AppDirectory As String = ""
  Dim folderPath As String = ""
  Dim Utilities As String() = New String() {""} ' array to hold the Utilities.txt file content
  Dim hlBuilder As New HyperLinkBuilder(Utilities)
  Dim ControlLibraries As String()
  Dim PUMLFolder As String = ""

  ' Arrays to hold the DB2 Declare to Member names
  ' these two array will share the same index
  Dim DB2Declares As New List(Of String)
  Dim MembersNames As New List(Of String)

  Dim ListOfDataGathering As New List(Of String)
  Dim NumberOfJobsToProcess As Integer = 0
  Dim ListOfJobs As New List(Of String)
  Dim ListOfLibraries As New List(Of String)              'array to hold JOBLIB, STEPLIB, JCLLIB names

  ' JCL
  'Dim DirectoryName As String = ""
  Dim FileNameOnly As String = ""
  Dim FileNameWithExtension As String = ""
  Dim tempCobFileName As String = ""
  Dim tempEZTFileName As String = ""
  Dim tempAsmFileName As String = ""
  Public Delimiter As String = ""
  Dim jControl As String = ""
  Dim jLabel As String = ""
  Dim jParameters As String = ""
  Dim procName As String = ""
  Dim jobName As String = ""
  Dim jobClass As String = ""
  Dim jobMsgClass As String = ""
  Dim JobSourceName As String = ""
  Dim JobAccountInfo As String = ""
  Dim JobProgrammerName As String = ""
  Dim JobTime As String = ""
  Dim JobSend As String = ""
  Dim JobRoute As String = ""
  Dim JobParm As String = ""
  Dim JobRegion As String = ""
  Dim JobCond As String = ""
  Dim JobJCLLib As String = ""
  Dim JobTyprun As String = ""
  Dim JobLib As String = ""
  Dim prevPgmName As String = ""
  Dim prevStepName As String = ""
  Dim prevDDName As String = ""
  Dim execName As String = ""
  Dim pgmName As String = ""
  Dim execCond As String = ""
  Dim DDName As String = ""
  Dim stepName As String = ""
  Dim InstreamProc As String = ""
  Dim CallPgmsFileName = ""


  Dim ddConcatSeq As Integer = 0
  Dim ddSequence As Integer = 0

  Dim SummaryRow As Integer = 0
  Dim JobRow As Integer = 0
  Dim JobCommentsRow As Integer = 0
  Dim ProgramsRow As Integer = 0
  Dim FilesRow As Integer = 0
  Dim RecordsRow As Integer = 0
  Dim FieldsRow As Integer = 0
  Dim CommentsRow As Integer = 0
  Dim EXECSQLRow As Integer = 0
  Dim EXECCICSRow As Integer = 0
  Dim IMSRow As Integer = 0
  Dim DataComRow As Integer = 0
  Dim ScreenMapRow As Integer = 0
  Dim CallsRow As Integer = 0
  Dim LibrariesRow As Integer = 0
  Dim InstreamsRow As Integer = 0

  Dim jclStmt As New List(Of String)
  Dim ListOfExecs As New List(Of String)        'array holding the executable programs
  Dim ListOfEasytrieveLoadAndGo As New Dictionary(Of String, String) 'array holding the names of the 'load and go' Easytrieve programs
  Dim DictOfInstreams As New Dictionary(Of String, String)    'key:job,step,pgm, value:content of DD*
  Dim AliasCobol As New Dictionary(Of String, String)   'program-id to filename

  Dim swIPFile As StreamWriter = Nothing        'Instream proc file, temporary
  Dim swPumlFile As StreamWriter = Nothing
  Dim swInstreamDatasetFile As StreamWriter = Nothing

  Dim LogFile As StreamWriter = Nothing
  Dim LogStmtFile As StreamWriter = Nothing
  Dim swCallPgmsFile As StreamWriter = Nothing

  ' load the Excel References
  ' Build out initial workbook 
  Dim workbook As New XLWorkbook
  ' Build out initial worksheets
  Dim SummaryWorksheet As IXLWorksheet = workbook.Worksheets.Add("Summary")
  Dim JobsWorksheet As IXLWorksheet = workbook.Worksheets.Add("Jobs")
  Dim JobCommentsWorksheet As IXLWorksheet = workbook.Worksheets.Add("JobComments")
  Dim ProgramsWorksheet As IXLWorksheet = workbook.Worksheets.Add("Programs")
  Dim FilesWorksheet As IXLWorksheet = workbook.Worksheets.Add("Files")
  Dim RecordsWorksheet As IXLWorksheet = workbook.Worksheets.Add("Records")
  Dim FieldsWorksheet As IXLWorksheet = workbook.Worksheets.Add("Fields")
  Dim CommentsWorksheet As IXLWorksheet = workbook.Worksheets.Add("Comments")
  Dim EXECSQLWorksheet As IXLWorksheet = workbook.Worksheets.Add("ExecSQL")
  Dim EXECCICSWorksheet As IXLWorksheet = workbook.Worksheets.Add("ExecCICS")
  Dim IMSWorksheet As IXLWorksheet = workbook.Worksheets.Add("IMS")
  Dim DataComWorksheet As IXLWorksheet = workbook.Worksheets.Add("DataCom")
  Dim ScreenMapWorksheet As IXLWorksheet = workbook.Worksheets.Add("ScreenMaps")
  Dim CallsWorksheet As IXLWorksheet = workbook.Worksheets.Add("Calls")
  Dim LibrariesWorksheet As IXLWorksheet = workbook.Worksheets.Add("Libraries")
  Dim InstreamsWorksheet As IXLWorksheet = workbook.Worksheets.Add("Instreams")

  Dim rngSummaryName As IXLRange
  Dim rngJobs As IXLRange
  Dim rngJobComments As IXLRange
  Dim rngPrograms As IXLRange
  Dim rngFiles As IXLRange
  Dim rngRecordsName As IXLRange
  Dim rngFieldsName As IXLRange
  Dim rngComments As IXLRange
  Dim rngEXECSQL As IXLRange
  Dim rngEXECCICS As IXLRange
  Dim rngIMS As IXLRange
  Dim rngDataCom As IXLRange
  Dim rngCalls As IXLRange
  Dim rngScreenMap As IXLRange
  Dim rngStats As IXLRange
  Dim rngLibraries As IXLRange
  Dim rngInstreams As IXLRange

  ' ClosedXML always uses .xlsx format
  Dim DefaultFormat As String = ".xlsx"

  ' indicate read-only 
  Dim SetAsReadOnly As Boolean = True


  Dim CntJobFiles As Integer = 0
  Dim CntProcFiles As Integer = 0
  Dim CntCOBOLFiles As Integer = 0
  Dim CntOutputFiles As Integer = 0
  Dim CntTelonFiles As Integer = 0
  Dim CntScreenMapFiles As Integer = 0
  Dim CntCopybookFiles As Integer = 0
  Dim CntDeclGenFiles As Integer = 0
  Dim CntEasytrieveFiles As Integer = 0
  Dim CntASMFiles As Integer = 0


  Dim ListOfTables As New List(Of String)
  ' Easytrieve fields
  Dim theProcName As String = ""
  ' COBOL fields
  Dim SourceType As String = ""
  Dim SourceCount As Integer = 0
  Dim SrcStmt As New List(Of String)
  Dim cWord As New List(Of String)
  Dim lWord As New List(Of String)                    'Word Level value for IF syncs with cWord
  Dim ListofCOBOLFiles As New List(Of String)        'array to hold all the source files instead of using file.exist()
  Dim ListofEasytrieveFiles As New List(Of String)    'array to hold all the Easytrieve files instead of using file.exist()
  Dim ListofAsmFiles As New List(Of String)           'array to hold all the Assembler files instead of using file.exist()
  Dim ListofCopybookFiles As New List(Of String)      'array to hold all the copybook files instead of using file.exist()
  Dim ListOfPrograms As New List(Of String)           'array to hold Program names and details
  Dim ListOfFiles As New List(Of String)              'array to hold File & DB2 Table names
  Dim ListOfDDs As New List(Of String)                'array to hold the DD entries for 1 JOB
  Dim ListOfRecordNames As New List(Of String)          'array to hold read/written records
  Dim ListOfRecords As New List(Of String)              'array to hold read/written records
  Dim ListOfFields As New List(Of String)             'array to hold fields for each record
  Dim ListOfReadIntoRecords As New List(Of String)    'array to hold Read Into Records
  Dim ListOfWriteFromRecords As New List(Of String)   'array to hold Write from records
  Dim ListOfComments As New List(Of String)           'array to hold comments from source (cobol & easytrieve)
  Dim ListOfCallPgms As New List(Of String)           'array to hold Call programs (sub routines)
  Dim ListOfEXECSQL As New List(Of String)            'array to hold EXEC SQL statments (cobol & easytrieve)
  Dim ListOfDB2Tables As New List(Of String)          'array to hold the DB2 Table names found
  Dim ListOfIMSPgms As New List(Of String)            'array to hold IMS Programs (PSPNames=Program Name)
  Dim ListOfDBDs As New List(Of String)               'array to hold the DBD usages values 
  Dim ListOfDBDNames As New List(Of String)           'array to hold the DBD Names
  Dim ListOfCICSMapNames As New List(Of String)       'array to hold the CICS Map names (cobol)
  Dim ListofScreenMaps As New List(Of String)            'array to hold the IMS Map names and literals
  Dim ListOfIMSMapNames As New List(Of String)        'array to hold the IMS Map Names
  Dim ListOfDataComs As New List(Of String)           'array to hold the DataComms

  Dim IFLevelIndex As New List(Of Integer)            'where in cWord the 'IF' is located
  Public VerbNames As New List(Of String)
  Public VerbCount As New List(Of Integer)
  Public COBOLCondWords As New List(Of String)
  Dim ProgramAuthor As String = ""
  Dim ProgramWritten As String = ""
  Dim IndentLevel As Integer = -1                  'how deep the if has gone
  Dim FirstWhenStatement As Boolean = False
  Public WithinReadStatement As Boolean = False
  Dim WithinReadConditionStatement As Boolean = False
  Dim WithinPerformCnt As Integer = 0
  Dim WithinIF As Boolean = False
  Dim pgmSeq As Integer = 0
  Dim pumlFile As StreamWriter = Nothing          'File holding the Plantuml commands
  Dim PSPFile As StreamWriter = Nothing           'File holding the PSP Names in PSPName until format
  Dim pumlMaxLineCnt As Integer = 1000
  Dim pumlLineCnt As Integer = 0
  Dim pumlPageCnt As Integer = 0
  Dim ScreenType As String = ""
  Dim numLoadAndGo As Integer = -1


  Public Structure ProgramInfo
    Public ProgramId As String
    Public IdentificationDivision As Integer
    Public EnvironmentDivision As Integer
    Public DataDivision As Integer
    Public WorkingStorage As Integer
    Public ProcedureDivision As Integer
    Public EndProgram As Integer
    Public SourceId As String
    Public Sub New(ByVal _ProgramId As String,
                   ByVal _IdentificationDivision As Integer,
                   ByVal _EnvironmentDivision As Integer,
                   ByVal _DataDivsision As Integer,
                   ByVal _WorkingStorage As Integer,
                   ByVal _ProcedureDivision As Integer,
                   ByVal _EndProgram As Integer,
                   ByVal _SourceId As String)
      ProgramId = _ProgramId
      IdentificationDivision = _IdentificationDivision
      EnvironmentDivision = _EnvironmentDivision
      DataDivision = _DataDivsision
      WorkingStorage = _WorkingStorage
      ProcedureDivision = _ProcedureDivision
      EndProgram = _EndProgram
      SourceId = _SourceId
    End Sub
  End Structure
  Dim listOfProgramInfo As New List(Of ProgramInfo)
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

  Private Sub btnAppFolder_Click(sender As Object, e As EventArgs) Handles btnAppFolder.Click
    ' browse for and select Application folder name that is within the Sandbox folder
    Dim bfd_AppFolder As New FolderBrowserDialog With {
      .Description = "Enter Application Folder name",
      .SelectedPath = InitDirectory
    }
    If bfd_AppFolder.ShowDialog() = DialogResult.OK Then
      Dim folders As String() = bfd_AppFolder.SelectedPath.Split("\")
      txtAppFolder.Text = "\" & folders(folders.Length - 1)
      Environment.SetEnvironmentVariable("ADDILite_Application", txtAppFolder.Text, EnvironmentVariableTarget.User)
    Else
      Exit Sub
    End If

  End Sub
  Private Sub btnDataGatheringForm_Click(sender As Object, e As EventArgs) Handles btnDataGatheringForm.Click
    If txtAppFolder.Text.Length <= 1 Then
      MessageBox.Show("Application Folder name Is required!")
      Exit Sub
    End If
    folderPath = InitDirectory & txtAppFolder.Text
    If Not Directory.Exists(folderPath) Then
      MessageBox.Show("Application Folder name does Not exist!" & vbLf & folderPath)
      Exit Sub
    End If
    ' Open file dialog
    Dim ofd_DataGatheringForm As New OpenFileDialog With {
      .InitialDirectory = folderPath,
      .Filter = "Spreadsheet|*.xlsx",
      .Title = "Open the Data Gathering Form"
    }
    If ofd_DataGatheringForm.ShowDialog() = DialogResult.OK Then
      txtDataGatheringForm.Text = ofd_DataGatheringForm.SafeFileName
      CntJobFiles = GetFileCount(folderPath & txtJCLJOBFolder.Text)
      CntProcFiles = GetFileCount(folderPath & txtProcFolder.Text)
      CntCOBOLFiles = GetFileCount(folderPath & txtCobolFolder.Text)
      CntCopybookFiles = GetFileCount(folderPath & txtCopybookFolder.Text)
      CntTelonFiles = GetFileCount(folderPath & txtTelonFolder.Text)
      CntDeclGenFiles = GetFileCount(folderPath & txtDECLGenFolder.Text)
      CntEasytrieveFiles = GetFileCount(folderPath & txtEasytrieveFolder.Text)
      CntASMFiles = GetFileCount(folderPath & txtASMFolder.Text)
      CntScreenMapFiles = GetFileCount(folderPath & txtScreenMapsFolder.Text)
      ' show count of files in each folder
      lblJOBFolder.Text = "JOBS (" & CntJobFiles & "):"
      lblPROCFolder.Text = "PROCS (" & CntProcFiles & "):"
      lblCOBOLFolder.Text = "COBOL (" & CntCOBOLFiles & "):"
      lblCopybooksFolder.Text = "COPYBOOKS (" & CntCopybookFiles & "):"
      lblTelonFolder.Text = "TELON (" & CntTelonFiles & "):"
      lblDECLGenFolder.Text = "DECLGEN (" & CntDeclGenFiles & "):"
      lblEasytrieveFolder.Text = "EASYTRIEVE (" & CntEasytrieveFiles & "):"
      lblASMFolder.Text = "ASM (" & CntASMFiles & "):"
      lblScreensFolder.Text = "SCREENS (" & CntScreenMapFiles & "):"
      PUMLFolder = folderPath & txtPUMLFolder.Text
    Else
      Exit Sub
    End If
  End Sub
  Function GetFileCount(ByRef myDirectoryName As String) As Integer
    ' Get the number of files in the directory
    Try
      Return My.Computer.FileSystem.GetFiles(myDirectoryName).Count

    Catch ex As Exception
      MessageBox.Show("GetFileCount Error:" & ex.Message)
      Return -1
    End Try
  End Function

  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub

  Private Sub txtDelimiter_TextChanged(sender As Object, e As EventArgs) Handles txtDelimiter.TextChanged
    Delimiter = txtDelimiter.Text
  End Sub

  Private Sub btnADDILite_Click(sender As Object, e As EventArgs) Handles btnADDILite.Click

    Dim start_time As DateTime = Now
    Dim stop_time As DateTime
    Dim elapsed_time As TimeSpan

    ' set up the base files Utilities and ControlLibrarues
    Dim UtilitiesFileName As String = folderPath & "\Utilities.txt"
    If Not File.Exists(UtilitiesFileName) Then
      MessageBox.Show("Caution! No Utilities.txt file found in folder:" & vbCrLf & folderPath)
      Utilities(0) = ""
    Else
      Utilities = File.ReadAllLines(UtilitiesFileName)
    End If

    ' Create a HyperLinkBuilder object to build hyperlinks for the Utilities
    hlBuilder = New HyperLinkBuilder(Utilities)

    Dim ControlLibrariesFileName As String = folderPath & "\ControlLibraries.txt"
    If Not File.Exists(ControlLibrariesFileName) Then
      MessageBox.Show("Caution! No ControlLibraries.txt file found in folder:" & vbCrLf & folderPath)
      ControlLibraries(0) = ""
    Else
      ControlLibraries = File.ReadAllLines(ControlLibrariesFileName)
    End If



    Delimiter = txtDelimiter.Text
    lblStatusMessage.Text = ""

    If Not Directory.Exists(folderPath & txtOutputFolder.Text) Then
      MessageBox.Show("OUTPUT folder name does not exist to write log file!" &
                      vbLf & folderPath & txtOutputFolder.Text)
      Exit Sub
    End If

    Dim logFileName As String = folderPath & txtOutputFolder.Text & "\ADDILite_log.txt"
    LogFile = My.Computer.FileSystem.OpenTextFileWriter(logFileName, False)
    LogFile.WriteLine(Date.Now & ",I,Program Starts," & Me.Text)
    LogFile.WriteLine(Date.Now & ",I,Data Gathering Form," & folderPath & "\" & txtDataGatheringForm.Text)
    LogFile.WriteLine(Date.Now & ",I,JOB Folder," & txtJCLJOBFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,PROC Folder" & txtProcFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,COBOL Folder," & txtCobolFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,Copybook Folder," & txtCopybookFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,DECLGEN Folder," & txtDECLGenFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,EASYTRIEVE," & txtEasytrieveFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,ASM Folder," & txtASMFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,TELON Folder," & txtTelonFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,Screen Map Folder," & txtScreenMapsFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,Output Folder," & txtOutputFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,PUML Folder," & txtPUMLFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,Expanded Folder," & txtExpandedFolder.Text)
    LogFile.WriteLine(Date.Now & ",I,Delimiter," & txtDelimiter.Text)
    LogFile.WriteLine(Date.Now & ",I,ScanModeOnly," & cbScanModeOnly.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption Jobs," & cbJOBS.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption Procs," & cbJobComments.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption Programs," & cbPrograms.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption Files," & cbFiles.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption Records," & cbRecords.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption Fields," & cbFields.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption Comments," & cbComments.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption EXECSQL," & cbexecSQL.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption EXECCICS," & cbexecCICS.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption IMS," & cbIMS.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption DataCom," & cbDataCom.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption ScreenMaps," & cbScreenMaps.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption Calls," & cbCalls.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption Libraries," & cbLibraries.Checked)
    LogFile.WriteLine(Date.Now & ",I,RunOption Instreams," & cbInstream.Checked)


    'validations
    If Not FileNamesAreValid() Then
      LogFile.WriteLine(Date.Now & ",S,Program Abnormally Ends,")
      LogFile.Close()
      MessageBox.Show("Folder/File Names are Not valid, see log.")
      Exit Sub
    End If

    ' Prepare for CallPgms file which holds all the Called Programs within the sources
    '  this file is processed as the last "JOB"
    CallPgmsFileName = folderPath & txtJCLJOBFolder.Text & "\CALLPGMS.JCL"
    ' Remove previous CallPgms.jcl file
    If File.Exists(CallPgmsFileName) Then
      Try
        File.Delete(CallPgmsFileName)
      Catch ex As Exception
        lblStatusMessage.Text = "Error deleting CallPgms.jcl file:" & ex.Message
        Exit Sub
      End Try
    End If

    ' Remove the previous #ADDI files
    lblStatusMessage.Text = "Removing previous ADDI files..."
    Dim dirInfo As New IO.DirectoryInfo(folderPath & txtCobolFolder.Text)
    Dim aryADDIFiles As IO.FileInfo() = dirInfo.GetFiles("#ADDI*.*", SearchOption.TopDirectoryOnly)
    For Each ADDIfile As IO.FileInfo In aryADDIFiles
      File.Delete(ADDIfile.FullName)
    Next

    ' Get the number of JOBS that will be processed
    NumberOfJobsToProcess = My.Computer.FileSystem.GetFiles(folderPath & txtJCLJOBFolder.Text).Count

    ' Load an AliasCobol array, if any
    lblStatusMessage.Text = "Loading Alias Cobol..."
    Dim AliasFileName As String = folderPath & "\COBOLAlias.csv"
    LogFile.WriteLine(Date.Now & ",I,Alias COBOL filename:," & AliasFileName)
    If Not File.Exists(AliasFileName) Then
      AliasCobol.Add("Empty", "Empty")
      LogFile.WriteLine(Date.Now & ",W,Alias COBOL file,Not Found")
    Else
      Dim AliasRows As String() = File.ReadAllLines(AliasFileName)
      For Each aliasrow In AliasRows
        Dim programidandfilename As String() = aliasrow.Split(",")
        If programidandfilename.Count > 1 Then
          If Not AliasCobol.ContainsKey(programidandfilename(0)) Then
            AliasCobol.Add(programidandfilename(0), programidandfilename(1))
          End If
        End If
      Next
      LogFile.WriteLine(Date.Now & ",I,Alias COBOL entries loaded:," & LTrim(Str(AliasCobol.Count)))
    End If

    ' ready the progress bar
    ProgressBar1.Minimum = 0
    ProgressBar1.Maximum = NumberOfJobsToProcess + 2
    ProgressBar1.Step = 1
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True

    Me.Cursor = Cursors.WaitCursor

    ' load the jobs to array list
    lblStatusMessage.Text = "Loading the JOBS..."
    LogFile.WriteLine(Date.Now & ",I,JCL Job files found," & LTrim(Str(NumberOfJobsToProcess)))
    For Each foundFile As String In My.Computer.FileSystem.GetFiles(folderPath & txtJCLJOBFolder.Text)
      ListOfJobs.Add(foundFile)
    Next

    lblStatusMessage.Text = "Initiating CloseXML..."

    ' Load the Data Gathering Form spreadsheet into the ListofDataGatheringForm array
    lblStatusMessage.Text = "Loading Data Gather Form to array..."
    Using dgfWorkbook As New XLWorkbook(folderPath & "\" & txtDataGatheringForm.Text)
      Dim dgfWorksheet As IXLWorksheet = dgfWorkbook.Worksheet(1)
      Dim lastRow As Integer = dgfWorksheet.LastRowUsed().RowNumber()
      For SummaryRow = 1 To lastRow
        Dim cellA = dgfWorksheet.Cell("A" & SummaryRow).GetValue(Of String)()
        If Val(cellA) > 0 Then
          Dim cellB = dgfWorksheet.Cell("B" & SummaryRow).GetValue(Of String)()
          Dim cellC = dgfWorksheet.Cell("C" & SummaryRow).GetValue(Of String)()
          ListOfDataGathering.Add(cellB & Delimiter & cellC)
        End If
      Next
    End Using
    SummaryRow = 0

    lblStatusMessage.Text = "Building cross-reference DB2 table names"
    'build a cross-reference table of DB2 Tablenames with source members
    For Each foundFile As String In My.Computer.FileSystem.GetFiles(folderPath & txtCobolFolder.Text)
      Dim memberLines As String() = File.ReadAllLines(foundFile)
      For index = 0 To memberLines.Count - 1
        If memberLines(index).IndexOf(" EXEC SQL DECLARE ") > -1 Then
          Dim srcWords As New List(Of String)
          Call GetSourceWords(memberLines(index), srcWords)
          MembersNames.Add(IO.Path.GetFileName(foundFile))
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

    ' Load the Screen Maps (Telon, IMS, CICS, or PC ScreenIO) files to array
    ' Decide which folder will hold the 'screen map' files. If its Telon files they are in
    '   the Telon folder. If other they will be in the Screens folder. This is because
    '   Telon members hold BOTH database info and Screen info. Other System (IMS/CICS) do not
    '   hold database info only screen info.
    Call LoadScreenMaps(folderPath & txtScreenMapsFolder.Text)
    Call LoadScreenMaps(folderPath & txtTelonFolder.Text)

    ' Build a list of source files so we don't have to use file exist function, just the list
    lblStatusMessage.Text = "Building list of COBOL source files..."
    Dim di As New IO.DirectoryInfo(folderPath & txtCobolFolder.Text)
    Dim aryFi As IO.FileInfo() = di.GetFiles("*.*")
    Dim fi As IO.FileInfo
    For Each fi In aryFi
      If fi.Name.ToUpper <> "DESKTOP.INI" Then
        ListofCOBOLFiles.Add(fi.Name.ToUpper)
      End If
    Next

    ' Build a list of copybook files so we don't have to use file exist function, just the list
    lblStatusMessage.Text = "Building list of copybook files..."
    Dim cbdi As New IO.DirectoryInfo(folderPath & txtCopybookFolder.Text)
    Dim aryCBFi As IO.FileInfo() = cbdi.GetFiles("*.*")
    Dim cbfi As IO.FileInfo
    For Each cbfi In aryCBFi
      If cbfi.Name.ToUpper <> "DESKTOP.INI" Then
        ListofCopybookFiles.Add(cbfi.Name.ToUpper)
      End If
    Next

    ' Build a list of Easytrieve files so we don't have to use file exist function, just the list
    lblStatusMessage.Text = "Building list of Easytrieve source files..."
    Dim eztDi As New IO.DirectoryInfo(folderPath & txtEasytrieveFolder.Text)
    Dim ezAryFiles As IO.FileInfo() = eztDi.GetFiles("*.*")
    Dim ezfile As IO.FileInfo
    For Each ezfile In ezAryFiles
      If ezfile.Name.ToUpper <> "DESKTOP.INI" Then
        ListofEasytrieveFiles.Add(ezfile.Name.ToUpper)
      End If
    Next

    ' Build a list of Assembler files so we don't have to use file exist function, just the list
    lblStatusMessage.Text = "Building list of Assembler source files..."
    Dim asmDi As New IO.DirectoryInfo(folderPath & txtASMFolder.Text)
    Dim asmAryFiles As IO.FileInfo() = asmDi.GetFiles("*.*")
    Dim asmfile As IO.FileInfo
    For Each asmfile In asmAryFiles
      If asmfile.Name.ToUpper <> "DESKTOP.INI" Then
        ListofAsmFiles.Add(asmfile.Name.ToUpper)
      End If
    Next

    ProgressBar1.PerformStep()
    ProgressBar1.Show()


    Dim ProgramsFileName = folderPath & txtOutputFolder.Text & "\ADDILite.xlsx"
    If File.Exists(ProgramsFileName) Then
      LogFile.WriteLine(Date.Now & ",I,Previous Model file deleted," & ProgramsFileName)
      Try
        File.Delete(ProgramsFileName)
      Catch ex As Exception
        LogFile.WriteLine(Date.Now & ",S,Error deleting Model," & ex.Message)
        lblStatusMessage.Text = "Error Deleting Model:" & ex.Message
        ProgressBar1.Visible = False
        Exit Sub
      End Try
    End If

    ' Create the Summary tab (aka Datagathering form details)
    CreateSummaryTab()

    ' Create, if any, all the in-stream data files as defined in the JOBS
    lblStatusMessage.Text = "Creating Instream data files from JOBS..."
    For Each JobFile In ListOfJobs
      Call CreateInStreamDataSets(JobFile)
    Next



    lblStatusMessage.Text = "Processing..."
    ' Process All the jobs in the JCL Folder.
    '  An addtional job could be created if should there be call subroutines
    Dim Jobcount As Integer = 0
    For Each JobFile In ListOfJobs
      Jobcount += 1
      FileNameOnly = IO.Path.GetFileNameWithoutExtension(JobFile)
      FileNameWithExtension = IO.Path.GetFileName(JobFile)
      lblProcessingJob.Text = "Processing Job #" & Jobcount & ": " & FileNameOnly
      LogFile.WriteLine(Date.Now & ",I,Processing Job," & IO.Path.GetFileNameWithoutExtension(JobFile))
      Call ProcessJOBFile(JobFile)
      Call ProcessExecFiles()
      ProgressBar1.PerformStep()
      ProgressBar1.Show()
      Call InitializeProgramVariables()
    Next

    Call ProcessCallPgms(CallPgmsFileName)

    Call CreateEXECSQLTab()
    Call CreateEXECCICSTab()
    Call CreateIMSTab()
    Call CreateDataComTab()
    Call CreateCallsTab()
    Call CreateIMSPSPNamesFile()
    Call CreateScreenMapTab()
    Call CreateLibrariesTab()
    Call CreateInstreamTab()


    ' Save Application Model Spreadsheet
    If Not cbScanModeOnly.Checked Then
      ' Format, Save and close Excel
      lblStatusMessage.Text = "Saving Spreadsheet"
      Call FormatWorksheets()
      workbook.SaveAs(ProgramsFileName)
      workbook.Dispose()
    End If
    GC.Collect()

    ProgressBar1.PerformStep()
    ProgressBar1.Show()

    stop_time = Now
    elapsed_time = stop_time.Subtract(start_time)

    LogFile.WriteLine(Date.Now & ",I,Program Ends," & elapsed_time.TotalMinutes.ToString("000.00") & " Minutes")
    LogFile.Close()
    Me.Cursor = Cursors.Default
    lblStatusMessage.Text = "Process Complete"

    System.Media.SystemSounds.Beep.Play()
    MessageBox.Show("Process Complete: " & elapsed_time.TotalMinutes.ToString("000.00") & " Minutes")
  End Sub

  Sub LoadScreenMaps(ByRef WorkFolder As String)
    ' This will load Screen maps from either TELON (SD/DR/BD), CICS, IMS, or PC(ScreenIO) to the ListOfScreenMaps array
    If WorkFolder.Length <= 0 Then
      Exit Sub
    End If

    For Each foundFile As String In My.Computer.FileSystem.GetFiles(WorkFolder)
      Dim memberLines As String() = File.ReadAllLines(foundFile)
      Dim srcWords As New List(Of String)
      Dim MapName As String = ""
      Dim PanelName As String = ""
      Dim PanelType As String = ""
      Dim literalsFound As String = ""
      Dim literalsFoundCnt As Integer = 0
      Dim literalsFoundMax As Integer = 6
      Dim MbrLine As String = ""
      Dim Text As String = ""
      ScreenType = ""

      For index = 0 To memberLines.Count - 1
        Text = memberLines(index) & Space(80)
        Text = Text.Substring(0, 72)
        MbrLine &= Text.Substring(0, 71).Trim
        Select Case Text.Substring(71, 1)
          Case "C", "X"
            Continue For
        End Select
        If MbrLine.Trim.Length = 0 Then
          Continue For
        End If

        ' Telon Screens pulled from (*SD, *DR, *BD)
        If ScreenType = "" Then
          If MbrLine.Trim.IndexOf("TELON") > -1 Then
            ScreenType = "TELON"
            MbrLine = ""
            Continue For
          End If
        End If
        If ScreenType = "TELON" And MbrLine.Substring(0, 1) <> "*" Then
          Call GetSourceWords(MbrLine, srcWords)
          Select Case srcWords(0)
            Case "SCREEN", "DRIVER", "BATCH"
              Dim screenWords As String() = srcWords(1).Split(",")
              MapName = screenWords(0)
              For Each screenWord In screenWords
                If screenWord.StartsWith("DESC='") Then
                  literalsFound = screenWord.Substring(6).Replace("'", "").Trim
                  literalsFoundCnt = 1
                  Exit For
                End If
              Next
              If literalsFound.Length > 45 Then
                literalsFound = AddNewLineAboutEveryNthCharacters(literalsFound, vbNewLine, 45)
              End If
              ListofScreenMaps.Add(IO.Path.GetFileName(foundFile) & Delimiter &
                                  ScreenType & Delimiter &
                                  MapName & Delimiter &
                                  literalsFound)
              Exit For
          End Select
        End If

        ' IMS screens
        If ScreenType = "" Then
          Call GetSourceWords(MbrLine, srcWords)
          If srcWords.Count >= 2 Then
            If srcWords(1) = "FMT" Then
              ScreenType = "IMS"
              MapName = srcWords(0)
              MbrLine = ""
              Continue For
            End If
          End If
        End If
        If ScreenType = "IMS" Then
          Call GetSourceWords(MbrLine, srcWords)
          If srcWords.Count >= 2 Then
            Dim startIndex As Integer = -1
            If srcWords(0) = "DFLD" Then
              startIndex = 1
            End If
            If srcWords(1) = "DFLD" Then
              startIndex = 2
            End If
            If startIndex > -1 Then
              If srcWords(startIndex).StartsWith("'") Then
                literalsFoundCnt += 1
                If literalsFoundCnt <= literalsFoundMax Then
                  Dim quoteWords As String() = srcWords(startIndex).Split("'")
                  literalsFound &= quoteWords(1).Trim & vbNewLine
                End If
              End If
            End If
          End If
          If MbrLine = "END" Or literalsFoundCnt > literalsFoundMax Then
            If literalsFound.Length > 2 Then
              literalsFound = literalsFound.Substring(0, literalsFound.Length - 2)    'remove last vbNewLine
            End If
            If literalsFound.Length > 45 Then
              literalsFound = AddNewLineAboutEveryNthCharacters(literalsFound, vbNewLine, 45)
            End If
            ListofScreenMaps.Add(IO.Path.GetFileName(foundFile) & Delimiter &
                                    ScreenType & Delimiter &
                                    MapName & Delimiter &
                                    literalsFound)
            If ListOfIMSMapNames.IndexOf(MapName) = -1 Then
              ListOfIMSMapNames.Add(MapName)
            End If
            Exit For
          End If
          MbrLine = ""
          Continue For
        End If

        ' CICS Screens
        If ScreenType = "" Then
          Call GetSourceWords(MbrLine, srcWords)
          If srcWords.Count >= 2 Then
            If srcWords(1) = "DFHMSD" Then
              ScreenType = "CICS"
              MapName = srcWords(0)
              MbrLine = ""
              literalsFoundCnt = 0
              literalsFound = ""
              Continue For
            End If
          End If
        End If
        If ScreenType = "CICS" Then
          Dim InitStringPosition As Integer = MbrLine.IndexOf("INITIAL='")
          If InitStringPosition > -1 Then
            Dim InitStringValue As String = MbrLine.Substring(InitStringPosition)
            literalsFoundCnt += 1
            If literalsFoundCnt <= literalsFoundMax Then
              Dim quoteWords As String() = InitStringValue.Split("'")
              literalsFound &= quoteWords(1).Trim & " "
            End If
          End If
          If literalsFoundCnt > literalsFoundMax Or MbrLine = "END" Then
            literalsFound = literalsFound.Trim
            If literalsFound.Length > 45 Then
              literalsFound = AddNewLineAboutEveryNthCharacters(literalsFound, vbNewLine, 45)
            End If
            ListofScreenMaps.Add(IO.Path.GetFileName(foundFile) & Delimiter &
                                  ScreenType & Delimiter &
                                  MapName & Delimiter &
                                  literalsFound)
            Exit For
          End If
        End If

        ' PC Screens (ScreenIO)
        If ScreenType = "" Then
          If Mid(MbrLine, 1, 6) <> Space(6) Then
            MbrLine = memberLines(index) & Space(80)
          End If
          If MbrLine.Length >= 70 Then
            ' store map/panel name and/or screen type (Pop-Up, Window, Main, etc.)
            If MbrLine.Substring(6, 4) = "*:- " Then
              ScreenType = "SCREENIO"
              If MbrLine.Substring(21, 12) = "Panel Name: " Then
                PanelName = MbrLine.Substring(33, 11).Trim
              End If
              If MbrLine.Substring(45, 11) = "PanelType: " Then
                PanelType = MbrLine.Substring(56, 12).Trim
              End If
              Continue For
            End If
          End If
        End If
        If ScreenType = "SCREENIO" Then
          If Mid(MbrLine, 1, 6) <> Space(6) Then
            MbrLine = memberLines(index) & Space(80)
          End If
          ' Get any comments when line has '*:> '
          If MbrLine.Substring(6, 4) = "*:> " Then
            literalsFound &= MbrLine.Substring(10, 60).Trim & vbNewLine
          End If
          ' store the screen map and exit loop when line has '*:+ '
          If MbrLine.Substring(6, 4) = "*:+ " Then
            If literalsFound.Length > 2 Then
              literalsFound = literalsFound.Substring(0, literalsFound.Length - 2)    'remove last vbNewLine
            End If
            If literalsFound.Length > 45 Then
              literalsFound = AddNewLineAboutEveryNthCharacters(literalsFound, vbNewLine, 45)
            End If
            ListofScreenMaps.Add(IO.Path.GetFileName(foundFile) & Delimiter &
                                PanelType & Delimiter &
                                PanelName & Delimiter &
                                literalsFound)
            Exit For
          End If
        End If

        MbrLine = ""
      Next index
    Next


  End Sub
  Sub ProcessCallPgms(ByRef CallPgmsFileName As String)
    If ListOfCallPgms.Count = 0 Then
      Exit Sub
    End If

    ' create the CallPgms.jcl file based on ListOfCallPgms
    swCallPgmsFile = New StreamWriter(CallPgmsFileName, False)
    Dim pgmCnt As Integer = 0
    swCallPgmsFile.WriteLine("//CALLPGMS JOB 'INTERNAL','SUBROUTINES CALLED'")
    For Each callpgm In ListOfCallPgms
      Dim execs As String() = callpgm.Split(Delimiter)
      If execs(3) = "Dynamic" Then  'do not analyze a dynamic call as we don't know name of program
        Continue For
      End If
      pgmCnt += 1
      swCallPgmsFile.WriteLine("//" & execs(2) & " EXEC PGM=" & execs(0).Replace(Delimiter, ""))
      swCallPgmsFile.WriteLine("//STEPLIB DD DSN=LIBRARY." & execs(1) & ",DISP=SHR")
    Next
    swCallPgmsFile.Close()

    'process the CallPgms file
    If File.Exists(CallPgmsFileName) Then
      lblProcessingJob.Text = "Processing CallPgms"
      LogFile.WriteLine(Date.Now & ",I,Processing Job," & CallPgmsFileName)
      FileNameOnly = IO.Path.GetFileNameWithoutExtension(CallPgmsFileName)
      Call ProcessJOBFile(CallPgmsFileName)
      Call ProcessExecFiles()
      ProgressBar1.PerformStep()
      ProgressBar1.Show()
      Call InitializeProgramVariables()
    Else
      LogFile.WriteLine(Date.Now & ",E,Call Pgms File not found?," & CallPgmsFileName)
    End If

  End Sub
  Function FileNamesAreValid() As Boolean
    FileNamesAreValid = False
    Select Case True
      Case txtDataGatheringForm.TextLength = 0
        LogFile.WriteLine(Date.Now & ",E,ERROR! Data Gathering Form name required,")
      Case Not IsValidFileNameOrPath(folderPath & "\" & txtDataGatheringForm.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! Data Gathering Form name has invalid characters," &
                          folderPath & "\" & txtDataGatheringForm.Text)
      Case Not File.Exists(folderPath & "\" & txtDataGatheringForm.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! Data Gathering Form not found," &
                          folderPath & "\" & txtDataGatheringForm.Text)

      Case txtJCLJOBFolder.TextLength = 0
        LogFile.WriteLine(Date.Now & ",E,ERROR! JCL JOBS Folder name required,")
      Case Not IsValidFileNameOrPath(folderPath & txtJCLJOBFolder.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! JCL JOBS Folder name has invalid characters,")
      Case Not Directory.Exists(folderPath & txtJCLJOBFolder.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! JCL JOBS folder does not exists,")

      Case txtCobolFolder.TextLength = 0
        LogFile.WriteLine(Date.Now & ",E,ERROR! Sources folder name required,")
      Case Not IsValidFileNameOrPath(folderPath & txtCobolFolder.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! Sources folder name has invalid characters,")
      Case Not Directory.Exists(folderPath & txtCobolFolder.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! Sources folder does not exists,")

      Case txtTelonFolder.TextLength = 0
        LogFile.WriteLine(Date.Now & ",E,ERROR! TELON folder name required,")
      Case Not IsValidFileNameOrPath(folderPath & txtTelonFolder.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! TELON folder name has invalid characters,")
      Case Not Directory.Exists(folderPath & txtTelonFolder.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! TELON folder name does not exists,")

      Case txtScreenMapsFolder.TextLength = 0
        LogFile.WriteLine(Date.Now & ",E,ERROR! SCREENS folder name required,")
      Case Not IsValidFileNameOrPath(folderPath & txtScreenMapsFolder.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! SCREENS folder name has invalid characters,")
      Case Not Directory.Exists(folderPath & txtScreenMapsFolder.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! SCREENS folder does not exists,")

      Case txtOutputFolder.TextLength = 0
        LogFile.WriteLine(Date.Now & ",E,ERROR! OUTPUTS folder name required,")
      Case Not IsValidFileNameOrPath(folderPath & txtOutputFolder.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! OUTPUTS folder has invalid characters,")
      Case Not Directory.Exists(folderPath & txtOutputFolder.Text)
        LogFile.WriteLine(Date.Now & ",E,ERROR! OUTPUTS folder does not exists,")

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

  Sub ProcessJOBFile(JobFile As String)
    'Load the Jobfile to the jclStmt List
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

    Dim WriteOutputResult As Integer = ProcessJCL()
    If WriteOutputResult = -1 Then
      MessageBox.Show("Error while building output. See log file")
    End If

  End Sub
  Function LoadJCLStatementsToArray(JobFile As String) As Integer '
    ' Clear out the JCL Statement array
    jclStmt.Clear()

    ' Load the JOB File into the array
    Dim JCL As New List(Of String)
    JCL = ReformatJCLAndLoadToArray(JobFile)

    ' PROCLIB member inclusion.
    ' Read all lines in the JOB array:
    ' if line has a PROC command (eg. //STEP PROC <name>) this is an instream PROC we'll ignore
    '   it and store its name as an INSTREAM PROC.
    ' if line has a EXEC command (eg. //COPY01   EXEC  DHSIMAGV,AUTOP1='PA903PA'), without a 'PGM=' parameter 
    '   this is an execute of a PROC and this is where to include/find the proc in PROCLIB folder.
    ' somehow need to remove the EXEC proc (and its continuation lines).
    ' Only going one level deep. eg, we are not supporting procs within procs within procs...
    '
    Dim jclWithProc As New List(Of String)
    Dim ListOfInstreamProcs As New List(Of String)
    Dim JCLParms As New Dictionary(Of String, String)
    Dim Command As Integer = 1
    Dim Label As Integer = 0
    Dim Parameters As Integer = 2
    Dim jclProcName As String = ""
    Dim JustLoadedPROC As Boolean = False
    For Each JCLLine In JCL
      Dim JCLStatement As String() = JCLLine.Split(Delimiter)
      ' IF JCLline is not PROC or EXEC but is part of a previous PROC load
      '   then put '++' to indicate part of the PROC
      '   else leave as is
      If JustLoadedPROC Then
        Select Case JCLStatement(Command)
          Case "DD", "COMMENT"
            Mid(JCLLine, 1, 2) = "++"
        End Select
      End If
      ' Place the JCLLine to the array 
      jclWithProc.Add(JCLLine)
      Select Case JCLStatement(Command)
        Case "PROC"
          'this must be an instream PROC
          jclProcName = JCLStatement(Label).Substring(2)
          ListOfInstreamProcs.Add(jclProcName)
          JustLoadedPROC = False
        Case "EXEC"
          JustLoadedPROC = False
          If JCLStatement(Parameters).Substring(0, 4) = "PGM=" Then
            Exit Select
          End If
          'this must be an exec <procname>, so need to load this proc here, IF not an instream proc
          '-need to save the JCL Parms for later substitue.
          Dim ParmValues As String() = JCLStatement(Parameters).Split(",")
          JCLParms = LoadJCLParms(ParmValues)
          Dim ProcName As String = folderPath & txtProcFolder.Text & "\" & ParmValues(0)
          If ListOfInstreamProcs.IndexOf(ParmValues(0)) = -1 Then
            Dim PROC As New List(Of String)
            PROC = ReformatJCLAndLoadToArray(ProcName)
            If PROC.Count = 0 Then
              LogFile.WriteLine(Date.Now & ",E,Missing PROC member," & ParmValues(0))
            End If
            For Each ProcLine In PROC
              Dim ProcLinePlus As String = "++" & ProcLine.Substring(2) 'replace leading // with ++ to indicate PROC
              Dim ProcLineParmsUpdated As String = ReplaceProcLineParms(ProcLinePlus, JCLParms)
              jclWithProc.Add(ProcLineParmsUpdated)
              JustLoadedPROC = True
            Next
          End If
      End Select
    Next

    jclStmt.AddRange(jclWithProc)
    LoadJCLStatementsToArray = jclStmt.Count

  End Function
  Function LoadJCLParms(ByRef ParmValues As String()) As Dictionary(Of String, String)
    'Load the JCL Parameters the JCL Line. The first occurence is not a parameter but a PROC name
    'i.e., DSNEXEC3,PGMLIB='PRD1.LINKLIB',PROGRAM=INSB610,SYSTEM=DSN 
    '   would return 3 key/value entries of PGMLIB, PROGRAM, and SYSTEM
    ' if the first ParmValues string has PROC= then remove it. 
    ParmValues(0) = ParmValues(0).Replace("PROC=", "")
    ' split the out the parmvalues to key/value pairs
    Dim theJCLParms As New Dictionary(Of String, String)
    For x As Integer = 1 To ParmValues.Count - 1 Step 1
      Dim KeyAndValue As String() = ParmValues(x).Split("=")
      If KeyAndValue.Count < 2 Then   'no key/value
        Continue For
      End If
      Dim theKey As String = "&" & KeyAndValue(0)   'place an & in front of keyword for later searching
      Dim theValue As String = KeyAndValue(1)
      If Not theJCLParms.ContainsKey(theKey) Then
        theJCLParms.Add(theKey, theValue)
      End If
    Next
    Return theJCLParms
  End Function
  Function ReplaceProcLineParms(ByRef ProcLinePlus As String, ByRef JCLParms As Dictionary(Of String, String)) As String
    ' replace the keyword=&parmValue or keyword='&parmValue' with the value from the JCLParms array
    ' split by delimiter/pipe
    Dim PROCStatement As String() = ProcLinePlus.Split(Delimiter)
    If PROCStatement.Count <= 2 Then
      Return ProcLinePlus
    End If

    ' do not substitue for PROC or COMMENT lines
    Select Case PROCStatement(1)
      Case "PROC", "COMMENT"
        Return ProcLinePlus
    End Select
    ' substitute jcl parameter
    Dim theLineReplaced As String = ProcLinePlus
    For Each pair As KeyValuePair(Of String, String) In JCLParms
      theLineReplaced = theLineReplaced.Replace(pair.Key, pair.Value) 'replace the value for the keyword
    Next
    Return theLineReplaced
  End Function
  Function ReformatJCLAndLoadToArray(ByRef Jobfile As String) As List(Of String)
    ' Load a JCL file to an Array which has
    ' -Remove continuations
    ' -drop Comments
    ' -keep lines only with '//' , '++', '/*'
    ' -parsed out as Label, Command, Parameters with Delimiter
    Dim JCL As New List(Of String)
    If Not File.Exists(Jobfile) Then
      ReformatJCLAndLoadToArray = JCL
      Exit Function
    End If

    Dim text1 As String = ""
    Dim continuation As Boolean = False
    Dim jStatement As String = ""
    Dim jclWords As New List(Of String)
    Dim comment As String = ""

    Dim JCLLines As String() = File.ReadAllLines(Jobfile)
    For Each JCLLine In JCLLines
      text1 = JCLLine.Replace(vbTab, Space(1))
      ' drop data (of an DD * statement) or not a JCL statement
      Select Case Mid(text1, 1, 2)
        Case "//", "++", "/*"
        Case Else
          Continue For
      End Select
      ' Keep columns 1-72, remove columns 73-80
      text1 = Microsoft.VisualBasic.Left(Mid(text1, 1) + Space(80), 72)
      ' remove '+' in column 72 (which used to mean continuation?)
      If Len(text1) >= 72 Then
        If Mid(text1, 72, 1) = "+" Then
          Mid(text1, 72, 1) = " "
        End If
      End If
      ' store the comments
      If Mid(text1, 1, 3) = "//*" Or Mid(text1, 1, 3) = "++*" Then
        If text1.IndexOf("//*PRODUCTION ") > -1 Then
          Continue For
        End If
        If text1.IndexOf("//*REP ") > -1 Then
          Continue For
        End If
        comment = text1.Replace("*", "").Replace("//", "").Replace("++", "").Trim
        If comment.Length = 0 Then
          Continue For
        End If
        JCL.Add(Mid(text1, 1, 2) & "*" & Delimiter & "COMMENT" & Delimiter & comment.Replace(Delimiter, " ").Trim)
        Continue For
      Else
        ' Drop simple IF statements in JCL
        If text1.IndexOf(" IF ") > -1 Then
          Continue For
        End If
        ' Drop simple ENDIF statements in JCL
        If text1.IndexOf(" ENDIF ") > -1 Then
          Continue For
        End If
        ' Drop the INCLUDE statement in JCL
        If text1.IndexOf(" INCLUDE ") > -1 Then
          Continue For
        End If
      End If
      ' remove leading slashes if this line is a continuation
      If continuation = True Then
        text1 = Trim(Mid(text1, 3))
      Else
        text1 = Trim(Mid(text1, 1))
      End If
      ' determine if there will be a continuation
      text1 &= Space(1)
      continuation = JCLContinued(text1, continuation)
      ' Build the JCL statement
      jStatement &= text1
      ' if NOT continuing building of the JCL statement then add it to the List
      If continuation = False Then
        If jStatement.Trim.Length > 0 Then
          GetJCLWords(jStatement, jclWords)
          jclWords(0) = AdjustProcName(jclWords, Jobfile)
          Select Case jclWords.Count
            Case 1
              Select Case jclWords(0)
                Case "//", "/*"
                Case Else
                  JCL.Add(jclWords(0) & Delimiter & Delimiter)
              End Select
            Case 2
              JCL.Add(jclWords(0) & Delimiter & jclWords(1) & Delimiter)
            Case 3
              JCL.Add(jclWords(0) & Delimiter & jclWords(1) & Delimiter & jclWords(2))
          End Select
        End If
        jStatement = ""
      End If
    Next
    ReformatJCLAndLoadToArray = JCL
  End Function
  Function AdjustProcName(ByRef theJCLwords As List(Of String), ByRef theSourceFile As String) As String
    If theJCLwords.Count < 2 Then
      Return theJCLwords(0)
    End If
    If theJCLwords(1) <> "PROC" Then
      Return theJCLwords(0)
    End If
    ' strip off just the filename with extension (if any), drop the pathinfo
    Dim theFileName = IO.Path.GetFileName(theSourceFile)
    Dim theProcName = theJCLwords(0).Replace("//", "")
    If theProcName = theFileName Then
      Return theJCLwords(0)
    End If
    ' On the PROC statement, the Proc Source Name and Proc Name is different, adjust to the source name
    LogFile.WriteLine(Date.Now & ",I,PROC Name adjusted to PROC Source Name," &
                      theJCLwords(0) & " vs " & theFileName)
    Return "//" & theFileName
  End Function
  Function JCLContinued(ByRef text As String, ByVal withinContinuation As Boolean) As Boolean
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
      If Mid(text, x, 2) = Space(2) And withinContinuation = True Then
        text = Mid(text, 1, x)                        'remove anything to the right of non continuation space+space
        JCLContinued = False
        Exit Function
      End If
      If Mid(text, x, 2) = ", " Then
        text = Mid(text, 1, x)                        'remove anything to the right of continuation comma+space
        JCLContinued = True
        Exit Function
      End If
    Next
  End Function

  Sub CreateInStreamDataSets(ByRef JobFile As String)
    'This will scan for 'DD *' JCL statement and create a source file with pattern
    '  <JOB name>_<Step name>_<DD name>
    ' This will ignore any 'DD DATA' JCL statements for now...
    ' Input: JOBS folder name (global variable)
    '        SOURCES folder name (global variable)
    ' Output: files written to the SOURCES folder name
    FileNameOnly = IO.Path.GetFileNameWithoutExtension(JobFile)
    FileNameWithExtension = IO.Path.GetFileName(JobFile)
    Dim JCLLines As String() = File.ReadAllLines(JobFile)
    Dim JCLWords As New List(Of String)
    Dim execIndex As Integer = -1
    jobName = ""
    stepName = ""
    DDName = ""
    For JCLIndex As Integer = 0 To JCLLines.Count - 1
      Dim JCLLine As String = JCLLines(JCLIndex)
      Dim tWord = JCLLine.Split(" ")
      JCLWords.Clear()
      For Each JCLword In tWord
        If JCLword.Trim.Length > 0 Then        'dropping empty words
          JCLWords.Add(JCLword.ToUpper)
        End If
      Next
      If JCLWords.Count < 3 Then
        Continue For
      End If
      If JCLWords(0).Length < 3 Then
        Continue For
      End If
      If JCLWords(0).Substring(0, 3) = "//*" Then
        Continue For
      End If
      If JCLWords(0).Length < 2 Then
        Continue For
      End If
      If JCLWords(0).Substring(0, 2) = "/*" Then
        Continue For
      End If
      Select Case JCLWords(1)
        Case "JOB"
          jobName = JCLWords(0).Replace("//", "").Trim()
        Case "EXEC"
          stepName = JCLWords(0).Replace("//", "").Trim()
          If JCLWords(2).IndexOf("PGM=EZTPA00") > -1 Then
            pgmName = "EZTPA00"
            execIndex = JCLIndex
          Else
            pgmName = ""
          End If
        Case "DD"
          DDName = JCLWords(0).Replace("//", "").Trim()
          If DDName.IndexOf(".") > -1 Then
            Dim stepanddd = DDName.Split(".")
            If stepanddd.Count >= 2 Then
              stepName = stepanddd(0)
              DDName = stepanddd(1)
            End If
            ' need to find pgmName for this DD Override by using the stepName
            pgmName = FindExecPgmName(JCLLines)
          End If
          If JCLWords(2) = "*" Then
            Call CreateInstreamDataset(JCLLines, JCLIndex)
            Continue For
          End If
          ' in case it is easytrieve and a SYSIN DD DSN
          If pgmName = "EZTPA00" And DDName = "SYSIN" Then
            Dim memberName As String = GrabPDSMemberName(JCLWords(2))
            If memberName.Length = 0 Then
              Continue For
            End If
            Dim myKey As String = FileNameOnly & "_" & jobName & "_" & stepName
            ListOfEasytrieveLoadAndGo.Add(myKey, memberName)
          End If
      End Select
    Next
  End Sub
  Function FindExecPgmName(ByRef JCLLines As String()) As String
    Dim JCLWords As New List(Of String)
    For JCLIndex As Integer = 0 To JCLLines.Count - 1
      Dim JCLLine As String = JCLLines(JCLIndex)
      Dim tWord = JCLLine.Split(" ")
      JCLWords.Clear()
      For Each JCLword In tWord
        If JCLword.Trim.Length > 0 Then        'dropping empty words
          JCLWords.Add(JCLword.ToUpper)
        End If
      Next
      If JCLWords.Count < 3 Then
        Continue For
      End If
      If JCLWords(0).Length < 3 Then
        Continue For
      End If
      If JCLWords(0).Substring(0, 3) = "//*" Then
        Continue For
      End If
      If JCLWords(0).Length < 2 Then
        Continue For
      End If
      If JCLWords(0).Substring(0, 2) = "/*" Then
        Continue For
      End If
      Select Case JCLWords(1)
        Case "EXEC"
          Dim findStepName As String = JCLWords(0).Replace("//", "").Trim()
          If findStepName = stepName Then
            Return GetParmPGM(JCLWords(2))
          End If
      End Select
    Next
    Return ""
  End Function
  Sub CreateInstreamDataset(ByRef JCLLines As String(), ByRef JCLIndex As Integer)
    ' This will write out the instream data set
    ' Input fields (Global): FileNameOnly, JobName StepName, DDName
    '      JCLLines() argument
    ' Output file will be named: <filenameonly>_<jobname>_<stepname>_<ddname>
    numLoadAndGo += 1
    Dim firstPartFileName As String = "#ADDI"
    Dim InstreamDatasetFileName = folderPath & txtCobolFolder.Text & "\" & firstPartFileName & LTrim(Str(numLoadAndGo))
    swInstreamDatasetFile = My.Computer.FileSystem.OpenTextFileWriter(InstreamDatasetFileName, False)

    ' Write the data after the 'DD *' until we reach a '//' or '/*' or end of array
    Dim isKey As String = FileNameOnly & Delimiter &
      stepName & Delimiter &
      LTrim(Str(numLoadAndGo))
    Dim isValue As String = ""
    For JCLIndex = JCLIndex + 1 To JCLLines.Count - 1
      If JCLLines(JCLIndex).Length >= 2 Then
        Select Case JCLLines(JCLIndex).Substring(0, 2)
          Case "//", "/*"
            Exit For
        End Select
        swInstreamDatasetFile.WriteLine(JCLLines(JCLIndex))
        isValue &= JCLLines(JCLIndex) & vbLf
      End If
    Next
    If isValue.Length > 0 Then
      DictOfInstreams.Add(isKey, isValue)
    End If
    If JCLIndex < JCLLines.Count - 1 Then
      JCLIndex -= 1
    End If
    swInstreamDatasetFile.Close()
    If pgmName = "EZTPA00" And DDName = "SYSIN" Then
      Dim myKey As String = FileNameOnly & "_" & jobName & "_" & stepName
      Dim myValue As String = "#ADDI" & LTrim(Str(numLoadAndGo))
      ListOfEasytrieveLoadAndGo.Add(myKey, myValue)
    End If
  End Sub
  Function GrabPDSMemberName(ByRef text As String) As String
    Dim OpenParenIndex As Integer = text.IndexOf("(")
    Dim CloseParenIndex As Integer = text.IndexOf(")")
    If OpenParenIndex < 0 Or CloseParenIndex < 0 Then
      Return ""
    End If
    Return text.Substring(OpenParenIndex + 1, (CloseParenIndex - OpenParenIndex - 1))
  End Function
  Sub LogStmtArray(ByRef theFileName As String, theStmtArray As List(Of String))
    ' write the stmt array to a file for debugging purposes
    Dim logStmtCount As Integer = -1
    Dim logStmtFileName As String = folderPath & txtOutputFolder.Text &
      "\Debug\" & theFileName & "_logStmt.txt"
    Try
      LogStmtFile = My.Computer.FileSystem.OpenTextFileWriter(logStmtFileName, False)
    Catch ex As Exception
      LogFile.WriteLine(Date.Now & ",E,Error creating Log stmt file," & ex.Message)
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
    ' example where statement is:" DISPLAY '*** CRCALCX REC READ        = ' WS-REC-READ.   "
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
  Function RemoveGDGValuesFromDSN(ByRef dsn As String) As String
    ' This will remove the (0), (+1), etc. numeric values from the dsn and return with just '()' empty parens
    ' Hopefully there is not a PDS with a numeric member name. 
    Dim OpenPIndex As Integer = dsn.IndexOf("(")
    Dim ClosePIndex As Integer = dsn.IndexOf(")")
    If OpenPIndex = -1 Or ClosePIndex = -1 Then
      Return dsn
    End If
    Dim lenWithinParens As Integer = ClosePIndex - OpenPIndex - 1
    ' check if any value between parens ie. it is '()'
    If lenWithinParens < 1 Then
      Return dsn
    End If
    Dim valueWithinParens As String = dsn.Substring(OpenPIndex + 1, lenWithinParens)
    If Not IsNumeric(valueWithinParens) Then
      Return dsn
    End If
    Return dsn.Substring(0, OpenPIndex) & "()"
  End Function
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
    Dim commaLocation As Integer = parameter.IndexOf(",")
    Select Case commaLocation
      Case -1
        Return RTrim(parameter)
      Case Else
        Return Microsoft.VisualBasic.Left(parameter, commaLocation)
    End Select
  End Function
  Function GetSecondJobParm(parameter As String) As String
    ' Extract the second parameter from the JOB Parameters
    ' This presumes NO commas within quotes!
    ' This presumes JOB card is valid syntax. There must be Acct Info, and Programmer-name parms minimum.
    Dim jobwords As String() = parameter.Split(",")
    If jobwords.Length = 0 Then
      Return ""
    End If
    If jobwords.Count >= 2 Then
      Return jobwords(1)
    End If
    Return ""
  End Function

  Function ProcessJCL() As Integer
    ' Store the details to the spreadsheet tabs.
    ' return of -1 means an error (at this time, no errors identified)
    ' return of 0 means all is okay

    ListOfDDs.Clear()       'in lieu of the swDDFile

    Dim ListOfSymbolics As New List(Of String)

    JobSourceName = FileNameOnly

    ' Values found on the JOB card
    jobName = ""
    jobClass = ""
    jobMsgClass = ""
    JobAccountInfo = ""
    JobProgrammerName = ""
    JobTime = ""
    JobCond = ""
    JobTyprun = ""
    JobRegion = ""
    ' Values found on the Jes2, route and send card(s)
    JobSend = ""
    JobRoute = ""
    JobParm = ""
    JobJCLLib = ""
    JobLib = ""

    For Each statement As String In jclStmt
      If statement.Substring(0, 2) = "//" Then
        procName = ""                 'not within a PROC
      End If
      Call GetLabelControlParms(statement, jLabel, jControl, jParameters)
      If Len(jControl) = 0 Then
        LogFile.WriteLine(Date.Now & ",E,JCL control not found,'" & statement & ": " & FileNameOnly & "'")
        Continue For
      End If

      Select Case jControl
        Case "COMMENT"
        Case "JOB"
          Call ProcessJOB()
        Case "PROC"
          procName = jLabel         'instream proc?
          ListOfSymbolics = LoadSymbolics(jParameters)
        Case "PEND"
        Case "EXEC"
          Call ProcessEXEC(True, ListOfSymbolics)
        Case "DD"
          Call ProcessDD(ListOfSymbolics)          'this writes the _dd.csv record
        Case "SET"
          Continue For
        Case "OUTPUT"
          JobSend = jParameters
        Case "IF"
          Continue For
        Case "ENDIF"
          Continue For
        Case "JCLLIB"
          Call ProcessJCLLIB()
        Case "EOF"

        Case Else
          If jLabel = "/*ROUTE" Then
            JobRoute = jControl & " " & jParameters
            Continue For
          End If
          If jLabel = "/*JOBPARM" Then
            JobParm = jControl
            Continue For
          End If
          If jLabel = "/*" Then
            Continue For
          End If
          If jLabel = "IF" And jControl = "RC" Then
            Continue For
          End If
          If jParameters = "PEND" Then
            Continue For
          End If
          If jParameters = "PROC" Then
            Continue For
          End If
          LogFile.WriteLine(Date.Now & ",E,Unknown JCL Control Value," & statement.Replace(",", ";") & " file:" & FileNameOnly)
          Continue For
      End Select

    Next


    Call CreateJobsTab()

    Call CreateJobCommentsTab(ListOfSymbolics)

    Call CreateProgramsTab()
    Call CreateFilesTab()

    Return 0
  End Function
  Function LoadSymbolics(ByRef jParameters As String) As List(Of String)
    Dim jparms As String() = jParameters.Split(",")
    Dim jsymbols As New List(Of String)
    For Each jvalue In jparms
      jsymbols.Add(jvalue)
    Next
    Return jsymbols
  End Function

  Sub GetLabelControlParms(statement As String,
                           ByRef jLabel As String,
                           ByRef jControl As String,
                           ByRef jParameters As String)
    'This will split out the three basic components of a JCL Statement
    'Each statement is expected to have either JOB, PROC, EXEC, DD, PEND, SET, OUTPUT, IF, ENDIF, etc.
    'Enforcement of JCL syntax is not done here except that there must be a space between
    ' Label, Control and Parmeters. There may not be a label, so adjustments are made
    'get previous jLabel by looking at last entry of the listofdds array

    Dim jLabelPrev As String = jLabel

    Dim jclWords As String() = statement.Split(Delimiter)

    ' remove the /* end of instream data indicator
    If jclWords(0) = "/*" Then
      jLabel = "/*"
      jControl = "EOF"
      jParameters = ""
      Exit Sub
    End If

    jLabel = jclWords(0)
    jControl = jclWords(1)
    jParameters = jclWords(2)

    If jLabel = "/*JOBPARM" Or jLabel = "/*ROUTE" Then
      Exit Sub
    End If

    jLabel = jLabel.Remove(0, 2) 'remove the  leading // or ++ symbols
    If Len(jLabel) = 0 Then
      jLabel = jLabelPrev
    End If

  End Sub
  Sub ProcessJOB()
    ' Extract out values from the JCL JOB card
    procName = ""
    ddSequence = 0
    jobName = jLabel
    jobMsgClass = GetParm(jParameters, "MSGCLASS=")
    jobClass = GetParm(jParameters, "CLASS=")
    JobTime = GetParm(jParameters, "TIME=")
    JobRegion = GetParm(jParameters, "REGION=")
    JobCond = GetParm(jParameters, "COND")
    JobAccountInfo = GetFirstParm(jParameters).Replace("'", "").Trim
    JobProgrammerName = GetSecondJobParm(jParameters).Replace("'", "").Trim

    InstreamProc = ""
  End Sub
  Sub ProcessJCLLIB()
    ' Grab/format the JCLLIB value(s)
    JobJCLLib = jParameters
    ' add libraries to array
    Dim jcllibs As String() = jParameters.Replace("ORDER=", "").Replace("(", "").Replace(")", "").Split(",")
    For Each jcllib In jcllibs
      If ListOfLibraries.IndexOf(jcllib & Delimiter & "JCLLIB") = -1 Then
        ListOfLibraries.Add(jcllib & Delimiter & "JCLLIB")
      End If
    Next
  End Sub
  Sub ProcessEXEC(ByVal NeedSourceType As Boolean, ListOfSymbolics As List(Of String))
    ' The "EXEC" control is for either PROC or a PGM
    ' For PROC it could be "EXEC <procname>" or "EXEC PROC=<procname"
    '   when it is a PROC we set the Procname and exit this routine.
    ' For PGM it is "EXEC PGM=<pgmname>"
    ' For PGM replace program name if it is a symbolic

    ' If by this time we haven't gotten a JOBname then set the JOB details
    If jobName.Length = 0 Then
      jobName = FileNameOnly
      jobClass = "?"
      jobMsgClass = "?"
    End If
    '
    stepName = jLabel

    ddSequence = 0
    execName = ""
    pgmName = ""
    execCond = ""

    pgmName = Trim(GetParmPGM(jParameters)).ToUpper
    execCond = GetParm(jParameters, "COND=").ToUpper

    ' Is this a PROC statement
    If pgmName.Length = 0 Then
      procName = GetParm(jParameters, "PROC=")
      If procName.Length = 0 Then
        procName = GetFirstParm(jParameters)
      End If
      pgmName = ""
      Exit Sub
    End If

    ' If pgmName is a symbolic name; replace it
    If pgmName.Substring(0, 1) = "&" Then
      pgmName = ReplaceSymbolics(pgmName, ListOfSymbolics)
    End If

    If stepName = "" Or stepName = "*" Then
      stepName = pgmName
    End If


    ' At this time we need to check if its an Easytrieve load-and-go program (EZTPA00)
    '  if so, we need to substitue the pgmName with the actual program from the DD SYSIN statement
    '  from a PDS member name of the DSN, 
    '  or from a Instream-data (ie DD *)
    '  which was identified during the ProcessJob routine.
    If pgmName = "EZTPA00" Then
      Dim myKey As String = FileNameOnly & "_" & jobName & "_" & stepName
      Dim myValue As String = Nothing
      If ListOfEasytrieveLoadAndGo.TryGetValue(myKey, myValue) Then
        pgmName = myValue
      End If
    End If

    ' If this NOT an IMS program? Get the program name.
    If pgmName <> "DFSRRC00" Then
      If NeedSourceType Then
        SourceType = GetSourceType(pgmName)       'note. SourceCount is also updated there
        execName = pgmName
      End If
      Exit Sub
    End If

    ' get the real program name from the IMS PARM phrase, 2nd value
    ' i.e., PARM='DLI,P2BPCSD1,P2BPCSD1'
    execName = pgmName
    Dim tempstr As String = GetParm(jParameters, "PARM=")
    If tempstr.Length = 0 Then
      pgmName = "IMS Unknown"
      If NeedSourceType Then
        SourceType = "Unknown"
      End If
      SourceCount = 0
      Exit Sub
    End If

    ' example content of tempstr:'DLI,P2BPCSD1,P2BPCSD1'
    Dim temparray As String() = tempstr.Split(",")
    pgmName = temparray(1)

    If NeedSourceType Then
      SourceType = GetSourceType(pgmName)       'note. SourceCount is also updated there
    End If

  End Sub

  Sub ProcessDD(ByRef ListOfSymbolics As List(Of String))
    ' Process the DD statement

    ' get the last jLabel if the previous was a comment line
    If jLabel = "*" Then
      If ListOfDDs.Count > 0 Then
        Dim lastDDEntries As String() = ListOfDDs(ListOfDDs.Count - 1).Split(Delimiter)
        jLabel = lastDDEntries(7)
      End If
    End If

    Dim db2 As String = ""
    If jLabel = "CAFIN" Then
      db2 = "DB2"
    End If

    ' Only substitue symbolics on DSN
    Dim dsn As String = GetParm(jParameters, "DSN=")
    If dsn.Length > 0 Then
      If dsn.Substring(0, 2) <> "&&" Then
        If dsn.IndexOf("&") > -1 Then
          dsn = ReplaceSymbolics(dsn, ListOfSymbolics)
        End If
      End If
    End If

    ' Handle JOBLIB
    If jLabel = "JOBLIB" Then
      JobLib = dsn
      If ListOfLibraries.IndexOf(JobLib & Delimiter & "JOBLIB") = -1 Then
        ListOfLibraries.Add(JobLib & Delimiter & "JOBLIB")
      End If
      Exit Sub
    End If

    ' Handle Steplib
    If jLabel = "STEPLIB" Then
      Select Case jobName
        Case "CALLPGMS", "ONLINE"
        Case Else
          Dim steplib = dsn
          If ListOfLibraries.IndexOf(steplib & Delimiter & "STEPLIB") = -1 Then
            ListOfLibraries.Add(steplib & Delimiter & "STEPLIB")
          End If
      End Select
    End If

    Dim reportID As String = ""
    Dim sysout As String = GetParm(jParameters, "SYSOUT=")
    Select Case sysout.Length
      Case 0

      Case 1
        Select Case jLabel
          Case "SYSOUT", "SYSPRINT", "SYSUDUMP"
            If sysout = "*" Then
              sysout = "SYSOUT=" & jobMsgClass
            Else
              sysout = "SYSOUT=" & sysout
            End If
            dsn = sysout
            Exit Select
        End Select
        If sysout = "*" Then
          sysout = "SYSOUT=" & jobMsgClass
          dsn = sysout
          Exit Select
        End If
        If dsn.Length = 0 Then
          dsn = "SYSOUT=" & sysout
          Exit Select
        End If

      Case Else
        Dim sysoutParms As String() = sysout.Replace("(", "").Replace(")", "").Split(",")
        If sysoutParms.Count <= 1 Then
          sysout = "SYSOUT=" & sysout
        Else
          sysout = "SYSOUT=" & sysoutParms(0)
          reportID = sysoutParms(1)
        End If

    End Select

    ' Adjust for GDG type of datasets; remove the numeric (ie, (0) or (+1))
    ' But be careful of PDS libraries
    dsn = RemoveGDGValuesFromDSN(dsn)


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

    ' when program name could not be determined, set to MISSING
    If pgmName.Length = 0 Then
      pgmName = "MISSING"
      SourceType = "Utility"
    End If

    ' Add to list of DD statements
    ListOfDDs.Add(JobSourceName & txtDelimiter.Text &
                       jobName & txtDelimiter.Text &
                       procName & txtDelimiter.Text &
                       stepName & txtDelimiter.Text &
                       execName & Delimiter &
                       pgmName & txtDelimiter.Text &
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
                       SourceCount)

    ' write out the list of programs (presume no duplicates)
    If ddSequence = 1 And ddConcatSeq = 0 Then
      Dim hlProcName As String = hlBuilder.CreateHyperLinkProcs(procName)
      Dim hlProgramSourceName As String = hlBuilder.CreateHyperLinkSources(pgmName, SourceType)
      Dim hlProgramFlowchartName As String = hlBuilder.CreateHyperLinkSVGFlowchart(pgmName)
      Dim hlProgramFlowchartP2PName As String = hlBuilder.CreateHyperLinkP2PFlowchart(pgmName)
      Dim hlProgramBRName As String = hlBuilder.CreateHyperLinkBRXLS(pgmName)
      ListOfPrograms.Add(JobSourceName & Delimiter &
                       jobName & Delimiter &
                       hlProcName & Delimiter &
                       stepName & Delimiter &
                       execName & Delimiter &
                       hlProgramSourceName & Delimiter &
                       SourceType & Delimiter &
                       hlProgramFlowchartName & Delimiter &
                       hlProgramFlowchartP2PName & Delimiter &
                       hlProgramBRName & Delimiter &
                       SourceCount & Delimiter &
                       execCond)
      Select Case SourceType
        Case "COBOL", "Easytrieve", "Assembler"
          ListOfExecs.Add(pgmName & Delimiter & SourceType)
      End Select
    End If

    prevDDName = ddName
    prevPgmName = pgmName
    prevStepName = stepName
  End Sub

  Function ReplaceSymbolics(ByRef dsn As String, ByRef ListOfSymbolics As List(Of String)) As String
    ' this will loop thru, as needed, replacing any symbolic with value from the ListOfSymbolics array
    ' find a symbolic which is terminated with a period, an &, or end of line
    If ListOfSymbolics.Count <= 0 Then
      Return dsn
    End If
    If dsn.Length = 0 Then
      Return dsn
    End If
    Dim x As Integer = 0
    Dim y As Integer = 0
    Dim symLength As Integer = 0
    Dim symName As String = ""
    Do Until x > dsn.Length - 1
      x = dsn.IndexOf("&", x)
      If x = -1 Then
        Exit Do
      End If
      'now find ending of symbolic
      For y = x + 1 To dsn.Length - 1
        Select Case dsn.Substring(y, 1)
          Case "&", "(", ")", ".", ","
            Exit For
        End Select
      Next
      If y > dsn.Length - 1 Then
        y = dsn.Length - 1
      End If
      symLength = y - x
      If symLength > 0 Then
        symName = dsn.Substring(x, symLength).Replace("&", "")
        For Each symbol In ListOfSymbolics
          If symbol.IndexOf(symName) > -1 Then
            Dim symSplit As String() = symbol.Split("=")
            Dim symValue As String = symSplit(1).Replace("'", "")
            dsn = dsn.Replace("&" & symName, symValue)
            Exit For
          End If
        Next
      Else
        Exit Do
      End If
      x += symLength
    Loop

    Return dsn
  End Function
  Function DetermineStartDisp(ByRef fileDisp As String()) As String
    ' determine start disp
    If fileDisp Is Nothing Then
      Return "OUTPUT"
    End If
    If fileDisp.Count >= 1 Then
      If fileDisp(0).Length = 0 Then
        Return "OUTPUT"
      Else
        Select Case fileDisp(0)
          Case "SHR"
            Return "INPUT"
          Case "MOD", "NEW"
            Return "OUTPUT"
          Case Else
            Return "INPUT"
        End Select
      End If
    Else
      Return "INPUT"
    End If
    Return ""
  End Function
  Function DetermineEndDisp(ByRef fileDisp As String()) As String
    If fileDisp Is Nothing Then
      Return "KEEP"
      Exit Function
    End If
    If fileDisp.Count >= 2 Then
      If fileDisp(1).Length = 0 Then
        Return "KEEP"
      Else
        Select Case fileDisp(1)
          Case "KEEP"
            Return "KEEP"
          Case "DELETE"
            Return "DELETE"
          Case "CATLG"
            Return "KEEP"
        End Select
      End If
    Else
      Return "KEEP"
    End If
    Return ""
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

  Sub CreateSummaryTab()

    'workbook = objExcel.Workbooks.Add
    'SummaryWorksheet = workbook.Sheets.Item(1)
    SummaryWorksheet.Name = "Summary"
    SummaryRow = 0
    SummaryWorksheet.Range("A1").Value = "Mainframe Documentation Project" & vbNewLine &
                                         "Data Gathering Form" & vbNewLine &
                                         IO.Path.GetFileName(folderPath & "\" & txtDataGatheringForm.Text) &
                                         vbNewLine &
                                         "Model Created:" & Date.Now & vbNewLine &
                                         "Accelerator: ADDILite, Version:" & ProgramVersion
    SummaryWorksheet.Cell("B1").Value = ""
    SummaryWorksheet.Cell("A2").Value = ""
    SummaryWorksheet.Cell("B2").Value = ""
    SummaryWorksheet.Cell("A3").Value = "Folder Locations:"
    Dim theFilename As String = QUOTE & "filename" & QUOTE
    Dim theOpenBracket As String = QUOTE & "[" & QUOTE
    Dim theFormula As String = "=LEFT(CELL(" & theFilename & "),FIND(" & theOpenBracket & ",CELL(" & theFilename & "))-2)"
    SummaryWorksheet.Cell("B3").FormulaA1 = theFormula
    SummaryWorksheet.Cell("A4").Value = "JOBS:"
    SummaryWorksheet.Cell("B4").Value = "/JOBS/"
    SummaryWorksheet.Cell("A5").Value = "PROCS:"
    SummaryWorksheet.Cell("B5").Value = "/PROCS/"
    SummaryWorksheet.Cell("A6").Value = "COBOL:"
    SummaryWorksheet.Cell("B6").Value = "/COBOL/"
    SummaryWorksheet.Cell("A7").Value = "EASYTRIEVE:"
    SummaryWorksheet.Cell("B7").Value = "/EASYTRIEVE/"
    SummaryWorksheet.Cell("A8").Value = "ASSEMBLER:"
    SummaryWorksheet.Cell("B8").Value = "/ASM/"
    SummaryWorksheet.Cell("A9").Value = "COPYBOOKS:"
    SummaryWorksheet.Cell("B9").Value = "/COPYBOOKS/"
    SummaryWorksheet.Cell("A10").Value = "FLOWCHARTS:"
    SummaryWorksheet.Cell("B10").Value = "/FLOWCHARTS/"
    SummaryWorksheet.Cell("A11").Value = "BUSINESS RULES:"
    SummaryWorksheet.Cell("B11").Value = "/BUSINESS RULES/"
    SummaryWorksheet.Cell("A12").Value = ""
    SummaryWorksheet.Cell("B12").Value = ""
    SummaryWorksheet.Cell("A13").Value = "Data Gathering Form Contents:"
    SummaryWorksheet.Cell("B13").Value = ""
    SummaryRow = 14
    For Each dgf In ListOfDataGathering
      SummaryRow += 1
      Dim dgfCol As String() = dgf.Split(Delimiter)
      SummaryWorksheet.Cell(SummaryRow, 1).Value = dgfCol(0)
      SummaryWorksheet.Cell(SummaryRow, 2).Value = dgfCol(1)
    Next

  End Sub
  Sub CreateJobsTab()
    ' Build the Jobs Worksheet
    ' Given a set of variables write 1 row on the JOB tab
    If Not cbJOBS.Checked Then
      Exit Sub
    End If
    LogFile.WriteLine(Date.Now & ",I,CreateJobsTab," & jobName)
    If JobRow = 0 Then
      ' Write the column headings row
      JobsWorksheet.Cell("A1").Value = "Flow"
      JobsWorksheet.Cell("B1").Value = "Job_Source"
      JobsWorksheet.Cell("C1").Value = "Job_Name"
      JobsWorksheet.Cell("D1").Value = "AccountInfo"
      JobsWorksheet.Cell("E1").Value = "ProgrammerName"
      JobsWorksheet.Cell("F1").Value = "Time"
      JobsWorksheet.Cell("G1").Value = "Class"
      JobsWorksheet.Cell("H1").Value = "MsgC"
      JobsWorksheet.Cell("I1").Value = "Send"
      JobsWorksheet.Cell("J1").Value = "Route"
      JobsWorksheet.Cell("K1").Value = "JobParm"
      JobsWorksheet.Cell("L1").Value = "Region"
      JobsWorksheet.Cell("M1").Value = "COND"
      JobsWorksheet.Cell("N1").Value = "JCLLIB"
      JobsWorksheet.Cell("O1").Value = "JOBLIB"
      JobsWorksheet.Cell("P1").Value = "Typrun"
      JobRow = 1
      JobsWorksheet.SheetView.FreezeRows(1)
    End If

    JobRow += 1
    Dim row As String = LTrim(Str(JobRow))
    If JobSourceName = "CALLPGMS" Then
      JobsWorksheet.Cell("A" & row).Value = ""
      JobsWorksheet.Cell("B" & row).Value = ""
    Else
      JobsWorksheet.Cell("A" & row).FormulaA1 = hlBuilder.CreateHyperLinkJobFlowchart(JobSourceName)
      JobsWorksheet.Cell("A" & row).Style.Font.Underline = XLFontUnderlineValues.Single
      JobsWorksheet.Cell("B" & row).FormulaA1 = hlBuilder.CreateHyperLinkJobSource(JobSourceName)
      JobsWorksheet.Cell("B" & row).Style.Font.Underline = XLFontUnderlineValues.Single
    End If
    JobsWorksheet.Cell("C" & row).Value = jobName
    JobsWorksheet.Cell("D" & row).Value = JobAccountInfo
    JobsWorksheet.Cell("E" & row).Value = JobProgrammerName
    JobsWorksheet.Cell("F" & row).Value = JobTime
    JobsWorksheet.Cell("G" & row).Value = jobClass
    JobsWorksheet.Cell("H" & row).Value = jobMsgClass
    JobsWorksheet.Cell("I" & row).Value = JobSend
    JobsWorksheet.Cell("J" & row).Value = JobRoute
    JobsWorksheet.Cell("K" & row).Value = JobParm
    JobsWorksheet.Cell("L" & row).Value = JobRegion
    JobsWorksheet.Cell("M" & row).Value = JobCond
    JobsWorksheet.Cell("N" & row).Value = JobJCLLib
    JobsWorksheet.Cell("O" & row).Value = JobLib
    JobsWorksheet.Cell("P" & row).Value = JobTyprun

    LogFile.WriteLine(Date.Now & ",I,CreateJobsTab," & "Complete")

  End Sub
  Sub CreateJobCommentsTab(ListOfSymbolics As List(Of String))
    ' Build the JobComments Worksheet for one JOB.
    ' Process through the JclStmt array. Look for an 'EXEC' command and then process backwards
    '  to find the first of the comments. Then string (vbLF) comments together and write
    '  an entry for that 'EXEC'
    If Not cbJobComments.Checked Then
      Exit Sub
    End If

    LogFile.WriteLine(Date.Now & ",I,CreateJobCommentsTab," & jclStmt.Count)
    'lblProcessingWorksheet.Text = "Processing Job Comments: " & FileNameOnly

    'build a list of Job comments
    If JobCommentsRow = 0 Then
      ' Write the column headings row
      JobCommentsWorksheet.Cell("A1").Value = "Source"
      JobCommentsWorksheet.Cell("B1").Value = "JobName"
      JobCommentsWorksheet.Cell("C1").Value = "Program"
      JobCommentsWorksheet.Cell("D1").Value = "StepName"
      JobCommentsWorksheet.Cell("E1").Value = "Comments above Program"
      JobCommentsWorksheet.SheetView.FreezeRows(1)
      JobCommentsRow = 1
    End If

    Dim ListOfJobComments As New List(Of String)      'this holds all comments for this job

    ' find the EXEC statement starting from bottom of the JclStmt array to top
    Dim row As String = ""
    For index = 0 To jclStmt.Count - 1
      Dim statement As String = jclStmt(index)
      Call GetLabelControlParms(statement, jLabel, jControl, jParameters)
      Select Case jControl
        Case "EXEC"
          stepName = jLabel
          Call ProcessEXEC(False, ListOfSymbolics)
          Dim comment As String = ""
          For pgmIndex As Integer = index - 1 To 0 Step -1
            Call GetLabelControlParms(jclStmt(pgmIndex), jLabel, jControl, jParameters)
            If jControl = "COMMENT" Then
              comment = jParameters.Replace("=", "") & vbLf & comment
            Else
              If jControl <> "PROC" Then
                Exit For
              End If
            End If
          Next
          If comment.EndsWith(vbLf) Then
            comment = comment.Remove(comment.Length - 1)
          End If
          If comment.Length > 0 Then
            'write the comment row 
            JobCommentsRow += 1
            JobCommentsWorksheet.Cell(JobCommentsRow, 1).Value = FileNameWithExtension
            JobCommentsWorksheet.Cell(JobCommentsRow, 2).Value = jobName
            JobCommentsWorksheet.Cell(JobCommentsRow, 3).Value = pgmName
            JobCommentsWorksheet.Cell(JobCommentsRow, 4).Value = stepName
            JobCommentsWorksheet.Cell(JobCommentsRow, 5).Value = comment
            JobCommentsWorksheet.Cell(JobCommentsRow, 5).Style.Alignment.WrapText = True
          End If
      End Select
    Next
    LogFile.WriteLine(Date.Now & ",I,CreateJobCommentsTab," & "Complete")
  End Sub
  Sub CreateProgramsTab()
    ' Build the Programs worksheet. Programs sheet is a list of all JCL Jobs with programs.
    If Not cbPrograms.Checked Then
      Exit Sub
    End If

    LogFile.WriteLine(Date.Now & ",I,CreateProgramsTab," & ListOfPrograms.Count)
    'lblProcessingWorksheet.Text = "Processing Programs: " & FileNameOnly & " : Rows = " & ListOfPrograms.Count

    If ProgramsRow = 0 Then
      ' Write the column headings row
      ProgramsWorksheet.Range("A1").Value = "Job_Source"
      ProgramsWorksheet.Range("B1").Value = "Job_Name"
      ProgramsWorksheet.Range("C1").Value = "Proc_Name"
      ProgramsWorksheet.Range("D1").Value = "StepName"
      ProgramsWorksheet.Range("E1").Value = "ExecName"
      ProgramsWorksheet.Range("F1").Value = "PgmName"
      ProgramsWorksheet.Range("G1").Value = "SourceType"
      ProgramsWorksheet.Range("H1").Value = "Flow"
      ProgramsWorksheet.Range("I1").Value = "P2P"
      ProgramsWorksheet.Range("J1").Value = "Business Rules"
      ProgramsWorksheet.Range("K1").Value = "Count"
      ProgramsWorksheet.Range("L1").Value = "ExecCOND"
      ProgramsWorksheet.SheetView.FreezeRows(1)
      ProgramsRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("L") - 1 ' Last column index (L = 12 - 1)

    Call MoveListOfToWorksheet(ProgramsRow, lastColIndex, ListOfPrograms, ProgramsWorksheet)

    LogFile.WriteLine(Date.Now & ",I,CreateProgramsTab," & "Complete")
  End Sub
  Sub MoveListOfToWorksheet(ByRef theRowIndex As Integer,
                            ByRef theColIndex As Integer,
                            ByRef theListOf As List(Of String),
                            ByRef theWorksheet As IXLWorksheet)
    ' Note. theRowIndex + 1 is the first row to write to
    For Each myRow As String In theListOf
      theRowIndex += 1
      Dim fields() As String = myRow.Split(New String() {Delimiter}, StringSplitOptions.None)
      For colIndex As Integer = 0 To theColIndex
        If fields(colIndex).StartsWith("=HYPERLINK") Then
          theWorksheet.Cell(theRowIndex, colIndex + 1).FormulaA1 = fields(colIndex)
          theWorksheet.Cell(theRowIndex, colIndex + 1).Style.Font.Underline = XLFontUnderlineValues.Single
        Else
          theWorksheet.Cell(theRowIndex, colIndex + 1).Value = fields(colIndex)
        End If
      Next
    Next
  End Sub

  Function ColumnLetterToNumber(ByVal colLetter As String) As Integer
    ' Converts Excel column letter(s) to column number (e.g., "Q" -> 17, "AA" -> 27)
    Dim colNum As Integer = 0
    colLetter = colLetter.ToUpperInvariant().Trim()
    For i As Integer = 0 To colLetter.Length - 1
      colNum = colNum * 26 + (Asc(colLetter(i)) - Asc("A"c) + 1)
    Next
    Return colNum
  End Function
  Sub CreateFilesTab()
    ' Build the Files Tab. This is a list of all Files (DD) in the JCL Jobs.
    If Not cbFiles.Checked Then
      Exit Sub
    End If

    LogFile.WriteLine(Date.Now & ",I,CreateFilesTab," & ListOfDDs.Count)
    'lblProcessingWorksheet.Text = "Processing Files: " & FileNameOnly & " : Rows = " & ListOfDDs.Count

    If FilesRow = 0 Then
      ' Write the column headings row
      FilesWorksheet.Range("A1").Value = "Job_Source"
      FilesWorksheet.Range("B1").Value = "Job_Name"
      FilesWorksheet.Range("C1").Value = "ProcName"
      FilesWorksheet.Range("D1").Value = "StepName"
      FilesWorksheet.Range("E1").Value = "ExecName"
      FilesWorksheet.Range("F1").Value = "PgmName"
      FilesWorksheet.Range("G1").Value = "DD"
      FilesWorksheet.Range("H1").Value = "DDSeq"
      FilesWorksheet.Range("I1").Value = "DDConcatSeq"
      FilesWorksheet.Range("J1").Value = "DatasetName"
      FilesWorksheet.Range("K1").Value = "StartDisp"
      FilesWorksheet.Range("L1").Value = "EndDisp"
      FilesWorksheet.Range("M1").Value = "AbendDisp"
      FilesWorksheet.Range("N1").Value = "RecFM"
      FilesWorksheet.Range("O1").Value = "LRECL"
      FilesWorksheet.Range("P1").Value = "DBMS"
      FilesWorksheet.Range("Q1").Value = "ReportId"
      FilesWorksheet.SheetView.FreezeRows(1)
      FilesRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("Q") - 1 ' Last column index (Q = 17 - 1)

    Call MoveListOfToWorksheet(FilesRow, lastColIndex, ListOfDDs, FilesWorksheet)

    LogFile.WriteLine(Date.Now & ",I,CreateFilesTab," & "Complete")
  End Sub

  Sub CreateInstreamTab()
    ' Build the Instream Tab. This is a list of all instream (DD *) in the JCL Jobs.
    ' This routine is structured differently than the other tabs (ie, Files, Programs, etc.)
    ' because it is a dictionary of key/value pairs instead of a list of strings.

    If Not cbJOBS.Checked Then
      Exit Sub
    End If
    '
    LogFile.WriteLine(Date.Now & ",I,CreateInstreamTab," & DictOfInstreams.Count)
    'lblProcessingWorksheet.Text = "Processing Instreams: " &
    '  "Rows = " & DictOfInstreams.Count
    If InstreamsRow = 0 Then
      ' Write the column headings row
      InstreamsWorksheet.Range("A1").Value = "Job_Source"
      InstreamsWorksheet.Range("B1").Value = "StepName"
      InstreamsWorksheet.Range("C1").Value = "StepSeq"
      InstreamsWorksheet.Range("D1").Value = "Instream Content"
      InstreamsWorksheet.SheetView.FreezeRows(1)
      InstreamsRow = 1
    End If

    If DictOfInstreams.Count = 0 Then
      Exit Sub
    End If

    ' convert List to Array 2D
    Dim myMaxRows As Integer = DictOfInstreams.Count - 1
    Dim myMaxcols As Integer = 3
    Dim tArray(myMaxRows, myMaxcols) As String
    For Each kvp As KeyValuePair(Of String, String) In DictOfInstreams
      Dim myKey As String = kvp.Key
      Dim myValue As String = kvp.Value
      Dim myKeys As String() = myKey.Split(Delimiter)
      Dim mySourceFile As String = myKeys(0)
      Dim myStepName As String = myKeys(1)
      Dim myStepSeq As String = myKeys(2)
      If myValue.EndsWith(vbLf) Then
        myValue = myValue.Remove(myValue.Length - 1, 1)
      End If
      InstreamsRow = +1
      InstreamsWorksheet.Cell(InstreamsRow, 1).Value = mySourceFile
      InstreamsWorksheet.Cell(InstreamsRow, 2).Value = myStepName
      InstreamsWorksheet.Cell(InstreamsRow, 3).Value = myStepSeq
      InstreamsWorksheet.Cell(InstreamsRow, 4).Value = myValue
    Next

    LogFile.WriteLine(Date.Now & ",I,CreateInsteamTab," & "Complete")
  End Sub
  Sub ProcessExecFiles()
    Dim SourceRecordsCount As Integer = 0
    Dim execCnt As Integer = 0
    Dim execCount As Integer = ListOfExecs.Count

    ' loop through the list of executables. Note we may be adding while processing (called members)
    For Each exec In ListOfExecs
      execCnt += 1
      Dim execs As String() = exec.Split(Delimiter)
      If execs.Count >= 2 Then
        exec = execs(0).Replace(Delimiter, "")
        SourceType = execs(1)
      End If
      If exec.Length = 0 Then
        LogFile.WriteLine(Date.Now & ",E,Source file name empty?," & FileNameOnly)
        Continue For
      End If
      lblProcessingSource.Text = "Processing Exec #" & execCnt & ":" & exec

      'Load the source file to the stmt array
      Select Case SourceType
        Case "COBOL"
          SourceRecordsCount = LoadCobolStatementsToArray(exec)
        Case "Easytrieve"
          SourceRecordsCount = LoadEasytrieveStatementsToArray(exec)
        Case "ASM"
          SourceRecordsCount = LoadAssemblerStatementsToArray(exec)
        Case Else
          Continue For
      End Select
      If SourceRecordsCount = -1 Then
        Continue For
      End If
      If SourceRecordsCount = 0 Then
        LogFile.WriteLine(Date.Now & ",E,No " & SourceType & " lines Found in file," & exec)
        Continue For
      End If

      If SrcStmt.Count = 0 Then
        LogFile.WriteLine(Date.Now & ",E,No statements load to SrcStmt Array," & exec)
        Continue For
      End If

      ' log file, optioned
      If cbLogStmt.Checked Then
        Call LogStmtArray(exec, SrcStmt)
      End If

      ' Analyze Source Statement array (SrcStmt) to get list of programs
      listOfProgramInfo.Clear()
      listOfProgramInfo = GetListOfProgramInfo(exec)      'list of programs within the executable source

      ' Analyze Source Statement array (SrcStmt) to get list of EXEC SQL statments
      Call GetListOfEXECSQLorIMS()

      Call GetListOfCICSMapNames()

      If cbDataCom.Checked Then
        Call GetListOfDataComs()
      End If

      If pgm.ProcedureDivision = -1 Then
        LogFile.WriteLine(Date.Now & ",E,Source is not complete," & exec)
        Continue For
      End If

      If cbScanModeOnly.Checked Then
        Continue For
      End If

      'write the output files/excel
      Select Case SourceType
        Case "COBOL"
          If WriteOutputCOBOL(exec) = -1 Then
            LogFile.WriteLine(Date.Now & ",E,Error while building COBOL output,")
          End If

        Case "Easytrieve"
          If WriteOutputEasytrieve(exec) = -1 Then
            LogFile.WriteLine(Date.Now & ",E,Error while building Easytrieve output," & exec)
          End If

      End Select

    Next
    If Not cbScanModeOnly.Checked Then
      Call CreateCommentsTab()
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
    tempCobFileName = folderPath & txtExpandedFolder.Text & "\" & CobolFile & "_expandedCOB.txt"

    ' Remove the temporary work file
    Try
      If My.Computer.FileSystem.FileExists(tempCobFileName) Then
        My.Computer.FileSystem.DeleteFile(tempCobFileName)
      End If
    Catch ex As Exception
      LogFile.WriteLine(Date.Now & ",E,Removal of Temp hlk error," & ex.Message)
      Return -1
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
    Dim FoundCobolFileName As String = CobolSourceExists(CobolFile)
    '
    If FoundCobolFileName.Length = 0 Then
      LogFile.WriteLine(Date.Now & ",E,COBOL Source not found," & CobolFile)
      Return -1
    End If
    ' Adjust CobolFile to the FOUND filename
    If CobolFile <> FoundCobolFileName Then
      LogFile.WriteLine(Date.Now & ",I,Cobol Alias Used," & CobolFile & "/Alias:" & FoundCobolFileName)
      CobolFile = FoundCobolFileName
    End If
    '
    LogFile.WriteLine(Date.Now & ",I,Processing COBOL Source," & FoundCobolFileName)

    ' Load the COBOL file into the working Array
    Dim CobolLines As String() = File.ReadAllLines(folderPath & txtCobolFolder.Text & "\" & FoundCobolFileName)

    ' If missing first 6 bytes (an asterisk in col 1), add the six bytes to the front of every line.
    Dim Missing6 As Boolean = False
    For index As Integer = 0 To CobolLines.Count - 1
      Select Case CobolLines(index).Length
        Case >= 15
          If (CobolLines(index).Substring(0, 15) = " IDENTIFICATION") Then
            Missing6 = True
            Exit For
          End If
        Case >= 4
          If (CobolLines(index).Substring(0, 4) = " ID ") Then
            Missing6 = True
            Exit For
          End If
      End Select
    Next
    If Missing6 = True Then
      LogFile.WriteLine(Date.Now & ",W,ADD MISSING 6 COBOL CHARACTERS," & FoundCobolFileName)
      For index As Integer = 0 To CobolLines.Count - 1
        CobolLines(index) = Space(6) & CobolLines(index)
      Next
    End If

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
      ' drop blank/empty lines
      If CobolLines(index).Trim.Length <= 6 Then
        Continue For
      End If
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
          Call ProcessComment(index, CobolLines(index), Division, FoundCobolFileName)
          Continue For
        End If
        If CobolLines(index).Substring(6, 1) = "/" And Division.Length > 0 Then
          Mid(CobolLines(index), 7, 1) = "*"
          Call ProcessComment(index, CobolLines(index), Division, FoundCobolFileName)
          Continue For
        End If
      End If
      ' if Division is Identification (or ID) and it is not 'program-id' then the cobol line 
      '   is treated as comments and so will we.
      If Division.ToUpper.Trim = "IDENTIFICATION" Or Division.ToUpper.Trim = "ID" Then
        Dim temppgmid As String = CobolLines(index).Substring(7).Trim
        If Mid(temppgmid, 1, 11) <> "PROGRAM-ID." Then
          Call ProcessComment(index, CobolLines(index), Division, FoundCobolFileName)
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
        ' Fix bad bytes and/or remove tab bytes
        CobolLines(index) = CobolLines(index).Replace(vbTab, " ").Replace(ChrW(26), " ")
        ' make a whole blank/empty line a comment line
        If Len(Trim(CobolLines(index))) = 0 Then
          swTemp.WriteLine(Space(6) & "*")
          Continue For
        End If

        Call FillInAreas(CobolLines(index),
                         SequenceNumberArea, IndicatorArea, AreaA, AreaB, CommentArea)

        ' Comment out DEBUG Lines
        If IndicatorArea = "D" Then
          IndicatorArea = "*"
          Mid(CobolLines(index), 7, 1) = "*"
        End If
        ' Comment out # Lines
        If IndicatorArea = "#" Then
          IndicatorArea = "*"
          Mid(CobolLines(index), 7, 1) = "*"
        End If
        ' Comment out this period in column 7 not sure why its there.
        If IndicatorArea = "." Then
          IndicatorArea = "*"
          Mid(CobolLines(index), 7, 1) = "*"
        End If
        ' special adjustment for slash in column 7, must be a Telon artifact
        If IndicatorArea = "/" Then
          IndicatorArea = "*"
          Mid(CobolLines(index), 7, 1) = "*"
        End If
        ' special adjustment for Continuation in column 7 but areaA and B are empty (HP COBOL)
        If IndicatorArea = "-" And AreaA.Trim.Length = 0 And AreaB.Trim.Length = 0 Then
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
        Dim IDirective As New List(Of String)
        Call GetSourceWords(IncludeDirective, IDirective)
        ' Checking for copy/include statement to process
        Dim CopyType As String = ""
        Select Case True
          Case IDirective(0) = "COPY"
            If IDirective.Count >= 2 Then
              CopybookName = Trim(IDirective(1).Replace(QUOTE, "").Replace(".", " "))
              If CopybookName.Substring(0, 1) = "\" Then
                CopybookName.Remove(0, 1)
              End If
              CopyType = IDirective(0)
            End If

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
              LogFile.WriteLine(Date.Now & ",E,Malformed SQL statement; missing END-EXEC," &
                              CobolLines(index) & " line#:" & index + 1)
              Return -1
            End If
            ' check to see if this an SQL INCLUDE or some other SQL command
            Dim execDirective As String() = combinedEXEC.Trim.Split(New Char() {" "c})
            If execDirective(1) = "SQL" And execDirective(2) = "INCLUDE" Then
              CopyType = "SQL"
              CopybookName = execDirective(3)
              ' comment out the COMBINED SQL INCLUDE statement(s) as a comment
              swTemp.WriteLine(Space(6) & "* " & combinedEXEC.Trim)
              index = endIndex
            Else
              '  ' write out these non-INCLUDE SQL statements
              swTemp.WriteLine(CobolLines(index))
              Continue For
            End If
          Case Else
            swTemp.WriteLine(CobolLines(index))
            Continue For
        End Select

        ' Expand copybooks/includes into the source
        NumberOfCopysFound += 1
        swTemp.WriteLine(Space(6) & "*" & CopyType & " " & CopybookName & " Begin Copy/Include")
        LogFile.WriteLine(Date.Now & ",I,Including COBOL " & CopyType & " copybook," & CopybookName)
        Call IncludeCopyMember(CopybookName, swTemp)
        swTemp.WriteLine(Space(6) & "*" & CopyType & " " & CopybookName & " End Copy/Include")
      Next
      swTemp.Close()

      ' check if we expanded any copybooks, if so we scan again for any copy/includes
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
    Dim numCobolStatements As Integer = 0

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
    Dim hlkcounter As Integer = -1
    Division = ""
    SrcStmt.Clear()

    For Each text1 As String In CobolLines
      hlkcounter += 1
      numCobolStatements += 1
      text1 = text1.Replace(vbTab, Space(4))                'replace TAB(S) with single space!
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
        ' special adjustment: should next line be a continuation this line
        '  cannot be end of sentence.
        If hlkcounter < CobolLines.Count - 1 Then
          If CobolLines(hlkcounter + 1).Length >= 7 Then
            If CobolLines(hlkcounter + 1).Substring(6, 1) = "-" Then
              PeriodIsPresent = False
            End If
          End If
        End If
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
        cStatement &= AreaAandB.Trim.Substring(1)
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
    Return numCobolStatements
  End Function
  Function GetListOfProgramInfo(ByRef exec As String) As List(Of ProgramInfo)
    ' Scan through the source looking for the programs.
    ' Each program could have multiple sub programs inline (especially COBOL).
    ' Also a program could call a sub program, which we will store out to a
    '  separate file (CallPgms.jcl) for later analysis.
    pgm.IdentificationDivision = -1
    pgm.ProcedureDivision = -1
    pgm.EnvironmentDivision = -1
    pgm.DataDivision = -1
    pgm.WorkingStorage = -1
    pgm.ProcedureDivision = -1
    pgm.EndProgram = -1
    pgm.ProgramId = ""
    pgm.SourceId = exec


    Select Case SourceType
      Case "COBOL"
        For stmtIndex As Integer = 0 To SrcStmt.Count - 1
          Select Case True
            Case SrcStmt(stmtIndex).Substring(0, 1) = "*"
              Continue For
            Case (SrcStmt(stmtIndex).IndexOf("IDENTIFICATION DIVISION.") > -1) Or
                (SrcStmt(stmtIndex).IndexOf("IDENTIFICATION  DIVISION.") > -1) Or
                (SrcStmt(stmtIndex).IndexOf("IDENTIFICATION   DIVISION.") > -1) Or
                (SrcStmt(stmtIndex).IndexOf("IDENTIFICATION    DIVISION.") > -1) Or
                (SrcStmt(stmtIndex).IndexOf("ID DIVISION.") > -1)
              If pgm.ProcedureDivision >= 1 Then
                pgm.EndProgram = stmtIndex - 1
                listOfProgramInfo.Add(pgm)
                pgm = Nothing
              End If
              pgm.IdentificationDivision = stmtIndex
              pgm.EnvironmentDivision = -1
              pgm.DataDivision = -1
              pgm.WorkingStorage = -1
              pgm.ProcedureDivision = -1
            Case (SrcStmt(stmtIndex).IndexOf("ENVIRONMENT DIVISION.") > -1) Or
                (SrcStmt(stmtIndex).IndexOf("ENVIRONMENT  DIVISION.") > -1) Or
                (SrcStmt(stmtIndex).IndexOf("ENVIRONMENT   DIVISION.") > -1) Or
                (SrcStmt(stmtIndex).IndexOf("ENVIRONMENT    DIVISION.") > -1) Or
                (SrcStmt(stmtIndex).IndexOf("ENVIRONMENT     DIVISION.") > -1) Or
                (SrcStmt(stmtIndex).IndexOf("ENVIRONMENT      DIVISION.") > -1)
              pgm.EnvironmentDivision = stmtIndex
            Case SrcStmt(stmtIndex).IndexOf("DATA DIVISION.") > -1
              pgm.DataDivision = stmtIndex
            Case SrcStmt(stmtIndex).IndexOf("WORKING-STORAGE SECTION.") > -1 And pgm.DataDivision > -1
              pgm.WorkingStorage = stmtIndex
            Case SrcStmt(stmtIndex).IndexOf("PROCEDURE DIVISION") > -1
              pgm.ProcedureDivision = stmtIndex
            Case SrcStmt(stmtIndex).IndexOf("PROGRAM-ID.") > -1
              Dim tmppgmid As String = SrcStmt(stmtIndex).Trim
              Dim tmppgmid2 As String() = tmppgmid.Substring(11).Replace(".", "").Replace("'", "").Trim.Split(" ")
              pgm.ProgramId = tmppgmid2(0)
              If pgm.ProgramId.Length = 0 Then
                ' since Program-id's value is not on same line, presume it is on next line
                '   else the source is invalid syntax.
                stmtIndex += 1
                pgm.ProgramId = SrcStmt(stmtIndex).Replace(".", "").Replace("'", "").Trim
              End If
          End Select
          If pgm.ProcedureDivision > -1 Then
            ' if there is a CALL verb anywhere on the line, analyze further
            If SrcStmt(stmtIndex).IndexOf(" CALL ") > -1 Then
              Call AddToListOfCallPgms(SrcStmt(stmtIndex))
            End If
          End If
        Next
        If Not IsNothing(pgm) Then
          pgm.EndProgram = SrcStmt.Count - 1
          listOfProgramInfo.Add(pgm)
        End If

      Case "Easytrieve"
        Dim srcWords As New List(Of String)
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
        If pgm.EnvironmentDivision = -1 Then
          pgm.EnvironmentDivision = pgm.IdentificationDivision
        End If
        listOfProgramInfo.Add(pgm)
    End Select

    Return listOfProgramInfo
  End Function
  Sub AddToListOfCallPgms(ByRef statement As String)
    ' Search for the verb CALL and determine what program it is calling.
    ' Also Identify as Static or Dynamic
    ' Input:
    '   COBOL statement string with all of its phrases
    ' Output is an entry added to the ListOfCallPgms array
    '   
    Dim CalledFileName As String = ""
    Dim CalledType As String = ""
    Dim CalledEntry As String = ""
    Dim CalledMember As String = ""
    Dim CalledSourceType As String = ""

    Call GetSourceWords(statement, cWord)

    For x As Integer = 0 To cWord.Count - 1
      If cWord(x) <> "CALL" Then
        x = IndexToNextVerb(cWord, x)
        If x = -1 Then
          Exit For
        Else
          Continue For
        End If
      End If

      ' What type of Call? Static or Dynamic
      'dynamic called routines indicated by lack of quote
      CalledEntry = ""
      CalledMember = cWord(x + 1)
      Select Case Mid(CalledMember, 1, 1)
        Case "'", QUOTE
          'Static Call
          CalledMember = CalledMember.Replace("'", "").Replace(QUOTE, "").ToUpper.Trim
          CalledType = "Static"
          ' if a utility, Mark source type as Utility and thus no source type is needed
          If Array.IndexOf(Utilities, CalledMember) > -1 Then
            CalledSourceType = "Utility"
          Else
            ' Get source type of Called Routine, first remove any extension and uppercase it
            Dim PartsOfCalledMember As String() = CalledMember.Split(".")
            CalledMember = PartsOfCalledMember(0).ToUpper
            CalledSourceType = GetSourceType(CalledMember)
          End If
        Case Else
          'Dynamic Call
          CalledSourceType = "n/a" ' no source type for dynamic calls
          CalledType = "Dynamic"
      End Select

      ' if not already in the list, add it
      CalledEntry = pgm.SourceId & Delimiter &
                pgm.ProgramId & Delimiter &
                CalledMember & Delimiter &
                CalledSourceType & Delimiter &
                CalledType
      If ListOfCallPgms.IndexOf(CalledEntry) = -1 Then
        ListOfCallPgms.Add(CalledEntry)
      End If

      ' Now skip to the next verb
      x = IndexToNextVerb(cWord, x)
      If x = -1 Then
        Exit For
      End If
    Next
  End Sub
  Sub GetListOfEXECSQLorIMS()
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
    Dim JustTheTable As String = ""
    For Each pgm In listOfProgramInfo
      Select Case SourceType
        Case "COBOL"
          For stmtIndex As Integer = pgm.DataDivision + 1 To pgm.EndProgram

            If SrcStmt(stmtIndex).IndexOf("'CBLTDLI'") > -1 Then
              If ListOfIMSPgms.IndexOf(pgm.ProgramId) = -1 Then
                ListOfIMSPgms.Add(pgm.ProgramId)
                Continue For
              End If
            End If

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
                  Statement = BuildTheStatement(x + 4, cWord)
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
                  Statement = BuildTheStatement(x + 5, cWord)
                  Call AddToListOfEXECSQL(execCnt, ExecSQL, Table, Cursor, Statement)
                  x += 4
                  Continue For
                End If

                If cWord(x) = "EXEC" And cWord(x + 1) = "SQL" And cWord(x + 2) = "DELETE" Then
                  ExecSQL = cWord(x + 2)
                  Table = cWord(x + 4)
                  Cursor = ""
                  Statement = BuildTheStatement(x + 5, cWord)
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

            If SrcStmt(stmtIndex).IndexOf("'CBLTDLI'") > -1 Then
              If ListOfIMSPgms.IndexOf(pgm.ProgramId) = -1 Then
                ListOfIMSPgms.Add(pgm.ProgramId)
                Continue For
              End If
            End If

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
                  End If
                Next y
                If Table.Length > 0 Then
                  If ListOfTables.IndexOf(Table) = -1 Then
                    ListOfTables.Add(Table)
                    Table = ""
                  End If
                End If
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
  Function BuildTheStatement(ByRef x As Integer, ByRef cWord As List(Of String)) As String
    ' Build the statement placing a space between each word
    Dim theStatement As String = ""
    For y As Integer = x To cWord.Count - 1
      If cWord(y) = "END-EXEC" Then
        Exit For
      End If
      theStatement &= cWord(y) & " "
    Next y
    Return theStatement.Trim
  End Function
  Sub AddToListOfEXECSQL(ByRef Execcnt As Integer,
                         ByRef execSql As String,
                         ByRef Table As String,
                         ByRef Cursor As String,
                         ByVal Statement As String)
    ' Add an entry to the ListOfEXECSQL array

    ' Need to remove any Delimiters within the fields and replace with &
    Statement = Statement.Replace(Delimiter, "&")
    ' Adjust for long statements
    If Statement.Length > 60 Then
      Statement = AddNewLineAboutEveryNthCharacters(Statement, vbNewLine, 60)
    End If
    ' Remove any commas in the Table and replace with new line
    Table = Table.Replace(",", vbNewLine).Trim

    Execcnt += 1
    ListOfEXECSQL.Add(FileNameOnly & Delimiter &
                      pgm.ProgramId & Delimiter &
                      execSql & Delimiter &
                      LTrim(Str(Execcnt)) & Delimiter &
                      Table & Delimiter &
                      Cursor & Delimiter &
                      Statement)
  End Sub
  Sub GetListOfCICSMapNames()
    Dim ExecCICS As String = ""
    Dim MapName As String = ""
    Dim execCnt As Integer = 0
    Dim NotFound As String = ""
    For Each pgm In listOfProgramInfo
      Select Case SourceType
        Case "COBOL"
          For stmtIndex As Integer = pgm.DataDivision + 1 To pgm.EndProgram

            If SrcStmt(stmtIndex).IndexOf("EXEC CICS") = -1 Then
              Continue For
            End If

            Call GetSourceWords(SrcStmt(stmtIndex).Replace("( ", "(").Replace(" (", "(").Replace(" )", ")"), cWord)

            For x = 0 To cWord.Count - 1
              If cWord(x) = "END-EXEC" Then
                Continue For
              End If

              If (x + 2) < cWord.Count - 1 Then
                If cWord(x) = "EXEC" And cWord(x + 1) = "CICS" And
                  (cWord(x + 2) = "RECEIVE" Or cWord(x + 2) = "SEND" Or cWord(x + 2) = "XCTL" Or cWord(x + 2) = "LINK") Then
                  ExecCICS = cWord(x + 2)
                  MapName = ""
                  NotFound = ""
                  For y = x + 2 To cWord.Count - 1
                    If cWord(y) = "END-EXEC" Then
                      Exit For
                    End If
                    If cWord(y).IndexOf("MAP(") > -1 Then
                      MapName = cWord(y).Replace("MAP('", "").Replace("')", "").Trim
                      Exit For
                    End If
                    If cWord(y).IndexOf("PROGRAM(") > -1 Then
                      'Is this a literal or variable name. Literal will have a quote mark
                      If cWord(y).IndexOf("'") > -1 Then
                        MapName = cWord(y).Replace("PROGRAM('", "").Replace("')", "").Trim
                        If CobolSourceExists(MapName).Length = 0 Then
                          NotFound = "NotFound"
                        End If
                      Else
                        ' search for variable name and get the VALUE clause, if any
                        Dim VariableName As String = cWord(y).Replace("PROGRAM(", "").Replace(")", "").Trim
                        MapName = GetVariableNameValue(VariableName)
                        Select Case MapName
                          Case "NOTFOUND", "NOVALUE"
                            MapName = VariableName & "=" & MapName
                            NotFound = "n/a"
                          Case Else
                            MapName = MapName.Replace("'", "")
                            If CobolSourceExists(MapName).Length = 0 Then
                              NotFound = "NotFound"
                            End If
                        End Select
                      End If
                      Exit For
                    End If
                  Next y
                  If MapName.Length > 0 Then
                    Call AddToListOfCICSMapNames(execCnt, pgm, ExecCICS, MapName, NotFound)
                  End If
                  x += 2
                  Continue For
                End If
              End If
            Next
          Next
      End Select
    Next


  End Sub
  Sub AddToListOfCICSMapNames(ByRef execCnt As Integer, ByRef pgm As ProgramInfo,
                              ByRef ExecCICS As String, ByRef MapName As String, ByRef NotFound As String)
    execCnt += 1
    ListOfCICSMapNames.Add(FileNameOnly & Delimiter &
                      pgm.SourceId & Delimiter &
                      pgm.ProgramId & Delimiter &
                      LTrim(Str(execCnt)) & Delimiter &
                      ExecCICS & Delimiter &
                      MapName & Delimiter &
                      NotFound)
  End Sub
  Sub GetListOfDataComs()
    Dim DataComStatement As String = ""
    Dim DataViewName As String = ""
    Dim DataViewName_2 As String = ""
    Dim WhereClause As String = ""
    Dim BatchOnly As String = ""
    Dim execCnt As Integer = 0
    Dim NotFound As String = ""
    ' are there any 'DATACOM SECTION. ' statements at all?
    If SrcStmt.IndexOf("DATACOM SECTION. ") = -1 Then
      Exit Sub
    End If
    ' 
    For Each pgm In listOfProgramInfo
      Select Case SourceType
        Case "COBOL"
          ' Check each Procedure Divisions for DATACOM commands
          For stmtIndex As Integer = pgm.ProcedureDivision + 1 To pgm.EndProgram

            Call GetSourceWords(SrcStmt(stmtIndex), cWord)

            DataComStatement = ""
            DataViewName = ""
            DataViewName_2 = ""
            WhereClause = ""

            Select Case cWord(0)
              Case "ENTER-DATACOM-DB"
                BatchOnly = "BATCH"
              Case "FOR"
                DataComStatement = cWord(0)
                DataViewName = GetForDataViewName(cWord)
                WhereClause = GetForWhereClause(cWord)
              Case "READ", "OBTAIN"
                Call GetReadDataviewAndWhereClause(cWord, DataComStatement, DataViewName, WhereClause)
              Case "WRITE", "REWRITE", "DELETE"
                DataComStatement = cWord(0)
                DataViewName = cWord(1)
              Case "LOCATE"
                Call GetLocateDataviewAndWhereClause(cWord, DataComStatement, DataViewName, WhereClause, DataViewName_2)
            End Select

            If DataComStatement.Length > 0 Then
              ListOfDataComs.Add(pgm.SourceId & Delimiter &
                                 pgm.ProgramId & Delimiter &
                                 DataComStatement & Delimiter &
                                 DataViewName & Delimiter &
                                 WhereClause & Delimiter &
                                 DataViewName_2)
            End If
          Next
      End Select
    Next

  End Sub
  Function GetForDataViewName(ByRef cWord As List(Of String)) As String
    ' look for the Dataview Name value as stated in the FOR statement
    Select Case cWord(1)
      Case "EACH"
        Return cWord(2)
      Case "FIRST", "ANY"
        Return cWord(3)
      Case Else
        Return cWord(1)
    End Select
  End Function
  Function GetForWhereClause(ByRef cWord As List(Of String)) As String
    ' look for the WHERE clause as stated in the FOR statement
    Dim WhereIndex As Integer = cWord.IndexOf("WHERE")
    If WhereIndex = -1 Then
      Return ""
    End If
    Dim ForWhereClause As String = "WHERE "
    For x As Integer = WhereIndex + 1 To cWord.Count - 1
      Select Case cWord(x)
        Case "HOLD", "COUNT", "ORDER", "WHEN"
          Exit For
        Case Else
          ForWhereClause &= cWord(x) & " "
      End Select
    Next
    Return ForWhereClause.Trim
  End Function
  Sub GetReadDataviewAndWhereClause(ByRef cWord As List(Of String),
                                    ByRef DataComStatement As String,
                                    ByRef DataViewName As String,
                                    ByRef WhereClause As String)
    ' This will return the DataComStatement, DataviewName, and WhereClause, if any, for the READ statement.

    ' POSSIBLE FLAW!!! the READ NEXT could be a VSAM file or other non-DATACOM file.

    ' For Datacom; the read statement has 7 formats.
    ' 1. Read [AND HOLD] WHERE, 2. Read Next, 3. Read Next Within Range, 4. Read Physical, 5. Read Previous,
    ' 6. Read Sequential, 7. Read Within Range WHERE

    Dim WhereIndex As Integer = cWord.IndexOf("WHERE")
    ' Handle NO where clause (could be formats 2-7)
    If WhereIndex = -1 Then
      Call GetReadNoWhereClauses(cWord, DataComStatement, DataViewName)
      Exit Sub
    End If

    ' Determine the DatacomStatement and the DataViewName
    ' Format 1. READ [AND HOLD] dataview-name
    '              WHERE ...conditions...
    ' Format 7. READ [AND HOLD] dataview-name WITHIN RANGE
    '              WHERE ...conditions...
    If cWord.Count >= 6 Then
      If (cWord(0) = "READ" Or cWord(0) = "OBTAIN") Then
        DataComStatement = cWord(0) & " "
        If cWord(1) = "AND" And cWord(2) = "HOLD" Then
          DataComStatement &= cWord(1) & Space(1) & cWord(2) & Space(1)
          DataViewName = cWord(3)
        End If
        If cWord(2) = "WITHIN" And cWord(3) = "RANGE" Then
          DataComStatement &= cWord(2) & Space(1) & cWord(3) & Space(1)
          DataViewName = cWord(1)
        End If
      End If
    End If

    ' string together the WHERE clause 
    For x As Integer = WhereIndex To cWord.Count - 1
      WhereClause &= Space(1) & cWord(x)
    Next
  End Sub
  Sub GetReadNoWhereClauses(ByRef cWord As List(Of String),
                                 ByRef DatacomStatement As String,
                                 ByRef DataViewName As String)
    ' no WHERE clause, check for formats 2-6.
    ' OBTAIN is a synonym for READ
    ' Format 2 READ NEXT [DUPLICATE] [AND HOLD] dataview-name
    If (cWord(0) = "READ" Or cWord(0) = "OBTAIN") And cWord(1) = "NEXT" Then
      If cWord.Count >= 6 Then
        If cWord(2) = "DUPLICATE" And cWord(3) = "AND" And cWord(4) = "HOLD" Then
          DatacomStatement = cWord(0) & " NEXT DUPLICATE AND HOLD"
          DataViewName = cWord(5)
          Exit Sub
        End If
      End If
      If cWord.Count >= 5 Then
        If cWord(2) = "AND" And cWord(3) = "HOLD" Then
          DatacomStatement = cWord(0) & " NEXT AND HOLD"
          DataViewName = cWord(4)
          Exit Sub
        End If
      End If
      If cWord.Count >= 4 Then
        If cWord(2) = "DUPLICATE" Then
          DatacomStatement = cWord(0) & " NEXT DUPLICATE"
          DataViewName = cWord(3)
          Exit Sub
        End If
      End If
    End If

    ' Format 3 READ [AND HOLD] NEXT dataview-name WITHIN RANGE
    If cWord.Count >= 7 Then
      If (cWord(0) = "READ" Or cWord(0) = "OBTAIN") And cWord(1) = "AND" And cWord(2) = "HOLD" And
        cWord(3) = "NEXT" And cWord(5) = "WITHIN" And cWord(6) = "RANGE" Then
        DatacomStatement = cWord(0) & " AND HOLD NEXT WITHIN RANGE"
        DataViewName = cWord(4)
        Exit Sub
      End If
    End If
    If cWord.Count >= 5 Then
      If (cWord(0) = "READ" Or cWord(0) = "OBTAIN") And cWord(1) = "NEXT" And
        cWord(3) = "WITHIN" And cWord(4) = "RANGE" Then
        DatacomStatement = cWord(0) & " NEXT WITHIN RANGE"
        DataViewName = cWord(2)
        Exit Sub
      End If
    End If
    '
    ' Format 4 READ PHYSICAL dataview-name
    If cWord.Count >= 3 Then
      If (cWord(0) = "READ" Or cWord(0) = "OBTAIN") And cWord(1) = "PHYSICAL" Then
        DatacomStatement = cWord(0) & Space(1) & cWord(1)
        DataViewName = cWord(2)
        Exit Sub
      End If
    End If
    '
    ' Format 5 READ [AND HOLD] PREVIOUS dataview-name
    If cWord.Count >= 5 Then
      If (cWord(0) = "READ" Or cWord(0) = "OBTAIN") And cWord(1) = "AND" And cWord(2) = "HOLD" And
        cWord(3) = "PHYSICAL" Then
        DatacomStatement = cWord(0) & Space(1) & cWord(1) & Space(1) & cWord(2) & Space(1) & cWord(3)
        DataViewName = cWord(4)
        Exit Sub
      End If
    End If
    If cWord.Count >= 3 Then
      If (cWord(0) = "READ" Or cWord(0) = "OBTAIN") And cWord(1) = "PREVIOUS" Then
        DatacomStatement = cWord(0) & Space(1) & cWord(1)
        DataViewName = cWord(2)
        Exit Sub
      End If
    End If
    '
    ' Format 6 READ SEQUENTIAL dataview-name
    If cWord.Count >= 3 Then
      If (cWord(0) = "READ" Or cWord(0) = "OBTAIN") And cWord(1) = "SEQUENTIAL" Then
        DatacomStatement = cWord(0) & Space(1) & cWord(1)
        DataViewName = cWord(2)
        Exit Sub
      End If
    End If

    DatacomStatement = "Unknown:"
    For Each unknownWord In cWord
      DatacomStatement &= unknownWord & Space(1)
    Next
    DataViewName = "Unknown"

  End Sub

  Sub GetLocateDataviewAndWhereClause(ByRef cWord As List(Of String),
                                    ByRef DataComStatement As String,
                                    ByRef DataViewName As String,
                                    ByRef WhereClause As String,
                                    ByRef DataViewName_2 As String)
    ' This will return the DataComStatement, DataviewName, and WhereClause, if any, for the LOCATE statement.
    ' The LOCATE statement has 8 formats.
    ' 1. Locate At, 2, Locate Next, 3. Locate Next Within Range, 4. Locate Physical, 5. Locate Previous,
    ' 6. Locate Sequential Where, 7. Locate Where, 8. Locate Within Range Where

    Dim WhereIndex As Integer = cWord.IndexOf("WHERE")
    ' Handle NO where clause (could be formats 1-5)
    If WhereIndex = -1 Then
      Call GetLocateNoWhereClauses(cWord, DataComStatement, DataViewName, DataViewName_2)
      Exit Sub
    End If

    ' Determine the DatacomStatement and the DataViewName
    ' Format 6. LOCATE SEQUENTIAL dataview-name
    '              WHERE ...conditions...
    ' Format 7. LOCATE dataview-name 
    '              WHERE ...conditions...
    ' Format 8. LOCATE dataview-name WITHIN RANGE
    '              WHERE ...conditions
    If cWord.Count >= 4 Then
      If cWord(0) = "LOCATE" Then
        DataComStatement = cWord(0)
        Select Case True
          Case cWord(1) = "SEQUENTIAL"
            DataComStatement &= Space(1) & cWord(1)
            DataViewName = cWord(2)
          Case cWord(2) = "WITHIN" And cWord(3) = "RANGE"
            DataComStatement &= Space(1) & cWord(2) & Space(1) & cWord(3)
            DataViewName = cWord(1)
          Case Else
            DataViewName = cWord(1)
        End Select
      End If
    End If

    ' string together the WHERE clause 
    For x As Integer = WhereIndex To cWord.Count - 1
      WhereClause &= Space(1) & cWord(x)
    Next
  End Sub
  Sub GetLocateNoWhereClauses(ByRef cWord As List(Of String),
                                 ByRef DatacomStatement As String,
                                 ByRef DataViewName As String,
                                 ByRef DataViewName_2 As String)
    ' no WHERE clause, check for formats 1-5.
    ' Format 1 LOCATE dataview-name-1 AT dataview-name-2
    If cWord(0) = "LOCATE" And cWord(2) = "AT" Then
      DatacomStatement = cWord(0) & Space(1) & cWord(2)
      DataViewName = cWord(1)
      DataViewName_2 = cWord(3)
      Exit Sub
    End If
    ' Format 2 LOCATE [NEXT] [DUPLICATE]
    '                        [KEY      ] dataview-name
    If cWord(0) = "LOCATE" And cWord(1) = "NEXT" And
         (cWord(2) = "DUPLICATE" Or cWord(2) = "DUP" Or cWord(2) = "KEY") Then
      DatacomStatement = cWord(0) & Space(1) & cWord(1) & Space(1) & cWord(2)
      DataViewName = cWord(3)
      Exit Sub
    End If
    If cWord(0) = "LOCATE" And
         (cWord(1) = "DUPLICATE" Or cWord(1) = "DUP" Or cWord(1) = "KEY") Then
      DatacomStatement = cWord(0) & Space(1) & cWord(1)
      DataViewName = cWord(2)
      Exit Sub
    End If

    ' Format 3 LOCATE NEXT dataview-name WITHIN RANGE
    If cWord.Count = 5 Then
      If cWord(0) = "LOCATE" And cWord(1) = "NEXT" And
        cWord(3) = "WITHIN" And cWord(4) = "RANGE" Then
        DatacomStatement = cWord(0) & Space(1) & cWord(1) & Space(1) & cWord(3) & Space(1) & cWord(4)
        DataViewName = cWord(2)
        Exit Sub
      End If
    End If
    '
    ' Format 4 LOCATE PHYSICAL [AND HOLD] dataview-name
    If cWord(0) = "LOCATE" And cWord(1) = "PHYSICAL" Then
      If cWord.Count = 5 Then
        If cWord(2) = "AND" And cWord(3) = "HOLD" Then
          DatacomStatement = cWord(0) & Space(1) & cWord(1) & Space(1) & cWord(2) & Space(1) & cWord(3)
          DataViewName = cWord(4)
          Exit Sub
        End If
      Else
        DatacomStatement = cWord(0) & Space(1) & cWord(1)
        DataViewName = cWord(2)
        Exit Sub
      End If
    End If
    '
    ' Format 5 LOCATE PREVIOUS dataview-name
    If cWord.Count = 3 Then
      If cWord(0) = "LOCATE" And cWord(1) = "PREVIOUS" Then
        DatacomStatement = cWord(0) & Space(1) & cWord(1)
        DataViewName = cWord(2)
        Exit Sub
      End If
    End If
    '

    DatacomStatement = "Unknown:"
    For Each unknownWord In cWord
      DatacomStatement &= unknownWord & Space(1)
    Next
    DataViewName = "Unknown"

  End Sub
  Function GetVariableNameValue(ByRef VariableName As String) As String
    ' search for variable name and get the VALUE clause, if any
    ' if no VALUE clause found; return "NOVALUE"
    ' else return "<VALUE CLAUSE>"
    ' if variableName is not found return "NOTFOUND"
    ' There is a special adjustment for syntax of "VALUE IS"
    For x As Integer = pgm.DataDivision To pgm.ProcedureDivision - 1
      If SrcStmt(x).IndexOf(VariableName) > -1 Then
        If SrcStmt(x).IndexOf("VALUE") > -1 Then
          Dim vWord As New List(Of String)
          Call GetSourceWords(SrcStmt(x), vWord)
          For y As Integer = 0 To vWord.Count - 1
            If vWord(y) = "VALUE" Then
              If vWord(y + 1) = "IS" Then
                If vWord(y + 2) = "SPACES" Or vWord(y + 2) = "SPACE" Then
                  Return "NOVALUE"
                End If
                Return vWord(y + 2)
              Else
                If vWord(y + 1) = "SPACES" Or vWord(y + 1) = "SPACE" Then
                  Return "NOVALUE"
                End If
                Return vWord(y + 1)
              End If
            End If
          Next
        Else
          Return "NOVALUE"
        End If
      End If
    Next
    Return "NOTFOUND"
  End Function
  Function CobolSourceExists(ByRef SourceFileName As String) As String
    'this will return FileName, FileName.cob, FileName.cbl, FileName.txt if exists in the Sources Directory
    ' Empty return means not found
    Select Case True
      Case ListofCOBOLFiles.IndexOf(SourceFileName) > -1
        Return SourceFileName
      Case ListofCOBOLFiles.IndexOf(SourceFileName & ".COB") > -1
        Return SourceFileName & ".COB"
      Case ListofCOBOLFiles.IndexOf(SourceFileName & ".CBL") > -1
        Return SourceFileName & ".CBL"
      Case ListofCOBOLFiles.IndexOf(SourceFileName & ".TXT") > -1
        Return SourceFileName & ".TXT"
    End Select
    ' See if there is an Alias
    Dim myKey As String = SourceFileName
    Dim myValue As String = Nothing
    If AliasCobol.TryGetValue(myKey, myValue) Then
      Return myValue
    End If
    '
    Return ""
  End Function

  Function CopybookSourceExists(ByRef SourceFileName As String) As String
    'this will return FileName, FileName.cpy, FileName.txt if exists in the Copybook Sources Directory
    ' Empty return means not found
    Select Case True
      Case ListofCopybookFiles.IndexOf(SourceFileName) > -1
        Return SourceFileName
      Case ListofCopybookFiles.IndexOf(SourceFileName & ".CPY") > -1
        Return SourceFileName & ".CPY"
      Case ListofCopybookFiles.IndexOf(SourceFileName & ".TXT") > -1
        Return SourceFileName & ".TXT"
    End Select
    '
    Return ""
  End Function

  Function EasytrieveSourceExists(ByRef SourceFileName As String) As String
    'this will return FileName, FileName.ezt, FileName.txt if exists in the Easytrieve Directory
    ' Empty return means not found
    Select Case True
      Case ListofEasytrieveFiles.IndexOf(SourceFileName) > -1
        Return SourceFileName
      Case ListofEasytrieveFiles.IndexOf(SourceFileName & ".EZT") > -1
        Return SourceFileName & ".EZT"
      Case ListofEasytrieveFiles.IndexOf(SourceFileName & ".TXT") > -1
        Return SourceFileName & ".TXT"
    End Select
    ' See if there is an Alias
    Return ""
  End Function

  Function ASMSourceExists(ByRef SourceFileName As String) As String
    'this will return FileName, FileName.asm, FileName.txt if exists in the Easytrieve Directory
    ' Empty return means not found
    Select Case True
      Case ListofAsmFiles.IndexOf(SourceFileName) > -1
        Return SourceFileName
      Case ListofAsmFiles.IndexOf(SourceFileName & ".ASM") > -1
        Return SourceFileName & ".ASM"
      Case ListofAsmFiles.IndexOf(SourceFileName & ".TXT") > -1
        Return SourceFileName & ".TXT"
    End Select
    Return ""
  End Function



  Function GetSourceType(ByRef FileName As String) As String
    ' Identify if this file is COBOL or Easytrieve or Utility or Assembler
    ' For Utility just check the Utilities array and get out.
    ' For COBOL, Easytrieve, Assembler, the file must exist in their source directory.
    ' the deal is it could have various file extensions.
    'GetSourceType = ""
    SourceCount = 0

    ' Utility ?
    If FileName.Trim.Length = 0 Then
      LogFile.WriteLine(Date.Now & ",E,Filename for GetSourcetype is empty," & FileNameOnly)
      Return "UTILITY"
    End If
    If Array.IndexOf(Utilities, FileName) > -1 Then
      Return "UTILITY"
    End If

    ' COBOL ?
    Dim FoundSourceFileName As String = CobolSourceExists(FileName)
    If FoundSourceFileName.Length > 0 Then
      Dim sourceLines As String() = File.ReadAllLines(folderPath & txtCobolFolder.Text & "\" & FoundSourceFileName)
      SourceCount = sourceLines.Count
      Return "COBOL"
    End If
    ' Easytrieve ?
    FoundSourceFileName = EasytrieveSourceExists(FileName)
    If FoundSourceFileName.Length > 0 Then
      Dim sourceLines As String() = File.ReadAllLines(folderPath & txtEasytrieveFolder.Text & "\" & FoundSourceFileName)
      SourceCount = sourceLines.Count
      Return "Easytrieve"
    End If
    ' Assembler ?
    FoundSourceFileName = ASMSourceExists(FileName)
    If FoundSourceFileName.Length > 0 Then
      Dim sourceLines As String() = File.ReadAllLines(folderPath & txtASMFolder.Text & "\" & FoundSourceFileName)
      SourceCount = sourceLines.Count
      Return "Assembler"
    End If

    LogFile.WriteLine(Date.Now & ",E,Source File Not Found," & FileName)
    Return "NotFound"

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
    ' Copy member found?
    Dim FoundCopyMember As String = CopybookSourceExists(CopyMember)
    If FoundCopyMember.Length = 0 Then
      swTemp.WriteLine(Space(6) & "*Copy Member not found:" & CopyMember)
      LogFile.WriteLine(Date.Now & ",E,Copy Member not found," & CopyMember)
      Exit Sub
    End If
    ' is the file empty?
    Dim mysize As Long = FileLen(folderPath & txtCopybookFolder.Text & "\" & FoundCopyMember)
    If mysize = 0 Then
      swTemp.WriteLine(Space(6) & "*Copy Member Empty:" & CopyMember)
      LogFile.WriteLine(Date.Now & ",E,Copy Member Empty," & CopyMember)
      Exit Sub
    End If

    Dim IncludeLines As String() = File.ReadAllLines(folderPath & txtCopybookFolder.Text & "\" & FoundCopyMember)

    ' If missing first 6 bytes, add the six bytes to the front of every line.
    Dim Missing6 As Boolean = False
    For index As Integer = 0 To IncludeLines.Count - 1
      Select Case IncludeLines(index).Length
        Case >= 15
          If (IncludeLines(index).Substring(0, 15) = " IDENTIFICATION") Then
            Missing6 = True
            Exit For
          End If
        Case >= 4
          If (IncludeLines(index).Substring(0, 4) = " ID ") Then
            Missing6 = True
            Exit For
          End If
      End Select
    Next
    Dim FirstByteAsterisk As Boolean = False
    If IncludeLines(0).Length > 0 Then
      If IncludeLines(0).Substring(0, 1) = "*" Then
        FirstByteAsterisk = True
      End If
    End If
    If Missing6 = True Or FirstByteAsterisk = True Then
      LogFile.WriteLine(Date.Now & ",W,ADD MISSING 6 COBOL CHARACTERS," & FoundCopyMember)
      For index As Integer = 0 To IncludeLines.Count - 1
        IncludeLines(index) = Space(6) & IncludeLines(index)
      Next
    End If

    ' append copymember to temp file and drop blank lines
    For Each line As String In IncludeLines
      If Len(Trim(line)) > 0 Then
        swTemp.WriteLine(line)
      End If
    Next
  End Sub
  Sub IncludeCopyMemberEasytrieve(ByVal CopyMember As String,
                        ByRef swTemp As StreamWriter)
    Dim FoundCopyMember As String = CopybookSourceExists(CopyMember)
    If FoundCopyMember.Length = 0 Then
      swTemp.WriteLine(Space(6) & "*Copy Member not found:" & CopyMember)
      LogFile.WriteLine(Date.Now & ",E,Copy Member not found," & CopyMember)
      Exit Sub
    End If
    Dim IncludeLines As String() = File.ReadAllLines(folderPath & txtCopybookFolder.Text & "\" & FoundCopyMember)

    ' append copymember to temp file and drop blank lines
    For Each line As String In IncludeLines
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
    Call CollectRecordsAndFieldsInfo()

  End Function
  Function WriteOutputCOBOL(ByRef exec As String) As Integer
    ' Write the output Pgm, Data, Procedure, copy files.
    ' return of -1 means an error
    ' return of 0 means all is okay

    Call CreateCobolFlowchart(SrcStmt, exec, PUMLFolder)

    ' Create a Records/Fields spreadsheet
    Call CollectRecordsAndFieldsInfo()

    Return 0

  End Function
  Function LoadEasytrieveStatementsToArray(ByRef exec As String) As Integer
    '*---------------------------------------------------------
    ' Load Easytrieve lines to a statements array. 
    '*---------------------------------------------------------
    '
    'Assign the Temporay File Name for this particular Easytrieve file
    '
    tempEZTFileName = folderPath & txtExpandedFolder.Text & "\" & exec & "_expandedEZT.txt"

    ' Remove the temporary work file
    Try
      If My.Computer.FileSystem.FileExists(tempEZTFileName) Then
        My.Computer.FileSystem.DeleteFile(tempEZTFileName)
      End If
    Catch ex As Exception
      LogFile.WriteLine(Date.Now & ",E,Removal of Temp EZT error," & ex.Message)
      Return -1
    End Try
    Dim FoundEasytrieveFileName As String = EasytrieveSourceExists(exec)
    If FoundEasytrieveFileName = 0 Then
      LogFile.WriteLine(Date.Now & ",E,EASYTRIEVE Source File Not Found?," & exec)
      Return -1
    End If
    LogFile.WriteLine(Date.Now & ",I,Processing Easytrieve Source," & FoundEasytrieveFileName)

    ' put all the lines into the array
    Dim EztLinesLoaded As String() = File.ReadAllLines(FoundEasytrieveFileName)

    ' Load the COMMENTS found in the program to the ListOfComments array
    ' We'll use division values same as COBOL
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
        statement = statement.Substring(0, statement.Length - 2).TrimEnd  'remove the ' +' continuation bytes
        statement &= " " & Trim(Microsoft.VisualBasic.Left(EztLinesLoaded(continuedIndex) & Space(72), 72))
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
    Dim EZTFileName As String = ""
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
          EZTFileName = ""
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
            EZTFileName = MembersNames(db2Index)
            Call IncludeCopyMemberEasytrieve(EZTFileName, swTemp)
          Else
            swTemp.WriteLine("*Crossref to DB2Declares not found in SOURCES folder")
            LogFile.WriteLine(Date.Now & ",E,Crossref to DB2Declares not found in SOURCES folder," &
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
          EZTFileName = statement.Substring(1)
          Call IncludeCopyMemberEasytrieve(EZTFileName, swTemp)
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

    Return reccnt
  End Function

  Function LoadAssemblerStatementsToArray(ByRef exec As String) As Integer
    '*---------------------------------------------------------
    ' Load Assembler lines to a statements array. 
    '*---------------------------------------------------------
    '
    'Assign the Temporay File Name for this particular Easytrieve file
    '
    tempAsmFileName = folderPath & txtExpandedFolder.Text & "\" & exec & "_expandedASM.txt"

    ' Remove the temporary work file
    Try
      If My.Computer.FileSystem.FileExists(tempAsmFileName) Then
        My.Computer.FileSystem.DeleteFile(tempAsmFileName)
      End If
    Catch ex As Exception
      LogFile.WriteLine(Date.Now & ",E,Removal of Temp ASM error," & ex.Message)
      Return -1
    End Try

    Dim FoundAsmFileName As String = ASMSourceExists(exec)
    If FoundAsmFileName = 0 Then
      LogFile.WriteLine(Date.Now & ",E,Assembler Source File Not Found?," & exec)
      Return -1
    End If
    LogFile.WriteLine(Date.Now & ",I,Processing Assembler Source," & FoundAsmFileName)

    ' put all the lines into the array
    Dim AsmLinesLoaded As String() = File.ReadAllLines(FoundAsmFileName)

    Dim statement As String = ""
    Dim newLine As String = ""
    Dim swTemp As StreamWriter = Nothing
    swTemp = New StreamWriter(tempAsmFileName, False)
    swTemp.WriteLine("; " & FoundAsmFileName)
    Dim reccnt As Integer = 0

    ' process the Asmlinesloaded array
    '
    For index As Integer = 0 To AsmLinesLoaded.Length - 1
      swTemp.WriteLine(AsmLinesLoaded(index))
    Next index
    swTemp.Close()

    ' load the Source Statement Array with final content
    SrcStmt.Clear()

    reccnt = 0
    AsmLinesLoaded = File.ReadAllLines(tempAsmFileName)
    For index = 0 To AsmLinesLoaded.Count - 1
      SrcStmt.Add(AsmLinesLoaded(index).Trim)
      reccnt += 1
    Next

    Return reccnt
  End Function


  Sub CollectRecordsAndFieldsInfo()
    ' This collects the information about Records & Fields
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
    ListOfPrograms.Clear()
    ListOfFields.Clear()
    '
    '
    ' Process each program module in this source file
    '
    For Each pgm In listOfProgramInfo
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
          Case "Dataview"
            ListOfRecordNames = GetListOfRecordNamesDataview(FileName, FileNameIndex)
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
          If CopybookName <> "NONE" Then
            CopybookName = hlBuilder.CreateHyperLinkCopybookSources(CopybookName)
          End If
          'write the filename,recordname,recordname index
          ListOfRecords.Add(FileNameOnly & Delimiter &
                            pgmName & Delimiter &
                            FileName & Delimiter &
                            FileNameDD & Delimiter &
                            FileNameType & Delimiter &
                              recordName & Delimiter &
                              CopybookName & Delimiter &
                              LTrim(Str(recordLength)) & Delimiter &
                              recordNameIndex & Delimiter &
                              recordNameLevel & Delimiter &
                              recordNameOpenMode & Delimiter &
                              recordNameRecFM & Delimiter &
                              recordNameMinLrecl & Delimiter &
                              recordNameMaxLrecl & Delimiter &
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

    Call CreateRecordsTab()
    Call CreateFieldsTab()


  End Sub
  Sub CreateRecordsTab()
    '
    ' Create the Excel Records sheet.
    '
    If Not cbRecords.Checked Then
      Exit Sub
    End If

    lblProcessingWorksheet.Text = "Processing Records: " & FileNameOnly & " : Rows = " & ListOfRecords.Count
    If RecordsRow = 0 Then
      ' Write the Column Headers row
      RecordsWorksheet.Range("A1").Value = "Source"
      RecordsWorksheet.Range("B1").Value = "Program"
      RecordsWorksheet.Range("C1").Value = "File/Table"
      RecordsWorksheet.Range("D1").Value = "DD"
      RecordsWorksheet.Range("E1").Value = "Type"
      RecordsWorksheet.Range("F1").Value = "RecordName"
      RecordsWorksheet.Range("G1").Value = "Copybook"
      RecordsWorksheet.Range("H1").Value = "Length"
      RecordsWorksheet.Range("I1").Value = "@Line"
      RecordsWorksheet.Range("J1").Value = "Level"
      RecordsWorksheet.Range("K1").Value = "Open"
      RecordsWorksheet.Range("L1").Value = "RecFM"
      RecordsWorksheet.Range("M1").Value = "FDMinLen"
      RecordsWorksheet.Range("N1").Value = "FDMaxLen"
      RecordsWorksheet.Range("O1").Value = "FDOrg"
      RecordsWorksheet.SheetView.FreezeRows(1)
      RecordsRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("O") - 1 ' 15 columns A-O

    Call MoveListOfToWorksheet(RecordsRow, lastColIndex, ListOfRecords, RecordsWorksheet)

  End Sub
  Sub CreateFieldsTab()
    '
    ' Create the Fields worksheet
    '
    If Not cbFields.Checked Then
      Exit Sub
    End If

    lblProcessingWorksheet.Text = "Processing Fields: " & FileNameOnly & " : Rows = " & ListOfFields.Count

    If FieldsRow = 0 Then
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
      FieldsWorksheet.SheetView.FreezeRows(1)
      FieldsRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("O") - 1 ' 15 columns A-O

    ' move the ListOfFields to the worksheet
    Call MoveListOfToWorksheet(FieldsRow, lastColIndex, ListOfFields, FieldsWorksheet)

  End Sub
  Sub FormatWorksheets()
    If SummaryRow > 0 Then
      ' Set font style and size for the header cell
      Dim rngSummaryName As IXLRange = SummaryWorksheet.Range("A1:A1")
      rngSummaryName.Style.Font.Bold = True
      rngSummaryName.Style.Font.FontSize = 16

      ' Define the data range
      rngSummaryName = SummaryWorksheet.Range("A13:B" & (SummaryRow).ToString())

      ' Auto-fit all columns in the worksheet
      SummaryWorksheet.Columns().AdjustToContents()
      ' Set the vertical alignment to Top
      rngSummaryName.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top)
    End If

    Call ApplyStandardWorksheetFormat(JobRow, "P", JobsWorksheet, rngJobs)
    Call ApplyStandardWorksheetFormat(JobCommentsRow, "E", JobCommentsWorksheet, rngJobComments)
    Call ApplyStandardWorksheetFormat(ProgramsRow, "L", ProgramsWorksheet, rngPrograms)
    Call ApplyStandardWorksheetFormat(FilesRow, "Q", FilesWorksheet, rngFiles)
    Call ApplyStandardWorksheetFormat(RecordsRow, "O", RecordsWorksheet, rngRecordsName)
    Call ApplyStandardWorksheetFormat(FieldsRow, "O", FieldsWorksheet, rngFieldsName)
    Call ApplyStandardWorksheetFormat(CommentsRow, "F", CommentsWorksheet, rngComments)
    Call ApplyStandardWorksheetFormat(EXECSQLRow, "G", EXECSQLWorksheet, rngEXECSQL)
    Call ApplyStandardWorksheetFormat(EXECCICSRow, "G", EXECCICSWorksheet, rngEXECCICS)
    Call ApplyStandardWorksheetFormat(IMSRow, "C", IMSWorksheet, rngIMS)
    Call ApplyStandardWorksheetFormat(DataComRow, "F", DataComWorksheet, rngDataCom)
    Call ApplyStandardWorksheetFormat(CallsRow, "E", CallsWorksheet, rngCalls)
    Call ApplyStandardWorksheetFormat(ScreenMapRow, "D", ScreenMapWorksheet, rngScreenMap)
    Call ApplyStandardWorksheetFormat(LibrariesRow, "B", LibrariesWorksheet, rngLibraries)
    Call ApplyStandardWorksheetFormat(InstreamsRow, "D", InstreamsWorksheet, rngInstreams)

    workbook.Worksheet(1).SetTabActive() ' Set the first worksheet as active

  End Sub

  Sub ApplyStandardWorksheetFormat(ByRef lastRow As Integer,
                                    ByVal lastCol As String,
                                    ByRef theWorksheet As IXLWorksheet,
                                    ByRef theRange As IXLRange)
    ' Format the worksheet
    ' first row bold the columns
    ' auto filter all columns
    ' Auto fit all columns
    ' Set vertical alignment to Top

    If lastRow = 0 Then
      Exit Sub
    End If

    theRange = theWorksheet.Range("A1:L1")
    theRange.Style.Font.Bold = True

    ' Auto filter all columns
    Dim usedRange As IXLRange = theWorksheet.RangeUsed

    ' Auto-fit all columns in the range
    theWorksheet.Columns().AdjustToContents()

    usedRange.SetAutoFilter()
    ' Set the vertical alignment to Top
    usedRange.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top)
  End Sub
  Sub CreateCommentsTab()
    '* Create the Comments worksheet from the listofcomments array
    If Not cbComments.Checked Then
      Exit Sub
    End If

    'Create a list of COMBINED comments; removed for now as not sure this feature is wanted.
    'Dim ListOfCombinedComments As New List(Of String)
    'Dim prevLineNum As Integer = -1
    'Dim currLineNum As Integer = 0
    'Dim prevProgram As String = ""
    'Dim currProgram As String = ""
    'Dim combinedComment As String = ""
    'For Each comment In ListOfComments
    '  Dim commentColumns As String() = comment.Split(Delimiter)
    '  currLineNum = Val(commentColumns(4))
    '  currProgram = commentColumns(1)
    '  If currLineNum - 1 <> prevLineNum Or currProgram <> prevProgram Then
    '    ListOfCombinedComments.Add(commentColumns(0) & txtDelimiter.Text &
    '                               commentColumns(1) & txtDelimiter.Text &
    '                               commentColumns(2) & txtDelimiter.Text &
    '                               commentColumns(3) & txtDelimiter.Text &
    '                               commentColumns(4) & txtDelimiter.Text &
    '                               combinedComment)
    '    combinedComment = commentColumns(5)
    '  Else
    '    combinedComment &= vbNewLine & commentColumns(5)
    '  End If
    '  prevLineNum = currLineNum
    '  prevProgram = currProgram
    'Next


    lblProcessingWorksheet.Text = "Processing Comments: " & FileNameOnly & " : Rows = " & ListOfComments.Count

    If CommentsRow = 0 Then
      ' Write the Column Headers row
      CommentsWorksheet.Range("A1").Value = "Source"
      CommentsWorksheet.Range("B1").Value = "Program"
      CommentsWorksheet.Range("C1").Value = "Type"
      CommentsWorksheet.Range("D1").Value = "Division"
      CommentsWorksheet.Range("E1").Value = "Line#"
      CommentsWorksheet.Range("F1").Value = "Comment"
      CommentsWorksheet.SheetView.FreezeRows(1)
      CommentsRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("F") - 1 ' 6 columns A-F

    Call MoveListOfToWorksheet(CommentsRow, lastColIndex, ListOfComments, CommentsWorksheet)

  End Sub
  Sub CreateEXECSQLTab()
    '* Create the ExecSQL worksheet from the listofexecsql array
    If Not cbexecSQL.Checked Then
      Exit Sub
    End If

    LogFile.WriteLine(Date.Now & ",I,CreateEXECSQLTab," & ListOfEXECSQL.Count)
    'lblProcessingWorksheet.Text = "Processing ExecSQL: " & FileNameOnly & " : Rows = " & ListOfEXECSQL.Count

    If EXECSQLRow = 0 Then
      ' Write the Column Headers row
      EXECSQLWorksheet.Range("A1").Value = "Source"
      EXECSQLWorksheet.Range("B1").Value = "Program"
      EXECSQLWorksheet.Range("C1").Value = "EXECSQL"
      EXECSQLWorksheet.Range("D1").Value = "Seq"
      EXECSQLWorksheet.Range("E1").Value = "Table"
      EXECSQLWorksheet.Range("F1").Value = "Cursor"
      EXECSQLWorksheet.Range("G1").Value = "Statement"
      EXECSQLWorksheet.SheetView.FreezeRows(1)
      EXECSQLRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("G") - 1 ' 7 columns A-G

    Call MoveListOfToWorksheet(EXECSQLRow, lastColIndex, ListOfEXECSQL, EXECSQLWorksheet)

    LogFile.WriteLine(Date.Now & ",I,CreateEXECSQLTab," & "Complete")
  End Sub
  Sub CreateEXECCICSTab()
    '* Create the ExecCICS worksheet from the listofCICSMapNames array
    If Not cbexecCICS.Checked Then
      Exit Sub
    End If

    LogFile.WriteLine(Date.Now & ",I,CreateEXECCICSTab," & ListOfCICSMapNames.Count)

    'lblProcessingWorksheet.Text = "Processing ExecCICS: " & FileNameOnly & " : Rows = " & ListOfCICSMapNames.Count

    If EXECCICSRow = 0 Then
      ' Write the Column Headers row
      EXECCICSWorksheet.Range("A1").Value = "Filename"
      EXECCICSWorksheet.Range("B1").Value = "SourceId"
      EXECCICSWorksheet.Range("C1").Value = "ProgramId"
      EXECCICSWorksheet.Range("D1").Value = "ExecSeq"
      EXECCICSWorksheet.Range("E1").Value = "ExecCICS"
      EXECCICSWorksheet.Range("F1").Value = "Map/Program"
      EXECCICSWorksheet.Range("G1").Value = "Program Found"
      EXECCICSWorksheet.SheetView.FreezeRows(1)
      EXECCICSRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("G") - 1 ' 7 columns A-G

    Call MoveListOfToWorksheet(EXECCICSRow, lastColIndex, ListOfCICSMapNames, EXECCICSWorksheet)

    LogFile.WriteLine(Date.Now & ",I,CreateEXECCICSTab," & "Complete")
  End Sub

  Sub CreateIMSTab()
    ' Create the IMS worksheet tab.
    ' This worksheet will hold PSP/Program and DBD Name(s)

    ' There are two possible inputs to this routine. (Background information)
    ' From a) text file from PSBMap programs or b) from a Telon text files.
    ' a) this text file (DBDNames.txt) is initiated by an initial pass of this model to create
    '    a list of PSP names (pspnames.txt). this is uploaded to the mainframe and then we 
    '    run a JCL job PSPJ which will create this DBDNames.txt file.
    ' b) these text files are the individual telon members. These are downloaded to the \TELON folder.
    '    We only look for the literal 'DBDNAME='. Note this does NOT mean all these DBDNames are actually
    '    used. These are what matched a TELON naming pattern (ie, P2BPBU*).
    '
    If Not cbIMS.Checked Then
      Exit Sub
    End If

    LogFile.WriteLine(Date.Now & ",I,CreateIMSTab," & "Start")

    Call AddToListOfDBDNames()
    Call AddtoListOfDBDNamesTelons()

    If IMSRow = 0 Then
      ' Write the Column Headers row
      IMSWorksheet.Range("A1").Value = "DBD Name"
      IMSWorksheet.Range("B1").Value = "PSP Name"
      IMSWorksheet.Range("C1").Value = "Type"
      IMSWorksheet.SheetView.FreezeRows(1)
      IMSRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("C") - 1 ' 3 columns A-C
    Call MoveListOfToWorksheet(IMSRow, lastColIndex, ListOfDBDs, IMSWorksheet)

    LogFile.WriteLine(Date.Now & ",I,CreateIMSTab," & "Complete")
  End Sub
  Sub CreateDataComTab()
    ' Create the DataComs worksheet tab.
    ' This worksheet will hold Datacom details such as Datacom, DataView-Name and WHERE statements
    '
    If Not cbDataCom.Checked Then
      Exit Sub
    End If
    LogFile.WriteLine(Date.Now & ",I,CreateDataComTab," & ListOfDataComs.Count)


    If DataComRow = 0 Then
      ' Write the Column Headers row
      DataComWorksheet.Range("A1").Value = "Source"
      DataComWorksheet.Range("B1").Value = "ProgramID"
      DataComWorksheet.Range("C1").Value = "DataCommand"
      DataComWorksheet.Range("D1").Value = "DataView"
      DataComWorksheet.Range("E1").Value = "Where"
      DataComWorksheet.Range("F1").Value = "DataView AT"
      DataComWorksheet.SheetView.FreezeRows(1)
      DataComRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("F") - 1 ' 6 columns A-F

    Call MoveListOfToWorksheet(DataComRow, lastColIndex, ListOfDataComs, DataComWorksheet)


    LogFile.WriteLine(Date.Now & ",I,CreateDataComTab," & "Complete")
  End Sub
  Sub CreateCallsTab()
    ' Create the CALLs worksheet tab.
    ' This worksheet will hold program CALLS of both Static and Dynamic calls.
    ' The input to this routine is the ListOfCallPgms array
    '
    If Not cbCalls.Checked Then
      Exit Sub
    End If
    LogFile.WriteLine(Date.Now & ",I,CreateCallsTab," & ListOfCallPgms.Count)

    If CallsRow = 0 Then
      ' Write the Column Headers row
      CallsWorksheet.Range("A1").Value = "Source-id"
      CallsWorksheet.Range("B1").Value = "Program-id"
      CallsWorksheet.Range("C1").Value = "Module"
      CallsWorksheet.Range("D1").Value = "Source Type"
      CallsWorksheet.Range("E1").Value = "Call Type"
      CallsWorksheet.SheetView.FreezeRows(1)
      CallsRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("E") - 1 ' 5 columns A-E

    Call MoveListOfToWorksheet(CallsRow, lastColIndex, ListOfCallPgms, CallsWorksheet)

    LogFile.WriteLine(Date.Now & ",I,CreateCallsTab," & "Complete")

  End Sub
  Sub CreateScreenMapTab()
    '* Create the Screen Map worksheet from the listofScreenMaps array
    If Not cbScreenMaps.Checked Then
      Exit Sub
    End If
    LogFile.WriteLine(Date.Now & ",I,CreateScreenMapTab," & ListofScreenMaps.Count)

    'lblProcessingWorksheet.Text = "Processing ScreenMaps: " & FileNameOnly & " : Rows = " & ListofScreenMaps.Count

    Dim cnt As Integer = 0
    Dim row As Integer = 0
    Dim Statement As String = ""
    Dim Table As String = ""
    If ScreenMapRow = 0 Then
      ' Write the Column Headers row
      ScreenMapWorksheet.Range("A1").Value = "MapSource"
      ScreenMapWorksheet.Range("B1").Value = "Type"
      ScreenMapWorksheet.Range("C1").Value = "Name"
      ScreenMapWorksheet.Range("D1").Value = "Literals"
      ScreenMapWorksheet.SheetView.FreezeRows(1)
      ScreenMapRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("D") - 1 ' 4 columns A-D

    Call MoveListOfToWorksheet(ScreenMapRow, lastColIndex, ListofScreenMaps, ScreenMapWorksheet)

    LogFile.WriteLine(Date.Now & ",I,CreateScreenMapTab," & "Complete")

  End Sub
  Sub CreateLibrariesTab()
    ' This will create the Libraries tab worksheet based on the ListOfLibraries array sorted
    If Not cbLibraries.Checked Then
      Exit Sub
    End If
    LogFile.WriteLine(Date.Now & ",I,CreateLibrariesTab," & ListOfLibraries.Count)

    Dim row As Integer = 0
    If LibrariesRow = 0 Then
      ' Write the Column Headers row
      LibrariesWorksheet.Range("A1").Value = "Library"
      LibrariesWorksheet.Range("B1").Value = "Type"
      LibrariesWorksheet.SheetView.FreezeRows(1)
      LibrariesRow = 1
    End If

    Dim lastColIndex As Integer = ColumnLetterToNumber("B") - 1 ' 2 columns A-B

    Call MoveListOfToWorksheet(LibrariesRow, lastColIndex, ListOfLibraries, LibrariesWorksheet)


    LogFile.WriteLine(Date.Now & ",I,CreateLibrariesTab," & "Complete")

  End Sub
  Sub AddToListOfDBDNames()
    ' Process the DBDNames.txt file, if exists, and put into ListOf array
    Dim DBDFileName = folderPath & txtCobolFolder.Text & "\DBDnames.txt"
    If Not File.Exists(DBDFileName) Then
      Exit Sub
    End If
    Dim DBDLines As String() = File.ReadAllLines(DBDFileName)
    ' load IMS DBDnames to spreadsheet.
    For IMSIndx As Integer = 0 To DBDLines.Count - 1
      Dim IMSColumns As String() = DBDLines(IMSIndx).Split(Delimiter)
      Dim pspName As String = IMSColumns(0)       'PSP Name
      Dim dbdName As String = IMSColumns(2)       'DBD Name
      If ListOfDBDs.IndexOf(dbdName & Delimiter & pspName & Delimiter & "SOURCE") = -1 Then
        ListOfDBDs.Add(dbdName & Delimiter & pspName & Delimiter & "SOURCE")
      End If
      If ListOfDBDNames.IndexOf(dbdName) = -1 Then
        ListOfDBDNames.Add(dbdName)
      End If
    Next
  End Sub
  Sub AddtoListOfDBDNamesTelons()
    ' Process the TELON files, if any exists and store in ListOf array
    For Each foundFile As String In My.Computer.FileSystem.GetFiles(folderPath & txtTelonFolder.Text)
      Dim memberLines As String() = File.ReadAllLines(foundFile)
      For index As Integer = 0 To memberLines.Count - 1
        If Len(memberLines(index)) = 0 Then
          Continue For
        End If
        If memberLines(index).Substring(0, 1) = "*" Then
          Continue For
        End If
        Dim dbdIndex As Integer = memberLines(index).IndexOf("DBDNAME=")
        If dbdIndex > -1 Then
          Dim dbdParms As String() = memberLines(index).Split(",")
          Dim dbdNames As String() = dbdParms(0).Split("=")
          If dbdNames.Count > 0 Then
            Dim dbdName As String = dbdNames(1)
            Dim pspName As String = IO.Path.GetFileNameWithoutExtension(foundFile)
            If ListOfDBDs.IndexOf(dbdName & Delimiter & pspName & Delimiter & "TELON") = -1 Then
              ListOfDBDs.Add(dbdName & Delimiter & pspName & Delimiter & "TELON")
            End If
            If ListOfDBDNames.IndexOf(dbdName) = -1 Then
              ListOfDBDNames.Add(dbdName)
            End If
          End If
        End If
      Next index
    Next

  End Sub

  Sub CreateIMSPSPNamesFile()
    ' Create the PSP Names text file.
    ' Intent is to upload this file to the Mainframe and run the PSPMAP IMS utility and REXX program
    '   which will return back a DBDNames file with PSP and DBD values which will load to 
    '   an IMS tab on the NEXT rerun of this model.

    ' Open the output file PSPNames.txt 
    Dim PSPFileName = folderPath & txtOutputFolder.Text & "\PSPNames.txt"

    ' Open output. Not worrying (try/catch) about subsequent writes
    Try
      PSPFile = My.Computer.FileSystem.OpenTextFileWriter(PSPFileName, False)
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error opening: " & PSPFileName)
      Exit Sub
    End Try

    ' Write every PSP entry, if any
    For Each PSP In ListOfIMSPgms
      PSPFile.WriteLine(Space(2) & PSP & Space(70))
    Next

    PSPFile.Close()
  End Sub

  Sub PumlPageBreak(ByRef exec As String)
    pumlPageCnt += 1
    ' Open the output file Puml 
    Dim pumlFileName As String = PUMLFolder & "\" & exec & ".puml"
    If pumlPageCnt > 1 Then
      pumlFileName = PUMLFolder & "\" & exec & "_" & LTrim(Str(pumlPageCnt)) & ".puml"
    End If

    ' Open and write at least one time. Not worrying (try/catch) about subsequent writes
    Try
      pumlFile = My.Computer.FileSystem.OpenTextFileWriter(pumlFileName, False)
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error opening PumlFile COBOL")
      Exit Sub
    End Try

    ' Write the top of file
    pumlFile.WriteLine("@startuml " & exec)
    pumlFile.WriteLine("header ADDILite(c), by IBM")
    pumlFile.Write("title Flowchart of COBOL Program: " & exec &
                       "\nProgram Author: " & ProgramAuthor &
                       "\nDate written: " & ProgramWritten)
    If pumlPageCnt > 1 Then
      pumlFile.WriteLine("\nPart: " & pumlPageCnt)
    Else
      pumlFile.WriteLine("")
    End If
    pumlLineCnt = 3
    WithinIF = False
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
    Dim PumlFileName = PUMLFolder & "\" & exec & ".puml"

    ' Open and write at least one time. Not worrying (try/catch) about subsequent writes
    Try
      pumlFile = My.Computer.FileSystem.OpenTextFileWriter(PumlFileName, False)
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error opening PumlFile for Easytrieve")
      Exit Sub
    End Try

    ' Write the top of file
    pumlLineCnt = 3
    pumlFile.WriteLine("@startuml " & exec)
    pumlFile.WriteLine("header ADDILite(c), by IBM")
    pumlFile.WriteLine("title Flowchart of Easytrieve Program: " & exec)

    For Each pgm In listOfProgramInfo
      pgmName = pgm.ProgramId

      For index As Integer = pgm.ProcedureDivision To pgm.EndProgram
        ' skip comments
        If SrcStmt(index).Length >= 1 Then
          If SrcStmt(index).Substring(0, 1) = "*" Then
            Continue For
          End If
        End If
        If SrcStmt(index).Length >= 2 Then
          If SrcStmt(index).Substring(0, 2) = "//" Then   'Inline JCL statement
            Continue For
          End If
          If SrcStmt(index).Substring(0, 2) = "/*" Then
            Continue For
          End If
        End If

        ' break the statement in words
        Call GetSourceWords(SrcStmt(index).Trim, cWord)

        If cWord.Count = 1 Then
          'potential one word proc name
          If SrcStmt(index).Trim.EndsWith(".") Then
            cWord(0) &= "."
          End If
        End If

        ' Process every VERB word in this statement 
        ' Every verb should/could be a plum object created.

        For wordIndex = 0 To cWord.Count - 1
          Select Case cWord(wordIndex)
            Case "JOB"
              Call ProcessPumlParagraphEasytrieve(ParagraphStarted)
              Exit For
            Case "SORT"
              Call ProcessPumlSortEasytrieve(wordIndex)
            Case "SELECT"
              Call ProcessPumlSelectEasytrieve(wordIndex)
            Case "IF"
              IFLevelIndex.Add(wordIndex)
              Call ProcessPumlIFEasytrieve(wordIndex)
            Case "ELSE", "OTHERWISE"
              Call ProcessPumlELSE(wordIndex)
            Case "END-IF", "END-IF."
              IndentLevel -= 1
              pumlLineCnt += 1
              pumlFile.WriteLine(Indent() & "endif")
            Case "CASE"
              Call ProcessPumlCase(wordIndex)
            Case "END-CASE", "END-CASE."
              IndentLevel -= 1
              pumlLineCnt += 2
              pumlFile.WriteLine(Indent() & "endif")
              IndentLevel -= 1
              pumlFile.WriteLine(Indent() & ":END-CASE;")
            Case "WHEN"
              Call ProcessPumlWHEN(wordIndex)
            Case "END-EVALUATE"
              Call ProcessPumlENDEVALUATE(wordIndex)
            Case "PERFORM"
              Call ProcessPumlPerformEasytrieve(wordIndex)
            Case "DO"
              Call ProcessPumlDO(wordIndex)
            Case "END-DO", "END-DO."
              Call ProcessPumlENDDO(wordIndex)
            Case "COMPUTE"
              Call ProcessPumlCOMPUTE(wordIndex)
            Case "GET"
              Call ProcessPumlGET(wordIndex)
            Case "GO"
              Call ProcessPumlGOTO(wordIndex)
            Case "EXEC", "DLI"
              ProcessPumlEXEC(wordIndex)
            Case "END-PROC", "END-PROC."
              IndentLevel = 1
              pumlLineCnt += 3
              pumlFile.WriteLine("end")
              pumlFile.WriteLine("}")
              pumlFile.WriteLine("")
              ParagraphStarted = False
            Case "REPORT"
              Call ProcessPumlParagraphEasytrieve(ParagraphStarted)
              Exit For
            Case "PROC"
              Call ProcessPumlParagraphEasytrieve(ParagraphStarted)
              Exit For
            Case "SEQUENCE"
              wordIndex = cWord.Count - 1
            Case "CONTROL"
              wordIndex = cWord.Count - 1
            Case "TITLE"
              If cWord.Count >= 5 Then
                If cWord(1) = "1" Or cWord(1) = "2" Then
                  pumlLineCnt += 1
                  pumlFile.WriteLine(Indent() & ":" & cWord(4).Replace("'", "").Replace("*", "").Trim & ";")
                End If
              End If
              wordIndex = cWord.Count - 1
            Case "LINE"
              wordIndex = cWord.Count - 1
            Case Else
              If cWord(wordIndex).IndexOf(".") > -1 Then
                If cWord.Count > 1 Then
                  If cWord(wordIndex + 1) = "PROC" Then
                    theProcName = cWord(wordIndex)
                    Continue For
                  End If
                Else
                  ' a paragraph name but the PROC word is not on same line; store the paragraph name
                  '  for when we eventually get that proc name
                  theProcName = cWord(wordIndex)
                End If
              Else
                Dim EndIndex As Integer = 0
                Dim MiscStatement As String = ""
                Call GetStatement(wordIndex, EndIndex, MiscStatement)
                pumlLineCnt += 1
                pumlFile.WriteLine(Indent() & ":" & MiscStatement.Trim & ";")
                wordIndex = EndIndex
              End If
          End Select
        Next wordIndex
      Next index

      If ParagraphStarted = True Then
        IndentLevel = 1
        pumlLineCnt += 2
        pumlFile.WriteLine("end")
        pumlFile.WriteLine("}")
        ParagraphStarted = False
      End If

    Next
    pumlLineCnt += 1
    pumlFile.WriteLine("@enduml")

    pumlFile.Close()
  End Sub


  Function GetListOfFiles() As List(Of String)
    ' Scan through the stmt array looking for all data "FILES"
    '   A "FILE" is something stated with either "SELECT" or "EXEC SQL DECLARE" or "DATA-VIEW"
    '  Store also the DDName and indicate if FILE or SQL or Dataview
    ' in format of: Filename,DDName,FILE,index
    '           or: Tablename,,SQL,index
    '           or: Dataview-name,DATA-BASE-IDENTIFICATION,Dataview,index
    Dim statement As String = ""
    Dim FDFileName As String = ""
    Dim srcWords As New List(Of String)
    Dim VaryingName As String = ""
    ListOfFiles.Clear()
    Select Case SourceType
      Case "COBOL"
        For stmtIndex As Integer = pgm.EnvironmentDivision + 1 To pgm.ProcedureDivision - 1
          statement = SrcStmt(stmtIndex)
          If statement.Length >= 1 Then
            If statement.Substring(0, 1) = "*" Then
              Continue For
            End If
          End If
          Call GetSourceWords(statement, srcWords)

          If srcWords(0) = "SELECT" Then
            Dim file_name_1 As String = ""
            If srcWords(1).Equals("OPTIONAL") Then
              FDFileName = srcWords(2)
            Else
              FDFileName = srcWords(1)
            End If
            VaryingName = ""
            For x As Integer = 0 To srcWords.Count - 1
              If srcWords(x) = "ASSIGN" Then
                If srcWords(x + 1) <> "TO" Then
                  DDName = srcWords(x + 1)
                  If DDName = "VARYING" Then
                    VaryingName = srcWords(x + 2)
                  End If
                Else
                  If x + 2 <= srcWords.Count - 1 Then
                    DDName = srcWords(x + 2)
                    If DDName = "VARYING" Then
                      VaryingName = srcWords(x + 3)
                    End If
                  End If
                End If
                '
                ' if ddname is varying we need to find that value in working storage.
                If VaryingName.Length > 0 Then
                  DDName = GetVariablesValue(VaryingName)
                  If DDName.Length = 0 Then
                    DDName = VaryingName
                  End If
                End If
                '
              End If
            Next
            ListOfFiles.Add(srcWords(1) & Delimiter &
                            DDName & Delimiter &
                            "FILE" & Delimiter &
                            LTrim(Str(stmtIndex)))
            Continue For
          End If

          If srcWords(0) = "DATA-VIEW" Then
            Dim dvDBIdIndex As Integer = srcWords.IndexOf("DATA-BASE-IDENTIFICATION")
            If dvDBIdIndex > -1 Then
              If srcWords(dvDBIdIndex + 1) = "IS" Then
                DDName = srcWords(dvDBIdIndex + 2)
              Else
                DDName = srcWords(dvDBIdIndex + 1)
              End If
            End If
            ListOfFiles.Add(srcWords(1) & Delimiter &
                            DDName & Delimiter &
                            "Dataview" & Delimiter &
                            LTrim(Str(stmtIndex)))
            Continue For
          End If

          If srcWords.Count >= 5 Then
            If srcWords(0) = "EXEC" And
              srcWords(1) = "SQL" And
              srcWords(2) = "DECLARE" And
              (srcWords(4) = "TABLE" Or srcWords(4) = "CURSOR") Then
              ListOfFiles.Add(srcWords(3) & Delimiter &
                            "" & Delimiter &
                            "SQL" & Delimiter &
                            LTrim(Str(stmtIndex)))
            End If
          End If
        Next

      Case "Easytrieve"
        For stmtIndex As Integer = pgm.EnvironmentDivision To pgm.ProcedureDivision - 1
          statement = SrcStmt(stmtIndex)
          If statement.Length >= 1 Then
            If statement.Substring(0, 1) = "*" Then
              Continue For
            End If
          End If
          Call GetSourceWords(statement, srcWords)
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

        Next

    End Select
    Return ListOfFiles
  End Function
  Function GetVariablesValue(ByRef VaryingName As String) As String
    'search through the Data Division for given VaryingName and grab its compile time value
    ' note. NOT handling Redefines
    Dim valueOfVariable As String = ""
    For index As Integer = pgm.DataDivision To pgm.ProcedureDivision - 1
      If SrcStmt(index).Substring(0, 1) = "*" Then
        Continue For
      End If
      If SrcStmt(index).IndexOf(VaryingName) > -1 Then
        'variable name is found. get it's level #.
        Dim vnWords As New List(Of String)
        Call GetSourceWords(SrcStmt(index), vnWords)
        Dim vIndex As Integer = vnWords.IndexOf("VALUE")
        If vIndex > -1 Then
          Return vnWords(vIndex + 1)
        End If
        ' The VALUE is not on the variable name; maybe its on is child fields
        Dim vnLevel As Integer = Val(vnWords(0))
        For valueIndex As Integer = index + 1 To pgm.ProcedureDivision - 1
          If SrcStmt(valueIndex).Substring(0, 1) = "*" Then
            Continue For
          End If
          Dim vaWords As New List(Of String)
          Call GetSourceWords(SrcStmt(valueIndex), vaWords)
          If Not IsNumeric(vaWords(0)) Then
            Exit For
          End If
          Dim vaLevel As Integer = vaWords(0)
          If vaLevel <= vnLevel Then
            Exit For
          End If
          Dim vaIndex As Integer = vaWords.IndexOf("VALUE")
          If vaIndex > -1 Then
            valueOfVariable &= vaWords(vaIndex + 1)
          End If
        Next
        Return valueOfVariable
      End If
    Next
    Return valueOfVariable
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
    For FDIndex As Integer = pgm.DataDivision + 1 To pgm.WorkingStorage - 1
      If SrcStmt(FDIndex).Length > 0 Then
        If SrcStmt(FDIndex).Substring(0, 1) = "*" Then
          Continue For
        End If
      End If
      Call GetSourceWords(SrcStmt(FDIndex), FDWords)
      If FDWords.Count >= 2 Then
        If FDWords(0) = "FILE" And FDWords(1) = "SECTION" Then
          Continue For
        End If
      End If
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
            For recIndex = FDIndex + 1 To pgm.WorkingStorage - 1
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
        For ReadIndex As Integer = pgm.ProcedureDivision + 1 To pgm.EndProgram
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
                    recordNameIndex = FindWSRecordNameIndex(pgm.WorkingStorage, recname)
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

    Return ListOfRecordNames
  End Function
  Function GetListOfRecordNamesSQL(ByRef filename As String, ByRef DeclareIndex As Integer) As List(Of String)
    ' starting at the DeclareIndex look for the record name (01 level)
    ' return: record name,index,level,recfm,minlrecl,maxlrecl,organization

    ' Presume that SQL only has one Record name (01 level)

    Dim SQLRecordNames As New List(Of String)
    Dim SelectWords As New List(Of String)

    Dim recnameOpenMode As String = GetOpenModeSQL(filename)
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
    Return SQLRecordNames
  End Function
  Function GetListOfRecordNamesDataview(ByRef filename As String, ByRef FileIndex As Integer) As List(Of String)
    ' Use the Data Division index to search Stmt array to get the Data-View record details,
    ' The next stmt record is the "01-Level" with the Dataview's record name.
    ' Could be multiple 01-levels(record names) for this file
    ' 0-Record Name,
    ' 1-index to record name
    ' 2-Level (DV)
    ' 3-OpenMode (I,O)
    ' 4-recfm (F)
    ' 5-minlrecl 
    ' 6-maxlrecl
    ' 7-organization (DataCom/DB)
    ' Note. variables with 'DV' mean DataView.
    Dim DVWords As New List(Of String)
    Dim ListOfRecordNames As New List(Of String)
    Dim DVDetails As String()
    Dim DVDetail As String = ""
    Dim DVDetaillevel As String = ""
    Dim DVDetailOpenMode As String = ""
    Dim DVDetailrecfm As String = ""
    Dim DVDetailminLrecl As String = ""
    Dim DVDetailmaxLrecl As String = ""
    Dim DVDetailorganization As String = ""
    Dim recname As String = ""
    Dim DVRecName As String = ""      'first DV 01-level (Workarea)
    Dim recIndex As Integer = 0
    For DVIndex As Integer = pgm.DataDivision To pgm.ProcedureDivision
      Call GetSourceWords(SrcStmt(DVIndex), DVWords)
      If DVWords.Count >= 1 Then
        If DVWords(0) = "DATA-VIEW" Then
          If DVWords(1) = filename Then
            Call GetSourceWords(SrcStmt(FileIndex), cWord)
            DVDetail = GetDVDetails(cWord)
            DVDetails = DVDetail.Split(Delimiter)
            DVDetaillevel = DVDetails(4)
            DVDetailOpenMode = DVDetails(5)
            DVDetailrecfm = DVDetails(6)
            DVDetailminLrecl = DVDetails(7)
            DVDetailmaxLrecl = DVDetails(8)
            DVDetailorganization = DVDetails(10)
            recname = DVDetails(11)
            recIndex = DVIndex
            Exit For
          End If
        End If
      End If
    Next

    ListOfReadIntoRecords.Add(recname)
    ListOfRecordNames.Add(recname & Delimiter &
                          LTrim(Str(recIndex)) & Delimiter &
                          "DV" & Delimiter &
                          "I-O" & Delimiter &
                          "F" & Delimiter &
                          LTrim(Str(0)) & Delimiter &
                          "0" & Delimiter &
                          "Datacom/DB")

    Return ListOfRecordNames
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
          LogFile.WriteLine(Date.Now & ",E,Redefine parent not found!," & fields.FieldName & "[" & FileNameOnly & "]")
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
    GetRecordLength = totalRecordLength
  End Function
  Function FindCopybookName(ByRef DataIndex As Integer, ByVal RecordName As String) As String
    ' Use the Data Division index to search Stmt array to get the Record  location,
    ' then look previous lines to see what the possible copybook name would be.
    ' look upward to see if we find 'COPY/INCLUDE/SQL INCLUDE' statement
    'Here are some examples:
    '  COBOL
    '*COPY CRCALC.
    '    EXEC SQL INCLUDE SQLCA END-EXEC.
    '*INCLUDE++ PM044016  (from Panvalet)
    '* Easytrieve
    '*%COPYBOOK SQL CRPVBI
    '*%COPYBOOK FILE PM044025
    ' we can stop searching at a NON-commented line.

    Dim CopyWords As New List(Of String)
    Dim myCopybookName As String = ""

    myCopybookName = ""
    For CopyIndex As Integer = DataIndex - 1 To pgm.DataDivision Step -1
      myCopybookName = FindCopyOrInclude(CopyIndex, CopyWords)
      If myCopybookName.Length > 0 Then
        Exit For
      End If
    Next
    If myCopybookName = "EXIT FOR" Then
      myCopybookName = ""
    End If
    If myCopybookName.Length > 0 Then
      Return myCopybookName
    End If

    ' if still no copybook found at this point then maybe the copy/include is
    ' placed after the record name; so search downward. ugh.
    myCopybookName = ""
    For CopyIndex As Integer = DataIndex + 1 To pgm.ProcedureDivision
      myCopybookName = FindCopyOrInclude(CopyIndex, CopyWords)
      If myCopybookName.Length > 0 Then
        Exit For
      End If
    Next
    If myCopybookName = "EXIT FOR" Then
      myCopybookName = ""
    End If
    If myCopybookName.Length = 0 Then
      myCopybookName = "NONE"
    End If
    Return myCopybookName
  End Function
  Sub ProcessPumlParagraph(ByRef ParagraphStarted As Boolean, ByRef statement As String, ByRef exec As String)
    If ParagraphStarted = True Then
      pumlLineCnt += 2
      pumlFile.WriteLine("end")
      pumlFile.WriteLine("")
    End If

    If pumlLineCnt > pumlMaxLineCnt Then
      pumlLineCnt += 1
      pumlFile.WriteLine("floating note left: Continued in Part " & pumlPageCnt + 1)
      pumlFile.WriteLine("@enduml")
      pumlFile.Close()
      PumlPageBreak(exec)
    End If

    pumlFile.WriteLine("start")
    pumlFile.WriteLine(":**" & Trim(statement.Replace(".", "")) & "**;")
    pumlLineCnt += 2
    ParagraphStarted = True
  End Sub
  Sub ProcessPumlParagraphEasytrieve(ByRef ParagraphStarted As Boolean)
    For index As Integer = 0 To cWord.Count - 1
      Select Case cWord(index)
        Case "JOB"
          If ParagraphStarted Then
            pumlLineCnt += 2
            IndentLevel = 1
            pumlFile.WriteLine("end")
            pumlFile.WriteLine("}")
            ParagraphStarted = False
          End If
          pumlLineCnt += 2
          pumlFile.WriteLine("partition **" & cWord(index) & "** {")
          pumlFile.WriteLine("start")
          ParagraphStarted = True
          IndentLevel = 1
          IFLevelIndex.Clear()
        Case "INPUT"
          If cWord(index + 1) <> "NULL" Then
            WithinReadStatement = True
            Dim StartIndex = index
            Dim EndIndex As Integer = StartIndex + 1
            Dim TogetherWords As String = StringTogetherWords(cWord, StartIndex, EndIndex)
            Dim ReadStatement As String = AddNewLineAboutEveryNthCharacters(TogetherWords, ESCAPENEWLINE, 30)
            pumlLineCnt += 1
            pumlFile.WriteLine(Indent() & ":" & ReadStatement.Trim & "/")
          End If
          index += 1
        Case "START", "FINISH"
          Dim StartIndex = index + 1
          Dim EndIndex As Integer = StartIndex
          Dim TogetherWords As String = StringTogetherWords(cWord, StartIndex, EndIndex)
          Dim ReadStatement As String = AddNewLineAboutEveryNthCharacters(TogetherWords, ESCAPENEWLINE, 30)
          pumlLineCnt += 1
          pumlFile.WriteLine(Indent() & ":PERFORM " & ReadStatement.Trim & "|")
          index += 1
        Case "REPORT"
          If ParagraphStarted Then
            IndentLevel = 1
            pumlLineCnt += 2
            pumlFile.WriteLine("end")
            pumlFile.WriteLine("}")
            ParagraphStarted = False
          End If
          pumlLineCnt += 2
          Dim theDD As String = ""
          If cWord.Count = 1 Then
            theDD = "SYSPRINT"
          Else
            theDD = cWord(index + 1)
          End If
          pumlFile.WriteLine("partition **" & cWord(index) & " " & theDD & "** {")
          pumlFile.WriteLine("start")
          ParagraphStarted = True
          IndentLevel = 1
          IFLevelIndex.Clear()
          Exit For
        Case "PROC"
          If ParagraphStarted Then
            pumlLineCnt += 2
            IndentLevel = 1
            pumlFile.WriteLine("end")
            pumlFile.WriteLine("}")
            ParagraphStarted = False
          End If
          If index > 0 Then     'proc name is on same line as PROC statement us it.
            theProcName = cWord(index - 1).Replace(".", "")
          Else
            If theProcName.Length = 0 Then
              theProcName = "ProcNameUndefine"
            End If
          End If
          theProcName = theProcName.Replace(".", "")
          pumlLineCnt += 2
          pumlFile.WriteLine("partition **" & theProcName & "** {")
          pumlFile.WriteLine("start")
          ParagraphStarted = True
          IndentLevel = 1
          IFLevelIndex.Clear()
          theProcName = ""
      End Select
    Next
  End Sub
  Sub ProcessPumlPerformEasytrieve(ByRef WordIndex As Integer)
    Dim EndIndex As Integer = 0
    Dim MiscStatement As String = ""
    Call GetStatement(WordIndex, EndIndex, MiscStatement)
    pumlLineCnt += 1
    pumlFile.WriteLine(Indent() & ":" & MiscStatement.Trim & "|")
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlSortEasytrieve(ByRef WordIndex As Integer)
    ' note that cWord is global
    Dim EndIndex As Integer = cWord.Count - 1
    Dim TogetherWords As String = StringTogetherWords(cWord, WordIndex, (cWord.Count - 1))
    Dim SortStatement As String = AddNewLineAboutEveryNthCharacters(TogetherWords, ESCAPENEWLINE, 30)

    pumlLineCnt += 2
    pumlFile.WriteLine("start")
    pumlFile.WriteLine(":" & SortStatement.Trim & ";")
    WordIndex = cWord.Count - 1
  End Sub
  Sub ProcessPumlIFEasytrieve(ByRef WordIndex As Integer)
    ' find the 'IF' aka Conditional statement
    ' Indentlevel is global
    Dim EndIndex As Integer = cWord.Count - 1
    Dim Statement As String = ""
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlLineCnt += 1
    pumlFile.WriteLine(Indent() & "if (" & Statement.Trim & ") then (yes)")
    IndentLevel += 1
    WordIndex = EndIndex
  End Sub

  Sub ProcessPumlSelectEasytrieve(ByRef WordIndex As Integer)
    Dim EndIndex As Integer = cWord.Count - 1
    Dim MiscStatement As String = StringTogetherWords(cWord, WordIndex, EndIndex)
    MiscStatement = AddNewLineAboutEveryNthCharacters(MiscStatement, ESCAPENEWLINE, 30)
    pumlLineCnt += 1
    pumlFile.WriteLine(Indent() & ":" & MiscStatement.Trim & ";")
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlELSE(ByRef WordIndex As Integer)
    IndentLevel -= 1
    pumlLineCnt += 1
    pumlFile.WriteLine(Indent() & "else (no)")
    IndentLevel += 1

  End Sub
  Sub ProcessPumlCase(ByRef wordIndex As Integer)
    'TODO: need to fix embedded Evaluates
    ' find the end of 'EVALUATE' / 'CASE' statement which should be at the first 'WHEN' clause
    'cWord is global
    'IndentLevel is global
    Dim Statement As String = ""
    Dim EndIndex As Integer = cWord.Count - 1
    Call GetStatement(wordIndex, EndIndex, Statement)
    pumlLineCnt += 1
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
      pumlLineCnt += 1
      pumlFile.WriteLine(Indent() & "if (" & Statement.Trim & ") then (yes)")
      IndentLevel += 1
    Else
      IndentLevel -= 1
      pumlLineCnt += 1
      pumlFile.WriteLine(Indent() & "elseif (" & Statement.Trim & ") then (yes)")
      IndentLevel += 1
    End If
    wordindex = EndIndex

  End Sub
  Sub ProcessPumlENDEVALUATE(ByRef wordindex As Integer)
    'TODO: Need to handle embedded end-evaluate
    FirstWhenStatement = False
    IndentLevel -= 1
    pumlLineCnt += 1
    pumlFile.WriteLine(Indent() & "endif")
  End Sub
  Sub ProcessPumlCOMPUTE(ByRef WordIndex As Integer)
    ' find the end of 'COMPUTE' statement
    Dim EndIndex As Integer = 0
    Dim Statement As String = ""
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlLineCnt += 1
    pumlFile.WriteLine(Indent() & ":" & Statement.Trim & ";")
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlGOTO(ByRef WordIndex As Integer)
    ' find the end of 'GO TO' statement
    Dim EndIndex As Integer = 0
    Dim Statement As String = ""
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlLineCnt += 2
    pumlFile.WriteLine(Indent() & "#pink:" & Statement.Trim & ";")
    pumlFile.WriteLine(Indent() & "detach")
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlGET(ByRef WordIndex As Integer)
    ' Next word is file/record name
    Dim EndIndex As Integer = WordIndex + 1
    Dim Statement As String = ""
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlLineCnt += 1
    pumlFile.WriteLine(Indent() & ":" & Statement.Trim & "/")
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
    pumlLineCnt += 1
    pumlFile.WriteLine(Indent() & "while (" & Statement.Trim & ") is (true)")
    IndentLevel += 1
    WordIndex = EndIndex
  End Sub

  Sub ProcessPumlENDDO(ByRef wordindex As Integer)
    IndentLevel -= 1
    pumlLineCnt += 1
    pumlFile.WriteLine(Indent() & "endwhile (Complete)")
  End Sub
  Sub ProcessPumlEXEC(ByRef WordIndex As Integer)
    Dim EndIndex As Integer = 0
    Dim EXECStatement As String = ""
    Call GetStatement(WordIndex, EndIndex, EXECStatement)
    pumlLineCnt += 1
    pumlFile.WriteLine(Indent() & ":" & EXECStatement.Trim & ">")
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
    EndIndex = IndexToNextVerb(cWord, WordIndex)
    If EndIndex = -1 Then
      EndIndex = cWord.Count - 1
    End If
    Dim WordsTogether As String = StringTogetherWords(cWord, WordIndex, EndIndex)
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
    Dim RecordingMode As String = "V"
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

        ' Find the FD clauses. At present only care about Recording Mode and Record Contains

        ' define and clear out the FD clause indexes
        Dim fdClauseName(10) As String
        Dim fdClauseFirstIndex(10) As Integer
        Dim fdClauseLastIndex(10) As Integer
        For x = 0 To 10
          fdClauseName(x) = ""
          fdClauseFirstIndex(x) = fdWords.Count
          fdClauseLastIndex(x) = fdWords.Count - 1
        Next
        ' find and set the FD clause name and starting word indexes
        Dim fdClauseIndex As Integer = -1
        Dim fdClauseIndexMax As Integer = -1
        For FDindex As Integer = 0 To fdWords.Count - 1
          Select Case fdWords(FDindex)
            Case "BLOCK", "LABEL", "VALUE", "DATA", "RECORDING", "LINAGE", "CODESET"
              Call AddToFDClause(fdWords(FDindex), FDindex, fdClauseIndex, fdClauseName, fdClauseFirstIndex, fdClauseLastIndex)
            Case "RECORD"
              If Not (fdWords(FDindex - 1) = "LABEL" Or fdWords(FDindex - 1) = "DATA") Then
                Call AddToFDClause(fdWords(FDindex), FDindex, fdClauseIndex, fdClauseName, fdClauseFirstIndex, fdClauseLastIndex)
              End If
          End Select
        Next
        fdClauseIndexMax = fdClauseIndex
        ' set the LAST index for each clause
        For FDIndex As Integer = 0 To fdClauseIndexMax
          fdClauseLastIndex(FDIndex) = fdClauseFirstIndex(FDIndex + 1) - 1
        Next
        'Now loop through the FD Clauses
        For fdClauseIndex = 0 To fdClauseIndexMax
          Select Case fdClauseName(fdClauseIndex)
            Case "RECORDING"
              RecordingMode = fdWords(fdClauseLastIndex(fdClauseIndex))
            Case "RECORD"
              For index = fdClauseFirstIndex(fdClauseIndex) + 1 To fdClauseLastIndex(fdClauseIndex)
                If IsNumeric(fdWords(index)) Then
                  If RecordSizeMinimum = 0 Then
                    RecordSizeMinimum = fdWords(index)
                  Else
                    RecordSizeMaximum = fdWords(index)
                  End If
                End If
              Next
          End Select
        Next

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
  Function GetDVDetails(ByRef cWord As List(Of String)) As String
    ' Analyze a File's Definition based on the Data-View syntax
    ' COBOL
    '   For the given Data-view (provided in cWord).
    '   This parses the Data-View statements.
    ' Easytrieve
    '   NOT APPLICABLE (I HOPE)
    ' Returns a string separated with delimiter:
    '0-FileNameOnly (Source)
    '1-pgmName (pgmid)
    '2-pgmSeq
    '3-file_name_1 (dataview-name)
    '4-Level  (DV)
    '5-OpenMode (Input/Output)
    '6-RecordingMode (Fixed) RECFM
    '7-RecordSizeMinimum
    '8-RecordSizeMaximum
    '9-assignment_name_1 (DBID)
    '10-organization (Datacom/DB)
    '11-record name (workarea)

    Dim file_name_1 As String = ""
    Dim level As String = "DV"
    Dim OpenMode As String = ""
    Dim RecordingMode As String = "F"
    Dim RecordSizeMinimum As Integer = 0
    Dim RecordSizeMaximum As Integer = 0
    Dim assignment_name_1 As String = ""
    Dim organization As String = ""
    Dim RecordName As String = ""
    Dim fdWords As New List(Of String)

    For dvIndex As Integer = 0 To cWord.Count - 1
      Select Case cWord(dvIndex)
        Case "DATA-VIEW"
          file_name_1 = cWord(dvIndex + 1)
        Case "WORKAREA"
          RecordName = cWord(dvIndex + 1)
        Case "ORGANIZATION"
          If cWord(dvIndex + 1) = "IS" Then
            organization = cWord(dvIndex + 2)
          Else
            organization = cWord(dvIndex + 1)
          End If
        Case "DATA-BASE-IDENTIFICATION"
          If cWord(dvIndex + 1) = "IS" Then
            assignment_name_1 = cWord(dvIndex + 2)
          Else
            assignment_name_1 = cWord(dvIndex + 1)
          End If
        Case "ACCESSED"
        Case "ELEMENTS"
        Case "FILE"
      End Select
    Next

    OpenMode = "I-O"

    Return (FileNameOnly & Delimiter &
                         pgmName & Delimiter &
                         LTrim(Str(pgmSeq)) & Delimiter &
                         file_name_1 & Delimiter &
                         level & Delimiter &
                         RTrim(OpenMode) & Delimiter &
                         RecordingMode & Delimiter &
                         LTrim(Str(RecordSizeMinimum)) & Delimiter &
                         LTrim(Str(RecordSizeMaximum)) & Delimiter &
                         assignment_name_1 & Delimiter &
                         organization & Delimiter &
                         RecordName)

  End Function

  Sub AddToFDClause(ByRef NameOfClause As String,
                    ByRef fdIndex As Integer,
                         ByRef fdClauseIndex As Integer,
                         ByRef fdClauseName() As String,
                         ByRef fdClauseFirstIndex() As Integer,
                         ByRef fdClauseLastIndex() As Integer)
    ' this will add an entry to the FDClause arrays
    fdClauseIndex += 1
    fdClauseName(fdClauseIndex) = NameOfClause
    fdClauseFirstIndex(fdClauseIndex) = fdIndex
    fdClauseLastIndex(fdClauseIndex) = -1
  End Sub
  Function GetIndexForRecordSize(ByVal index As Integer, ByRef fdWords As List(Of String)) As Integer
    GetIndexForRecordSize = index
    Select Case True
      Case IsNumeric(fdWords(index)) : Exit Select
      Case IsNumeric(fdWords(index + 1)) : GetIndexForRecordSize += 1
      Case IsNumeric(fdWords(index + 2)) : GetIndexForRecordSize += 2
      Case IsNumeric(fdWords(index + 3)) : GetIndexForRecordSize += 3
      Case Else
        'MessageBox.Show("Unknown 'IS VARYING' syntax@" & pgmName & "FD:" & fdWords.ToString)
        Dim tempx As String = pgmName & " FDwords(0)=" & fdWords(0) & " index=" & Str(index)
        LogFile.WriteLine(Date.Now & ",E,Unknown 'IS VARYING' syntax," & tempx)
        Exit Select
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
            Case "I-O,"
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
              If x >= 1 Then
                If srcWords(x - 2) = "EXEC" And srcWords(x - 1) = "SQL" Then
                  Exit For
                End If
              End If
              MessageBox.Show("Never found open mode for file:" &
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
    Dim srcWords As New List(Of String)
    Dim ListOfOpenModes As New List(Of String)

    ' parse out the TABLE portion of the given SQL filename
    Dim myTableName As String = ""
    If filename.Contains(".") Then
      Dim myFileNameParts As String() = filename.Split(".")
      If myFileNameParts.Count > 1 Then
        myTableName = myFileNameParts(1)
      Else
        myTableName = filename
      End If
    Else
      myTableName = filename
    End If


    For Index As Integer = pgm.ProcedureDivision + 1 To pgm.EndProgram
      If SrcStmt(Index).Substring(0, 1) = "*" Then
        Continue For
      End If
      Call GetSourceWords(SrcStmt(Index), srcWords)
      If srcWords.Count <= 6 Then
        Continue For
      End If
      Dim endExecIndex As Integer = -1
      For cblIndex = 0 To srcWords.Count - 1
        If (cblIndex + 1) > (srcWords.Count - 1) Then
          Exit For
        End If
        If srcWords(cblIndex) = "EXEC" And srcWords(cblIndex + 1) = "SQL" Then
          ' find the next END-EXEC
          If endExecIndex = -1 Then
            For x = cblIndex + 2 To srcWords.Count - 1
              If srcWords(x) = "END-EXEC" Then
                endExecIndex = x
              End If
            Next x
            If endExecIndex = -1 Then
              endExecIndex = srcWords.Count - 1
            End If
          End If

          ' first thing is to figure out the table name location
          ' SELECT uses FROM next word, DELETE uses FROM next word
          ' INSERT uses INTO next word, UPDATE uses next word
          Dim tableNameIndx As Integer = -1
          Select Case srcWords(cblIndex + 2)
            Case "SELECT", "DELETE"
              tableNameIndx = srcWords.IndexOf("FROM") + 1
            Case "INSERT"
              tableNameIndx = srcWords.IndexOf("INTO") + 1
            Case "UPDATE"
              tableNameIndx = cblIndex + 3
            Case "FETCH", "OPEN", "CLOSE", "COMMIT", "DECLARE", "SET"
              cblIndex = endExecIndex
              Continue For
          End Select
          If tableNameIndx = -1 Then
            MessageBox.Show("GetOpenModeSQL: Unknown SQL statement:" & SrcStmt(Index))
            Exit For
          End If
          Dim tableName As String = srcWords(tableNameIndx)
          Dim filenamefound As Boolean = False
          ' find any part of that filename (could have DB Qualifier on it)
          If filename.Contains(tableName) Then
            filenamefound = True
            If ListOfOpenModes.IndexOf(srcWords(cblIndex + 2)) = -1 Then
              ListOfOpenModes.Add(srcWords(cblIndex + 2))
            End If
          End If
          cblIndex = endExecIndex
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
        If SrcStmt(index).Contains(myTableName) Then
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
    Return modes.Trim()
  End Function
  Sub GetSourceWords(ByVal statement As String, ByRef srcWords As List(Of String))
    ' takes the stmt and breaks into words and drops blanks
    srcWords.Clear()
    ' example of statement:" DISPLAY '*** CRCALCX REC READ        = ' WS-REC-READ.   "
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
  Function IsParagraph(ByRef CobolWords As List(Of String)) As Boolean
    ' Identify if the stmt is a paragraph or a section name.
    If CobolWords.Count <> 1 Then
      If CobolWords.Count = 2 Then
        If CobolWords(1) = "SECTION" Then
          Return True
        End If
        If CobolWords(1) = "EXIT" Then      'exit is on same line as paragraph name...ugh!
          Return True
        End If
      End If
      Return False
    End If
    Select Case CobolWords(0)
      Case "GOBACK", "EXIT"
        Return False
    End Select
    Return True
  End Function

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
    ' Use the Data Division working storage index to search Stmt array to get the WS record index/location,
    ' 
    Dim FDWords As New List(Of String)
    Dim RecordWords As New List(Of String)
    Dim FDIndex As Integer = -1
    '
    For FDIndex = DataIndex + 1 To pgm.ProcedureDivision
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
      Return FDIndex
    End If
    Return -1
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
    ' look for various types of copy/include commands once found return copybook name
    Call GetSourceWords(SrcStmt(CopyIndex), CopyWords)
    If CopyWords.Count >= 3 Then
      If CopyWords(0) = "*COPY" And CopyWords(2) = "BEGIN" Then
        Return CopyWords(1)
      End If
    End If
    If CopyWords.Count >= 5 Then
      If CopyWords(1) = "EXEC" And CopyWords(2) = "SQL" And CopyWords(3) = "INCLUDE" Then
        Return CopyWords(4)
      End If
    End If
    If CopyWords.Count >= 2 Then
      If CopyWords(0) = "*INCLUDE++" Then
        Return CopyWords(1)
      End If
      If CopyWords(0) = "*%COPYBOOK" Then
        Return CopyWords(2)
      End If
    End If
    If IsNumeric(CopyWords(0)) Then
      Return "EXIT FOR"
    End If
    Select Case CopyWords(0)
      Case "FILE", "FD", "WORKING-STORAGE", "LOCAL-STORAGE", "LINKAGE"
        Return "EXIT FOR"
    End Select
    Return ""
  End Function
  Function StringTogetherWords(CobWords As List(Of String), ByRef StartCondIndex As Integer, ByRef EndCondIndex As Integer) As String
    ' string together from startofconditionindex to endofconditionindex
    ' cWord is a global variable
    Dim wordsStrungTogether As String = ""
    For condIndex As Integer = StartCondIndex To EndCondIndex
      wordsStrungTogether &= CobWords(condIndex) & " "
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
      Return ""
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
    Return condStatementCR
  End Function
  Function IndexToNextVerb(cobWords As List(Of String), ByRef StartCondIndex As Integer) As Integer
    ' cWord is a global variable
    ' VerbNames is a global variable
    ' find ending index to next COBOL verb in cWord
    Dim EndCondIndex As Integer = -1
    Dim VerbIndex As Integer = -1
    For EndCondIndex = StartCondIndex + 1 To cobWords.Count - 1
      If WithinReadStatement = True Then
        Select Case cobWords(EndCondIndex)
          Case "AT", "END", "NOT"
            Return EndCondIndex
          Case "NEXT"
            Continue For
        End Select
      End If
      VerbIndex = VerbNames.IndexOf(cobWords(EndCondIndex))
      If VerbIndex > -1 Then
        Return EndCondIndex - 1
      End If
    Next
    ' there is not another verb in this statement
    Return -1
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
    lblStatusMessage.Text = ""

    ' Initialize all beginning Variables and tables.
    ' JCL
    FileNameOnly = ""
    tempCobFileName = ""
    tempEZTFileName = ""
    jControl = ""
    jLabel = ""
    jParameters = ""
    procName = ""
    jobName = ""
    jobClass = ""
    jobMsgClass = ""
    prevPgmName = ""
    prevStepName = ""
    prevDDName = ""
    pgmName = ""
    DDName = ""
    stepName = ""
    InstreamProc = ""
    ddConcatSeq = 0
    ddSequence = 0
    jclStmt.Clear()
    ListOfExecs.Clear()
    ' COBOL fields
    SourceType = ""
    SrcStmt.Clear()
    cWord.Clear()
    ListOfPrograms.Clear()
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
    FirstWhenStatement = False
    WithinReadStatement = False
    WithinReadConditionStatement = False
    WithinIF = False
    pgmSeq = 0

  End Sub
  Function GetInitialDirectory(ByRef InitDirectory As String) As String
    ' get/set the initial directory (aka Sandbox) which holds all the application directories
    ' This should come from the My.Settings.InitDirectory properties, if path is set

    ' try the properties default value (my folder as distributed)
    InitDirectory = My.Settings.InitDirectory
    If Directory.Exists(InitDirectory) Then
      Return InitDirectory
    End If
    MessageBox.Show("Initial Directory of Sandbox not found:" &
                    vbCrLf & InitDirectory & vbCrLf &
                    "You will now be prompted to locate a Sandbox directory")

    ' prompt for and select an initial directory name;
    '  you can create the initial folder here
    '    but you cannot create a sandbox directory
    Dim bfd_InitFolder As New FolderBrowserDialog With
      {
        .Description = "Enter Directory where Sandbox folders will reside",
        .ShowNewFolderButton = True,
        .SelectedPath = Environment.SpecialFolder.Personal
      }
    Select Case bfd_InitFolder.ShowDialog
      Case DialogResult.OK
        InitDirectory = bfd_InitFolder.SelectedPath
        My.Settings.InitDirectory = InitDirectory                             'also now save to distributed 
        My.Settings.Save()
        lblInitDirectory.Text = InitDirectory
        Environment.SetEnvironmentVariable("ADDILite", InitDirectory, EnvironmentVariableTarget.User)
        Return InitDirectory
    End Select
    MessageBox.Show("Initial Directory set cancelled. Try Sandbox button.")
    Return ""
  End Function
  Private Sub btnSandbox_Click(sender As Object, e As EventArgs) Handles btnSandbox.Click
    'this will set / get the sandbox folder. This is the home directory of sandbox folders (applications)

    Dim bfd_InitFolder As New FolderBrowserDialog With {
            .Description = "Enter Directory where Sandbox folders will reside",
            .ShowNewFolderButton = True,
            .SelectedPath = InitDirectory
            }
    Select Case bfd_InitFolder.ShowDialog
      Case DialogResult.OK
        InitDirectory = bfd_InitFolder.SelectedPath
        My.Settings.InitDirectory = InitDirectory                             'also now save
        My.Settings.Save()
        lblInitDirectory.Text = InitDirectory
        Environment.SetEnvironmentVariable("ADDILite_Sandbox", InitDirectory, EnvironmentVariableTarget.User)
      Case DialogResult.Cancel
        Exit Sub
    End Select
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

    COBOLCondWords.Add("IF")
    COBOLCondWords.Add("<NOT>IF")           'special for Type of Rules determination
    COBOLCondWords.Add("<NOT><NOT>IF")           'special for Type of Rules determination
    COBOLCondWords.Add("THEN")
    COBOLCondWords.Add("IS")
    COBOLCondWords.Add("THAN")
    COBOLCondWords.Add("GREATER")
    COBOLCondWords.Add("LESS")
    COBOLCondWords.Add("EQUAL")
    COBOLCondWords.Add("TO")
    COBOLCondWords.Add("OR")
    COBOLCondWords.Add("AND")
    COBOLCondWords.Add("NOT")
    COBOLCondWords.Add("=")
    COBOLCondWords.Add("<=")
    COBOLCondWords.Add(">=")
    COBOLCondWords.Add(">")
    COBOLCondWords.Add("<")
    COBOLCondWords.Add("NUMERIC")
    COBOLCondWords.Add("ALPHABETIC")
    COBOLCondWords.Add("ALPHABETIC-LOWER")
    COBOLCondWords.Add("ALPHABETIC-UPPER")
    COBOLCondWords.Add("POSITIVE")
    COBOLCondWords.Add("NEGATIVE")
    COBOLCondWords.Add("DBCS")
    COBOLCondWords.Add("KANJI")
    COBOLCondWords.Add("SPACES")
    COBOLCondWords.Add("SPACE")
    COBOLCondWords.Add("ZEROES")
    COBOLCondWords.Add("ZERO")
    COBOLCondWords.Add("ZEROS")
    COBOLCondWords.Add("HIGH-VALUES")
    COBOLCondWords.Add("LOW-VALUES")
    COBOLCondWords.Add("ADDRESS")
    COBOLCondWords.Add("OF")
    COBOLCondWords.Add("NULL")
    COBOLCondWords.Add("NULLS")
    COBOLCondWords.Add("SELF")

    ' need to set up the Initial directory value.
    ' if we couldn't set initial directory program will have terminated
    InitDirectory = GetInitialDirectory(InitDirectory)
    lblInitDirectory.Text = InitDirectory
    If InitDirectory.Length = 0 Then
      Exit Sub
    End If

    btnADDILite.Enabled = True
    btnDataGatheringForm.Enabled = True

    txtJCLJOBFolder.Text = "\JOBS"
    txtProcFolder.Text = "\PROCS"
    txtCobolFolder.Text = "\COBOL"
    txtCopybookFolder.Text = "\COPYBOOKS"
    txtDECLGenFolder.Text = "\DECLGEN"
    txtEasytrieveFolder.Text = "\EASYTRIEVE"
    txtASMFolder.Text = "\ASM"
    txtTelonFolder.Text = "\TELON"
    txtScreenMapsFolder.Text = "\SCREENS"
    txtOutputFolder.Text = "\OUTPUT"
    txtPUMLFolder.Text = "\PUML"
    txtExpandedFolder.Text = "\EXPANDED"

  End Sub

End Class