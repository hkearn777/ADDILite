Imports System.IO
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
  Dim ProgramVersion As String = "v1.0"
  'Change-History.
  ' 2023/09/25 v1.1 hk New program
  '
  Const JOBCARD As String = " JOB "
  Const PROCCARD As String = " PROC "
  Const PENDCARD As String = " PEND "
  Const EXECCARD As String = " EXEC "
  Const DDCARD As String = " DD "
  Const DDCARDNOLABEL As String = "DD "
  Const SETCARD As String = " SET "
  Const SETCARDNOLABEL As String = "SET "
  Const OUTPUTCARD As String = " OUTPUT "
  Const IFCARD As String = "IF "
  Const ENDIFCARD As String = "ENDIF"
  Const QUOTE As Char = Chr(34)       'double-quote

  ' Arrays to hold the DB2 Declare to Member names
  ' these two array will share the same index
  Dim DB2Declares As New List(Of String)
  Dim MembersNames As New List(Of String)

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
  Dim pgmName As String = ""
  Dim DDName As String = ""
  Dim stepName As String = ""
  Dim InstreamProc As String = ""

  Dim ddConcatSeq As Integer = 0
  Dim ddSequence As Integer = 0
  Dim jobSequence As Integer = 0
  Dim procSequence As Integer = 0
  Dim execSequence As Integer = 0

  Dim RecordsRow As Integer = 0
  Dim FieldsRow As Integer = 0

  Dim jclStmt As New List(Of String)
  Dim ListOfExecs As New List(Of String)        'array holding the executable programs

  Dim swIPFile As StreamWriter = Nothing        'Instream proc file, temporary
  Dim swDDFile As StreamWriter = Nothing
  Dim swPumlFile As StreamWriter = Nothing
  Dim LogFile As StreamWriter = Nothing

  ' load the Excel References
  Dim objExcel As New Microsoft.Office.Interop.Excel.Application
  Dim workbook As Microsoft.Office.Interop.Excel.Workbook
  Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim RecordsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim FieldsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim rngRecordName As Microsoft.Office.Interop.Excel.Range
  Dim rngRecordsName As Microsoft.Office.Interop.Excel.Range
  Dim rngFieldsName As Microsoft.Office.Interop.Excel.Range

  ' COBOL fields
  Dim SourceType As String = ""
  Dim SrcStmt As New List(Of String)
  Dim cWord As New List(Of String)
  Dim ListOfFiles As New List(Of String)              'array to hold File & DB2 Table names
  Dim ListOfRecordNames As New List(Of String)          'array to hold read/written records
  Dim ListOfRecords As New List(Of String)              'array to hold read/written records
  Dim ListOfFields As New List(Of String)             'array to hold fields for each record
  Dim ListOfReadIntoRecords As New List(Of String)    'array to hold Read Into Records
  Dim ListOfWriteFromRecords As New List(Of String)   'array to hold Write from records
  Dim IFLevelIndex As New List(Of Integer)     'where in cWord the 'IF' is located
  Dim VerbNames As New List(Of String)
  Dim VerbCount As New List(Of Integer)
  Dim ProgramAuthor As String = ""
  Dim ProgramWritten As String = ""
  Dim IndentLevel As Integer = -1                  'how deep the if has gone
  Dim FirstWhenStatement As Boolean = False
  Dim WithinReadStatement As Boolean = False
  Dim WithinReadConditionStatement As Boolean = False
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

  Private Sub btnJCLJOBFilename_Click(sender As Object, e As EventArgs) Handles btnJCLJOBFilename.Click
    Dim ofd_InFile As New OpenFileDialog
    ' browse for and select file name
    ofd_InFile.Filter = "JCL|*.jcl|All|*.*"
    ofd_InFile.Title = "Select the JCL JOB file"
    If ofd_InFile.ShowDialog() = DialogResult.OK Then
      txtJCLJOBFilename.Text = ofd_InFile.FileName
      DirectoryName = Path.GetDirectoryName(ofd_InFile.FileName)
      FileNameOnly = Path.GetFileNameWithoutExtension(ofd_InFile.FileName)
    End If
  End Sub

  Private Sub btnJCLProclibFolder_Click(sender As Object, e As EventArgs) Handles btnJCLProclibFolder.Click
    ' browse for and select folder name
    Dim bfd_ProclibFolder As New FolderBrowserDialog With {
      .Description = "Enter Proclib folder name",
      .SelectedPath = DirectoryName
    }
    If bfd_ProclibFolder.ShowDialog() = DialogResult.OK Then
      txtJCLProclibFoldername.Text = bfd_ProclibFolder.SelectedPath
    End If
  End Sub

  Private Sub btnSourceFolder_Click(sender As Object, e As EventArgs) Handles btnSourceFolder.Click
    ' browse for and select folder name
    Dim bfd_SourceFolder As New FolderBrowserDialog With {
      .Description = "Enter Source folder name",
      .SelectedPath = DirectoryName
    }
    If bfd_SourceFolder.ShowDialog() = DialogResult.OK Then
      txtSourceFolderName.Text = bfd_SourceFolder.SelectedPath
    End If
  End Sub

  Private Sub btnOutputFolder_Click(sender As Object, e As EventArgs) Handles btnOutputFolder.Click
    ' browse for and select folder name
    Dim bfd_OutputFolder As New FolderBrowserDialog With {
      .Description = "Enter OUTPUT folder name",
      .SelectedPath = DirectoryName
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
    Delimiter = txtDelimiter.Text

    ' ready the progress bar
    ProgressBar1.Minimum = 0
    ProgressBar1.Maximum = 11
    ProgressBar1.Step = 1
    ProgressBar1.Value = 0

    Me.Cursor = Cursors.WaitCursor

    Dim logFileName As String = txtOutputFoldername.Text & "\" & FileNameOnly & "_log.txt"
    LogFile = My.Computer.FileSystem.OpenTextFileWriter(logFileName, False)
    LogFile.WriteLine(Date.Now & ",Program Starts," & Me.Text)
    LogFile.WriteLine(Date.Now & ",JCL JOB Filename," & txtJCLJOBFilename.Text)
    LogFile.WriteLine(Date.Now & ",JCL Proclib Folder," & txtJCLProclibFoldername.Text)
    LogFile.WriteLine(Date.Now & ",Source Folder," & txtSourceFolderName.Text)
    LogFile.WriteLine(Date.Now & ",Output Folder," & txtOutputFoldername.Text)
    LogFile.WriteLine(Date.Now & ",Delimiter," & txtDelimiter.Text)

    'validations
    If Not FileNamesAreValid() Then
      LogFile.WriteLine(Date.Now & ",File Names are not Valid,")
      Me.Cursor = Cursors.Default
      Exit Sub
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

    objExcel.Visible = False

    Dim ProgramsFileName = DirectoryName & "\" & FileNameOnly & ".xlsx"
    If File.Exists(ProgramsFileName) Then
      LogFile.WriteLine(Date.Now & ",Previous Excel file deleted," & ProgramsFileName)
      File.Delete(ProgramsFileName)
    End If


    ProcessJOBFile()

    ProgressBar1.PerformStep()
    ProgressBar1.Show()

    ProcessSourceFiles()

    ProgressBar1.PerformStep()

    '
    ' Save and close
    '
    Dim DefaultFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault
    workbook.SaveAs(ProgramsFileName, DefaultFormat)
    workbook.Close()
    objExcel.Quit()

    LogFile.WriteLine(Date.Now & ",Program Ends,")
    LogFile.Close()
    ProgressBar1.PerformStep()
    Me.Cursor = Cursors.Default
  End Sub
  Function FileNamesAreValid() As Boolean
    FileNamesAreValid = False
    Select Case True
      Case txtJCLJOBFilename.TextLength = 0
        LogFile.WriteLine(Date.Now & ",JCL JOB File name required,")
      Case Not IsValidFileNameOrPath(txtJCLJOBFilename.Text)
        LogFile.WriteLine(Date.Now & ",JCL JOB File name has invalid characters,")
      Case Not My.Computer.FileSystem.FileExists(txtJCLJOBFilename.Text)
        LogFile.WriteLine(Date.Now & ",JCL JOB File name not found,")

      Case txtOutputFoldername.TextLength = 0
        LogFile.WriteLine(Date.Now & ",OutFolder name required,")
      Case Not IsValidFileNameOrPath(txtOutputFoldername.Text)
        LogFile.WriteLine(Date.Now & ",OutFolder has invalid characters,")
      Case Not IsValidFileNameOrPath(txtSourceFolderName.Text)
        LogFile.WriteLine(Date.Now & ",Source folder name has invalid characters,")
      Case Not IsValidFileNameOrPath(txtJCLProclibFoldername.Text)
        LogFile.WriteLine(Date.Now & ",Proclib folder name has invalid characters,")
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
  Sub ProcessJOBFile()
    'Load the infile to the jclStmt List
    Dim jclRecordsCount As Integer = LoadJCLStatementsToArray()
    If jclRecordsCount = 0 Then
      MessageBox.Show("No JCL records")
      Exit Sub
    End If

    If jclStmt.Count = 0 Then
      MessageBox.Show("No JCL statements found on inFile")
    End If

    ProgressBar1.PerformStep()

    If WriteOutput() = -1 Then
      MessageBox.Show("Error while building output. See log file")
    End If

    ProgressBar1.PerformStep()

  End Sub
  Function LoadJCLStatementsToArray() As Integer
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
    Dim procIndex As Integer = 0
    Dim continuation As Boolean = False
    LoadJCLStatementsToArray = 0
    Dim debugCount As Integer = -1

    Dim JCLLines As String() = File.ReadAllLines(txtJCLJOBFilename.Text)
    '
    ' Write any Instream PROC(s) to a Proclib and drop it and drop empty lines
    ' Not handling Procs within Procs
    '
    Dim NumberOfInstreamProcsFound As Integer = 0
    Dim swTemp As StreamWriter = Nothing

    Dim ipName As String = ""
    '
    ' Create a Proc file from the Instream proc, if any
    ' And drop the instream proc lines 
    ' And remove any columns 73-80 values
    '
    swTemp = New StreamWriter(tempFileName, False)
    For index As Integer = 0 To JCLLines.Count - 1
      Dim jline As String = JCLLines(index) & Space(72)
      jline = jline.Substring(0, 72).Trim
      If jline.Substring(0, 2) = "/*" Then
        swTemp.WriteLine(jline)
        Continue For
      End If
      If jline.Substring(0, 3) = "//*" Then
        swTemp.WriteLine(jline)
        Continue For
      End If
      Dim procLocation As Integer = jline.IndexOf(" PROC ")
      If procLocation = -1 Then
        swTemp.WriteLine(jline)
        Continue For
      End If
      ipName = Trim(jline.Substring(2, procLocation - 3 + 1))
      Dim ipFileName = txtJCLProclibFoldername.Text & "\" & ipName & ".jcl"
      LogFile.WriteLine(Date.Now & ",Instream Proc Created," & ipFileName)
      swIPFile = New StreamWriter(ipFileName, False)
      For ipIndex As Integer = index To JCLLines.Count - 1
        Dim pline As String = JCLLines(ipIndex) & Space(72)
        pline = pline.Substring(0, 72).Trim
        swIPFile.WriteLine(pline)
        If pline.IndexOf(" PEND ") > -1 Then
          index = ipIndex
          Exit For
        End If
      Next
      swIPFile.Close()
    Next
    swTemp.Close()
    '
    ' Now include all PROCs from the ProcLib and place after the EXEC PROC statement
    '
    Dim jclWords As New List(Of String)
    JCLLines = File.ReadAllLines(tempFileName)
    swTemp = New StreamWriter(tempFileName, False)
    For index As Integer = 0 To JCLLines.Count - 1
      If JCLLines(index).Substring(0, 2) = "/*" Then
        swTemp.WriteLine(JCLLines(index))
        Continue For
      End If
      If JCLLines(index).Substring(0, 3) = "//*" Then
        swTemp.WriteLine(JCLLines(index))
        Continue For
      End If
      Call GetJCLWords(JCLLines(index), jclWords)
      If jclWords.Count >= 2 Then
        jControl = jclWords(1)
      End If
      If jControl <> "EXEC" Then
        swTemp.WriteLine(JCLLines(index))
        Continue For
      End If
      procName = GetParm(JCLLines(index), "PROC=")
      ' write all lines for this EXEC statement
      For contIndex As Integer = index To JCLLines.Count - 1
        swTemp.WriteLine(JCLLines(contIndex))
        If JCLLines(contIndex).Substring(0, 3) = "//*" Then
          Continue For
        End If
        If Microsoft.VisualBasic.
            Right(Trim(JCLLines(contIndex).PadRight(80).Substring(0, 70)), 1) <> "," Then
          index = contIndex
          Exit For
        End If
      Next

      If procName.Length = 0 Then
        Continue For
      End If
      ' if it is an "EXEC PROC=" copy the PROC file into here
      Dim ProcFileName As String = txtJCLProclibFoldername.Text & "\" & procName & ".jcl"
      LogFile.WriteLine(Date.Now & ",Processing PROC source," & ProcFileName)
      Dim procLines As String() = File.ReadAllLines(ProcFileName)
      For procIndex = 0 To procLines.Count - 1
        Dim jline As String = "++" & Mid(procLines(procIndex), 3) & Space(72)
        jline = jline.Substring(0, 72).Trim
        swTemp.WriteLine(jline)
      Next
    Next
    swTemp.Close()
    '
    ' Load JCL lines to a JCL statements array. 
    ' Basically dealing with continuations and removing comments
    '
    JCLLines = File.ReadAllLines(tempFileName)

    For index As Integer = 0 To JCLLines.Count - 1
      LoadJCLStatementsToArray += 1
      debugCount += 1
      text1 = JCLLines(index).Replace(vbTab, Space(1))
      ' drop comments
      If Mid(text1, 1, 3) = "//*" Or Mid(text1, 1, 3) = "++*" Then
        Continue For
      End If
      ' drop data (of an DD * statement) or not a JCL statement
      If Mid(text1, 1, 2) = "//" Or Mid(text1, 1, 2) = "++" Then
      Else
        Continue For
      End If
      If Mid(text1, 1, 14) = "//SEND OUTPUT " Then
        Continue For
      End If
      If Mid(text1, 1, 9) = "/*JOBPARM" Then
        Continue For
      End If
      ' format only the good stuff out of the line (no slashes, no comments)
      text1 = Trim(Microsoft.VisualBasic.Left(Mid(text1, 3) + Space(70), 70))
      ' determine if there will be a continuation
      If Microsoft.VisualBasic.Right(text1, 1) = "," Then
        continuation = True
      Else
        continuation = False
      End If
      ' Build the JCL statement
      jStatement &= text1
      ' if NOT continuing building of the JCL statement then add it to the List
      If continuation = False Then
        jclStmt.Add(jStatement)
        jStatement = ""
      End If
    Next

  End Function
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
    ' determine the value of the Parm looking for ending "," or ")"
    For x As Integer = FirstCharacter + SearchForThis.Length To SearchWithinThis.Length - 1
      ByteValue = SearchWithinThis.Substring(x, 1)
      Select Case ByteValue
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

    For Each statement As String In jclStmt
      Call GetLabelControlParms(statement, jLabel, jControl, jParameters)
      If Len(jControl) = 0 Then
        MessageBox.Show("JCL control not found:" & statement)
        WriteOutput = -1
        Exit For
      End If

      Select Case jControl
        Case "JOB"
          Call ProcessJOB()
        Case "PROC"
        Case "PEND"
        Case "EXEC"
          Call ProcessEXEC()
        Case "DD"
          Call ProcessDD()
        Case "SET"
          Continue For
        Case "OUTPUT"
          Continue For
        Case "IF"
          Continue For
        Case "ENDIF"
          Continue For
        Case Else
          MessageBox.Show("Unknown JCL control value:" & statement)
          Exit For
      End Select

    Next
    ' close the files

    swDDFile.Close()
    ProgressBar1.PerformStep()


    Call CreatePuml()
    ProgressBar1.PerformStep()

    Call CreatePrograms()
    ProgressBar1.PerformStep()

  End Function

  Sub GetLabelControlParms(statement As String,
                           ByRef jLabel As String,
                           ByRef jControl As String,
                           ByRef jParameters As String)
    'This will split out the three basic components of a JCL Statement
    'Each statement must have either JOB, PROC, EXEC, DD, PEND, SET, OUTPUT, IF, ENDIF
    Dim jLabelPrev As String = jLabel
    jControl = ""
    jParameters = ""
    For sPos As Integer = 1 To Len(statement)
      Select Case True
        Case Mid(statement, sPos, Len(JOBCARD)) = JOBCARD
          jLabel = RTrim(Mid(statement, 1, sPos - 1))
          jControl = Mid(statement, sPos + 1, 3)
          jParameters = Trim(Mid(statement, sPos + 5))
          Exit For
        Case Mid(statement, sPos, Len(PROCCARD)) = PROCCARD
          jLabel = RTrim(Mid(statement, 1, sPos - 1))
          jControl = Mid(statement, sPos + 1, 4)
          jParameters = Trim(Mid(statement, sPos + 6))
          Exit For
        Case Mid(statement, sPos, Len(PENDCARD)) = PENDCARD
          jLabel = RTrim(Mid(statement, 1, sPos - 1))
          jControl = Mid(statement, sPos + 1, 4)
          jParameters = Trim(Mid(statement, sPos + 6))
          Exit For
        Case Microsoft.VisualBasic.Left(statement, 5) = "PEND "
          jLabel = ""
          jControl = "PEND"
          jParameters = Trim(Mid(statement, 6))
          Exit For
        Case Mid(statement, sPos, Len(EXECCARD)) = EXECCARD
          jLabel = RTrim(Mid(statement, 1, sPos - 1))
          jControl = Mid(statement, sPos + 1, 4)
          jParameters = Trim(Mid(statement, sPos + 6))
          Exit For
        Case Mid(statement, sPos, Len(DDCARD)) = DDCARD
          jLabel = RTrim(Mid(statement, 1, sPos - 1))
          jControl = Mid(statement, sPos + 1, 2)
          jParameters = Trim(Mid(statement, sPos + 4))
          Exit For
        Case Mid(statement, sPos, Len(SETCARD)) = SETCARD
          jLabel = RTrim(Mid(statement, 1, sPos - 1))
          jControl = Mid(statement, sPos + 1, 3)
          jParameters = Trim(Mid(statement, sPos + 5))
          Exit For
        Case Mid(statement, 1, Len(DDCARDNOLABEL)) = DDCARDNOLABEL    'concat DD
          jLabel = jLabelPrev
          jControl = Mid(statement, 1, 2)
          jParameters = Trim(Mid(statement, 4))
          Exit For
        Case Mid(statement, 1, Len(SETCARDNOLABEL)) = SETCARDNOLABEL
          jLabel = ""
          jControl = Mid(statement, 1, 3)
          jParameters = Trim(Mid(statement, 5))
          Exit For
        Case Mid(statement, sPos, Len(OUTPUTCARD)) = OUTPUTCARD
          jLabel = RTrim(Mid(statement, 1, sPos - 1))
          jControl = Mid(statement, sPos + 1, 6)
          jParameters = Trim(Mid(statement, sPos + 8))
          Exit For
        Case Mid(statement, sPos, Len(IFCARD)) = IFCARD
          jLabel = ""
          jControl = Mid(statement, sPos, 2)
          jParameters = Trim(Mid(statement, sPos + 3))
          Exit For
        Case Mid(statement, sPos, Len(ENDIFCARD)) = ENDIFCARD
          jLabel = ""
          jControl = Trim(Mid(statement, sPos, 5))
          jParameters = ""
          Exit For

      End Select
    Next
  End Sub
  Sub ProcessJOB()
    jobSequence += 1
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
    '
    ddSequence = 0

    pgmName = Trim(GetParmPGM(jParameters))
    If pgmName.Length = 0 Then
      procSequence += 1
      procName = GetParm(jParameters, "PROC=")
      If procName.Length = 0 Then
        procName = GetFirstParm(jParameters)
      End If
      pgmName = ""
    Else
      execSequence += 1
    End If

    If pgmName.Length > 0 Then
      SourceType = GetSourceType(pgmName)
    End If

    stepName = jLabel
    'swExecFile.WriteLine(jobName & Delimiter &
    '                     LTrim(Str(jobSequence)) & Delimiter &
    '                     procName & Delimiter &
    '                     LTrim(Str(procSequence)) & Delimiter &
    '                     stepName & Delimiter &
    '                     pgmName & Delimiter &
    '                     LTrim(Str(execSequence)) & Delimiter &
    '                     RTrim(jParameters))


    ' write a PROC statement in the DD File
    'If pgmName.Length = 0 Then       'just so the DD file shows the PROC
    '  swDDFile.WriteLine(jobName & txtDelimiter.Text &
    '               LTrim(Str(jobSequence)) & txtDelimiter.Text &
    '               procName & txtDelimiter.Text &
    '               LTrim(Str(procSequence)) & txtDelimiter.Text &
    '               InstreamProc & Delimiter &
    '               stepName & txtDelimiter.Text &
    '               "" & txtDelimiter.Text &
    '               "" & txtDelimiter.Text &
    '               "<Instream PROC>" & txtDelimiter.Text &
    '               "" & txtDelimiter.Text &
    '               RTrim(jParameters))

    'End If
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
                       SourceType)
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
    swPumlFile.WriteLine("header ParseJCL(c), by IBM")
    swPumlFile.WriteLine("title Flowchart of JOB: " & FileNameOnly)

    ' Read the DD CSV file back in and load to array

    Dim FileName = txtOutputFoldername.Text & "/" & FileNameOnly & "_DD.csv"
    If Not File.Exists(FileName) Then
      Exit Sub
    End If
    '
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
      If csvCnt = 1 Then        'skip column heading row
        Continue Do
      End If
      stepName = csvRecord(4)
      pgmName = csvRecord(5)
      If pgmName.Length = 0 Then
        Continue Do
      End If
      Dim DDName As String = csvRecord(7)
      Dim DDSeq As String = csvRecord(8)
      jParameters = csvRecord(9)
      Dim dsn As String = GetParm(jParameters, "DSN=")

      Dim disp As String = GetParm(jParameters, "DISP=")
      Dim InOrOut As String = " <-left- "
      If disp.Length >= 4 Then
        If disp.Substring(0, 4) = "(NEW" Then
          InOrOut = " -right->"
        End If
      End If
      If disp.Length >= 2 Then
        If disp.Substring(0, 2) = "(," Then
          InOrOut = " -right-> "
        End If
      End If

      If Val(DDSeq) = 1 Then
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

  Sub CreatePrograms()

    ' Build the spreadsheet. First sheet is a list of all programs and DD's.

    workbook = objExcel.Workbooks.Add
    worksheet = workbook.Sheets.Item(1)
    worksheet.Name = "Programs"

    ' Write the column headings row
    worksheet.Range("A1").Value = "JobName"
    worksheet.Range("B1").Value = "JobSeq"
    worksheet.Range("C1").Value = "ProcName"
    worksheet.Range("D1").Value = "ProcSeq"
    worksheet.Range("E1").Value = "StepName"
    worksheet.Range("F1").Value = "PgmName"
    worksheet.Range("G1").Value = "PgmSeq"
    worksheet.Range("H1").Value = "DD"
    worksheet.Range("I1").Value = "DDSeq"
    worksheet.Range("J1").Value = "DDConcatSeq"
    worksheet.Range("K1").Value = "DatasetName"
    worksheet.Range("L1").Value = "StartDisp"
    worksheet.Range("M1").Value = "EndDisp"
    worksheet.Range("N1").Value = "AbendDisp"
    worksheet.Range("O1").Value = "RecFM"
    worksheet.Range("P1").Value = "LRECL"
    worksheet.Range("Q1").Value = "DBMS"
    worksheet.Range("R1").Value = "ReportId"
    worksheet.Range("S1").Value = "ReportDescription"
    worksheet.Range("T1").Value = "SourceType"

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
    Dim rowcnt As Integer = 1
    ProgressBar1.PerformStep()

    Do While Not csvFile.EndOfData
      csvRecord = csvFile.ReadFields

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
      rowcnt += 1
      row = LTrim(Str(rowcnt))
      worksheet.Range("A" & row).Value = jobName
      worksheet.Range("B" & row).Value = LTrim(Str(jobSequence))
      worksheet.Range("C" & row).Value = procName
      worksheet.Range("D" & row).Value = LTrim(Str(procSequence))
      worksheet.Range("E" & row).Value = stepName
      worksheet.Range("F" & row).Value = pgmName
      worksheet.Range("G" & row).Value = LTrim(Str(execSequence))
      worksheet.Range("H" & row).Value = DDName
      worksheet.Range("I" & row).Value = LTrim(Str(ddSequence))
      worksheet.Range("J" & row).Value = LTrim(Str(ddConcatSeq))
      worksheet.Range("K" & row).Value = dsn
      worksheet.Range("L" & row).Value = startDisp
      worksheet.Range("M" & row).Value = endDisp
      worksheet.Range("N" & row).Value = abendDisp
      worksheet.Range("O" & row).Value = dcbRecFM
      worksheet.Range("P" & row).Value = dcbLrecl
      worksheet.Range("Q" & row).Value = db2
      worksheet.Range("R" & row).Value = reportID
      worksheet.Range("S" & row).Value = reportDescription
      worksheet.Range("T" & row).Value = SourceType
      ' load up a list of programs to analyze
      If ddSequence = 1 And ddConcatSeq = 0 Then
        Select Case pgmName
          Case "IEFBR14", "SORT", "IEBGENER", "IEBCOPY", "IDCAMS", "DSNUTILB",
               "SRCHPRNT", "CMSAUTO1"
          Case Else
            If ListOfExecs.IndexOf(pgmName) = -1 Then
              ListOfExecs.Add(pgmName & Delimiter & SourceType)
            End If
        End Select
      End If
    Loop
    ' Format the Sheet - first row bold the columns
    rngRecordName = worksheet.Range("A1:T1")
    rngRecordName.Font.Bold = True
    ' data area autofit all columns
    rngRecordName = worksheet.Range("A1:T" & row)
    'rngRecordName.AutoFilter()
    workbook.Worksheets("Programs").Range("A1").AutoFilter
    rngRecordName.Columns.AutoFit()
    ' ignore error flag that numbers being loaded into a text field
    objExcel.ErrorCheckingOptions.NumberAsText = False
    '
    ProgressBar1.PerformStep()

  End Sub
  Sub ProcessSourceFiles()
    Dim SourceRecordsCount As Integer = 0
    For Each exec In ListOfExecs
      Dim execs As String() = exec.Split(Delimiter)
      exec = execs(0)
      SourceType = execs(1)
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

      ' Analyze Source Statement array (SrcStmt)
      listOfPrograms.Clear()
      listOfPrograms = GetListOfPrograms()      'list of programs within the exec source

      If pgm.ProcedureDivision = -1 Then
        LogFile.WriteLine(Date.Now & ",Source is not complete," & exec)
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
    ProgressBar1.PerformStep()

  End Sub
  Function LoadCobolStatementsToArray(ByRef CobolFile) As Integer
    '*---------------------------------------------------------
    ' Load COBOL lines to a Cobol statements array. 
    '*---------------------------------------------------------
    '
    'Assign the TempFileName for this particular cobolfile
    '
    tempCobFileName = DirectoryName & "\" & CobolFile & "_expandedCOB.txt"

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
    Dim CompilerDirective As String = ""
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
    Dim CobolFileName = txtSourceFolderName.Text & "\" & CobolFile
    If Not File.Exists(CobolFileName) Then
      LogFile.WriteLine(Date.Now & ",Source not found," & CobolFileName)
      LoadCobolStatementsToArray = -1
      Exit Function
    End If
    LogFile.WriteLine(Date.Now & ",Processing Source," & CobolFileName)

    Dim CobolLines As String() = File.ReadAllLines(CobolFileName)
    '
    ' Identify if this file is COBOL or Easytrieve
    '
    For index As Integer = 0 To CobolLines.Count - 1
      If Len(Trim(CobolLines(index))) = 0 Then
        Continue For
      End If
      If CobolLines(index).IndexOf("IDENTIFICATION DIVISION.") > -1 Then
        SourceType = "COBOL"
        Exit For
      End If
      If CobolLines(index).Length >= 5 Then
        If CobolLines(index).Substring(0, 5) = "PARM " Then
          SourceType = "Easytrieve"
          Exit For
        End If
      End If
    Next
    '
    ' Expand all copy/include members into a single file, we also drop empty lines
    '
    Do
      NumberOfScans += 1
      NumberOfCopysFound = 0
      debugCnt = 0
      swTemp = New StreamWriter(tempCobFileName, False)
      For index As Integer = 0 To CobolLines.Count - 1
        debugCnt += 1
        ' make a blank/empty line a comment line
        If Len(Trim(CobolLines(index))) = 0 Then
          swTemp.WriteLine(Space(6) & "*")
          Continue For
        End If

        Call FillInAreas(CobolLines(index),
                         SequenceNumberArea, IndicatorArea, AreaA, AreaB, CommentArea)
        ' write the comment line back out
        If IndicatorArea = "*" Then
          swTemp.WriteLine(CobolLines(index))
          Continue For
        End If

        ' get the Compiler directive, if any
        AreaAandB = AreaA & AreaB

        If AreaAandB.Trim.Length >= 5 Then
          If AreaAandB.Trim.Substring(0, 5) = "EJECT" Then
            Continue For
          End If
        End If
        If AreaAandB.Trim.Length >= 4 Then
          If AreaAandB.Trim.Substring(0, 4) = "SKIP" Then
            Continue For
          End If
        End If

        CompilerDirective = AreaAandB.ToUpper
        Dim tDirective As String() = CompilerDirective.Trim.Split(New Char() {" "c})
        ' determine if within Data Division section.
        If tDirective.Count >= 2 Then
          If tDirective(1) = "DIVISION." Then
            If tDirective(0) = "DATA" Then
              WithinDataDivision = True
            Else
              WithinDataDivision = False
            End If
          End If
        End If
        ' Checking for copy/include statement to process
        Select Case True
          Case tDirective(0) = "COPY"
          Case tDirective(0) = "++INCLUDE"
            CopybookName = Trim(tDirective(1).Replace(".", " "))
          Case tDirective(0) = "EXEC"
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
            Continue For
        End Select

        ' Expand copybooks/includes into the source
        NumberOfCopysFound += 1

        If Len(CopybookName) > 8 Then
          MessageBox.Show("somehow Member name > 8:" & CopybookName)
          LoadCobolStatementsToArray = -1
          Exit Function
        End If
        Dim CopybookFileName As String = txtSourceFolderName.Text &
                                         "\" & CopybookName
        LogFile.WriteLine(Date.Now & ",Including COBOL copybook," & CopybookFileName)
        Call IncludeCopyMember(CopybookFileName, swTemp)
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
      'If IndicatorArea = "*" Then
      ' Continue For
      'End If
      AreaAandB = AreaA & AreaB
      CompilerDirective = AreaAandB.ToUpper
      Dim tDirective As String() = CompilerDirective.Trim.Split(New Char() {" "c})
      If tDirective(0) = "REPLACE" Then
        Call ReplaceAll(AreaAandB, CobolLines, cIndex)
      End If
    Next
    '
    ' Process the WHOLE/ALL the cobol lines now that copybooks are now embedded
    ' and replace is done.
    ' This is also where we concatenate the lines, as needed, into a single statement.
    '
    Dim hlkcounter As Integer = 0
    Dim Division As String = ""
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

      CompilerDirective = DropDuplicateSpaces(AreaAandB.ToUpper)
      Dim tDirective As String() = CompilerDirective.Trim.Split(New Char() {" "c})
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
    pgm.IdentificationDivision = -1
    pgm.ProcedureDivision = -1
    pgm.EnvironmentDivision = -1
    pgm.DataDivision = -1
    pgm.ProcedureDivision = -1
    pgm.EndProgram = -1
    pgm.ProgramId = ""
    Select Case SourceType
      Case "COBOL"
        For stmtIndex As Integer = 0 To SrcStmt.Count - 1
          Select Case True
            Case SrcStmt(stmtIndex).IndexOf("IDENTIFICATION DIVISION.") > -1
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
            Case SrcStmt(stmtIndex).IndexOf("PROCEDURE DIVISION.") > -1
              pgm.ProcedureDivision = stmtIndex
            Case SrcStmt(stmtIndex).IndexOf("PROGRAM-ID.") > -1
              pgm.ProgramId = SrcStmt(stmtIndex).Substring(11).Replace(".", "").Trim
          End Select
        Next
        If Not IsNothing(pgm) Then
          pgm.EndProgram = SrcStmt.Count - 1
          listOfPrograms.Add(pgm)
        End If

      Case "Easytrieve"
        pgm.EndProgram = SrcStmt.Count - 1
        pgm.IdentificationDivision = 0
        For stmtIndex As Integer = 0 To SrcStmt.Count - 1
          Select Case True
            Case SrcStmt(stmtIndex).IndexOf("FILE") > -1 Or
                SrcStmt(stmtIndex).IndexOf("SQL") > -1
              pgm.EnvironmentDivision = stmtIndex
              pgm.DataDivision = stmtIndex
            Case SrcStmt(stmtIndex).IndexOf("JOB") > -1 Or
                 SrcStmt(stmtIndex).IndexOf("SORT") > -1
              pgm.ProcedureDivision = stmtIndex
            Case SrcStmt(stmtIndex).IndexOf("PROGRAM-ID.") > -1
              pgm.ProgramId = SrcStmt(stmtIndex).Substring(13).Trim
              Exit For
          End Select
        Next
        listOfPrograms.Add(pgm)
    End Select

    GetListOfPrograms = listOfPrograms
  End Function
  Function GetSourceType(ByRef FileName As String) As String
    ' Identify if this file is COBOL or Easytrieve
    Dim SourceFileName As String = txtSourceFolderName.Text & "\" & FileName
    GetSourceType = ""
    If Not File.Exists(SourceFileName) Then
      GetSourceType = ""
      Exit Function
    End If
    Dim CobolLines As String() = File.ReadAllLines(SourceFileName)
    For index As Integer = 0 To CobolLines.Count - 1
      If Len(Trim(CobolLines(index))) = 0 Then
        Continue For
      End If
      If CobolLines(index).IndexOf("IDENTIFICATION DIVISION.") > -1 Then
        GetSourceType = "COBOL"
        Exit Function
      End If
      If CobolLines(index).Length >= 6 Then
        Select Case CobolLines(index).Substring(0, 4)
          Case "PARM", "FILE", "SORT", "JOB "
            GetSourceType = "Easytrieve"
            Exit Function
        End Select
      End If
    Next
    LogFile.WriteLine(Date.Now & ",Unknow Type of Source File," & SourceFileName)

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
      LogFile.WriteLine(Date.Now & ",Copy Member not found," & CopyMember)
      Exit Sub
    End If
    Dim IncludeLines As String() = File.ReadAllLines(CopyMember)
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
    Dim tWord = tLine.Split(New Char() {" "c})
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
        Case "IDENTIFICATION",
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
    'Call CreatePumlEasytrieve(exec)

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

    ' Call CreateComponentsFile()

  End Function
  Function LoadEasytrieveStatementsToArray(ByRef exec As String) As Integer
    '*---------------------------------------------------------
    ' Load Easytrieve lines to a statements array. 
    '*---------------------------------------------------------
    '
    'Assign the Temporay File Name for this particular Easytrieve file
    '
    tempEZTFileName = DirectoryName & "\" & exec & "_expandedEZT.txt"

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
      LogFile.WriteLine(Date.Now & ",Source File Not Found," & FileName)
      LoadEasytrieveStatementsToArray = -1
      Exit Function
    Else
      LogFile.WriteLine(Date.Now & ",Processing Source," & FileName)
    End If

    ' put all the lines into the array
    Dim EztLinesLoaded As String() = File.ReadAllLines(FileName)

    Dim statement As String = ""
    Dim newLine As String = ""
    Dim swTemp As StreamWriter = Nothing
    swTemp = New StreamWriter(tempEZTFileName, False)
    Dim reccnt As Integer = 0

    ' process the eztlinesloaded array
    '  we will drop empty/blank lines, trim off leading spaces, and combine continued lines
    For index As Integer = 0 To EztLinesLoaded.Length - 1
      If Trim(EztLinesLoaded(index)).Length = 0 Then
        Continue For
      End If
      If Mid(EztLinesLoaded(index), 1, 1) = "*" Then
        swTemp.WriteLine(statement)
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
          swTemp.WriteLine("*" & statement & " Begin Include")
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
    SrcStmt.Add("*PROGRAM-ID. " & exec)
    reccnt = 0
    EztLinesLoaded = File.ReadAllLines(tempEZTFileName)
    For index = 0 To EztLinesLoaded.Count - 1
      SrcStmt.Add(EztLinesLoaded(index))
      reccnt += 1
    Next

    LoadEasytrieveStatementsToArray = reccnt
  End Function
  Sub CreateRecordFile()
    ' This creates the *_filename.xlxs file which will hold all things about DATA.
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
          Dim fields As New fieldInfo("", "", "", "", "DISPLAY", 0, 0, 0, -1, -1, "")
          Dim FieldSeq As Integer = 0
          For Each fields In List_Fields
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
                                  fields.Length)
          Next fields
        Next recname
      Next file
    Next pgm

    '
    ' Create the Excel Records sheet.
    '
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
    If ListOfRecords.Count > 0 Then
      For Each record In ListOfRecords
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
      Next

      ' Format the Records Sheet - first row bold the columns

      rngRecordsName = RecordsWorksheet.Range("A1:O1")
      rngRecordsName.Font.Bold = True
      ' data area autofit all columns
      rngRecordsName = RecordsWorksheet.Range("A1:O" & row)
      workbook.Worksheets("Records").Range("A1").AutoFilter
      rngRecordsName.Columns.AutoFit()
      ' ignore error flag that numbers being loaded into a text field
      objExcel.ErrorCheckingOptions.NumberAsText = False
    End If

    '
    ' Create the Fields worksheet
    '
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
      FieldsRow = 1
    End If
    '
    ' write the Fields data
    '
    row = LTrim(Str(FieldsRow))
    If ListOfFields.Count > 0 Then
      For Each FieldRow In ListOfFields
        FieldsRow += 1
        DelimText = FieldRow.Split(Delimiter)
        row = LTrim(Str(FieldsRow))
        If DelimText.Count >= 14 Then
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
        End If
      Next
      ' Format the Sheet - first row bold the columns
      rngFieldsName = FieldsWorksheet.Range("A1:N1")
      rngFieldsName.Font.Bold = True
      ' data area autofit all columns
      rngFieldsName = FieldsWorksheet.Range("A1:N" & row)
      workbook.Worksheets("Fields").Range("A1").AutoFilter
      rngFieldsName.Columns.AutoFit()
      ' ignore error flag that numbers being loaded into a text field
      objExcel.ErrorCheckingOptions.NumberAsText = False
    End If
    '
    ' how to navigate back to the first sheet '*here????
    ' how to autofilter'*here ????

  End Sub
  Sub CreatePumlCOBOL(ByRef exec As String)
    Dim EndCondIndex As Integer = -1
    Dim StartCondIndex As Integer = -1
    Dim ParagraphStarted As Boolean = False
    Dim condStatement As String = ""
    Dim condStatementCR As String = ""
    Dim imperativeStatement As String = ""
    Dim imperativeStatementCR As String = ""
    Dim statement As String = ""
    Dim vwordIndex As Integer = -1

    ' Open the output file Puml 
    Dim PumlFileName = txtOutputFoldername.Text & "\" & exec & ".puml"

    ' Open and write at least one time. Not worrying (try/catch) about subsequent writes
    Try
      pumlFile = My.Computer.FileSystem.OpenTextFileWriter(PumlFileName, False)
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error opening PumlFile")
      Exit Sub
    End Try

    ' Write the top of file
    pumlFile.WriteLine("@startuml " & FileNameOnly)
    pumlFile.WriteLine("header ParseCob(c), by IBM")
    pumlFile.WriteLine("title Flowchart of Program: " & FileNameOnly &
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

        ' break the statement in words
        Call GetSourceWords(SrcStmt(index).Trim, cWord)

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
              IFLevelIndex.Add(wordIndex)
              Call ProcessPumlIF(wordIndex)
            Case "ELSE"
              Call ProcessPumlELSE(wordIndex)
            Case "END-IF"
              IndentLevel -= 1
              pumlFile.WriteLine(Indent() & "endif")
            Case "EVALUATE"
              Call ProcessPumlEVALUATE(wordIndex)
            Case "WHEN"
              Call ProcessPumlWHEN(wordIndex)
            Case "END-EVALUATE"
              Call ProcessPumlENDEVALUATE(wordIndex)
            Case "PERFORM"
              Call ProcessPumlPERFORM(wordIndex)
            Case "END-PERFORM"
              Call ProcessPumlENDPERFORM(wordIndex)
            Case "COMPUTE"
              Call ProcessPumlCOMPUTE(wordIndex)
            Case "READ"
              Call ProcessPumlREAD(wordIndex)
            Case "AT", "END", "NOT"
              ProcessPumlReadCondition(wordIndex)
            Case "END-READ"
              ProcessPumlENDREAD(wordIndex)
            Case "GO"
              Call ProcessPumlGOTO(wordIndex)
            Case "EXEC"
              ProcessPumlEXEC(wordIndex)
            Case Else
              Dim EndIndex As Integer = 0
              Dim MiscStatement As String = ""
              Call GetStatement(wordIndex, EndIndex, MiscStatement)
              pumlFile.WriteLine(Indent() & ":" & MiscStatement.Trim & ";")
              wordIndex = EndIndex
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
            DDName = srcWords(srcWords.IndexOf("ASSIGN") + 2)
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
    Dim DataWords As New List(Of String)
    Dim EndOf01 As Integer = -1
    ' Now find the end of this 01-level
    For stmtindex As Integer = StartOf01 + 1 To pgm.ProcedureDivision
      Call GetSourceWords(SrcStmt(stmtindex), DataWords)
      Select Case DataWords(0)
        Case "FD", "01", "LINKAGE", "PROCEDURE"
          EndOf01 = stmtindex - 1
          Exit For
      End Select
    Next
    If EndOf01 = -1 Then
      MessageBox.Show("Not able to find End of 01, start of 01-Level:" & StartOf01)
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
    List_Fields.Clear()
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
      Dim fields As New fieldInfo("", "", "", "", "DISPLAY", 0, 0, 0, -1, -1, "")

      ' level clause
      fields.Level = Microsoft.VisualBasic.Right("000" & fieldWords(0), 2)
      If fieldWords(0) = "01" Then
        fields.StartPos = 1
      End If

      ' field name clause
      If fieldWords.IndexOf("REDEFINES") > -1 Then
        fields.FieldName = fieldWords(1)
        fields.Redefines = fieldWords(3)
      Else
        fields.FieldName = fieldWords(1)
        fields.Redefines = ""
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
      ' Do not care about VALUE clause
      List_Fields.Add(fields)
    Next
    ' determine the length of the 01-level (Record)
    ' Tally up all lengths except for those under a REDEFINE
    Dim withinRedefines As Boolean = False
    Dim redefinesLevel As String = "00"
    Dim totalRecordLength As Integer = 0
    For Each fields In List_Fields
      If withinRedefines And fields.Level > redefinesLevel Then
        Continue For
      End If
      withinRedefines = False
      If fields.Redefines.Length > 0 Then
        withinRedefines = True
        redefinesLevel = fields.Level
        totalRecordLength += fields.Length
        Continue For
      End If
      If fields.OccursMinimumTimes > -1 Then
        totalRecordLength += fields.Length * fields.OccursMinimumTimes
      Else
        totalRecordLength += fields.Length
      End If
    Next
    ' put final length on the 01-level. Messy, I know
    For Each fields In List_Fields
      If fields.Level = 1 Then
        fields.Length = totalRecordLength
        Exit For
      End If
    Next

    GetRecordLength = totalRecordLength
  End Function
  Function FindCopybookName(ByRef DataIndex As Integer, ByVal RecordName As String) As String
    ' Use the Data Division index to search Stmt array to get the Record  location,
    ' then look previous lines to see what the possible copybook name would be.

    Dim CopyWords As New List(Of String)
    'Dim RecordWords As New List(Of String)
    ' look upward to see if we find 'COPY/INCLUDE/SQL INCLUDE' statement
    ' here are some examples:
    '*COPY CRCALC.
    '    EXEC SQL INCLUDE SQLCA END-EXEC.
    '*INCLUDE++ PM044016
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
  Sub ProcessPumlIF(ByRef WordIndex As Integer)
    ' find the 'IF' aka Conditional statement
    ' Indentlevel is global
    Dim EndIndex As Integer = 0
    Dim Statement As String = ""
    Call GetStatement(WordIndex, EndIndex, Statement)
    pumlFile.WriteLine(Indent() & "if (" & Statement.Trim & ") then (yes)")
    IndentLevel += 1
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlELSE(ByRef WordIndex As Integer)
    ' Does current 'ELSE' belong to my 'IF' or another 'IF'???
    ' cWord is global
    ' IndentLevel is global
    'presuming a well-formed if/else/end-if
    IndentLevel -= 1
    pumlFile.WriteLine(Indent() & "else (no)")
    IndentLevel += 1

  End Sub
  Sub ProcessPumlEVALUATE(ByRef wordIndex As Integer)
    'TODO: need to fix embedded Evaluates
    ' find the end of 'EVALUATE' statement which should be at the first 'WHEN' clause
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
    pumlFile.WriteLine(Indent() & ":" & Statement.Trim & ";")
    'pumlFile.WriteLine("detach")
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlREAD(ByRef WordIndex As Integer)
    'TODO: need to fix embedded READ
    ' find the end of 'READ' statement which should be at either AT, END, NOT or END-READ
    'cWord is global
    'IndentLevel is global

    ' Format 1:Sequential Read
    WithinReadStatement = True
    Dim StartIndex = WordIndex
    Dim EndIndex As Integer = -1

    For EndIndex = WordIndex + 1 To cWord.Count - 1
      Select Case cWord(EndIndex)
        Case "AT", "END", "NOT", "END-READ"
          Exit For
      End Select
    Next
    If EndIndex > cWord.Count - 1 Then
      EndIndex = cWord.Count - 1
    End If
    EndIndex -= 1
    Dim TogetherWords As String = StringTogetherWords(StartIndex, EndIndex)
    Dim ReadStatement As String = AddNewLineAboutEvery30Characters(TogetherWords)
    pumlFile.WriteLine(Indent() & ":" & ReadStatement.Trim & "/")
    WordIndex = EndIndex
    IndentLevel += 1

    ''Format 2:random retrieval
  End Sub
  Sub ProcessPumlReadCondition(ByRef WordIndex As Integer)
    If WithinReadStatement = False Then
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
  Sub ProcessPumlPERFORM(ByRef WordIndex As Integer)
    ' find the end of 'PERFORM' statement
    ' Out-of-line Perform is when procedure-name-1 IS specified.
    ' In-line Perform is when procedure-name-1 IS NOT specified.
    ' When in-line there must be delimited with the END-PERFORM clause
    '
    Dim EndIndex As Integer = 0
    Dim Statement As String = ""
    Dim EndPerformFound As Boolean = False
    ' looking for END-PERFORM, if PERFORM is found then there is no END-PERFORM
    For EndIndex = WordIndex + 1 To cWord.Count - 1
      If cWord(EndIndex) = "PERFORM" Then
        EndPerformFound = False
        Exit For
      End If
      If cWord(EndIndex) = "END-PERFORM" Then
        EndPerformFound = True
        Exit For
      End If
    Next
    EndIndex -= 1
    Call GetStatement(WordIndex, EndIndex, Statement)
    If EndPerformFound = False Then
      pumlFile.WriteLine(Indent() & ":" & Statement.Trim & "|")
    Else
      pumlFile.WriteLine(Indent() & "while (" & Statement.Trim & ") is (true)")
      IndentLevel += 1
    End If
    WordIndex = EndIndex
  End Sub
  Sub ProcessPumlENDPERFORM(ByRef wordIndex As Integer)
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
    Dim WordsTogether As String = StringTogetherWords(WordIndex, EndIndex)
    statement = AddNewLineAboutEvery30Characters(WordsTogether)
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

        RecordingMode = "V"
        index = fdWords.IndexOf("RECORDING")
        If index > -1 Then
          Select Case True
            Case fdWords(index + 1) = "MODE" And fdWords(index + 2) = "IS"
              RecordingMode = fdWords(index + 3)
            Case fdWords(index + 1) = "MODE" Or fdWords(index + 1) = "IS"
              RecordingMode = fdWords(index + 2)
            Case Else
              RecordingMode = fdWords(index + 1)
          End Select
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
              Case fdWords(index + 1).Equals("IS") And
                   fdWords(index + 2).Equals("VARYING")
                index2 = index + 3
                index2 = GetIndexForRecordSize(index2, fdWords)
                If fdWords(index2 + 1) = "TO" Then
                  index3 = index2 + 2
                End If
                Exit For
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
                         Level & Delimiter &
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
              GetOpenMode &= "INPUT "
              Exit For
            Case "OUTPUT"
              GetOpenMode &= "OUTPUT "
              Exit For
            Case "I-O"
              GetOpenMode &= "I/O "
              Exit For
            Case "EXTEND"
              GetOpenMode &= "EXTEND "
              Exit For
            ' these cases below indicate this was not an OPEN verb
            Case "READ"
              Exit For
            Case "CLOSE"
              Exit For
            Case "SORT"
              GetOpenMode &= "SORT "
              Exit For
            Case "MERGE"
              GetOpenMode &= "MERGE "
              Exit For
            Case "USING"
              GetOpenMode &= "SORTIN"
              Exit For
            Case "GIVING"
              GetOpenMode &= "SORTOUT"
              Exit For
            Case "OPEN"
              MessageBox.Show("Never found open mode!:" &
                              file_name_1 & ":" & SrcStmt(Index))
              Exit For
            Case "PUT"
              GetOpenMode &= "WRITE"
          End Select
        Next x
      Next fnIndex
    Next Index

  End Function
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
              GetOpenModeSQL &= srcWords(cblIndex + 2) & " "
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
          GetOpenModeSQL &= srcWords(4) & " "
        End If
      End If
    Next
    GetOpenModeSQL.Trim()
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
      If FDWords.Count >= 1 Then
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
  Function AddNewLineAboutEvery30Characters(ByRef condStatement As String) As String
    ' add "\n" about every 30 characters
    Dim condStatementCR As String = ""
    Dim bytesMoved As Integer = 0
    If condStatement.Length > 30 Then
      'condStatementCR = condStatement.Substring(0, 30) & "\n"
      'StartCondIndex = 31
      For condIndex As Integer = 0 To condStatement.Length - 1
        If condStatement.Substring(condIndex, 1) = Space(1) And bytesMoved > 29 Then
          condStatementCR &= "\n"
          bytesMoved = 0
        End If
        condStatementCR &= condStatement.Substring(condIndex, 1)
        bytesMoved += 1
      Next
    Else
      condStatementCR = condStatement
    End If
    AddNewLineAboutEvery30Characters = condStatementCR
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
    IndexToNextVerb = cWord.Count - 1
  End Function

  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Me.Text = "ADDILite " & ProgramVersion
    lblCopybookMessage.Text = ""
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