Imports System.IO

Module FlowchartCOBOL
  Const ESCAPENEWLINE As String = "\n"

  'Dim Delimiter As String = "|"
  Dim newLine As String = System.Environment.NewLine


  Dim idx As Integer = 0
  Dim stmt As String = ""
  Dim IndentLevel As Integer = 0
  Dim cWord As New List(Of String)
  Dim ListOfStatements As New List(Of String)

  Dim currentParagraph As String = ""
  Dim StmtLevel As Integer = 0
  Dim VerbNames As New List(Of String)
  Dim VerbCount As New List(Of Integer)
  Dim Condition As String = ""
  Dim Imperative As String = ""
  Dim ImperativeNum As Integer = 0
  Dim WithinReadStatement As Integer = 0
  Dim ConditionalReadCnt As Integer = False
  Dim EarlyTerminate As Boolean = False

  Dim pumlFile As StreamWriter = Nothing
  Dim pumlMaxLineCnt As Integer = 1000
  Dim pumlLineCnt As Integer = 0
  Dim pumlPageCnt As Integer = 0
  Dim ProgramID As String = ""
  Dim ProgramAuthor As String = ""
  Dim ProgramWritten As String = ""
  Dim WithinIF As Integer = 0
  Dim WithinEvaluate As Integer = 0
  Dim WithinSearch As Integer = 0
  Dim WithinStart As Integer = 0
  Dim FirstWhenStatement As Boolean = False
  Dim OutputFolder As String = ""

  Public Sub CreateCobolFlowchart(logStmt As List(Of String), exec As String, outFolder As String)
    Call InitializeVariables()
    OutputFolder = outFolder
    Dim FoundProcedure As Boolean = False

    For logIndex As Integer = 0 To logStmt.Count - 1
      If EarlyTerminate Then
        Exit For
      End If
      idx = GetLogIndex(logIndex)
      stmt = GetLogStmt(logStmt(logIndex))
      ' Drop Empty lines
      If stmt.Length = 0 Then
        Continue For
      End If
      ' Drop Comments
      If stmt.Substring(0, 1) = "*" Then
        Continue For
      End If
      ' if this statment only has a period; just skip this
      If stmt.Trim = "." Then
        Continue For
      End If
      If stmt.Length >= 10 Then
        If stmt.Substring(0, 10) = "PROGRAM-ID" Then
          ProgramID = stmt.Substring(11).Replace(".", "").Trim
          Continue For
        End If
      End If
      If stmt.Length >= 8 Then
        If stmt.Substring(0, 6) = "AUTHOR" Then
          ProgramAuthor = stmt.Substring(7).Replace(".", "").Trim
          Continue For
        End If
      End If
      If stmt.Length >= 13 Then
        If stmt.Substring(0, 12) = "DATE-WRITTEN" Then
          ProgramWritten = stmt.Substring(13).Replace(".", "").Trim
          Continue For
        End If
      End If
      ' Only look for PROCEDURE DIVISION
      If Not FoundProcedure Then
        If stmt.Contains("PROCEDURE DIVISION") Then
          FoundProcedure = True
        End If
        Continue For
      End If
      ' split statement into cobol words 
      Call form1.GetSourceWords(stmt, cWord)

      ' If paragraph name, store it (maybe use later)
      If Form1.IsParagraph(cWord) Then
        currentParagraph = cWord(0)
        If cWord.Count = 2 Then     'add SECTION or EXIT
          currentParagraph &= " " & cWord(1)
        End If
        StmtLevel = 0
        Call AddToListOfStatements(currentParagraph)
        Continue For
      End If

      ' Take a COBOL Statement and format to cobol phrases with indent levels.
      StmtLevel = 1
      WithinIF = 0
      WithinReadStatement = 0
      WithinEvaluate = 0
      WithinSearch = 0
      ConditionalReadCnt = 0
      FirstWhenStatement = False
      For cWordIndex As Integer = 0 To cWord.Count - 1
        If EarlyTerminate Then
          Exit For
        End If
        Select Case cWord(cWordIndex)
          Case "IF"
            WithinIF += 1
            cWordIndex = Process_IF(cWordIndex)
            StmtLevel += 1
          Case "ELSE"
            StmtLevel -= 1
            cWordIndex = Process_ELSE(cWordIndex)
            StmtLevel += 1
          Case "END-IF"
            StmtLevel -= 1
            cWordIndex = Process_ENDIF(cWordIndex)
            WithinIF -= 1
          Case "READ", "RETURN"
            WithinReadStatement += 1
            cWordIndex = Process_Read(cWordIndex)
          Case "SEARCH"
            WithinSearch += 1
            FirstWhenStatement = True
            cWordIndex = Process_Search(cWordIndex)
          Case "START"
            cWordIndex = Process_Start(cWordIndex)
          Case "END-READ"
            WithinReadStatement -= 1
            cWordIndex = Process_EndRead(cWordIndex)
          Case "EVALUATE"
            WithinEvaluate += 1
            FirstWhenStatement = True
            cWordIndex = Process_Imperative(cWordIndex)
            StmtLevel += 1
          Case "WHEN"
            If FirstWhenStatement Then
              FirstWhenStatement = False
            Else
              StmtLevel -= 1
            End If
            cWordIndex = Process_WHEN(cWordIndex)
            If WithinSearch > 0 Then
              WithinIF += 1
            End If
            StmtLevel += 1
          Case "END-EVALUATE"
            StmtLevel -= 2
            cWordIndex = Process_ENDEVALUATE(cWordIndex)
            WithinEvaluate -= 1
          Case "END-SEARCH"
            StmtLevel -= 1
            cWordIndex = Process_ENDIF(cWordIndex)
            WithinIF -= 1
            WithinSearch -= 1
          Case "CONTINUE"
            cWordIndex = Process_CONTINUE(cWordIndex)
          Case "EXEC"
            cWordIndex = Process_EXEC(cWordIndex)
          Case ""
            'skip empty word (can happen if period symbol is preceeded by a space)
          Case Else
            cWordIndex = Process_Imperative(cWordIndex)
        End Select
      Next

      'all verbs processed from the statement
      'write any pending/implied end-* statements
      For y As Integer = WithinIF To 1 Step -1
        For x As Integer = StmtLevel - 1 To 1 Step -1
          StmtLevel -= 1
          Call AddToListOfStatements("END-IF")
          If y = 1 Then
            Exit For
          End If
        Next x
      Next y
      If WithinEvaluate > 0 Then
        For x As Integer = WithinEvaluate To 1 Step -1
          StmtLevel = x
          Call AddToListOfStatements("END-EVALUATE")
        Next
      End If
      If WithinSearch > 0 Then
        For x As Integer = WithinSearch To 1 Step -1
          StmtLevel = x
          Call AddToListOfStatements("END-SEARCH")
        Next
      End If
    Next

    ' next step is to print out the listOfStatements Array to PUML formatted statements
    ' note the listOfStatements has 3 parts:
    '  (0)-Index to the log file
    '  (1)-statement level (indentation). 
    '    value 0 - Paragraph name
    '    value -1- Early termination indicator
    '    values 1-N
    '  (2)-Statement (or error message, if level=-1
    Dim ParagraphStarted As Boolean = False
    'Dim fileout As String = WorkFolder & "\" & txtMember.Text & "_statements.txt"
    'pumlFile = My.Computer.FileSystem.OpenTextFileWriter(fileout, False)

    pumlLineCnt = pumlMaxLineCnt + 1
    pumlPageCnt = 0
    PumlPageBreak(exec)

    For Each Statement In ListOfStatements
      Dim parts As String() = Statement.Split(Form1.Delimiter)
      Select Case Val(parts(1))
        Case > 0        'statements
          IndentLevel = Val(parts(1)) * 2
          stmt = parts(2)
          Call WritePumlStatement()
        Case 0          'paragraph name
          Call ProcessPumlParagraph(ParagraphStarted, parts(2), exec)
        Case < 0        'early terminate
          'pumlFile.WriteLine(parts(0) & " " & parts(1) & " ** " & parts(2))
          pumlLineCnt += 2
          pumlFile.WriteLine("end")
          pumlFile.WriteLine("note Right: " & parts(2))
          Exit For
      End Select
    Next

    If ParagraphStarted = True And Not EarlyTerminate Then
      pumlLineCnt += 2
      pumlFile.WriteLine("end")
      pumlFile.WriteLine("}")
      ParagraphStarted = False
    End If

    pumlLineCnt += 1
    pumlFile.WriteLine("@enduml")
    pumlFile.Close()

  End Sub
  Sub WritePumlStatement()
    If stmt.Length = 0 Then
      MessageBox.Show("stmt.length is 0")
      Exit Sub
    End If
    Call Form1.GetSourceWords(stmt, cWord)
    Select Case cWord(0)
      Case "IF"
        pumlLineCnt += 1
        pumlFile.WriteLine(Indent() & "#lightblue:if (" & stmt.Trim.Replace("IF ", "").Replace(" THEN", "") & "?) then (yes)")
      Case "ELSE"
        pumlLineCnt += 1
        pumlFile.WriteLine(Indent() & "else (no)")
      Case "END-IF", "END-EVALUATE"
        pumlLineCnt += 1
        pumlFile.WriteLine(Indent() & "endif")
      Case "EVALUATE", "SEARCH"
        pumlLineCnt += 1
        pumlFile.WriteLine(Indent() & ":" & stmt.Trim & ";")
        FirstWhenStatement = True
      Case "WHEN"
        If FirstWhenStatement Then
          FirstWhenStatement = False
          Condition = "if"
        Else
          Condition = "elseif"
        End If
        pumlLineCnt += 1
        pumlFile.WriteLine(Indent() & "#lightblue:" & Condition & " (" & stmt.Replace("WHEN", "").Trim & "?) then (yes)")
      Case "GO"
        pumlLineCnt += 2
        pumlFile.WriteLine(Indent() & "#lightgreen:" & stmt.Trim & ";")
        pumlFile.WriteLine(Indent() & "stop")
      Case "READ", "RETURN", "RELEASE", "DISPLAY", "ACCEPT", "WRITE", "REWRITE", "UPDATE", "DELETE", "SORT", "START",
           "OPEN", "CLOSE"
        pumlLineCnt += 1
        pumlFile.WriteLine(Indent() & ":" & stmt.Trim & "/")
      Case "PERFORM"
        pumlLineCnt += 1
        pumlFile.WriteLine(Indent() & ":" & stmt.Trim & "|")
      Case "CALL"
        pumlLineCnt += 1
        pumlFile.WriteLine(Indent() & ":" & stmt.Trim & ">")
      Case "STOP", "GOBACK"
        pumlLineCnt += 1
        pumlFile.WriteLine(Indent() & "#lightCyan:" & stmt.Trim & ";")
      Case "END-SEARCH"

      Case Else
        Dim myStmt As String = Form1.AddNewLineAboutEveryNthCharacters(stmt.Trim, ESCAPENEWLINE, 45)
        pumlLineCnt += 1
        pumlFile.WriteLine(Indent() & ":" & myStmt & "]")

    End Select
  End Sub
  Function Process_IF(ByRef cWordIndex As Integer) As Integer
    ' gather the IF statement words and then add to the List of Statements array
    ' the ending Index is returned
    Dim EndIndex As Integer = GetEndIndex(cWordIndex, cWord.Count)
    GetStatement(cWordIndex, EndIndex, Condition)
    AddToListOfStatements(Condition)
    Return EndIndex
  End Function
  Function Process_ELSE(ByRef cWordIndex As Integer) As Integer
    'check if the ELSE is correct by searching backwards looking for same stmtLevel with an "IF"
    For x As Integer = ListOfStatements.Count - 1 To 0 Step -1
      Dim stuff As String() = ListOfStatements(x).Split(Form1.Delimiter)
      If Val(stuff(1)) = StmtLevel Then
        Dim words As String() = stuff(2).Split(" ")
        If words(0) <> "IF" Then
          StmtLevel = -1
          EarlyTerminate = True
          Call AddToListOfStatements("ELSE has no IF" & ESCAPENEWLINE &
                                     " Nested Level:" & stuff(1) & ESCAPENEWLINE &
                                     " Statement:" & LTrim(Str(idx)) & ESCAPENEWLINE &
                                     " Verb:" & LTrim(Str(cWordIndex)))
          Return cWordIndex
        Else
          Exit For
        End If
      End If
    Next x

    Dim elseCondition As String = "ELSE"
    Call AddToListOfStatements(elseCondition)
    Return cWordIndex
  End Function
  Function Process_ENDIF(ByRef cWordIndex As Integer) As Integer
    'Dim endIfCondition As String = cWord(cWordIndex)
    Call AddToListOfStatements("END-IF")
    Return cWordIndex
  End Function

  Function Process_ENDEVALUATE(ByRef cWordIndex As Integer) As Integer
    'Dim endIfCondition As String = cWord(cWordIndex)
    Call AddToListOfStatements("END-EVALUATE")
    Return cWordIndex
  End Function


  Function Process_WHEN(ByRef cWordIndex As Integer) As Integer
    ' gather the WHEN (aka IF) statement words and then add to the List of Statements array
    ' the ending Index is returned
    Dim EndIndex As Integer = GetEndIndex(cWordIndex, cWord.Count)
    GetStatement(cWordIndex, EndIndex, Condition)
    AddToListOfStatements(Condition)
    Return EndIndex
  End Function

  Function Process_Read(ByRef cWordIndex As Integer) As Integer
    Dim EndIndex As Integer = 0
    WithinReadStatement += 1
    ' Check if this is a Conditional Read statement
    ' loop till it is neither: "AT", "END', or "NOT"
    Dim ConditionalRead As Boolean = False
    Dim EndOfReadConditionIndex As Integer = -1
    Dim StartOfReadConditionIndex As Integer = cWordIndex + 2
    Dim ReadConditionalStatement As String = ""
    Dim NumberOfConditionals As Integer = 0
    For x As Integer = StartOfReadConditionIndex To cWord.Count - 1
      Select Case cWord(x)
        Case "AT", "END", "NOT"
          ConditionalRead = True
          EndOfReadConditionIndex = x
          ReadConditionalStatement &= cWord(x) & " "
          NumberOfConditionals += 1
        Case "NEXT", "RECORD", "INTO"
          'x += 1
          EndIndex = x
        Case "END-READ", "END-RETURN"
          Exit For
        Case Else
          Exit For
      End Select
    Next
    If ConditionalRead Then
      ConditionalReadCnt += 1
    End If
    ' write out the "Read <file-name>" statement
    'Imperative = cWord(cWordIndex) & " " & cWord(cWordIndex + 1)
    If EndOfReadConditionIndex = -1 Then
      EndOfReadConditionIndex = EndIndex
    End If
    If (EndOfReadConditionIndex - NumberOfConditionals < 0) Or
      (EndOfReadConditionIndex - NumberOfConditionals > cWord.Count) Then
      MessageBox.Show("indexes for stringtogether is wrong")
    End If
    If (EndOfReadConditionIndex - NumberOfConditionals) = 0 Then
      GetStatement(cWordIndex, EndIndex, Imperative)    'the whole read statement there is no conditions
    Else
      Imperative = Form1.StringTogetherWords(cWord, cWordIndex, EndOfReadConditionIndex - NumberOfConditionals)
    End If
    If Imperative.Length = 0 Then
      'MessageBox.Show("Read imperative length is 0")
      Imperative = "NOOP READ"
    End If
    AddToListOfStatements(Imperative)
    ' write out the "AT END" or "NOT AT END" 
    If ConditionalRead Then
      AddToListOfStatements("IF " & ReadConditionalStatement.Trim & " THEN")
      EndIndex = EndOfReadConditionIndex
      StmtLevel += 1
      WithinIF += 1
    End If
    Return EndIndex
  End Function
  Function Process_EndRead(ByRef cWordIndex As Integer) As Integer
    If ConditionalReadCnt > 0 Then
      StmtLevel -= 1
      AddToListOfStatements("END-IF")
      ConditionalReadCnt -= 1
      WithinIF -= 1
    Else
      AddToListOfStatements("END-READ")
    End If
    WithinReadStatement -= 1
    Return cWordIndex
  End Function

  Function Process_Search(ByRef cWordIndex As Integer) As Integer
    ' Search has two sections: a found (WHEN) and an Optional NOT found condition (AT END).
    ' this could be terminated with an End-Search or end of statement marker('.')

    ' find the WHEN index
    Dim WhenIndex As Integer = cWord.IndexOf("WHEN", cWordIndex + 1)

    ' write out the "Search <table-name> at end imperatives" statement (first section)
    Imperative = Form1.StringTogetherWords(cWord, cWordIndex, WhenIndex - 1)
    Imperative = Form1.AddNewLineAboutEveryNthCharacters(Imperative, ESCAPENEWLINE, 30)
    AddToListOfStatements(Imperative)

    Return WhenIndex - 1
  End Function

  Function Process_Start(ByRef cWordIndex As Integer) As Integer
    WithinStart += 1
    ' START does not always have an INVLAID option
    Dim EndIndex As Integer = cWord.IndexOf("END-START", cWordIndex + 1)
    If EndIndex = -1 Then
      EndIndex = cWord.Count - 1
    End If
    Dim InvalidIndex As Integer = cWord.IndexOf("INVALID", cWordIndex + 1)
    If InvalidIndex = -1 Then
      Imperative = Form1.StringTogetherWords(cWord, cWordIndex, EndIndex)
      AddToListOfStatements(Imperative)
      Return EndIndex
    End If
    ' Find the beginning and ending indexes of the INVALID condition
    Dim EndInvalidIndex As Integer = 0
    If cWord(InvalidIndex + 1) = "KEY" Then
      EndInvalidIndex = InvalidIndex + 1
    End If
    If cWord(InvalidIndex - 1) = "NOT" Then
      InvalidIndex -= 1
    End If
    ' Write the START command w/o the INVALID condition.
    Imperative = Form1.StringTogetherWords(cWord, cWordIndex, InvalidIndex - 1)
    AddToListOfStatements(Imperative)

    ' write the INVALID condition
    Condition = Form1.StringTogetherWords(cWord, InvalidIndex, EndInvalidIndex)
    AddToListOfStatements("IF " & Condition & " THEN")
    WithinIF += 1
    StmtLevel += 1
    Return EndInvalidIndex
  End Function

  Function Process_Imperative(ByRef cWordIndex As Integer) As Integer
    ' gather the statement words and then add to the List of Statements array
    ' the ending Index is returned
    Dim EndIndex As Integer = GetEndIndex(cWordIndex, cWord.Count)
    GetStatement(cWordIndex, EndIndex, Imperative)
    AddToListOfStatements(Imperative)
    Return EndIndex
  End Function

  Function Process_CONTINUE(ByRef cWordIndex As Integer) As Integer
    ' Continue verb is just one word
    Imperative = cWord(cWordIndex)
    AddToListOfStatements(Imperative)
    Return cWordIndex
  End Function

  Function Process_EXEC(ByRef cWordIndex As Integer) As Integer
    ' Grab all from EXEC to END-EXEC
    Dim EndIndex As Integer = cWord.IndexOf("END-EXEC", cWordIndex + 1)
    If EndIndex = -1 Then
      EndIndex = cWord.Count - 1
    End If
    Imperative = Form1.StringTogetherWords(cWord, cWordIndex, EndIndex)
    AddToListOfStatements(Imperative)
    Return EndIndex
  End Function


  Sub AddToListOfStatements(ByRef Statement As String)
    Dim idxStatement As String = LTrim(Str(idx))
    If Statement.Length = 0 Then
      MessageBox.Show("addtolistofstatement statement length is 0")
    End If
    ListOfStatements.Add(idxStatement & Form1.Delimiter & StmtLevel & Form1.Delimiter & Statement)
    If ListOfStatements.Count > 40000 Then
      MessageBox.Show("Over 40,000K ListOfStatements?")
    End If
  End Sub
  Function GetEndIndex(ByRef Wordindex As Integer, ByRef count As Integer) As Integer
    ' determine the end Index to the next verb
    ' adjust if no ending verb to last word as the index
    Dim EndIndex = Form1.IndexToNextVerb(cWord, Wordindex)
    If EndIndex = -1 Then
      EndIndex = count - 1
    End If
    Return EndIndex
  End Function
  Function GetLogIndex(indx As String) As Integer
    Return indx
  End Function
  Function GetLogStmt(logStatement As String) As String
    Return logStatement
  End Function

  ' routines from addilite

  Sub GetStatement(ByRef WordIndex As Integer, ByRef EndIndex As Integer, ByRef statement As String)
    ' get the whole COBOL statement of this verb by looking for the next verb
    'Dim StartIndex As Integer = WordIndex
    EndIndex = Form1.IndexToNextVerb(cWord, WordIndex)
    If EndIndex = -1 Then
      EndIndex = cWord.Count - 1
    End If
    Dim WordsTogether As String = Form1.StringTogetherWords(cWord, WordIndex, EndIndex)
    statement = Form1.AddNewLineAboutEveryNthCharacters(WordsTogether, ESCAPENEWLINE, 30)
  End Sub

  'Function StringTogetherWords(ByRef StartCondIndex As Integer, ByRef EndCondIndex As Integer) As String
  '  ' string together from startofconditionindex to endofconditionindex
  '  ' cWord is a global variable
  '  Dim wordsStrungTogether As String = ""
  '  Try
  '    For condIndex As Integer = StartCondIndex To EndCondIndex
  '      wordsStrungTogether &= cWord(condIndex) & " "
  '    Next

  '  Catch ex As Exception
  '    MessageBox.Show("error:" & ex.Message)
  '  End Try
  '  StringTogetherWords = wordsStrungTogether.TrimEnd
  'End Function
  'Function IndexToNextVerb(ByRef StartCondIndex As Integer) As Integer
  '  ' cWord is a global variable
  '  ' VerNames is a global variable
  '  ' find ending index to next COBOL verb in cWord
  '  Dim EndCondIndex As Integer = -1
  '  Dim VerbIndex As Integer = -1
  '  For EndCondIndex = StartCondIndex + 1 To cWord.Count - 1
  '    If WithinReadStatement > 0 Then
  '      Select Case cWord(EndCondIndex)
  '        Case "AT", "END", "NOT"
  '          Return EndCondIndex
  '        Case "NEXT"
  '          Continue For
  '      End Select
  '    End If
  '    VerbIndex = VerbNames.IndexOf(cWord(EndCondIndex))
  '    If VerbIndex > -1 Then
  '      Return EndCondIndex - 1
  '    End If
  '  Next
  '  ' there is not another verb in this statement
  '  Return -1
  'End Function

  'Function AddNewLineAboutEveryNthCharacters(ByRef condStatement As String,
  '                                          ByRef theNewLine As String,
  '                                          ByVal Size As Integer) As String
  '  ' add "\n" or vbnewline (theNewLine) about every SIZE number of characters
  '  Dim condStatementCR As String = ""
  '  Dim bytesMoved As Integer = 0
  '  If condStatement.Length = 0 Then
  '    Return ""
  '  End If
  '  If condStatement.Length > Size Then
  '    For condIndex As Integer = 0 To condStatement.Length - 1
  '      If condStatement.Substring(condIndex, 1) = Space(1) And bytesMoved > (Size - 1) Then
  '        condStatementCR &= theNewLine
  '        bytesMoved = 0
  '      End If
  '      condStatementCR &= condStatement.Substring(condIndex, 1)
  '      bytesMoved += 1
  '    Next
  '  Else
  '    condStatementCR = condStatement
  '  End If
  '  Return condStatementCR
  'End Function

  'Sub GetSourceWords(ByVal statement As String, ByRef srcWords As List(Of String))
  '  ' takes the stmt and breaks into words and drops blanks
  '  srcWords.Clear()
  '  statement = statement.Trim
  '  Dim WithinQuotes As Boolean = False
  '  Dim word As String = ""
  '  Dim aByte As String = ""
  '  For x As Integer = 0 To statement.Length - 1
  '    aByte = statement.Substring(x, 1)
  '    If aByte = "'" Then
  '      WithinQuotes = Not WithinQuotes
  '    End If
  '    If aByte = " " Then
  '      If WithinQuotes Then
  '        word &= aByte
  '      Else
  '        If word.Trim.Length > 0 Then
  '          srcWords.Add(word.ToUpper)
  '          word = ""
  '        End If
  '      End If
  '    Else
  '      word &= aByte
  '    End If
  '  Next
  '  If word.EndsWith(".") Then
  '    word = word.Remove(word.Length - 1)
  '    srcWords.Add(word.ToUpper)
  '    word = ""
  '  End If
  '  If word.Length > 0 Then
  '    srcWords.Add(word)
  '  End If
  'End Sub
  Sub ProcessPumlParagraph(ByRef ParagraphStarted As Boolean, ByRef statement As String, ByRef exec As String)
    If ParagraphStarted = True Then
      pumlLineCnt += 3
      pumlFile.WriteLine("end")
      pumlFile.WriteLine("}")
      pumlFile.WriteLine("")
    End If

    If pumlLineCnt > pumlMaxLineCnt Then
      pumlLineCnt += 1
      pumlFile.WriteLine("floating note left: Continued in Part " & pumlPageCnt + 1)
      pumlFile.WriteLine("@enduml")
      pumlFile.Close()
      PumlPageBreak(exec)
    End If

    pumlFile.WriteLine("partition **" & Trim(statement.Replace(".", "")) & "** {")
    pumlFile.WriteLine("start")
    'pumlFile.WriteLine("#yellow:**" & Trim(statement.Replace(".", "")) & "**;")
    pumlLineCnt += 2
    ParagraphStarted = True
  End Sub

  Sub PumlPageBreak(ByRef exec As String)
    pumlPageCnt += 1
    ' Open the output file Puml 
    Dim pumlFileName As String = OutputFolder & "\" & exec & ".puml"
    If pumlPageCnt > 1 Then
      pumlFileName = OutputFolder & "\" & exec & "_" & LTrim(Str(pumlPageCnt)) & ".puml"
    End If

    ' Open and write at least one time. Not worrying (try/catch) about subsequent writes
    Try
      pumlFile = My.Computer.FileSystem.OpenTextFileWriter(pumlFileName, False)
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error opening PumlFile COBOL")
      Exit Sub
    End Try

    ' Write the top of file
    If pumlPageCnt > 1 Then
      pumlFile.WriteLine("@startuml " & exec & "_" & LTrim(Str(pumlPageCnt)))
    Else
      pumlFile.WriteLine("@startuml " & exec)
    End If
    pumlFile.WriteLine("header ADDILite(c), by IBM")
    pumlFile.Write("title Flowchart of COBOL Program: " & exec &
                       "\nProgram Author: " & ProgramAuthor &
                       "\nDate written: " & ProgramWritten)
    If pumlPageCnt > 1 Then
      pumlFile.WriteLine("\nPart: " & pumlPageCnt)
    Else
      pumlFile.WriteLine("")
    End If
    If EarlyTerminate Then
      pumlFile.WriteLine("note right #FFAAAA: LOGIC ERROR IN CODE")
    End If
    pumlLineCnt = 3
    WithinIF -= 1
  End Sub
  Function Indent() As String
    If IndentLevel > 0 Then
      Return Space(IndentLevel * 2)
    End If
    Indent = ""
  End Function
  Public Sub InitializeVariables()
    currentParagraph = ""
    StmtLevel = 0
    Condition = ""
    Imperative = ""
    ImperativeNum = 0
    WithinReadStatement = 0
    ConditionalReadCnt = False
    EarlyTerminate = False
    pumlMaxLineCnt = 1000
    pumlLineCnt = 0
    pumlPageCnt = 0
    ProgramID = ""
    ProgramAuthor = ""
    ProgramWritten = ""
    WithinIF = 0
    WithinEvaluate = 0
    WithinSearch = 0
    WithinStart = 0
    FirstWhenStatement = 0
    OutputFolder = ""

    ListOfStatements.Clear()

    ' This area is the COBOL Verb array with counts. 
    ' **BE SURE TO KEEP VerbNames AND VerbCount ARRAYS IN SYNC!!!**
    ' Flow commands
    'VerbNames.Add("GO")
    'VerbNames.Add("ALTER")
    'VerbNames.Add("CALL")
    'VerbNames.Add("PERFORM")
    'VerbNames.Add("EVALUATE")
    'VerbNames.Add("WHEN")
    'VerbNames.Add("CONTINUE")
    'VerbNames.Add("IF")
    'VerbNames.Add("ELSE")
    'VerbNames.Add("GOBACK")
    'VerbNames.Add("STOP")
    'VerbNames.Add("CHAIN")
    '' I/O
    'VerbNames.Add("OPEN")
    'VerbNames.Add("READ")
    'VerbNames.Add("WRITE")
    'VerbNames.Add("REWRITE")
    'VerbNames.Add("CLOSE")
    'VerbNames.Add("EXEC")
    'VerbNames.Add("COMMIT")
    'VerbNames.Add("CANCEL")
    'VerbNames.Add("DELETE")
    'VerbNames.Add("MERGE")
    'VerbNames.Add("SORT")
    'VerbNames.Add("RETURN")
    'VerbNames.Add("NEXT")
    '' Maths
    'VerbNames.Add("COMPUTE")
    'VerbNames.Add("ADD")
    'VerbNames.Add("SUBTRACT")
    'VerbNames.Add("MULTIPLY")
    'VerbNames.Add("DIVIDE")
    '' Misc
    'VerbNames.Add("MOVE")
    'VerbNames.Add("DISABLE")
    'VerbNames.Add("DISPLAY")
    'VerbNames.Add("ENABLE")
    'VerbNames.Add("END-READ")
    'VerbNames.Add("END-EVALUATE")
    'VerbNames.Add("END-IF")
    'VerbNames.Add("END-INVOKE")
    'VerbNames.Add("END-PERFORM")
    'VerbNames.Add("END-SET")
    'VerbNames.Add("ENTER")
    'VerbNames.Add("ENTRY")
    'VerbNames.Add("EXAMINE")
    'VerbNames.Add("EXECUTE")
    'VerbNames.Add("EXHIBIT")
    'VerbNames.Add("EXIT")
    'VerbNames.Add("GENERATE")
    'VerbNames.Add("INITIALIZE")
    'VerbNames.Add("INITIATE")
    'VerbNames.Add("INSPECT")
    'VerbNames.Add("INVOKE")
    'VerbNames.Add("NOTE")
    'VerbNames.Add("OTHERWISE")
    'VerbNames.Add("READY")
    'VerbNames.Add("RECEIVE")
    'VerbNames.Add("RECOVER")
    'VerbNames.Add("RELEASE")
    'VerbNames.Add("RESET")
    'VerbNames.Add("ROLLBACK")
    'VerbNames.Add("SEARCH")
    'VerbNames.Add("SEND")
    'VerbNames.Add("SERVICE")
    'VerbNames.Add("SET")
    'VerbNames.Add("START")
    'VerbNames.Add("STRING")
    'VerbNames.Add("SUPPRESS")
    'VerbNames.Add("TERMINATE")
    'VerbNames.Add("TRANSFORM")
    'VerbNames.Add("UNLOCK")
    'VerbNames.Add("UNSTRING")

    '' Flow commands
    'VerbCount.Add(0)    'GO
    'VerbCount.Add(0)    'ALTER
    'VerbCount.Add(0)    'CALL
    'VerbCount.Add(0)    'PERFORM
    'VerbCount.Add(0)    'EVALUATE
    'VerbCount.Add(0)    'WHEN
    'VerbCount.Add(0)    'CONTINUE
    'VerbCount.Add(0)    'IF
    'VerbCount.Add(0)    'ELSE
    'VerbCount.Add(0)    'GOBACK
    'VerbCount.Add(0)    'STOP
    'VerbCount.Add(0)    'CHAIN
    '' I/O
    'VerbCount.Add(0)    'OPEN
    'VerbCount.Add(0)    'READ
    'VerbCount.Add(0)    'WRITE
    'VerbCount.Add(0)    'REWRITE
    'VerbCount.Add(0)    'CLOSE
    'VerbCount.Add(0)    'EXEC
    'VerbCount.Add(0)    'COMMIT
    'VerbCount.Add(0)    'CANCEL
    'VerbCount.Add(0)    'DELETE
    'VerbCount.Add(0)    'MERGE
    'VerbCount.Add(0)    'SORT
    'VerbCount.Add(0)    'RETURN
    'VerbCount.Add(0)    'NEXT
    '' Maths
    'VerbCount.Add(0)    'COMPUTE
    'VerbCount.Add(0)    'ADD
    'VerbCount.Add(0)    'SUBTRACT
    'VerbCount.Add(0)    'MULTIPLY
    'VerbCount.Add(0)    'DIVIDE
    '' Misc
    'VerbCount.Add(0)    'MOVE
    'VerbCount.Add(0)    'DISABLE
    'VerbCount.Add(0)    'DISPLAY
    'VerbCount.Add(0)    'ENABLE
    'VerbCount.Add(0)    'END-READ
    'VerbCount.Add(0)    'END-EVALUATE
    'VerbCount.Add(0)    'END-IF
    'VerbCount.Add(0)    'END-INVOKE
    'VerbCount.Add(0)    'END-PERFORM
    'VerbCount.Add(0)    'END-SET
    'VerbCount.Add(0)    'ENTER
    'VerbCount.Add(0)    'ENTRY
    'VerbCount.Add(0)    'EXAMINE
    'VerbCount.Add(0)    'EXECUTE
    'VerbCount.Add(0)    'EXHIBIT
    'VerbCount.Add(0)    'EXIT
    'VerbCount.Add(0)    'GENERATE
    'VerbCount.Add(0)    'INITIALIZE
    'VerbCount.Add(0)    'INITIATE
    'VerbCount.Add(0)    'INSPECT
    'VerbCount.Add(0)    'INVOKE
    'VerbCount.Add(0)    'NOTE
    'VerbCount.Add(0)    'OTHERWISE
    'VerbCount.Add(0)    'READY
    'VerbCount.Add(0)    'RECEIVE
    'VerbCount.Add(0)    'RECOVER
    'VerbCount.Add(0)    'RELEASE
    'VerbCount.Add(0)    'RESET
    'VerbCount.Add(0)    'ROLLBACK
    'VerbCount.Add(0)    'SEARCH
    'VerbCount.Add(0)    'SEND
    'VerbCount.Add(0)    'SERVICE
    'VerbCount.Add(0)    'SET
    'VerbCount.Add(0)    'START
    'VerbCount.Add(0)    'STRING
    'VerbCount.Add(0)    'SUPPRESS
    'VerbCount.Add(0)    'TERMINATE
    'VerbCount.Add(0)    'TRANSFORM
    'VerbCount.Add(0)    'UNLOCK
    'VerbCount.Add(0)    'UNSTRING

  End Sub
End Module
