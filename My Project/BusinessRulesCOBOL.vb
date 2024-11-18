Imports System.IO
Imports ADDILite.Form1
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Module BusinessRulesCOBOL
  'BusinessRulesCobol.
  'Create a spreadsheet of the COBOL's Business Rules (BR). 
  ' An 'IF' verb begins is a Rule statements
  ' All other verbs are Business statements
  ' ASSUMPTION!!! A COBOL statment will never have an 'IF' within the statement without it STARTING with an 'IF'
  '   So if we encounter an 'IF' at first word of a statement this begins a Business Rule
  'Written by Howard Kearney
  'Change-history.
  '  2024-09-24 v1.6.2 remove equal sign in value of spreadsheet
  '  2024-08-21 v1.6.1 Clear Arrays
  '  2024-07-03 hk New code
  '  2024-07-04 hk removed blank lines from business rules

  Dim BRobjExcel As New Microsoft.Office.Interop.Excel.Application
  Dim BRWorkbook As Microsoft.Office.Interop.Excel.Workbook
  Dim BusinessRulesWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim BusinessFieldsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim rngBusinessRules As Microsoft.Office.Interop.Excel.Range
  Dim rngBusinessFields As Microsoft.Office.Interop.Excel.Range
  Dim DefaultFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault
  Dim SetAsReadOnly = Microsoft.Office.Interop.Excel.XlFileAccess.xlReadOnly

  Dim pumlFile As StreamWriter = Nothing


  Dim ListOfRules As New List(Of String)
  Dim ListOfParagraphs As New List(Of String)
  Dim ListOfFieldNames As New List(Of String)
  Dim ListOfBusinessFieldNames As New List(Of String)
  Dim ParagraphReferenceList As New List(Of String)

  Dim idx As Integer = 0
  Dim stmt As String = ""
  Dim cWord As New List(Of String)
  Dim currentParagraph As String = "root"
  Dim RuleNum As Integer = 0
  Dim subRuleNum As Integer = 0
  Dim Rule As String = ""
  Dim Business As String = ""
  Dim newLine As String = System.Environment.NewLine
  Dim delimiter As String = "|"

  Public Sub CreateCOBOLBusinessRules(ByRef srcStmt As List(Of String), ByRef exec As String,
                                      ByRef outFolder As String,
                                      ByRef outPumlFolder As String,
                                      ByRef pgm As ProgramInfo,
                                      ByRef ListOfFields As List(Of String))
    ' Initialize Arrays
    ListOfRules.Clear()
    ListOfParagraphs.Clear()
    ListOfFieldNames.Clear()
    ListOfBusinessFieldNames.Clear()
    ParagraphReferenceList.Clear()

    ' get just the FieldNames (this is used for Rules type determination)
    For Each entry In ListOfFields
      Dim fieldAttribs As String() = entry.Split(Form1.Delimiter)
      ListOfFieldNames.Add(fieldAttribs(9))
    Next

    ' grab all the paragraph names
    For srcIndex As Integer = pgm.ProcedureDivision + 1 To pgm.EndProgram
      Call Form1.GetSourceWords(srcStmt(srcIndex), cWord)
      If Form1.IsParagraph(cWord) Then
        ListOfParagraphs.Add(cWord(0))
      End If
    Next

    ' process the source
    Dim FoundProcedure As Boolean = False
    For srcIndex As Integer = pgm.ProcedureDivision + 1 To pgm.EndProgram
      idx = GetSrcIndex(srcIndex)
      stmt = GetSrcStmt(srcStmt(srcIndex))
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
      ' split statement into cobol words 
      Call Form1.GetSourceWords(stmt, cWord)

      ' If paragraph name, store it (used in P2P Puml)
      If Form1.IsParagraph(cWord) Then
        currentParagraph = cWord(0)
        Continue For
      End If

      Rule = ""
      Business = ""
      Form1.WithinReadStatement = False
      subRuleNum = 0

      ' main loop to process a COBOL sentence
      For cWordIndex As Integer = 0 To cWord.Count - 1
        If cWordIndex = 0 And cWord(0) = "IF" Then
          'this is the begining of a rule
          RuleNum += 1
          subRuleNum = 0
        End If
        If cWord(0) = "IF" Then
          Select Case cWord(cWordIndex)
            Case "IF"
              Call ProcessBR_IF(cWordIndex)
            Case "ELSE"
              Call ProcessBR_ELSE()
            Case "END-IF"
              Call ProcessBR_ENDIF()
            Case Else
              If cWord(0) = "READ" Then
                Form1.WithinReadStatement = True
              End If
              If cWord(0) = "END-READ" Then
                Form1.WithinReadStatement = False
              End If
              Call AddToBusiness(Business, cWordIndex)
          End Select
        Else
          Call AddToBusiness(Business, cWordIndex)
        End If
      Next

      If Business.Length > 0 Then
        Call AddToListOfRules()
      End If
      Rule = ""
      Business = ""
    Next

    Call CreateBRExcelWorksheet(outFolder, exec)

    Call CreateBRPuml(outPumlFolder, exec)

  End Sub
  'Sub ProcessParagraphPerform(cWordIndex As Integer, currentParagraph As String)
  '  ' for PERFORM verb there are many syntax. looking for the one that references paragraph names.
  '  ' find a paragraph name
  '  ' No Need to find the end of 'PERFORM' statement
  '  ' From the manual------
  '  ' The PERFORM statement is: 
  '  ' - An out-of-line PERFORM statement When procedure-name-1 is specified. 
  '  ' - An in-line PERFORM statement When procedure-name-1 is omitted.
  '  '
  '  ' An in-line PERFORM must be delimited by the END-PERFORM phrase. 
  '  '
  '  ' The in-line and out-of-line formats cannot be combined. For example, if procedure-name-1 is specified, imperative statements and the END-PERFORM 'phrase must not be specified. 
  '  '
  '  ' The PERFORM statement formats are: 
  '  ' - Basic PERFORM 
  '  ' - TIMES phrase PERFORM 
  '  ' - UNTIL phrase PERFORM 
  '  ' - VARYING phrase PERFORM
  '  '
  '  'Dim EndIndex As Integer = 0
  '  'Dim Statement As String = ""

  '  ' BASIC Perform
  '  ' If the next phrase/word is a procedure-name there is no end-perform
  '  '  even if there is a TIMES or VARYING phrase
  '  If ListOfParagraphs.IndexOf(cWord(cWordIndex + 1)) > -1 Then
  '    Dim ParagraphReferenceText = currentParagraph & delimiter & cWord(cWordIndex + 1)
  '    If ParagraphReferenceList.IndexOf(ParagraphReferenceText) = -1 Then
  '      ParagraphReferenceList.Add(ParagraphReferenceText)
  '    End If
  '    Exit Sub
  '  End If

  '  '' looking for Conditional (UNTIL) phrase
  '  'For EndIndex = cWordIndex + 1 To cWord.Count - 1
  '  '  If cWord(EndIndex) = "UNTIL" Then
  '  '    Exit Sub
  '  '  End If
  '  'Next

  '  '' looking for TIMES phrase
  '  'For EndIndex = cWordIndex + 1 To cWord.Count - 1
  '  '  If cWord(EndIndex) = "TIMES" Then
  '  '    Exit Sub
  '  '  End If
  '  'Next

  '  '' all other combinations of the PERFORM should have an END-PERFORM
  '  ''  this is just a start of a series of commands, not really a PERFORM with a loop
  '  'For EndIndex = cWordIndex + 1 To cWord.Count - 1
  '  '  If cWord(EndIndex) = "END-PERFORM" Then
  '  '    Exit Sub
  '  '  End If
  '  'Next

  'End Sub
  'Sub ProcessParagraphGoTo(cWordIndex As Integer, currentParagraph As String)

  '  If ListOfParagraphs.IndexOf(cWord(cWordIndex + 2)) > -1 Then
  '    Dim ParagraphReferenceText = currentParagraph & delimiter & cWord(cWordIndex + 1)
  '    If ParagraphReferenceList.IndexOf(ParagraphReferenceText) = -1 Then
  '      ParagraphReferenceList.Add(ParagraphReferenceText)
  '    End If
  '    Exit Sub
  '  End If

  'End Sub

  Sub AddToBusiness(ByRef Business As String, ByRef cWordIndex As Integer)
    ' if a verb; append a newline so verb starts on a new line
    If (Form1.VerbNames.IndexOf(cWord(cWordIndex)) > -1) And (Business.Length > 0) Then
      Business &= newLine
    End If
    ' append the word to the Business text
    Business &= cWord(cWordIndex) & " "

    Call AddToParagraphReferenceList(cWordIndex)
  End Sub
  Sub AddToParagraphReferenceList(cWordIndex As Integer)
    ' this will check if the verb is a possible paragraph reference. if so, store that reference and
    ' the current paragraph we are in.
    ' this will update the ParagraphReference list
    Dim NameIndex As Integer = 0
    Select Case cWord(cWordIndex)
      Case "PERFORM"
        NameIndex = cWordIndex + 1
      Case "GO"
        NameIndex = cWordIndex + 2
      Case Else
        Exit Sub
    End Select
    If ListOfParagraphs.IndexOf(cWord(NameIndex)) > -1 Then
      Dim ParagraphReferenceText = currentParagraph & delimiter & cWord(NameIndex)
      If ParagraphReferenceList.IndexOf(ParagraphReferenceText) = -1 Then
        ParagraphReferenceList.Add(ParagraphReferenceText)
      End If
    End If
  End Sub

  Sub ProcessBR_IF(ByRef cWordIndex As Integer)
    If Business.Length > 0 Then
      Call AddToListOfRules()     'this writes out just business without a rule
    End If
    Dim EndIndex As Integer = 0
    Rule = GetRule(cWordIndex, EndIndex)
    cWordIndex = EndIndex
  End Sub
  Sub ProcessBR_ELSE()
    Dim notRule As String = "<NOT>" & Rule
    Call AddToListOfRules()
    Rule = notRule
  End Sub
  Sub ProcessBR_ENDIF()
    Call AddToListOfRules()
  End Sub
  Sub AddToListOfRules()
    subRuleNum += 1
    Dim RuleStatement As String = Form1.AddNewLineAboutEveryNthCharacters(Rule, newLine, 45).Replace(Form1.Delimiter, "")
    Dim BusinessStatement As String = Business.Replace(Form1.Delimiter, "")
    Dim idxStatement As String = LTrim(Str(idx))
    Dim RuleNumStatement As String = LTrim(Str(RuleNum))
    Dim SubRuleNumStatement As String = LTrim(Str(subRuleNum))
    If RuleStatement.Length = 0 Then
      RuleNumStatement = ""
      SubRuleNumStatement = ""
    End If

    ' Analyze the fields referenced in the rule and see if we can determine the source types of the fields.
    ' It would be either: Business field, Technical Field, or Mixture of both.
    ' A business field would be from/to a file. Or a field within a copybook.
    ' A technical field would be NOT a business field.
    Dim TypeOfRule As String = FindTypeOfRule(Rule)


    ListOfRules.Add(idxStatement & Form1.Delimiter &
                    currentParagraph & Form1.Delimiter &
                    RuleNumStatement & Form1.Delimiter &
                    SubRuleNumStatement & Form1.Delimiter &
                    RuleStatement & Form1.Delimiter &
                    BusinessStatement & Form1.Delimiter &
                    TypeOfRule)
    Rule = ""
    Business = ""
  End Sub
  Function FindTypeOfRule(ByRef rule As String) As String
    ' 2. search through the List of Fields Array looking for these fields. 
    ' remove any special condition symbols
    If rule.Length = 0 Then
      Return ""
    End If
    Dim ruleWords As String = rule.Replace("(", "").
      Replace(")", " ").
      Replace(" + ", " ").
      Replace(" - ", " ").
      Replace(" * ", " ").
      Replace(" / ", " ").
      Replace(" \ ", " ").
      Replace(" MOD ", "")
    Dim WordsInRule As New List(Of String)
    Call Form1.GetSourceWords(ruleWords, WordsInRule)
    Dim BusinessRuleCount As Integer = 0
    Dim TechnicalRuleCount As Integer = 0

    For Each RuleWord In WordsInRule
      If RuleWord.StartsWith("'") Then
        Continue For
      End If
      If RuleWord.StartsWith(Chr(34)) Then
        Continue For
      End If
      If IsNumeric(RuleWord) Then
        Continue For
      End If
      If Form1.COBOLCondWords.IndexOf(RuleWord) > -1 Then      'COBOL reserved condition words
        Continue For
      End If
      If ListOfFieldNames.IndexOf(RuleWord) > -1 Then
        BusinessRuleCount += 1
        If ListOfBusinessFieldNames.IndexOf(RuleWord & Form1.Delimiter & "Business") = -1 Then
          ListOfBusinessFieldNames.Add(RuleWord & Form1.Delimiter & "Business")
        End If
      Else
        TechnicalRuleCount += 1
        If ListOfBusinessFieldNames.IndexOf(RuleWord & Form1.Delimiter & "Technical") = -1 Then
          ListOfBusinessFieldNames.Add(RuleWord & Form1.Delimiter & "Technical")
        End If
      End If
    Next
    If BusinessRuleCount > 0 And TechnicalRuleCount = 0 Then
      Return "Business"
    End If
    If BusinessRuleCount = 0 And TechnicalRuleCount > 0 Then
      Return "Technical"
    End If
    If BusinessRuleCount = 0 And TechnicalRuleCount = 0 Then
      Return "Unknown"
    End If
    Return "Both"
  End Function
  Function GetRule(ByRef WordIndex As Integer, ByRef EndIndex As Integer) As String
    EndIndex = Form1.IndexToNextVerb(cWord, WordIndex)
    If EndIndex = -1 Then
      EndIndex = cWord.Count - 1
    End If
    Return Form1.StringTogetherWords(cWord, WordIndex, EndIndex)
  End Function

  Function GetSrcIndex(SrcIndex As Integer) As Integer
    Return SrcIndex
  End Function
  Function GetSrcStmt(SrcStatement As String) As String
    Return SrcStatement
  End Function

  Sub CreateBRExcelWorksheet(outfolder As String, exec As String)
    Dim BRRow As Integer = 1
    Dim BRFieldsRow As Integer = 1

    ' remove previous excel file
    Dim ProgramsFileName = outfolder & "\" & exec & "_BR.xlsx"
    If File.Exists(ProgramsFileName) Then
      Try
        File.Delete(ProgramsFileName)
      Catch ex As Exception
        MessageBox.Show("Error deleting " & ProgramsFileName & ":" & ex.Message)
        Exit Sub
      End Try
    End If

    ' Hide Excel app from view
    BRobjExcel.Visible = False

    ' Create spreadsheet's workbook and first worksheet
    BRWorkbook = BRobjExcel.Workbooks.Add
    BusinessRulesWorksheet = BRWorkbook.Sheets.Item(1)
    BusinessRulesWorksheet.Name = "BusinessRules"
    BusinessRulesWorksheet.Range("A1").Value = "Source"
    BusinessRulesWorksheet.Range("B1").Value = "Stmt#"
    BusinessRulesWorksheet.Range("C1").Value = "Paragraph"
    BusinessRulesWorksheet.Range("D1").Value = newLine & "Rule#"
    BusinessRulesWorksheet.Range("E1").Value = "Rule" & newLine & "Statement"
    BusinessRulesWorksheet.Range("F1").Value = "Type"
    BusinessRulesWorksheet.Range("G1").Value = "Business" & newLine & "Statement"
    BusinessRulesWorksheet.Activate()
    BusinessRulesWorksheet.Application.ActiveWindow.SplitRow = 1
    BusinessRulesWorksheet.Application.ActiveWindow.FreezePanes = True


    ' Write the Excel rows
    For Each brEntry In ListOfRules
      Dim BusinessRulesColumns As String() = brEntry.Split(Form1.Delimiter)
      BRRow += 1
      Dim row As String = LTrim(Str(BRRow))
      If BusinessRulesColumns.Count >= 7 Then
        Dim rulestatement As String = BusinessRulesColumns(2) & "." & BusinessRulesColumns(3)
        If rulestatement = "." Then
          rulestatement = ""
        End If
        BusinessRulesWorksheet.Range("A" & row).Value = exec
        BusinessRulesWorksheet.Range("B" & row).Value = BusinessRulesColumns(0) 'Stmt#
        BusinessRulesWorksheet.Range("C" & row).Value = BusinessRulesColumns(1) 'Paragraph Name
        BusinessRulesWorksheet.Range("D" & row).Value = rulestatement 'Rule#.subrule#
        BusinessRulesWorksheet.Range("E" & row).Value = BusinessRulesColumns(4) 'Rule Text
        BusinessRulesWorksheet.Range("F" & row).Value = BusinessRulesColumns(6) 'Type of Rule
        BusinessRulesWorksheet.Range("G" & row).Value = BusinessRulesColumns(5) 'Business Text
      End If
    Next

    ' Format Excel Business Rules worksheet
    If BRRow > 1 Then
      Dim row As Integer = LTrim(Str(BRRow))
      ' Format the Sheet - first row bold the columns
      rngBusinessRules = BusinessRulesWorksheet.Range("A1:G1")
      rngBusinessRules.Font.Bold = True
      ' data area autofit all columns
      BRWorkbook.Worksheets("BusinessRules").Range("A1").AutoFilter
      rngBusinessRules = BusinessRulesWorksheet.Range("A1:G" & row)
      rngBusinessRules.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
      rngBusinessRules.Columns.AutoFit()
      rngBusinessRules.Rows.AutoFit()
      ' ignore error flag that numbers being loaded into a text field
      BRobjExcel.ErrorCheckingOptions.NumberAsText = False
    End If

    ListOfBusinessFieldNames.Sort()

    ' Create second worksheet (list of field names used)
    BusinessFieldsWorksheet = BRWorkbook.Sheets.Add(After:=BRWorkbook.Worksheets(BRWorkbook.Worksheets.Count))
    BusinessFieldsWorksheet.Name = "BusinessFields"
    BusinessFieldsWorksheet.Range("A1").Value = "Field Names"
    BusinessFieldsWorksheet.Range("B1").Value = "Type of Rule"
    BusinessFieldsWorksheet.Activate()
    BusinessFieldsWorksheet.Application.ActiveWindow.SplitRow = 1
    BusinessFieldsWorksheet.Application.ActiveWindow.FreezePanes = True

    ' Write the Excel rows
    For Each brFieldsEntry In ListOfBusinessFieldNames
      brFieldsEntry = brFieldsEntry.Replace("=", "")
      Dim BusinessFieldsColumns As String() = brFieldsEntry.Split(Form1.Delimiter)
      BRFieldsRow += 1
      Dim row As String = LTrim(Str(BRFieldsRow))
      If BusinessFieldsColumns.Count >= 2 Then
        BusinessFieldsWorksheet.Range("A" & row).Value = BusinessFieldsColumns(0) 'Field Name
        BusinessFieldsWorksheet.Range("B" & row).Value = BusinessFieldsColumns(1) 'Type
      End If
    Next

    ' Format Excel Business Fields worksheet
    If BRFieldsRow > 1 Then
      Dim row As Integer = LTrim(Str(BRFieldsRow))
      ' Format the Sheet - first row bold the columns
      rngBusinessFields = BusinessFieldsWorksheet.Range("A1:B1")
      rngBusinessFields.Font.Bold = True
      BRWorkbook.Worksheets("BusinessFields").Range("A1").AutoFilter
      ' data area autofit all columns
      rngBusinessFields = BusinessFieldsWorksheet.Range("A1:B" & row)
      rngBusinessFields.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
      rngBusinessFields.Columns.AutoFit()
      rngBusinessFields.Rows.AutoFit()
      ' ignore error flag that numbers being loaded into a text field
      BRobjExcel.ErrorCheckingOptions.NumberAsText = False
    End If

    ' Position to first worksheet
    BusinessRulesWorksheet.Select(1)
    BusinessRulesWorksheet.Activate()

    ' Save worksheet and close
    BRWorkbook.SaveAs(ProgramsFileName, DefaultFormat,,, SetAsReadOnly)
    BRWorkbook.Close()
    BRobjExcel.Quit()

    GC.Collect()        'to ensure all excel objects are removed
  End Sub


  Sub CreateBRPuml(OutputFolder As String, exec As String)
    ' Open PUML and write headers. Not worrying (try/catch) about subsequent writes
    Dim pumlFileName As String = OutputFolder & "\" & exec & "_P2P.puml"
    Try
      pumlFile = My.Computer.FileSystem.OpenTextFileWriter(pumlFileName, False)
      pumlFile.WriteLine("@startuml " & exec & "_P2P")
      pumlFile.WriteLine("header ADDILite(c), by IBM")
      pumlFile.WriteLine("title Paragraph to Paragraph diagram of Program: " & exec)
      pumlFile.WriteLine("")
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error opening PumlFile COBOL")
      Exit Sub
    End Try

    ' write the details
    For Each Entry In ParagraphReferenceList
      Dim References As String() = Entry.Split(delimiter)
      pumlFile.WriteLine("(" & References(0) & ") ---> (" & References(1) & ")")
    Next

    ' close the file
    pumlFile.WriteLine("@enduml")
    pumlFile.Close()
  End Sub

End Module
