Public Class HyperLinkBuilder
  Private Utilities As String()
  Private Const QUOTE As Char = Chr(34)

  Public Sub New(utilities As String())
    Me.Utilities = utilities
  End Sub

  Public Function CreateHyperLinkSources(text As String, ByRef SourceType As String) As String
    ' Based on Sourcetype, determine which CreateHyperlink routine to call
    Select Case SourceType
      Case "COBOL"
        Return CreateHyperLinkCobolSources(text)
      Case "Easytrieve"
        Return CreateHyperLinkEasytreiveSources(text)
      Case "Assembler"
        Return CreateHyperLinkAssemblerSources(text)
      Case Else
        Return "n/a"
    End Select
  End Function
  Public Function CreateHyperLinkCobolSources(text As String) As String
    If Array.IndexOf(Utilities, text) > -1 Then
      Return "n/a"
    End If
    Return "=HYPERLINK(Summary!B3&Summary!B6&" &
             QUOTE & text &
            QUOTE & ", " & QUOTE & text & QUOTE & ")"
  End Function
  Public Function CreateHyperLinkEasytreiveSources(text As String) As String
    If Array.IndexOf(Utilities, text) > -1 Then
      Return "n/a"
    End If
    Return "=HYPERLINK(Summary!B3&Summary!B7&" &
             QUOTE & text &
            QUOTE & ", " & QUOTE & text & QUOTE & ")"
  End Function
  Public Function CreateHyperLinkAssemblerSources(text As String) As String
    If Array.IndexOf(Utilities, text) > -1 Then
      Return "n/a"
    End If
    Return "=HYPERLINK(Summary!B3&Summary!B8&" &
             QUOTE & text &
            QUOTE & ", " & QUOTE & text & QUOTE & ")"
  End Function
  Public Function CreateHyperLinkCopybookSources(text As String) As String
    Return "=HYPERLINK(Summary!B3&Summary!B9&" &
             QUOTE & text &
            QUOTE & ", " & QUOTE & text & QUOTE & ")"
  End Function
  Public Function CreateHyperLinkBRXLS(text As String) As String
    If Array.IndexOf(Utilities, text) > -1 Then
      Return "n/a"
    End If
    Return "=HYPERLINK(Summary!B3&Summary!B11&" &
             QUOTE & text & "_BR.xlsx" &
            QUOTE & ", " & QUOTE & text & QUOTE & ")"
  End Function
  Public Function CreateHyperLinkSVGFlowchart(text As String) As String
    If Array.IndexOf(Utilities, text) > -1 Then
      Return "n/a"
    End If
    Return "=HYPERLINK(Summary!B3&Summary!B10&" &
            QUOTE & text & ".svg" &
            QUOTE & ", " & QUOTE & "view" & QUOTE & ")"
  End Function
  Public Function CreateHyperLinkP2PFlowchart(text As String) As String
    If Array.IndexOf(Utilities, text) > -1 Then
      Return "n/a"
    End If
    Return "=HYPERLINK(Summary!B3&Summary!B10&" &
            QUOTE & text & "_P2P.svg" &
            QUOTE & ", " & QUOTE & "view" & QUOTE & ")"
  End Function
  Public Function CreateHyperLinkJobFlowchart(text As String) As String
    Return "=HYPERLINK(Summary!B3&Summary!B10&" &
            QUOTE & text & "_JOB.svg" &
            QUOTE & ", " & QUOTE & "view" & QUOTE & ")"
  End Function
  Public Function CreateHyperLinkProcs(text As String) As String
    Return "=HYPERLINK(Summary!B3&Summary!B5&" &
            QUOTE & text &
            QUOTE & ", " & QUOTE & text & QUOTE & ")"
  End Function
  Public Function CreateHyperLinkJobSource(text As String) As String
    Return "=HYPERLINK(Summary!B3&Summary!B4&" &
            QUOTE & text &
            QUOTE & ", " & QUOTE & text & QUOTE & ")"
  End Function

End Class
