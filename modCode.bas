Attribute VB_Name = "modCode"
Global fso As FileSystemObject 'Object Variable for the File System Object
Global tso As TextStream 'Object Variable for the File System Object
Global ColorizeCode As Boolean

Public Sub FormatCurrentText(RTB As RichTextBox, VBStyle As Boolean)
Dim NumberOfLines As Long
Dim TmpLineNum As Long
Dim Textlines() As String
Dim PreCursorPos As Long
Dim FormattedLine As New CVBLine
Dim LineStartOffset As Long
    ' store the cursor position before we start the operation
    PreCursorPos = RTB.SelStart
    If Not VBStyle Then
        RTB.SelStart = 0
        RTB.SelLength = Len(RTB.Text)
        RTB.SelColor = 0
        RTB.SelStart = PreCursorPos
        Exit Sub
    End If
    ' Fill the TextLines() array with the text
    Textlines = Split(RTB.Text, vbCrLf)
    ' Get the number of lines in this text
    NumberOfLines = UBound(Textlines)
    ' Flag that this is not a change to avoid triggering line formatting for each change
    NoChange = True
    ' Empty the textbox
    RTB.Text = ""
    ' Fill it again with formatted data
    For TmpLineNum = LBound(Textlines) To NumberOfLines
        LineStartOffset = Len(RTB.Text)
        FormattedLine.ColLine RTB, Textlines(TmpLineNum), LineStartOffset
        LineStartOffset = Len(RTB.Text)
        ' add a CRLF (only if this is not to the end of the last line of the text)
        If TmpLineNum < NumberOfLines Then
            RTB.SelStart = LineStartOffset
            RTB.SelText = vbCrLf
        End If
    Next
    ' restore the nochange flag
    NoChange = False
End Sub


