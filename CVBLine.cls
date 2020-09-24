VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVBLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function LockWindowUpdate Lib "user32" ( _
        ByVal hwndLock As Long _
) As Long

Private mCommentColor As Long
Private mKeyWordColor As Long
Private mQuotedColor As Long
Private mRegTxtColor As Long
Private CurrentLineStartOffset As Long

Private Const KeyWords = "*#Const*#Else*#ElseIf*#End If*#If*Alias*And*Any*As*Base*Binary*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Function*Get*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer*Is*LBound*Left*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Mid*Name*New*Next*Not*Nothing*Object*On*Open*Option*Or*Output*Print*Private*Property*Public*Put*Random*Read*ReDim*Resume*Return*Right*RSet*Seek*Select*Set*Single*Spc*Split*Static*Stop*String*StrReplace*Sub*Tab*Then*Then*To*True*Type*UBound*Unlock*Variant*Wend*While*With*Xor*"

Public Property Get CommentColor() As Long
    CommentColor = mCommentColor
End Property

Public Property Let CommentColor(Color As Long)
    mCommentColor = Color
    SaveSetting App.Title, "Colors", "CommentColor", mCommentColor
    PropertyChanged = "CommentColor"
End Property

Public Property Get KeyWordColor() As Long
    KeyWordColor = mKeyWordColor
End Property

Public Property Let KeyWordColor(Color As Long)
    mKeyWordColor = Color
    SaveSetting App.Title, "Colors", "KeyWordColor", mKeyWordColor
    PropertyChanged = "KeyWordColor"
End Property

Public Property Get QuotedColor() As Long
    QuotedColor = mQuotedColor
End Property

Public Property Let QuotedColor(Color As Long)
    mQuotedColor = Color
    SaveSetting App.Title, "Colors", "QuotedColor", mQuotedColor
    PropertyChanged = "QuotedColor"
End Property

Public Property Get RegTxtColor() As Long
    RegTxtColor = mRegTxtColor
End Property

Public Property Let RegTxtColor(Color As Long)
    mRegTxtColor = Color
    SaveSetting App.Title, "Colors", "RegTxtColor", mRegTxtColor
    PropertyChanged = "RegTxtColor"
End Property

Public Function ColLine(RTB As RichTextBox, TheLine As String, LineStartOffset As Long) As Boolean
Dim CommentStrings() As String
Dim CommentNum As Integer
Dim QuotedStrings() As String
Dim QuotedStrNum As Integer
Dim InQuotes As Boolean
Dim StrPos As Integer
Dim RightPreCommentSpace As String
    
'    If LockWindowUpdate(RTB.hwnd) = 0 Then
'        ' already locked by someone else!
'        Exit Function
'    End If
    
    ' Hilite the whole line
    RTB.SelStart = LineStartOffset
    RTB.SelLength = Len(TheLine)
    
    ' If the line is blank, exit quick, otherwise delete it
    If Trim$(TheLine) = "" Then
        Exit Function
    Else
        RTB.SelText = ""
    End If
    
    ' first, find out if this whole line is a comment
    If Left$(Trim$(TheLine), 4) = "Rem " Then
        RTB.SelColor = mCommentColor
        RTB.SelText = TheLine
        Exit Function
    End If
    
    ' parse the line to separate comments and deal with comments
    CommentStrings = Split(TheLine, "'")
    
    ' parse the non-comment part of the line into quoted strings
    QuotedStrings = Split(CommentStrings(0), Chr(34))
    ' store the spaces that precede the comment if there is any
    If UBound(CommentStrings) > 0 Then
        RightPreCommentSpace = Space$(Len(CommentStrings(0)) - Len(RTrim$(CommentStrings(0))))
        CommentStrings(0) = RTrim$(CommentStrings(0))   ' this will spare a few useless loops
    End If
    
    For QuotedStrNum = 0 To UBound(QuotedStrings)
        ProcessQuotedString RTB, QuotedStrings(QuotedStrNum), InQuotes
        InQuotes = Not InQuotes
    Next QuotedStrNum
    
    ' add the spaces before the comment
    RTB.SelText = RightPreCommentSpace
    ' do the comment part of the line
    RTB.SelColor = mCommentColor
    For CommentNum = 1 To UBound(CommentStrings)
        If CommentNum < UBound(CommentStrings) Then
            RTB.SelText = "' " & CommentStrings(CommentNum)
        Else
            ' use RTB.SelText = "'" & CommentStrings(CommentNum) instead to not format comments ****
            RTB.SelText = "' " & LTrim$(CommentStrings(CommentNum))
        End If
    Next CommentNum
    ' unlock text update
'    LockWindowUpdate (0&)
End Function

Private Sub Class_Initialize()
    mCommentColor = GetSetting(App.Title, "Colors", "CommentColor", vbGreen)
    mKeyWordColor = GetSetting(App.Title, "Colors", "KeyWordColor", vbBlue)
    mQuotedColor = GetSetting(App.Title, "Colors", "QuotedColor", vbBlack)
    mRegTxtColor = GetSetting(App.Title, "Colors", "RegTxtColor", vbBlack)
End Sub

Private Function GetLineNumber(RTB As RichTextBox, txtoffset As Long) As Long
Dim CRLFPos As Long
Dim Done As Boolean
    While Not Done
        CRLFPos = InStr(CRLFPos + 1, RTB.Text, vbCrLf)
        If CRLFPos = 0 Then
            CurrentLineStartOffset = 1
            Exit Function
        ElseIf CRLFPos < txtoffset Then
            GetLineNumber = GetLineNumber + 1
            CurrentLineStartOffset = CRLFPos + 1
        Else
            If CurrentLineStartOffset = 0 Then CurrentLineStartOffset = 1
            Done = True
        End If
    Wend
End Function

Private Sub Class_Terminate()
    SaveSetting App.Title, "Colors", "CommentColor", mCommentColor
    SaveSetting App.Title, "Colors", "QuotedColor", mQuotedColor
    SaveSetting App.Title, "Colors", "KeyWordColor", mKeyWordColor
    SaveSetting App.Title, "Colors", "RegTxtColor", mRegTxtColor
End Sub

Private Sub ProcessKeyWords(RTBox As RichTextBox, TheString As String)
Dim Words() As String
Dim WordNum As Integer
Dim IsLeftSpace As Boolean
Dim StrPos As Integer
Static WaitingForEnd As Boolean
    Words = Split(TheString, " ")
    IsLeftSpace = True
    RTBox.SelStart = RTBox.SelStart + Len(RTBox.SelText)
    For WordNum = 0 To UBound(Words)
        If Words(WordNum) = "" Then
            If IsLeftSpace Then
                RTBox.SelText = RTBox.SelText & " "
            End If
        Else
            IsLeftSpace = False ' from now on, skip extra spaces
            If Words(WordNum) = "BEGIN" Then
                ' this is to avoid confusing "End" to end the program
                ' and "END" which comes after "BEGIN".
                WaitingForEnd = True
            ElseIf UCase(Words(WordNum)) = "END" And WaitingForEnd Then
                WaitingForEnd = False
                RTBox.SelColor = mRegTxtColor
                RTBox.SelText = "END"
            Else
                If Len(Words(WordNum)) = 1 Then
                    ' no keyword is one char long, expedite.
                    RTBox.SelColor = mRegTxtColor
                    If WordNum = UBound(Words) Then
                        RTBox.SelText = Words(WordNum)
                    Else
                        RTBox.SelText = Words(WordNum) & " "
                    End If
                Else
                    Select Case Left$(Words(WordNum), 1)
                        Case "("
                            Select Case Right$(Words(WordNum), 1)
                                Case "("
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum) & "*", 1, _
                                                    Len(Words(WordNum)) - 2)) _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        RTBox.SelText = Words(WordNum)  ' no terminating space
                                    Else
                                        RTBox.SelText = "("
                                        RTBox.SelStart = RTBox.SelStart + 1
                                        RTBox.SelColor = mKeyWordColor
                                        RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 2)
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 2
                                        RTBox.SelText = "("
                                        RTBox.SelColor = mRegTxtColor
                                    End If
                                    
                                Case ")"
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum) & "*", 1, _
                                                    Len(Words(WordNum)) - 2)) _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        RTBox.SelText = Words(WordNum)  ' no terminating space
                                    Else
                                        RTBox.SelText = "("
                                        RTBox.SelStart = RTBox.SelStart + 1
                                        RTBox.SelColor = mKeyWordColor
                                        RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 2)
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 2
                                        RTBox.SelText = ")"
                                        RTBox.SelColor = mRegTxtColor
                                    End If
                                Case ","
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum) & "*", 1, _
                                                    Len(Words(WordNum)) - 2)) _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        RTBox.SelText = Words(WordNum)  ' no terminating space
                                    Else
                                        RTBox.SelText = "("
                                        RTBox.SelStart = RTBox.SelStart + 1
                                        RTBox.SelColor = mKeyWordColor
                                        RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 2)
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 2
                                        RTBox.SelText = ","
                                        RTBox.SelColor = mRegTxtColor
                                    End If
                                Case Else
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum), 2, _
                                                    Len(Words(WordNum)) - 1)) & "*" _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        If WordNum < UBound(Words) Then
                                            RTBox.SelText = Words(WordNum)  ' no terminating space
                                        Else
                                            RTBox.SelText = Words(WordNum) & " "
                                        End If
                                    Else
                                        RTBox.SelText = "("
                                        'RTBox.SelStart = RTBox.SelStart + 1
                                        RTBox.SelColor = mKeyWordColor
                                        If WordNum < UBound(Words) Then
                                            RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 1) & " "
                                        Else
                                            RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 1)
                                        End If
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 1
                                    End If
                            End Select
                        
                        Case ")"
                            Select Case Right$(Words(WordNum), 1)
                                Case "("
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum) & "*", 1, _
                                                    Len(Words(WordNum)) - 2)) _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        RTBox.SelText = Words(WordNum)  ' no terminating space
                                    Else
                                        RTBox.SelText = ")"
                                        RTBox.SelStart = RTBox.SelStart + 1
                                        RTBox.SelColor = mKeyWordColor
                                        RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 2)
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 2
                                        RTBox.SelText = "("
                                        RTBox.SelColor = mRegTxtColor
                                    End If
                                    
                                Case ")"
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum) & "*", 1, _
                                                    Len(Words(WordNum)) - 2)) _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        RTBox.SelText = Words(WordNum)  ' no terminating space
                                    Else
                                        RTBox.SelText = ")"
                                        RTBox.SelStart = RTBox.SelStart + 1
                                        RTBox.SelColor = mKeyWordColor
                                        RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 2)
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 2
                                        RTBox.SelText = ")"
                                        RTBox.SelColor = mRegTxtColor
                                    End If
                                Case ","
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum) & "*", 1, _
                                                    Len(Words(WordNum)) - 2)) _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        RTBox.SelText = Words(WordNum)  ' no terminating space
                                    Else
                                        RTBox.SelText = ")"
                                        RTBox.SelStart = RTBox.SelStart + 1
                                        RTBox.SelColor = mKeyWordColor
                                        RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 2)
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 2
                                        RTBox.SelText = ","
                                        RTBox.SelColor = mRegTxtColor
                                    End If
                                Case Else
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum) & "*", 1, _
                                                    Len(Words(WordNum)) - 2)) _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        RTBox.SelText = Words(WordNum)  ' no terminating space
                                    Else
                                        RTBox.SelText = "("
                                        RTBox.SelStart = RTBox.SelStart + 1
                                        RTBox.SelColor = mKeyWordColor
                                        RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 1)
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 1
                                    End If
                            End Select
                        
                        Case ","
                            Select Case Right$(Words(WordNum), 1)
                                Case "("
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum) & "*", 1, _
                                                    Len(Words(WordNum)) - 2)) _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        RTBox.SelText = Words(WordNum)  ' no terminating space
                                    Else
                                        RTBox.SelText = ", "
                                        RTBox.SelStart = RTBox.SelStart + 2
                                        RTBox.SelColor = mKeyWordColor
                                        RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 2)
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 2
                                        RTBox.SelText = "("
                                        RTBox.SelColor = mRegTxtColor
                                    End If
                                    
                                Case ")"
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum) & "*", 1, _
                                                    Len(Words(WordNum)) - 2)) _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        RTBox.SelText = Words(WordNum)  ' no terminating space
                                    Else
                                        RTBox.SelText = "( "
                                        RTBox.SelStart = RTBox.SelStart + 2
                                        RTBox.SelColor = mKeyWordColor
                                        RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 2)
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 2
                                        RTBox.SelText = ")"
                                        RTBox.SelColor = mRegTxtColor
                                    End If
                                Case ","
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum) & "*", 1, _
                                                    Len(Words(WordNum)) - 2)) _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        RTBox.SelText = Words(WordNum)  ' no terminating space
                                    Else
                                        RTBox.SelText = ", "
                                        RTBox.SelStart = RTBox.SelStart + 2
                                        RTBox.SelColor = mKeyWordColor
                                        RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 2)
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 2
                                        RTBox.SelText = ","
                                        RTBox.SelColor = mRegTxtColor
                                    End If
                                Case Else
                                    ' get the position of words(wordnum) w/out extra chars in the keywords string
                                    StrPos = InStr(UCase(KeyWords), _
                                            "*" & UCase(Mid$(Words(WordNum) & "*", 1, _
                                                    Len(Words(WordNum)) - 2)) _
                                            )
                                    RTBox.SelColor = mRegTxtColor
                                    If StrPos = 0 Then  ' if it's not found
                                        RTBox.SelText = Words(WordNum)  ' no terminating space
                                    Else
                                        RTBox.SelText = "("
                                        RTBox.SelStart = RTBox.SelStart + 1
                                        RTBox.SelColor = mKeyWordColor
                                        RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 1)
                                        RTBox.SelStart = RTBox.SelStart + Len(Words(WordNum)) - 1
                                    End If
                            End Select
                        Case Else
                            Select Case Right(Words(WordNum), 1)
                                Case ")"
                                    StrPos = InStr(UCase(KeyWords), "*" & Left$(UCase(Words(WordNum)), Len(Words(WordNum)) - 1) & "*")
                                    If StrPos = 0 Then
                                        RTBox.SelColor = mRegTxtColor
                                        If WordNum < UBound(Words) Then
                                            RTBox.SelText = Words(WordNum) & " "
                                        Else
                                            RTBox.SelText = Words(WordNum)
                                        End If
                                    Else
                                        RTBox.SelColor = mKeyWordColor
                                        If WordNum < UBound(Words) Then
                                            RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 1)
                                            RTBox.SelColor = mRegTxtColor
                                            RTBox.SelText = ") "
                                        Else
                                            RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 1)
                                            RTBox.SelColor = mRegTxtColor
                                            RTBox.SelText = ")"
                                        End If
                                    End If
                                Case "("
                                    StrPos = InStr(UCase(KeyWords), "*" & Left$(UCase(Words(WordNum)), Len(Words(WordNum)) - 1) & "*")
                                    If StrPos = 0 Then
                                        RTBox.SelColor = mRegTxtColor
                                        If WordNum < UBound(Words) Then
                                            RTBox.SelText = Words(WordNum) & " "
                                        Else
                                            RTBox.SelText = Words(WordNum)
                                        End If
                                    Else
                                        RTBox.SelColor = mKeyWordColor
                                        If WordNum < UBound(Words) Then
                                            RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 1) & "( "
                                        Else
                                            RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 1) & "("
                                        End If
                                    End If
                                Case ","
                                    StrPos = InStr(UCase(KeyWords), "*" & Left$(UCase(Words(WordNum)), Len(Words(WordNum)) - 1) & "*")
                                    If StrPos = 0 Then
                                        RTBox.SelColor = mRegTxtColor
                                        If WordNum < UBound(Words) Then
                                            RTBox.SelText = Words(WordNum) & " "
                                        Else
                                            RTBox.SelText = Words(WordNum)
                                        End If
                                    Else
                                        RTBox.SelColor = mKeyWordColor
                                        If WordNum < UBound(Words) Then
                                            RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 1) & ", "
                                        Else
                                            RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)) - 1) & ","
                                        End If
                                    End If
                                Case Else
                                    StrPos = InStr(UCase(KeyWords), "*" & UCase(Words(WordNum)) & "*")
                                    If StrPos = 0 Then
                                        RTBox.SelColor = mRegTxtColor
                                        If WordNum < UBound(Words) Then
                                            RTBox.SelText = Words(WordNum) & " "
                                        Else
                                            RTBox.SelText = Words(WordNum)
                                        End If
                                    Else
                                        RTBox.SelColor = mKeyWordColor
                                        If WordNum < UBound(Words) Then
                                            RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum))) & " "
                                        Else
                                            RTBox.SelText = Mid$(KeyWords, StrPos + 1, Len(Words(WordNum)))
                                        End If
                                    End If
                            End Select
                    End Select
                End If  ' If Len(Words(WordNum)) = 1 Then
            End If  ' If Words(WordNum) = "BEGIN" Then
        End If  ' If Words(WordNum) = "" Then
    Next WordNum
End Sub

Private Sub ProcessQuotedString(RTBox As RichTextBox, TheString As String, InQuotes As Boolean)
    If InQuotes Then
        RTBox.SelColor = mQuotedColor
        RTBox.SelText = Chr(34) & TheString & Chr(34)
    Else
        ProcessKeyWords RTBox, TheString
    End If
End Sub


