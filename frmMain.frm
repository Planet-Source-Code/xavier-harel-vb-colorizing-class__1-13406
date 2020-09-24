VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdColorSettings 
      Caption         =   "Color settings"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdColorize 
      Caption         =   "Colorize"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "Open file"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CMDLG 
      Left            =   0
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   2640
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdColorize_Click()
    ColorizeCode = True
    FormatCurrentText RichTextBox1, True
End Sub

Private Sub cmdColorSettings_Click()
    frmColorSettings.Show vbModal
End Sub

Private Sub cmdOK_Click()
    Unload frmColorSettings
    Unload Me
End Sub

Private Sub cmdOpenFile_Click()
Dim sFileName As String 'Holds the filename to open
    
    On Error GoTo errHandler
    'Retrieve the filename to open
    With CMDLG
        .Filter = "Text Files (*.txt)|*.txt|VB files (*.frm; *.bas; *.cls)|" & _
                "*.frm; *.bas; *.cls|All Files (*.*)|*.*"
        .FilterIndex = 1 'default is text files
        .DialogTitle = "Open Text File"
        .ShowOpen
        If .FileName = "" Then Exit Sub
        sFileName = .FileName
    End With
    Set tso = fso.OpenTextFile(sFileName, ForReading) 'Open the text file
    RichTextBox1.Text = tso.ReadAll 'Read the text from the file into the textbox (default: no word wrap)
    Exit Sub
errHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub Form_Load()
    Set fso = New FileSystemObject
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then Exit Sub
    If ScaleHeight < 1300 Then Height = 1600
    If ScaleWidth < 1500 Then ScaleWidth = 1500
    RichTextBox1.Top = 120
    RichTextBox1.Left = 120
    RichTextBox1.Height = ScaleHeight - 1200
    RichTextBox1.Width = ScaleWidth - 240
    cmdOK.Top = ScaleHeight - 1000
    cmdOK.Left = ScaleWidth / 2 - cmdOK.Width - 50
    cmdColorize.Top = ScaleHeight - 1000
    cmdColorize.Left = ScaleWidth / 2 + 50
    cmdOpenFile.Top = ScaleHeight - 1000
    cmdOpenFile.Left = cmdOK.Left - cmdOpenFile.Width - 100
    cmdColorSettings.Top = ScaleHeight - 1000
    cmdColorSettings.Left = cmdColorize.Left + cmdColorize.Width + 100
End Sub

Private Sub RichTextBox1_Change()
    If ColorizeCode Then
        Static CursorLineNumAtLastChange As Long    ' 0-based
        Dim CurrentCursorLineNum As Long            ' 0-based
        Dim TxtLines() As String
        Dim OffsetOfCurrentLine As Long
        Dim CurLine As New CVBLine
        Dim CurrentCursorPosInText As Long
        CurrentCursorPosInText = RichTextBox1.SelStart
        TxtLines = Split(RichTextBox1.Text, vbCrLf)
        For CurrentCursorLineNum = 0 To UBound(TxtLines)
            If CurrentCursorPosInText > OffsetOfCurrentLine Then
                If CurrentCursorPosInText <= OffsetOfCurrentLine + Len(TxtLines(CurrentCursorLineNum)) Then
                    Exit For
                Else
                    OffsetOfCurrentLine = OffsetOfCurrentLine + _
                                          Len(TxtLines(CurrentCursorLineNum)) + _
                                          2 ' 2 for the length of vbCRLF between lines
                End If
            Else
                Exit For
            End If
        Next CurrentCursorLineNum
        If CurrentCursorLineNum <> CursorLineNumAtLastChange Then
            CurLine.ColLine RichTextBox1, _
                            TxtLines(CurrentCursorLineNum), _
                            OffsetOfCurrentLine
        End If
        CursorLineNumAtLastChange = CurrentCursorLineNum
        RichTextBox1.SelStart = CurrentCursorPosInText
    End If  ' If ColorizeCode Then
    bDocChanged = True
End Sub
