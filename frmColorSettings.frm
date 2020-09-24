VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColorSettings 
   Caption         =   "Form2"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   LinkTopic       =   "Form2"
   ScaleHeight     =   2490
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHidden 
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Save && E&xit"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame fraColors 
      Caption         =   "Color settings"
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin MSComctlLib.ImageCombo icboColors 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "ImageCombo1"
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblRegText 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Regular text color"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblQuotes 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quoted color"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblKeywords 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Keyword color"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblComments 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comments color"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblClickOnColor 
         Caption         =   "Click on the textbox to adjust color"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
   End
   Begin MSComctlLib.ImageList ilst1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "frmColorSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const COLOR_LIST = "Black,Maroon,Green,Olive,Navy,Purple,Teal,Gray,Silver,Red,Lime,Yellow,Blue,Fuschia,Aqua,White"
Dim Colors() As String  'Array to hold each color name
Private Const COLORS_COUNT = 16
Private CallingLabel As Label
Private VBColors As New CVBLine

' Evaluates a string containing a color name and returns a long representing that color.
' It can be used to set color properties.
Private Function GetColorFromString(ByVal sColor As String) As Long
    ' Make color string is all uppercase
    sColor = StrConv(sColor, vbUpperCase)
    
    Select Case sColor
        Case "BLACK": GetColorFromString = &H0
        Case "MAROON": GetColorFromString = &H80
        Case "GREEN": GetColorFromString = 49152    ' &HC000
        Case "OLIVE": GetColorFromString = 32896    ' &H8080
        Case "NAVY": GetColorFromString = &H800000
        Case "PURPLE": GetColorFromString = &H800080
        Case "TEAL": GetColorFromString = &H808000
        Case "GRAY": GetColorFromString = &H808080
        Case "SILVER": GetColorFromString = &HC0C0C0
        Case "RED": GetColorFromString = &HFF
        Case "LIME": GetColorFromString = 65280     ' &HFF00
        Case "YELLOW": GetColorFromString = 65535   ' &HFFFF
        Case "BLUE": GetColorFromString = &HFF0000
        Case "FUSCHIA": GetColorFromString = &HFF00FF
        Case "AQUA": GetColorFromString = &HFFFF00
        Case "WHITE": GetColorFromString = &HFFFFFF
        Case Else: GetColorFromString = &H0 'Black (default)
    End Select
End Function

Private Sub InitColors(InitColor As Long)
Dim i As Integer
Dim ColToSelect As Integer
Dim colorUpper As String
Dim colorProper As String 'Proper case for text
Dim CurColor As Long
    ' Split COLOR_LIST constant into the Color() array.
    Colors = Split(COLOR_LIST, ",") 'VB 6 only
    ' clean the listimage and listbox is needed
    If ilst1.ListImages.Count Then
        ilst1.ListImages.Clear
        icboColors.ComboItems.Clear
    End If
    ' Loop through array, creating a picture of each color
    For i = 0 To COLORS_COUNT - 1 'Step 1
      CreateColorImage picHidden, ilst1, Colors(i)
      CurColor = GetColorFromString(Colors(i))
      If CurColor = InitColor Then ColToSelect = i
    Next
    ' Initialize imagelist for combo box
    icboColors.ImageList = ilst1
    ' Loop through and add each picture created earlier
    For i = 0 To COLORS_COUNT - 1
      colorProper = StrConv(Colors(i), vbProperCase)
      colorUpper = StrConv(Colors(i), vbUpperCase)
      icboColors.ComboItems.Add , colorUpper, colorProper, colorUpper
    Next
    ' select the color that was passed to this sub
    If ColToSelect Then
        icboColors.ComboItems(ColToSelect + 1).Selected = True
    Else
        icboColors.ComboItems(1).Selected = True
    End If
    'icboColors.ComboItems(ColToSelect + 1).Text
    icboColors.Refresh
End Sub

Private Sub cmdExit_Click()
    VBColors.CommentColor = lblComments.ForeColor
    VBColors.KeyWordColor = lblKeywords.ForeColor
    VBColors.QuotedColor = lblQuotes.ForeColor
    VBColors.RegTxtColor = lblRegText.ForeColor
    Hide
End Sub

Private Sub cmdOK_Click()
    HideLabelsShowIcboColors False
    lblClickOnColor.Caption = "Click on the textbox to adjust color"
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

Public Sub ColorLabels()
Dim TmpVBColors As New CVBLine
    lblComments.ForeColor = TmpVBColors.CommentColor
    lblQuotes.ForeColor = TmpVBColors.QuotedColor
    lblKeywords.ForeColor = TmpVBColors.KeyWordColor
    lblRegText.ForeColor = TmpVBColors.RegTxtColor
    Set TmpVBColors = Nothing
End Sub

Private Sub CreateColorImage(picBox As VB.PictureBox, iList As MSComctlLib.ImageList, sColor As String)
    ' 1- Make color string all uppercase
    sColor = StrConv(sColor, vbUpperCase)
    With picBox   'picHidden
        .AutoRedraw = True  'Ensure "True" painting
        
        ' Set image size in pixels. A height of 16 will leave no border between items
        ' in the image cbo.
        .Width = picBox.ScaleX(32, vbPixels, vbTwips)
        .Height = picBox.ScaleY(12, vbPixels, vbTwips)
        
        ' Flatten box to eliminate pic distortion (ensures the created image
        ' is the dimensions specified).
        .BorderStyle = 0
        .Appearance = 0
        
        ' set the backcolor of the picture box to sColor.
        .BackColor = GetColorFromString(sColor)
        
        ' Using the Line function, draw a 1-pixel black line around the box to look better.
        ' (The white image won't bleed into background of the image cbo).
        .ForeColor = vbBlack
        picBox.Line (0, 0)-(picBox.Width - picBox.ScaleX(1, vbPixels, vbTwips), _
                picBox.Height - picBox.ScaleY(1, vbPixels, vbTwips)), , B
        
        ' Set the image (the way the picture box looks) to the box's picture property,
        ' and add it to the image list, using the color string as the image key in the
        ' image list.
        .Picture = picBox.Image
        iList.ListImages.Add , sColor, .Picture
        
        .Cls    'Clear picture box
        .Picture = Nothing  'Clear the picture property
    End With
End Sub

Private Sub Form_Load()
    ColorLabels
End Sub

Private Sub HideLabelsShowIcboColors(DoIt As Boolean)
    icboColors.Visible = DoIt
    lblComments.Visible = Not DoIt
    lblKeywords.Visible = Not DoIt
    lblQuotes.Visible = Not DoIt
    lblRegText.Visible = Not DoIt
    If Not DoIt Then
        lblClickOnColor.Caption = "Click on the textbox to adjust color"
        cmdOK.Enabled = False
        cmdExit.Default = True
    Else
        cmdOK.Enabled = True
        cmdOK.Default = True
    End If
End Sub

Private Sub icboColors_Click()
Dim TmpVBCol As New CVBLine
    ' Update the cmdbutton with the image combo's selected color.
    CallingLabel.ForeColor = GetColorFromString(icboColors.SelectedItem.Text)
    HideLabelsShowIcboColors False
End Sub

Private Sub lblComments_Click()
    HideLabelsShowIcboColors True
    InitColors VBColors.CommentColor
    Set CallingLabel = lblComments
    lblClickOnColor.Caption = "Select a color for comments"
End Sub

Private Sub lblKeywords_Click()
    HideLabelsShowIcboColors True
    InitColors VBColors.KeyWordColor
    Set CallingLabel = lblKeywords
    lblClickOnColor.Caption = "Select a color for keywords"
End Sub

Private Sub lblQuotes_Click()
    HideLabelsShowIcboColors True
    InitColors VBColors.QuotedColor
    Set CallingLabel = lblQuotes
    lblClickOnColor.Caption = "Select a color for quoted strings"
End Sub

Private Sub lblRegText_Click()
    HideLabelsShowIcboColors True
    InitColors VBColors.RegTxtColor
    Set CallingLabel = lblRegText
    lblClickOnColor.Caption = "Select a color for regular text"
End Sub


