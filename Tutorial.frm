VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Rich Text Box / Common Dialog Tutorial"
   ClientHeight    =   8160
   ClientLeft      =   165
   ClientTop       =   -1500
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   375
      Left            =   10320
      TabIndex        =   7
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Character Count"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Set Back Colour"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Set Font"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save File"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   7680
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load File"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random Colours"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13150
      _Version        =   393217
      BackColor       =   16777215
      ScrollBars      =   3
      TextRTF         =   $"Tutorial.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'set up a loop 1000 times
For i = 1 To 1000
    
    'make a random color to type the text in
    RichTextBox1.SelColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    
    
    'this command lets the computer share resources,
    'so other programs still work when the loop is going
    '(if you don't get it, try taking this command out!)
    DoEvents
    
    'the command Chr(10) tells the computer to
    'move down one line
    RichTextBox1.SelText = Chr(10)
    
        'add the text "Its Colour not Color!"
    RichTextBox1.SelText = "Its Colour not Color!"
     
    'the command Chr(10) tells the computer to
    'move down one line (same as above)
    RichTextBox1.SelText = Chr(10)
    
    'add the text "You crazy guys, you spell funny!"
    RichTextBox1.SelText = "You crazy guys, you spell funny!"
    
    'the command Chr(10) tells the computer to
    'move down one line (same as above)
    RichTextBox1.SelText = Chr(10)



'loop back to top
Next i


End Sub

Private Sub Command2_Click()

'declare a variable that we
'can use to open a file using
'the common dialog box
Dim FileToOpen

'if there is an error just go on
On Error Resume Next

'this stops the "Open as read only" checkbox
'from going on to the CommonDialog (take it out
'if you want the checkbox)
CommonDialog1.Flags = &H4

'OK, this is a bit messy but I'll try to explain it.
'the common dialogs filter is the bit where you select file type.
'someone has to tell the filter what files go in and here it is!
'"All files (*.*)|*.*|Text Files (*.txt)|*.txt|"
'   ^     ^   ^    ^     ^     ^   ^       ^
' display the text |  show "text files"    |
'                  |                       |
'           tell computer            tell computer
'          what all files is       what text files are

'I know its really tricky, if you have any further questions
'about the filter, or anything else, email me cbsg@interact.net.au
'i'll always write back, no matter if the question sounds basic
'theres always about 50 other people thinking the same thing!
CommonDialog1.Filter = "All files (*.*)|*.*|Text Files (*.txt)|*.txt|"


'show the Open Dialog box
CommonDialog1.ShowOpen

'call the variable "FileToOpen" and name it
'the file that has been selected in the
'common dialog Open box.
FileToOpen = CommonDialog1.FileName

'Load the file (that we have called
'FileToOpen) into the RichTextBox
RichTextBox1.LoadFile (FileToOpen)



End Sub

Private Sub Command3_Click()

'declare variable naming the file to save
Dim FileToSave

'if there is an error just go on anyway
'(!THIS IS NOT A GOOD WAY TO HANDLE ERRORS!
'how is the user going to know there was
'an error?
'I'll make a tutorial on error handling
'if people want.)
On Error Resume Next

'turn on the over write promt so that
'the user is asked if they want to overwrite
'an existing file
CommonDialog1.Flags = &H2

'show the save dialog
CommonDialog1.ShowSave

'name FileToSave the file that the user
'has just chosen/made in the save dialog box
FileToSave = CommonDialog1.FileName

'tell richbox to save the file using
'the FileToSave variable we have just
'defined.
RichTextBox1.SaveFile (FileToSave)

End Sub

Private Sub Command4_Click()

'this is hard to explain, but basicly it tells
'the common dialog box what fonts it is to list.
'NOTE: if you take this command out, common dialog
'will crash and tell you that there is no fonts
'to display
CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth

  'tell the common dialog to show the font
  'selecting dialog
  CommonDialog1.ShowFont
  
  'set rich text box's font to common dialog's font name,
  'which was just chosen from the dialog
  RichTextBox1.SelFontName = CommonDialog1.FontName
  
  'set rich text box's font size to what is selected in
  'the common dialog font box
  RichTextBox1.Font.Size = CommonDialog1.FontSize
  
  'set rich text box's font bold to what is selected in
  'the common dialog font box
  RichTextBox1.Font.Bold = CommonDialog1.FontBold
  
  'set rich text box's font italic to what is selected in
  'the common dialog font box
  RichTextBox1.Font.Italic = CommonDialog1.FontItalic
  
  'set rich text box's font underline to what is selected in
  'the common dialog font box
  RichTextBox1.Font.Underline = CommonDialog1.FontUnderline
  
  'set rich text box's font strikethrough to what is selected in
  'the common dialog font box
  'NOTE: Strikethrough  and  FontStrikethru are spelt differently
  'must just be some lazy thing
  RichTextBox1.Font.Strikethrough = CommonDialog1.FontStrikethru
  
  'Change the colour of selected text using the colour defined
  'in the font box
  RichTextBox1.SelColor = CommonDialog1.Color
  

End Sub

Private Sub Command5_Click()
'this command tells the common dialog to enable
'the custom colours button
CommonDialog1.Flags = cdlCCRGBInit

'this command tells the common dialog to
'show the colour selection box
CommonDialog1.ShowColor
  
'set rich text boxes back color to the colour selected
'in the common dialog
RichTextBox1.BackColor = CommonDialog1.Color

End Sub

Private Sub Command6_Click()
'Define a variable for holding the
'information on how many characters there are
Dim CharCount

'using the Len (length) command, count
'how many characters make up rich text box's text
'now name this number CharCount
CharCount = Len(RichTextBox1.Text)

'show a message box showing how many characters there
'are, as well as a message, title and information icon
MsgBox "There are a total of " & CharCount & " Characters in the RichTextBox control.", vbInformation, "Richtext box / CMDLG Tutorial"
End Sub

Private Sub Command7_Click()

    'Exit the program using the END command
    End
    
End Sub
