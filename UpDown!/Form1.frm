VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Up Down Writer !"
   ClientHeight    =   3120
   ClientLeft      =   165
   ClientTop       =   915
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "RomanC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "RomanC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Save"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAbout_Click()
  MsgBox "Updown writer!!! " + vbCrLf + _
         "Whatever you write in first text box" + vbCrLf + _
         "will be written in next text box up down side!" + vbCrLf + _
         "================================================" + vbCrLf + _
         "<<< Mahdi Shakouri Rad : Mahdi_Rad@Yahoo.com >>>", , "About!"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuNew_Click()
Text1 = ""
Text2 = ""
End Sub

Private Sub mnuSave_Click()
  CommonDialog1.ShowSave
  Open CommonDialog1.FileName For Output As 1
    Print #1, Text2
  Close 1
End Sub

Private Sub Text1_Change()
Select Case Right$(Text1, 1)
  Case "A", "a"
    Text2 = "V" + Text2
  Case "B", "b"
    Text2 = "q" + Text2
  Case "C", "c"
    Text2 = ")" + Text2
  Case "D", "d"
    Text2 = "p" + Text2
  Case "E", "e"
    Text2 = "3" + Text2
  Case "F", "f"
    Text2 = "ƒ" + Text2
  Case "G", "g"
    Text2 = "6" + Text2
  Case "H", "h"
    Text2 = "y" + Text2
  Case "I"
    Text2 = "I" + Text2
  Case "i"
    Text2 = "!" + Text2
  Case "J", "j"
    Text2 = "[" + Text2
  Case "K", "k"
    Text2 = "k" + Text2
  Case "L"
    Text2 = "7" + Text2
  Case "l"
    Text2 = "l" + Text2
  Case "M"
    Text2 = "W" + Text2
  Case "m"
    Text2 = "w" + Text2
  Case "N"
    Text2 = "N" + Text2
  Case "n"
    Text2 = "u" + Text2
  Case "O"
    Text2 = "O" + Text2
  Case "o"
    Text2 = "o" + Text2
  Case "P", "p"
    Text2 = "d" + Text2
  Case "Q", "q"
    Text2 = "b" + Text2
  Case "R", "r"
    Text2 = "J" + Text2
  Case "S"
    Text2 = "S" + Text2
  Case "s"
    Text2 = "s" + Text2
  Case "T", "t"
    Text2 = "+" + Text2
  Case "U", "u"
    Text2 = "n" + Text2
  Case "V", "v"
    Text2 = "A" + Text2
  Case "W"
    Text2 = "M" + Text2
  Case "w"
    Text2 = "m" + Text2
  Case "X"
    Text2 = "X" + Text2
  Case "x"
    Text2 = "x" + Text2
  Case "Y", "y"
    Text2 = "£" + Text2
  Case "Z"
    Text2 = "Z" + Text2
  Case "z"
    Text2 = "z" + Text2
  Case "!"
    Text2 = "i" + Text2
  Case " "
    Text2 = " " + Text2
  Case ","
    Text2 = "'" + Text2
  Case "'"
    Text2 = "," + Text2
  Case "0"
    Text2 = "0" + Text2
  Case "1"
    Text2 = "1" + Text2
  Case "2"
    Text2 = "Z" + Text2
  Case "3"
    Text2 = "E" + Text2
  Case "4"
    Text2 = "h" + Text2
  Case "5"
    Text2 = "5" + Text2
  Case "6"
    Text2 = "9" + Text2
  Case "7"
    Text2 = "L" + Text2
  Case "8"
    Text2 = "8" + Text2
  Case "9"
    Text2 = "6" + Text2
    
End Select
End Sub
