VERSION 5.00
Begin VB.Form Katakana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KATAKANA ENTRY"
   ClientHeight    =   4275
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Kana to Clipboard"
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "CitrusFruits"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Label Label5 
      Caption         =   "appear in the lower box."
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "box for the Japanese Katakana to"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Enter English Charactors in the top"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "By Richard Nicol"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Katakana Entry"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Menu cmdHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Katakana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By Richard Nicol

'eminent_uk@hotmail.com

Option Explicit


Private Sub cmdHelp_Click()

help.Show

End Sub

Private Sub Command1_Click()

   Clipboard.Clear
   Clipboard.SetText (Text2.Text)

End Sub

Private Sub Text1_Change()
Dim before
Dim after
Dim i
Dim x

before = Text1.Text

For i = 1 To Len(before)

If UCase(Mid(before, i, 1)) = vbCrLf Then
    after = after & vbCrLf
End If
If UCase(Mid(before, i, 1)) = " " Then
    after = after & " "
End If
If UCase(Mid(before, i, 1)) = "A" Then
    after = after & "3"
End If
If UCase(Mid(before, i, 1)) = "E" Then
    after = after & "5"
End If
If UCase(Mid(before, i, 1)) = "I" Then
    after = after & "e"
End If
If UCase(Mid(before, i, 1)) = "O" Then
    after = after & "6"
End If
If UCase(Mid(before, i, 1)) = "U" Then
    after = after & "4"
End If
If UCase(Mid(before, i, 1)) = "Q" Then
    after = after & "Z"
End If
If UCase(Mid(before, i, 1)) = "-" Then
    after = after & "["
End If
If UCase(Mid(before, i, 1)) = "." Then
    after = after & "Y"
End If

If UCase(Mid(before, i, 2)) = "BA" Then
    after = after & "F"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "BE" Then
    after = after & "`"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "BI" Then
    after = after & "V"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "BU" Then
    after = after & """"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "BO" Then
    after = after & "="
    x = 1
End If

If UCase(Mid(before, i, 2)) = "DA" Then
    after = after & "Q"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "DE" Then
    after = after & "W"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "JI" Then
    after = after & "A"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "DI" Then
    after = after & "A"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "DO" Then
    after = after & "S"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "DU" Then
    after = after & "D"
    x = 1
End If

If UCase(Mid(before, i, 2)) = "GA" Then
    after = after & "T"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "GE" Then
    after = after & "*"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "GI" Then
    after = after & "G"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "GO" Then
    after = after & "B"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "GU" Then
    after = after & "H"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "HA" Then
    after = after & "f"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "HE" Then
    after = after & "^"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "HI" Then
    after = after & "v"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "HO" Then
    after = after & "-"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "HU" Or UCase(Mid(before, i, 2)) = "FU" Then
    after = after & "2"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "KA" Then
    after = after & "t"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "KE" Then
    after = after & ":"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "KI" Then
    after = after & "g"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "KO" Then
    after = after & "b"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "KU" Then
    after = after & "h"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "MA" Then
    after = after & "j"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "ME" Then
    after = after & "/"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "MI" Then
    after = after & "n"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "MO" Then
    after = after & "m"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "MU" Then
    after = after & "]"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "NN" Then
    after = after & "y"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "NA" Then
    after = after & "u"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "NE" Then
    after = after & ","
    x = 1
End If
If UCase(Mid(before, i, 2)) = "NI" Then
    after = after & "i"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "NO" Then
    after = after & "k"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "NU" Then
    after = after & "1"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "PA" Then
    after = after & "U"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "PE" Then
    after = after & "<"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "PI" Then
    after = after & "I"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "PO" Then
    after = after & "K"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "PU" Then
    after = after & "!"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "RA" Then
    after = after & "o"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "RE" Then
    after = after & ";"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "RI" Then
    after = after & "l"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "RO" Then
    after = after & "N"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "RU" Then
    after = after & "."
    x = 1
End If
If UCase(Mid(before, i, 2)) = "SA" Then
    after = after & "x"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "SE" Then
    after = after & "p"
    x = 1
End If
If UCase(Mid(before, i, 3)) = "SHI" Then
    after = after & "d"
    x = 2
End If
If UCase(Mid(before, i, 2)) = "SI" Then
    after = after & "d"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "SO" Then
    after = after & "c"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "SU" Then
    after = after & "r"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "TA" Then
    after = after & "q"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "TE" Then
    after = after & "w"
    x = 1
End If
If UCase(Mid(before, i, 3)) = "CHI" Then
    after = after & "a"
    x = 2
End If
If UCase(Mid(before, i, 2)) = "TI" Then
    after = after & "a"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "TO" Then
    after = after & "s"
    x = 1
End If
If UCase(Mid(before, i, 3)) = "TSU" Then
    after = after & "z"
    x = 2
End If
If UCase(Mid(before, i, 2)) = "TU" Then
    after = after & "z"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "WA" Then
    after = after & "0"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "WO" Then
    after = after & "}"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "WU" Then
    after = after & "4"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "YA" Then
    after = after & "7"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "YO" Then
    after = after & "9"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "YU" Then
    after = after & "8"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "ZA" Then
    after = after & "X"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "ZE" Then
    after = after & "P"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "ZI" Then
    after = after & "D"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "ZO" Then
    after = after & "C"
    x = 1
End If
If UCase(Mid(before, i, 2)) = "ZU" Then
    after = after & "R"
    x = 1
End If

i = i + x
x = 0
Next i



Text2.Text = after

End Sub
