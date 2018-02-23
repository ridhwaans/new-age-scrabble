VERSION 5.00
Begin VB.Form frmDictionary 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6735
   Icon            =   "frmDictionary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDictionary.frx":1CCA
   ScaleHeight     =   3240
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label lblOutput 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   6015
   End
   Begin VB.Label lblSearch 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a word to look up:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "frmDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearch_Click()
    txtSearch.Text = Replace$(txtSearch.Text, Space(1), Space(0))

    If txtSearch.Text <> "" Then    'If a word is entered in the textbox
        intword = FreeFile

        'Opens the dictionary text file
        Open App.Path & "\game_files\dictionary\" + Left(txtSearch.Text, 1) + ".txt" For Input As #intword    'previous filepath "\game_files\dictionary.txt"

        Do
            Input #intword, strNext
            If StrComp((UCase(txtSearch.Text)), (UCase(strNext)), vbTextCompare) = 0 Then    'Compares word entered to each word in the dictionary
                lblOutput.Caption = txtSearch.Text & " is a VALID word"
                Exit Do
            Else
                lblOutput.Caption = txtSearch.Text & " is an INVALID word"
            End If
        Loop Until EOF(intword)    'Check each word until end of the dictionary is reached
    Else
        lblOutput.Caption = ""
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = Chr(13) Then    'if ENTER or carriage return then call print command button
        Call cmdSearch_Click
    End If
End Sub

