VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup"
   ClientHeight    =   3225
   ClientLeft      =   9105
   ClientTop       =   2880
   ClientWidth     =   4620
   Icon            =   "frmDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDialog.frx":1CCA
   ScaleHeight     =   3225
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdChange 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtPlayer2Name 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "Player 2"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtPlayer1Name 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "Player 1"
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblComputerName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Computer AI"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblVersus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Versus"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click(Index As Integer)
    If txtPlayer2Name.Visible = True Then
        txtPlayer2Name.Visible = False
        lblComputerName.Visible = True
    Else
        txtPlayer2Name.Visible = True
        lblComputerName.Visible = False
    End If
End Sub

Private Sub Form_Activate()
    txtPlayer1Name.SetFocus
    txtPlayer2Name.Visible = True
    lblComputerName.Visible = False
End Sub

Private Sub OKButton_Click()
    Do While ((Left(txtPlayer1Name.Text, 1)) = " ")
        txtPlayer1Name.Text = Mid(txtPlayer1Name.Text, 2, Len(txtPlayer1Name.Text))
    Loop

    Do While ((Left(txtPlayer2Name.Text, 1)) = " ")
        txtPlayer2Name.Text = Mid(txtPlayer2Name.Text, 2, Len(txtPlayer2Name.Text))
    Loop

    If lblComputerName.Visible = True Then
        If txtPlayer1Name.Text = "" Then
            MsgBox "Enter a name for Player 1", vbCritical, "Setup"
        Else
            'Start player vs computer!!! yay
            blnTwoPlayer = False    'Means its player vs computer A.I. mode
            strPlayer(1) = "Computer AI"
            GoTo Start
        End If
    ElseIf txtPlayer2Name.Visible = True Then
        If txtPlayer1Name.Text = "" Then
            MsgBox "Enter a name for Player 1", vbCritical, "Setup"
        ElseIf txtPlayer2Name.Text = "" Then
            MsgBox "Enter a name for Player 2", vbCritical, "Setup"
        Else
            blnTwoPlayer = True
            strPlayer(1) = txtPlayer2Name.Text
Start:
            strPlayer(0) = txtPlayer1Name.Text
            Unload Me
            frmMainMenu.Hide


            intBestWord = 0
            intGameTimeElapsed = 0
            strGameSummary(0) = ""
            strGameSummary(1) = "Best word: N/A"
            strGameSummary(2) = ""
            strGameSummary(3) = ""
            strGameSummary(4) = ""
            strGameSummary(5) = ""
            
            frmGameOver.lblGameSummary.Caption = ""
            frmGameOver.lstPlayer(0).ListItems.Clear
            frmGameOver.lstPlayer(1).ListItems.Clear
            frmGameOver.Hide

            frmScrabble.Show
        End If
    End If

End Sub
