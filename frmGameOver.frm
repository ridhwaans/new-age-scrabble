VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGameOver 
   BorderStyle     =   0  'None
   Caption         =   "Game Over"
   ClientHeight    =   8985
   ClientLeft      =   5535
   ClientTop       =   2160
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGameOver.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lstPlayer 
      Height          =   3495
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   2400
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Word Played"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Score"
         Object.Width           =   3440
      EndProperty
   End
   Begin MSComctlLib.ListView lstPlayer 
      Height          =   3495
      Index           =   1
      Left            =   8400
      TabIndex        =   4
      Top             =   2400
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Word Played"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Score"
         Object.Width           =   3440
      EndProperty
   End
   Begin VB.Image imgQuit 
      Height          =   570
      Left            =   5520
      Picture         =   "frmGameOver.frx":18B862
      Top             =   7920
      Width           =   2310
   End
   Begin VB.Label lblSummary 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lblGameSummary 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   5160
      TabIndex        =   6
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Image imgNewGame 
      Height          =   1215
      Left            =   6960
      Picture         =   "frmGameOver.frx":18BAB0
      Top             =   6480
      Width           =   2325
   End
   Begin VB.Image imgMainMenu 
      Height          =   1215
      Left            =   4080
      Picture         =   "frmGameOver.frx":18BEE3
      Top             =   6480
      Width           =   2325
   End
   Begin VB.Label lblPlayerScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   1
      Left            =   9360
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblPlayerScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblPlayer2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8760
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label lblPlayer1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
   End
End
Attribute VB_Name = "frmGameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowScrollBar Lib "user32" _
                                       (ByVal hwnd As Long, ByVal wBar As Long, _
                                        ByVal bShow As Long) As Long


Private Sub Form_Load()
    lblPlayer1.Caption = strPlayer(0)
    lblPlayer2.Caption = strPlayer(1)
    ShowScrollBar lstPlayer(0).hwnd, 0, False
    ShowScrollBar lstPlayer(1).hwnd, 0, False

    lblGameSummary.Caption = ""
    strGameSummary(1) = "Best word: N/A"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMainMenu.Picture = LoadPicture(App.Path & "\game_files\gameover_screen\main_menu_black.gif")
    imgNewGame.Picture = LoadPicture(App.Path & "\game_files\gameover_screen\new_game_black.gif")
    imgQuit.Picture = LoadPicture(App.Path & "\game_files\gameover_screen\quit_black.gif")
End Sub

Private Sub imgMainMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub imgNewGame_Click()
    frmSetup.Show vbModal
End Sub

Private Sub imgQuit_Click()
    Dim intDecision As Integer

    intDecision = MsgBox("Quit New-Age Scrabble?", vbYesNo + vbQuestion, "Quit")
    If intDecision = vbYes Then
        End
    End If
End Sub

Private Sub imgMainMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMainMenu.Picture = LoadPicture(App.Path & "\game_files\gameover_screen\main_menu_red.gif")
End Sub

Private Sub imgNewGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgNewGame.Picture = LoadPicture(App.Path & "\game_files\gameover_screen\new_game_red.gif")
End Sub

Private Sub imgQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgQuit.Picture = LoadPicture(App.Path & "\game_files\gameover_screen\quit_red.gif")
End Sub
