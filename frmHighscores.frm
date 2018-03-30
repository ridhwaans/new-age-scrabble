VERSION 5.00
Begin VB.Form frmHighscores 
   ClientHeight    =   10110
   ClientLeft      =   8685
   ClientTop       =   450
   ClientWidth     =   6735
   Icon            =   "frmHighscores.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmHighscores.frx":1CCA
   ScaleHeight     =   10110
   ScaleWidth      =   6735
   Begin VB.ListBox lstHighscores 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5250
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Image imgReturn 
      Height          =   570
      Left            =   1560
      Picture         =   "frmHighscores.frx":E09E6
      Top             =   7920
      Width           =   3450
   End
End
Attribute VB_Name = "frmHighscores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                                     (ByVal hwnd As Long, ByVal wMsg As Long, _
                                      ByVal wParam As Long, lParam As Any) As Long
Private Const LB_SETTABSTOPS = &H192

Private Sub Form_Load()
    ReDim TabStop(0 To 2) As Long

    TabStop(0) = 60
    TabStop(1) = 60

    intHighScores = FreeFile
    Open App.Path & "\game_files\highscores_screen\highscores.txt" For Input As intHighScores

    Do
        Input #intHighScores, strname, intScore

        Call SendMessage(lstHighscores.hwnd, LB_SETTABSTOPS, 0&, ByVal 0&)
        Call SendMessage(lstHighscores.hwnd, LB_SETTABSTOPS, 3, TabStop(0))
        lstHighscores.AddItem strname & Chr(9) & intScore
        lstHighscores.Refresh
    Loop Until EOF(intHighScores)

    Close #intHighScores
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReturn.Picture = LoadPicture(App.Path & "\game_files\highscores_screen\return_black.gif")
End Sub

Private Sub imgReturn_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub imgReturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReturn.Picture = LoadPicture(App.Path & "\game_files\highscores_screen\return_red.gif")
End Sub
