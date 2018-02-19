VERSION 5.00
Begin VB.Form frmPauseMenu 
   BorderStyle     =   0  'None
   Caption         =   "Paused"
   ClientHeight    =   6735
   ClientLeft      =   10905
   ClientTop       =   3000
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPauseMenu.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgQuit 
      Height          =   585
      Left            =   960
      Picture         =   "frmPauseMenu.frx":62E4C
      Top             =   4800
      Width           =   2550
   End
   Begin VB.Image imgRestart 
      Height          =   570
      Left            =   120
      Picture         =   "frmPauseMenu.frx":630DB
      Top             =   3480
      Width           =   4185
   End
   Begin VB.Image imgResume 
      Height          =   570
      Left            =   480
      Picture         =   "frmPauseMenu.frx":634EE
      Top             =   2160
      Width           =   3585
   End
End
Attribute VB_Name = "frmPauseMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
End Sub

Private Sub imgRestart_Click()
    Dim intDecision As Integer

    intDecision = MsgBox("End game and start new game?", vbYesNo + vbQuestion, "Restart Game")
    If intDecision = vbYes Then
        Unload Me
        frmDictionary.Hide
        Unload frmScrabble
        frmScrabble.Show
    End If
End Sub

Private Sub imgResume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
    imgResume.Picture = LoadPicture(App.Path & "\game_files\pause_menu\resume_red.gif")
End Sub

Private Sub imgRestart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
    imgRestart.Picture = LoadPicture(App.Path & "\game_files\pause_menu\restart_red.gif")
End Sub

Private Sub imgQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
    imgQuit.Picture = LoadPicture(App.Path & "\game_files\pause_menu\quit_red.gif")
End Sub

Private Sub imgQuit_Click()
    Dim intDecision As Integer

    intDecision = MsgBox("End the game?", vbYesNo + vbQuestion, "Quit Game")
    If intDecision = vbYes Then

        Call EndGame

        Unload Me
        frmDictionary.Hide
        Unload frmScrabble
        frmGameOver.Show
    End If
End Sub

Private Sub imgResume_Click() 'Resume game
    frmScrabble.tmrTimeElapsed.Enabled = True
    frmScrabble.cmdPause.Picture = LoadPicture(App.Path & "\game_files\game_screen\pause_black.bmp")
    Unload Me
End Sub

Private Sub ResetButtons()
    imgResume.Picture = LoadPicture(App.Path & "\game_files\pause_menu\resume_black.gif")
    imgRestart.Picture = LoadPicture(App.Path & "\game_files\pause_menu\restart_black.gif")
    imgQuit.Picture = LoadPicture(App.Path & "\game_files\pause_menu\quit_black.gif")
End Sub

