VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   8055
   ClientLeft      =   7590
   ClientTop       =   540
   ClientWidth     =   8970
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":1CCA
   ScaleHeight     =   8055
   ScaleWidth      =   8970
   Begin VB.Image imgHighscores 
      Height          =   585
      Left            =   1440
      Picture         =   "frmMainMenu.frx":18D52E
      Top             =   4560
      Width           =   6000
   End
   Begin VB.Image imgQuit 
      Height          =   585
      Left            =   3120
      Picture         =   "frmMainMenu.frx":18DAD4
      Top             =   6840
      Width           =   2550
   End
   Begin VB.Image imgHelp 
      Height          =   645
      Left            =   3120
      Picture         =   "frmMainMenu.frx":18DD63
      Top             =   5640
      Width           =   2550
   End
   Begin VB.Image imgPlay 
      Height          =   585
      Left            =   3120
      Picture         =   "frmMainMenu.frx":18DFFC
      Top             =   3480
      Width           =   2550
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
End Sub

Private Sub imgHelp_Click()
    Unload Me
    frmHelp.Show
End Sub

Private Sub imgHighscores_Click()
    Unload Me
    frmHighscores.Show
End Sub

Private Sub imgPlay_Click()
    frmSetup.Show vbModal
End Sub

Private Sub imgPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
    imgPlay.Picture = LoadPicture(App.Path & "\game_files\main_menu\play_red.gif")
End Sub

Private Sub imgHighscores_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
    imgHighscores.Picture = LoadPicture(App.Path & "\game_files\main_menu\highscores_red.gif")
End Sub

Private Sub imgHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
    imgHelp.Picture = LoadPicture(App.Path & "\game_files\main_menu\help_red.gif")
End Sub

Private Sub imgQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
    imgQuit.Picture = LoadPicture(App.Path & "\game_files\main_menu\quit_red.gif")
End Sub

Private Sub imgQuit_Click()
    Dim intDecision As Integer

    intDecision = MsgBox("Quit New-Age Scrabble?", vbYesNo + vbQuestion, "Quit")
    If intDecision = vbYes Then
        End
    End If
End Sub

Private Sub ResetButtons()
    imgPlay.Picture = LoadPicture(App.Path & "\game_files\main_menu\play_black.gif")
    imgHighscores.Picture = LoadPicture(App.Path & "\game_files\main_menu\highscores_black.gif")
    imgHelp.Picture = LoadPicture(App.Path & "\game_files\main_menu\help_black.gif")
    imgQuit.Picture = LoadPicture(App.Path & "\game_files\main_menu\quit_black.gif")
End Sub
