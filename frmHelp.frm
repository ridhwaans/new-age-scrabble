VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6720
   ClientLeft      =   9270
   ClientTop       =   3375
   ClientWidth     =   4500
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHelp.frx":1CCA
   ScaleHeight     =   6720
   ScaleWidth      =   4500
   Begin VB.Image imgReturn 
      Height          =   375
      Left            =   1080
      Picture         =   "frmHelp.frx":64B16
      Top             =   4920
      Width           =   2385
   End
   Begin VB.Image imgCredits 
      Height          =   420
      Left            =   840
      Picture         =   "frmHelp.frx":64D1B
      Top             =   3960
      Width           =   2970
   End
   Begin VB.Image imgControls 
      Height          =   420
      Left            =   720
      Picture         =   "frmHelp.frx":64F91
      Top             =   3000
      Width           =   3195
   End
   Begin VB.Image imgRules 
      Height          =   375
      Left            =   1320
      Picture         =   "frmHelp.frx":6524A
      Top             =   2040
      Width           =   1980
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
End Sub

Private Sub imgControls_Click()
    Unload Me
    frmControls.Show
End Sub

Private Sub imgCredits_Click()
    frmAbout.Show vbModal
End Sub

Private Sub imgRules_Click()
    Unload Me
    frmRules.Show
End Sub

Private Sub imgRules_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
    imgRules.Picture = LoadPicture(App.Path & "\game_files\help_menu\rules_red.gif")
End Sub

Private Sub imgControls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
    imgControls.Picture = LoadPicture(App.Path & "\game_files\help_menu\controls_red.gif")
End Sub

Private Sub imgCredits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
    imgCredits.Picture = LoadPicture(App.Path & "\game_files\help_menu\credits_red.gif")
End Sub

Private Sub imgReturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
    imgReturn.Picture = LoadPicture(App.Path & "\game_files\help_menu\return_red.gif")
End Sub

Private Sub imgReturn_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub ResetButtons()
    imgRules.Picture = LoadPicture(App.Path & "\game_files\help_menu\rules_black.gif")
    imgControls.Picture = LoadPicture(App.Path & "\game_files\help_menu\controls_black.gif")
    imgCredits.Picture = LoadPicture(App.Path & "\game_files\help_menu\credits_black.gif")
    imgReturn.Picture = LoadPicture(App.Path & "\game_files\help_menu\return_black.gif")
End Sub

