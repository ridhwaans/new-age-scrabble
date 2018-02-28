VERSION 5.00
Begin VB.Form frmControls 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5595
   ClientLeft      =   9540
   ClientTop       =   3825
   ClientWidth     =   3690
   Icon            =   "frmControls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmControls.frx":1CCA
   ScaleHeight     =   5595
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtControls 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Image imgReturn 
      Height          =   375
      Left            =   600
      Picture         =   "frmControls.frx":4562E
      Top             =   4680
      Width           =   2385
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim intRules As Integer
    Dim strSentence As String

    intRules = FreeFile
    Open App.Path & "\game_files\controls_screen\controls.txt" For Input As intRules

    Do
        Input #intRules, strSentence
        txtControls.Text = txtControls.Text + strSentence + vbNewLine
    Loop Until EOF(intRules)

    Close #intRules
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReturn.Picture = LoadPicture(App.Path & "\game_files\controls_screen\return_black.gif")
End Sub

Private Sub imgReturn_Click()
    Unload Me
    frmHelp.Show
End Sub

Private Sub imgReturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReturn.Picture = LoadPicture(App.Path & "\game_files\controls_screen\return_red.gif")
End Sub


Private Sub txtControls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReturn.Picture = LoadPicture(App.Path & "\game_files\controls_screen\return_black.gif")
End Sub
