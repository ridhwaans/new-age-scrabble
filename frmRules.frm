VERSION 5.00
Begin VB.Form frmRules 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7905
   ClientLeft      =   7380
   ClientTop       =   2190
   ClientWidth     =   8820
   Icon            =   "frmRules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRules.frx":1CCA
   ScaleHeight     =   7905
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRules 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Image imgReturn 
      Height          =   375
      Left            =   3120
      Picture         =   "frmRules.frx":18D52E
      Top             =   6840
      Width           =   2385
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim intRules As Integer
    Dim strSentence As String

    intRules = FreeFile
    Open App.Path & "\game_files\rules_screen\rules.txt" For Input As intRules

    Do
        Input #intRules, strSentence
        txtRules.Text = txtRules.Text + strSentence + vbNewLine
    Loop Until EOF(intRules)

    Close #intRules
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReturn.Picture = LoadPicture(App.Path & "\game_files\rules_screen\return_black.gif")
End Sub

Private Sub imgReturn_Click()
    Unload Me
    frmHelp.Show
End Sub

Private Sub imgReturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReturn.Picture = LoadPicture(App.Path & "\game_files\rules_screen\return_red.gif")
End Sub

Private Sub txtRules_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReturn.Picture = LoadPicture(App.Path & "\game_files\rules_screen\return_black.gif")
End Sub
