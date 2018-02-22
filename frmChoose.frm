VERSION 5.00
Begin VB.Form frmChoose 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4815
   ClientLeft      =   8265
   ClientTop       =   4485
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoose.frx":0000
   ScaleHeight     =   4815
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblChoose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a letter"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   25
      Left            =   5520
      Picture         =   "frmChoose.frx":D300
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   24
      Left            =   4800
      Picture         =   "frmChoose.frx":E15B
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   23
      Left            =   4080
      Picture         =   "frmChoose.frx":EFE3
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   22
      Left            =   3360
      Picture         =   "frmChoose.frx":FE47
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   21
      Left            =   2640
      Picture         =   "frmChoose.frx":10C93
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   20
      Left            =   1920
      Picture         =   "frmChoose.frx":11A92
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   19
      Left            =   1200
      Picture         =   "frmChoose.frx":128E0
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   18
      Left            =   4800
      Picture         =   "frmChoose.frx":13752
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   17
      Left            =   4080
      Picture         =   "frmChoose.frx":145EE
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   16
      Left            =   3360
      Picture         =   "frmChoose.frx":1545B
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   15
      Left            =   2640
      Picture         =   "frmChoose.frx":162B0
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   14
      Left            =   1920
      Picture         =   "frmChoose.frx":17086
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   13
      Left            =   1200
      Picture         =   "frmChoose.frx":17EFF
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   12
      Left            =   4800
      Picture         =   "frmChoose.frx":18D81
      Top             =   2040
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   11
      Left            =   4080
      Picture         =   "frmChoose.frx":19BF9
      Top             =   2040
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   10
      Left            =   3360
      Picture         =   "frmChoose.frx":1AA87
      Top             =   2040
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   9
      Left            =   2640
      Picture         =   "frmChoose.frx":1B8D2
      Top             =   2040
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   8
      Left            =   1920
      Picture         =   "frmChoose.frx":1C703
      Top             =   2040
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   7
      Left            =   1200
      Picture         =   "frmChoose.frx":1D524
      Top             =   2040
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   6
      Left            =   4800
      Picture         =   "frmChoose.frx":1E3B5
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   5
      Left            =   4080
      Picture         =   "frmChoose.frx":1F266
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   4
      Left            =   3360
      Picture         =   "frmChoose.frx":200B2
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   3
      Left            =   2640
      Picture         =   "frmChoose.frx":20F63
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   2
      Left            =   1920
      Picture         =   "frmChoose.frx":21E02
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   1
      Left            =   1200
      Picture         =   "frmChoose.frx":22C47
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image imgLetter 
      Height          =   780
      Index           =   0
      Left            =   480
      Picture         =   "frmChoose.frx":23A77
      Top             =   1200
      Width           =   720
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgLetter_Click(Index As Integer)
    intTile = Index
    Select Case Index
    Case 0
        strTilePath = "\game_files\word_tiles\tile_A"
    Case 1
        strTilePath = "\game_files\word_tiles\tile_B"
    Case 2
        strTilePath = "\game_files\word_tiles\tile_C"
    Case 3
        strTilePath = "\game_files\word_tiles\tile_D"
    Case 4
        strTilePath = "\game_files\word_tiles\tile_E"
    Case 5
        strTilePath = "\game_files\word_tiles\tile_F"
    Case 6
        strTilePath = "\game_files\word_tiles\tile_G"
    Case 7
        strTilePath = "\game_files\word_tiles\tile_H"
    Case 8
        strTilePath = "\game_files\word_tiles\tile_I"
    Case 9
        strTilePath = "\game_files\word_tiles\tile_J"
    Case 10
        strTilePath = "\game_files\word_tiles\tile_K"
    Case 11
        strTilePath = "\game_files\word_tiles\tile_L"
    Case 12
        strTilePath = "\game_files\word_tiles\tile_M"
    Case 13
        strTilePath = "\game_files\word_tiles\tile_N"
    Case 14
        strTilePath = "\game_files\word_tiles\tile_O"
    Case 15
        strTilePath = "\game_files\word_tiles\tile_P"
    Case 16
        strTilePath = "\game_files\word_tiles\tile_Q"
    Case 17
        strTilePath = "\game_files\word_tiles\tile_R"
    Case 18
        strTilePath = "\game_files\word_tiles\tile_S"
    Case 19
        strTilePath = "\game_files\word_tiles\tile_T"
    Case 20
        strTilePath = "\game_files\word_tiles\tile_U"
    Case 21
        strTilePath = "\game_files\word_tiles\tile_V"
    Case 22
        strTilePath = "\game_files\word_tiles\tile_W"
    Case 23
        strTilePath = "\game_files\word_tiles\tile_X"
    Case 24
        strTilePath = "\game_files\word_tiles\tile_Y"
    Case 25
        strTilePath = "\game_files\word_tiles\tile_Z"
    End Select
    frmChoose.Hide
End Sub
