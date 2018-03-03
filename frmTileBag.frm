VERSION 5.00
Begin VB.Form frmTileBag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tile Bag"
   ClientHeight    =   3540
   ClientLeft      =   8610
   ClientTop       =   8820
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTileBag.frx":0000
   ScaleHeight     =   3540
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmTileBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strDisplay As String

Private Sub Form_Load()
    For intCount = 0 To 8
        strDisplay = strDisplay + strTileValue(intCount) + " x" + Str(intTilesRemaining(intCount)) + vbNewLine
    Next intCount

    lblDisplay(0).Caption = strDisplay
    strDisplay = ""

    For intCount = 9 To 17
        strDisplay = strDisplay + strTileValue(intCount) + " x" + Str(intTilesRemaining(intCount)) + vbNewLine
    Next intCount

    lblDisplay(1).Caption = strDisplay
    strDisplay = ""

    For intCount = 18 To 26
        If intCount = 26 Then
            strDisplay = strDisplay + "Blank x" + Str(intTilesRemaining(intCount)) + vbNewLine
        Else
            strDisplay = strDisplay + strTileValue(intCount) + " x" + Str(intTilesRemaining(intCount)) + vbNewLine
        End If
    Next intCount

    lblDisplay(2).Caption = strDisplay
    strDisplay = ""
End Sub
