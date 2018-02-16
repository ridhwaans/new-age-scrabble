VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLoading 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4500
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLoading.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoading.frx":000C
   ScaleHeight     =   4500
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pgbLoading 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   3240
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   0
      Top             =   2640
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    tmrLoad.Enabled = True
End Sub

Private Sub tmrLoad_Timer()
    pgbLoading.value = pgbLoading.value + 1
    If pgbLoading.value = 100 Then    'When the loading bar reaches 100% based on the 25ms timer, go to main menu
        tmrLoad.Enabled = False
        Unload Me
        frmMainMenu.Show
    End If
End Sub

