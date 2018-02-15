VERSION 5.00
Begin VB.Form frmWebsite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vintage Gaming"
   ClientHeight    =   10920
   ClientLeft      =   -30
   ClientTop       =   360
   ClientWidth     =   16800
   Icon            =   "frmWebsite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   728
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1120
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picholder 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      Height          =   10695
      Left            =   0
      ScaleHeight     =   709
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1101
      TabIndex        =   2
      Top             =   0
      Width           =   16575
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   10695
         Left            =   0
         Picture         =   "frmWebsite.frx":1CCA
         ScaleHeight     =   10695
         ScaleWidth      =   16575
         TabIndex        =   3
         Top             =   0
         Width           =   16575
         Begin VB.Label lblContact 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   13320
            TabIndex        =   8
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblBlog 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   12360
            TabIndex        =   7
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblTrial 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   11520
            TabIndex        =   6
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lblAbout 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   10440
            TabIndex        =   5
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblHome 
            BackStyle       =   0  'Transparent
            Height          =   495
            Left            =   9240
            TabIndex        =   4
            Top             =   720
            Width           =   975
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   10935
      Left            =   16560
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   10680
      Width           =   16575
   End
End
Attribute VB_Name = "frmWebsite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RECT
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Ori_Pos As RECT, Tmp_Pos As RECT, Tmp_Sel As RECT, Ori_Sel As RECT
Attribute Tmp_Pos.VB_VarUserMemId = 1073938432
Attribute Tmp_Sel.VB_VarUserMemId = 1073938432
Attribute Ori_Sel.VB_VarUserMemId = 1073938432

Private Sub Form_Load()
    With pic
        .ScaleMode = vbPixels
        .AutoRedraw = True
        .AutoSize = True
        .Picture = LoadPicture(App.Path & "\game_files\website_screen\home_webpage.bmp")
        ScrollAdjust
        Ori_Pos.Top = .Top
        Ori_Pos.Left = .Left
        Ori_Pos.Height = .Height
        Ori_Pos.Width = .Width
    End With

    HScroll1.Visible = True
    VScroll1.Visible = True
End Sub


Private Sub HScroll1_Change()
    pic.Left = -HScroll1.value
End Sub

Private Sub lblAbout_Click()
    pic.Picture = LoadPicture(App.Path & "\game_files\website_screen\about_webpage.bmp")
End Sub

Private Sub lblBlog_Click()
    pic.Picture = LoadPicture(App.Path & "\game_files\website_screen\blog_webpage.bmp")
End Sub

Private Sub lblContact_Click()
    pic.Picture = LoadPicture(App.Path & "\game_files\website_screen\contact_webpage.bmp")
End Sub

Private Sub lblHome_Click()
    pic.Picture = LoadPicture(App.Path & "\game_files\website_screen\home_webpage.bmp")
End Sub

Private Sub VScroll1_Change()
    pic.Top = -VScroll1.value
End Sub

Private Sub ScrollAdjust()
    With pic
        If .ScaleWidth < Picholder.ScaleWidth Then
            .Left = (Picholder.Width - .Width) / 2
            HScroll1.Enabled = False
        Else
            HScroll1.Min = 0
            HScroll1.Max = .Width - Picholder.Width
            HScroll1.SmallChange = 1
            HScroll1.LargeChange = 10
            HScroll1.Enabled = True
        End If
        If .ScaleHeight < Picholder.ScaleHeight Then
            .Top = (Picholder.Height - .Height) / 2
            VScroll1.Enabled = False
        Else
            VScroll1.Min = 0
            VScroll1.Max = .Height - Picholder.Height
            VScroll1.SmallChange = 50
            VScroll1.LargeChange = 100
            VScroll1.Enabled = True
        End If
    End With
End Sub

