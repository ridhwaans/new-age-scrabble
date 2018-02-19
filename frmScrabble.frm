VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10d.ocx"
Begin VB.Form frmScrabble 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New-Age Scrabble"
   ClientHeight    =   12390
   ClientLeft      =   4245
   ClientTop       =   375
   ClientWidth     =   15630
   Icon            =   "frmScrabble.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmScrabble.frx":1CCA
   ScaleHeight     =   12390
   ScaleWidth      =   15630
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swfComputing 
      Height          =   735
      Left            =   2040
      TabIndex        =   15
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
      _cx             =   1296
      _cy             =   1296
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.CommandButton cmdSubmit 
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      Picture         =   "frmScrabble.frx":F694
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10920
      Width           =   1695
   End
   Begin VB.CommandButton cmdPause 
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      Picture         =   "frmScrabble.frx":12FEE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   11640
      Width           =   735
   End
   Begin VB.CommandButton cmdRecall 
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Picture         =   "frmScrabble.frx":173C2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Timer tmrTimeElapsed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   11760
   End
   Begin VB.CommandButton cmdPass 
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   3960
      Picture         =   "frmScrabble.frx":1DC14
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   11640
      Width           =   735
   End
   Begin VB.CommandButton cmdExchange 
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      Picture         =   "frmScrabble.frx":21FE8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdShuffle 
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      Picture         =   "frmScrabble.frx":2883A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton cmdDictionary 
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Picture         =   "frmScrabble.frx":2F08C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Turn time left"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   11520
      Width           =   1695
   End
   Begin VB.Label lblTilesRemaining 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   11040
      Width           =   495
   End
   Begin VB.Image imgBag 
      Height          =   720
      Left            =   3960
      Picture         =   "frmScrabble.frx":358DE
      Top             =   10800
      Width           =   720
   End
   Begin VB.Label lblMessages 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Scrabble!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   6240
      Width           =   4695
   End
   Begin VB.Label lblPlayer2Score 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblPlayer1Score 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblPlayer2Name 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label lblPlayer1Name 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label lblTimeElapsed 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3:00"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   11880
      Width           =   1575
   End
   Begin VB.Image imgLetterTile 
      Height          =   780
      Index           =   6
      Left            =   3120
      Top             =   9960
      Width           =   720
   End
   Begin VB.Image imgLetterTile 
      Height          =   780
      Index           =   5
      Left            =   2400
      Top             =   9960
      Width           =   720
   End
   Begin VB.Image imgLetterTile 
      Height          =   780
      Index           =   4
      Left            =   1680
      Top             =   9960
      Width           =   720
   End
   Begin VB.Image imgLetterTile 
      Height          =   780
      Index           =   3
      Left            =   960
      Top             =   9960
      Width           =   720
   End
   Begin VB.Image imgLetterTile 
      Height          =   780
      Index           =   2
      Left            =   2760
      Top             =   9240
      Width           =   720
   End
   Begin VB.Image imgLetterTile 
      Height          =   780
      Index           =   1
      Left            =   2040
      Top             =   9240
      Width           =   720
   End
   Begin VB.Image imgLetterTile 
      Height          =   780
      Index           =   0
      Left            =   1320
      Top             =   9240
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   224
      Left            =   14880
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   223
      Left            =   14160
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   222
      Left            =   13440
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   221
      Left            =   12720
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   220
      Left            =   12000
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   219
      Left            =   11280
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   218
      Left            =   10560
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   217
      Left            =   9840
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   216
      Left            =   9120
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   215
      Left            =   8400
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   214
      Left            =   7680
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   213
      Left            =   6960
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   212
      Left            =   6240
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   211
      Left            =   5520
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   210
      Left            =   4800
      Top             =   11600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   209
      Left            =   14880
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   208
      Left            =   14160
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   207
      Left            =   13440
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   206
      Left            =   12720
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   205
      Left            =   12000
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   204
      Left            =   11280
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   203
      Left            =   10560
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   202
      Left            =   9840
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   201
      Left            =   9120
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   200
      Left            =   8400
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   199
      Left            =   7680
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   198
      Left            =   6960
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   197
      Left            =   6240
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   196
      Left            =   5520
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   195
      Left            =   4800
      Top             =   10770
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   194
      Left            =   14880
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   193
      Left            =   14160
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   192
      Left            =   13440
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   191
      Left            =   12720
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   190
      Left            =   12000
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   189
      Left            =   11280
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   188
      Left            =   10560
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   187
      Left            =   9840
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   186
      Left            =   9120
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   185
      Left            =   8400
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   184
      Left            =   7680
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   183
      Left            =   6960
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   182
      Left            =   6240
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   181
      Left            =   5520
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   180
      Left            =   4800
      Top             =   9900
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   179
      Left            =   14880
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   178
      Left            =   14160
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   177
      Left            =   13440
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   176
      Left            =   12720
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   780
      Index           =   175
      Left            =   12000
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   174
      Left            =   11280
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   173
      Left            =   10560
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   172
      Left            =   9840
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   171
      Left            =   9120
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   170
      Left            =   8400
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   169
      Left            =   7680
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   168
      Left            =   6960
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   167
      Left            =   6240
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   166
      Left            =   5520
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   165
      Left            =   4800
      Top             =   9100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   164
      Left            =   14880
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   163
      Left            =   14160
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   162
      Left            =   13440
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   161
      Left            =   12720
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   160
      Left            =   12000
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   159
      Left            =   11280
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   158
      Left            =   10560
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   157
      Left            =   9840
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   156
      Left            =   9120
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   155
      Left            =   8400
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   154
      Left            =   7680
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   153
      Left            =   6960
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   152
      Left            =   6240
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   151
      Left            =   5520
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   150
      Left            =   4800
      Top             =   8270
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   149
      Left            =   14880
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   148
      Left            =   14160
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   147
      Left            =   13440
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   146
      Left            =   12720
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   145
      Left            =   12000
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   144
      Left            =   11280
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   143
      Left            =   10560
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   142
      Left            =   9840
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   141
      Left            =   9120
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   140
      Left            =   8400
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   139
      Left            =   7680
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   138
      Left            =   6960
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   137
      Left            =   6240
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   136
      Left            =   5520
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   135
      Left            =   4800
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   134
      Left            =   14880
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   133
      Left            =   14160
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   132
      Left            =   13440
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   131
      Left            =   12720
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   130
      Left            =   12000
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   129
      Left            =   11280
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   128
      Left            =   10560
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   127
      Left            =   9840
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   126
      Left            =   9120
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   125
      Left            =   8400
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   124
      Left            =   7680
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   123
      Left            =   6960
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   122
      Left            =   6240
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   121
      Left            =   5520
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   120
      Left            =   4800
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   119
      Left            =   14880
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   118
      Left            =   14160
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   117
      Left            =   13440
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   116
      Left            =   12720
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   115
      Left            =   12000
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   114
      Left            =   11280
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   113
      Left            =   10560
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   112
      Left            =   9840
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   111
      Left            =   9120
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   110
      Left            =   8400
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   109
      Left            =   7680
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   108
      Left            =   6960
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   107
      Left            =   6240
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   106
      Left            =   5520
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   105
      Left            =   4800
      Top             =   5780
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   104
      Left            =   14880
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   103
      Left            =   14160
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   102
      Left            =   13440
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   101
      Left            =   12720
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   100
      Left            =   12000
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   99
      Left            =   11280
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   98
      Left            =   10560
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   97
      Left            =   9840
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   96
      Left            =   9120
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   95
      Left            =   8400
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   94
      Left            =   7680
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   93
      Left            =   6960
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   92
      Left            =   6240
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   91
      Left            =   5520
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   90
      Left            =   4800
      Top             =   4970
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   89
      Left            =   14880
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   88
      Left            =   14160
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   87
      Left            =   13440
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   86
      Left            =   12720
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   85
      Left            =   12000
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   84
      Left            =   11280
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   83
      Left            =   10560
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   82
      Left            =   9840
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   81
      Left            =   9120
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   80
      Left            =   8400
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   79
      Left            =   7680
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   78
      Left            =   6960
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   77
      Left            =   6240
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   76
      Left            =   5520
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   75
      Left            =   4800
      Top             =   4100
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   74
      Left            =   14880
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   73
      Left            =   14160
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   72
      Left            =   13440
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   71
      Left            =   12720
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   70
      Left            =   12000
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   69
      Left            =   11280
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   68
      Left            =   10560
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   67
      Left            =   9840
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   66
      Left            =   9120
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   65
      Left            =   8400
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   64
      Left            =   7680
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   63
      Left            =   6960
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   62
      Left            =   6240
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   61
      Left            =   5520
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   60
      Left            =   4800
      Top             =   3290
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   59
      Left            =   14880
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   58
      Left            =   14160
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   57
      Left            =   13440
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   56
      Left            =   12720
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   55
      Left            =   12000
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   54
      Left            =   11280
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   53
      Left            =   10560
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   52
      Left            =   9840
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   51
      Left            =   9120
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   50
      Left            =   8400
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   49
      Left            =   7680
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   48
      Left            =   6960
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   47
      Left            =   6240
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   46
      Left            =   5520
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   45
      Left            =   4800
      Top             =   2450
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   44
      Left            =   14880
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   43
      Left            =   14160
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   42
      Left            =   13440
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   41
      Left            =   12720
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   40
      Left            =   12000
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   39
      Left            =   11280
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   38
      Left            =   10560
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   37
      Left            =   9840
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   36
      Left            =   9120
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   35
      Left            =   8400
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   34
      Left            =   7680
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   33
      Left            =   6960
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   32
      Left            =   6240
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   31
      Left            =   5520
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   30
      Left            =   4800
      Top             =   1570
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   29
      Left            =   14880
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   28
      Left            =   14160
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   27
      Left            =   13440
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   26
      Left            =   12720
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   25
      Left            =   12000
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   24
      Left            =   11280
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   23
      Left            =   10560
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   22
      Left            =   9840
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   21
      Left            =   9120
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   20
      Left            =   8400
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   19
      Left            =   7680
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   18
      Left            =   6960
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   17
      Left            =   6240
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   16
      Left            =   5520
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   15
      Left            =   4800
      Top             =   760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   14
      Left            =   14880
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   13
      Left            =   14160
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   12
      Left            =   13440
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   11
      Left            =   12720
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   10
      Left            =   12000
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   9
      Left            =   11280
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   8
      Left            =   10560
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   7
      Left            =   9840
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   6
      Left            =   9120
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   5
      Left            =   8400
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   4
      Left            =   7680
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   3
      Left            =   6960
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   2
      Left            =   6240
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   810
      Index           =   1
      Left            =   5520
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   780
      Index           =   0
      Left            =   4800
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgBoardGrid 
      Height          =   12600
      Left            =   4800
      Picture         =   "frmScrabble.frx":3617A
      Top             =   -120
      Width           =   10845
   End
End
Attribute VB_Name = "frmScrabble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDictionary_Click()
    If frmScrabble.MouseIcon = 0 Then    'If cursor is not holding a tile, you can refer to the dictionary
        frmDictionary.Show vbModal
    End If
End Sub

Private Sub cmdDictionary_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetButtons    'Reset the button to the color black if the mouse is not over the button
    cmdDictionary.Picture = LoadPicture(App.Path & "\game_files\game_screen\dictionary_red.bmp")    'Change dictionary button to red when mouse over
End Sub

Private Sub cmdExchange_Click()
    Dim intDecision As Integer
    Dim intTemporary(0 To 6) As Integer
    Dim intTemporary2(0 To 6) As Integer
    Dim random As Integer

    If blnTwoPlayer = False And intTurn = 1 Then    'If computer's turn then
        GoTo Exchange:
    End If

    If frmScrabble.MouseIcon = 0 Then    'If cursor is not holding a tile, you can exchange tiles
        intDecision = MsgBox("Exchange all seven rack tiles for new tiles?", vbYesNo + vbQuestion, "Exchange Tiles")
        If intDecision = vbYes Then
Exchange:
            cmdExchange.Enabled = False

            For intCount = 0 To 6    'Hold back existing tiles at hand
                intTemporary(intCount) = intPlayerTiles(intTurn, intCount)
                intTemporary2(intCount) = intTilesInPlay(intTurn, intCount)
            Next intCount

            For intCount = 0 To 6    'Get new tiles from bag
                Do
                    Randomize
                    random = Int(Rnd * 100)
                    intPlayerTiles(intTurn, intCount) = intTileBank(random)
                Loop Until (intTileBank(random) <> -1)    'Get a random letter so long as there is letter quantity remaining

                intTileBank(random) = -1
                intTilesInPlay(intTurn, intCount) = random

                intTilesRemaining(intPlayerTiles(intTurn, intCount)) = intTilesRemaining(intPlayerTiles(intTurn, intCount)) - 1    'Subtract one tile from bag

                Call IdentifyTile(intPlayerTiles(intTurn, intCount))    'Identifies letter based on number (example: A= 0, B= 1 etc)

                imgLetterTile(intCount).Picture = LoadPicture(App.Path & strTilePath & ".gif")    'Load the letter image on tile
            Next intCount

            For intCount = 0 To 6    'Return tiles at hand back to the tile bank
                intTileBank(intTemporary2(intCount)) = intTemporary(intCount)
            Next intCount

            For intCount = 0 To 6    'Return tiles at hand back to the bag
                intTilesRemaining(intTemporary(intCount)) = intTilesRemaining(intTemporary(intCount)) + 1
            Next intCount
        End If

        blnExchangeDisabled = True    'Exchange is allowed only once per turn
    End If
End Sub

Private Sub cmdExchange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetButtons    'Reset the button to the color black if the mouse is not over the button
    cmdExchange.Picture = LoadPicture(App.Path & "\game_files\game_screen\exchange_red.bmp")    'Change exchange button to red when mouse over
End Sub

Private Sub DisableRecall()
    Dim blnEnable As Boolean

    For intCount = 0 To 6    'Checks if all the seven player tiles are on the rack or on the board
        If imgLetterTile(intCount).Picture = 0 Then    'Rack tile is empty or not
            blnEnabled = True
        End If
    Next intCount

    'If at least one player tile is on the board while game in progress, Recall is allowed,
    'Else if player tile rack is full, Recall is disabled
    If blnEnabled = True Then
        cmdRecall.Enabled = True
        blnEnable = False
    Else
        cmdRecall.Enabled = False
    End If
End Sub
Private Sub DisableExchange()
    Dim blnDisable As Boolean

    For intCount = 0 To 6    'Checks if all the seven player tiles are on the rack or on the board
        If imgLetterTile(intCount).Picture = 0 Then
            blnDisable = True
        End If
    Next intCount

    'All seven player tiles must be on the rack and not on the gameboard to allow Exchange
    If blnDisable = True Then
        cmdExchange.Enabled = False
        cmdShuffle.Enabled = False
        blnDisable = False
    Else
        If blnExchangeDisabled = False Then
            cmdExchange.Enabled = True
        End If
        cmdShuffle.Enabled = True
    End If
End Sub

Private Sub cmdPass_Click()
    Dim intDecision As Integer
    If frmScrabble.MouseIcon = 0 Then    'If cursor is not holding a tile
        intDecision = MsgBox("Skip your turn?", vbYesNo + vbQuestion, "Pass")
        If intDecision = vbYes Then

            For intCount = 0 To 224
                If imgTile(intCount).Picture <> 0 Then    'Remove tiles placed on gameboard and return them to rack
                    If blnGridTilesPlayed(intCount) = False Then
                        imgTile(intCount).Picture = Nothing
                    End If
                End If
            Next intCount

            Call NewTurn
            cmdPass.Picture = LoadPicture(App.Path & "\game_files\game_screen\pass_black.bmp")
        End If
    End If
End Sub

Private Sub RecallTiles()
    If frmScrabble.MouseIcon = 0 Then    'If cursor is not holding a tile
        For intCount = 0 To 6
            If imgLetterTile(intCount).Picture = 0 Then
                Call IdentifyTile(intPlayerTiles(intTurn, intCount))    'Identifies letter based on letter number (example: A=0)
                imgLetterTile(intCount).Picture = LoadPicture(App.Path & strTilePath & ".gif")    'Return all seven player tiles to rack
            End If
        Next intCount

        For intCount = 0 To 224
            If imgTile(intCount).Picture <> 0 Then    'Remove tiles placed on gameboard and return them to rack
                If blnGridTilesPlayed(intCount) = False Then
                    imgTile(intCount).Picture = Nothing

                    If intCount = intBlankTile(0) Then
                        intBlankTile(0) = -1
                    ElseIf intCount = intBlankTile(1) Then
                        intBlankTile(1) = -1
                    End If

                End If
            End If
        Next intCount

        Call DisableExchange
        cmdRecall.Enabled = False
    End If
End Sub

Private Sub cmdRecall_Click()
    Call RecallTiles
End Sub

Private Sub cmdRecall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetButtons    'Reset buttons to black color
    cmdRecall.Picture = LoadPicture(App.Path & "\game_files\game_screen\recall_red.bmp")    'Change Recall button to red color if mouse over
End Sub

Private Sub cmdShuffle_Click()
    Dim i, j, tmp As Integer

    If frmScrabble.MouseIcon = 0 Then    'If cursor is not holding a tile
        Randomize
        For i = 0 To 6
            j = Int((7 - i) * Rnd + i)    'Shuffle all the elements in the array
            tmp = intPlayerTiles(intTurn, i)
            intPlayerTiles(intTurn, i) = intPlayerTiles(intTurn, j)
            intPlayerTiles(intTurn, j) = tmp
        Next i

        For intCount = 0 To 6
            Call IdentifyTile(intPlayerTiles(intTurn, intCount))
            imgLetterTile(intCount).Picture = LoadPicture(App.Path & strTilePath & ".gif")    'Reload player tile images after shuffling
        Next intCount
    End If
End Sub

Private Sub cmdShuffle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetButtons    'Change all buttons to black color
    cmdShuffle.Picture = LoadPicture(App.Path & "\game_files\game_screen\shuffle_red.bmp")    'Change Shuffle button to red color if mouse over
End Sub

Private Function crossCheck(ByRef rToken As Integer, ByRef cToken As Integer, ByRef numCounter As Integer, ByRef numLocation() As Integer, ByRef strMainWord) As Boolean    'Checks tile's cross direction for other words
    Dim strCrossWords(1 To 15) As String

    crossCheck = True    'Initially true

    If (rToken > cToken) Or ((rToken = 0) And (cToken = 0)) Then    'Horizontal word, check for vertical crosswords

        For intLoop = 1 To numCounter
            strCrossWords(intLoop) = strTileValue(intGridTiles(numLocation(intLoop)))
            'Up
            If (numLocation(intLoop) >= 15) Then
                If (imgTile(numLocation(intLoop) - 15).Picture <> 0) Then
                    For intCount = (numLocation(intLoop)) To (numLocation(intLoop) Mod 15) Step -15
                        If (intCount <> numLocation(intLoop)) Then
                            If (imgTile(intCount).Picture <> 0) Then
                                strCrossWords(intLoop) = strTileValue(intGridTiles(intCount)) + strCrossWords(intLoop)
                            Else
                                GoTo Exit_Up:
                            End If
                        End If
                    Next intCount
                End If
            End If

Exit_Up:
            'Down
            If (numLocation(intLoop) <= 209) Then
                If (imgTile(numLocation(intLoop) + 15).Picture <> 0) Then
                    For intCount = (numLocation(intLoop)) To ((numLocation(intLoop) Mod 15) + 210) Step 15
                        If (intCount <> numLocation(intLoop)) Then
                            If (imgTile(intCount).Picture <> 0) Then
                                strCrossWords(intLoop) = strCrossWords(intLoop) + strTileValue(intGridTiles(intCount))
                            Else
                                GoTo Exit_Down:
                            End If
                        End If
                    Next intCount
                End If
            End If

Exit_Down:

        Next intLoop

    ElseIf (rToken < cToken) Then    'Vertical word, check for horizontal crosswords

        For intLoop = 1 To numCounter
            strCrossWords(intLoop) = strTileValue(intGridTiles(numLocation(intLoop)))
            'Left
            If ((numLocation(intLoop) Mod 15) <> 0) Then
                If (imgTile(numLocation(intLoop) - 1).Picture <> 0) Then
                    For intCount = (numLocation(intLoop)) To (Int(numLocation(intLoop) / 15) * 15) Step -1
                        If (intCount <> numLocation(intLoop)) Then
                            If (imgTile(intCount).Picture <> 0) Then
                                strCrossWords(intLoop) = strTileValue(intGridTiles(intCount)) + strCrossWords(intLoop)
                            Else
                                GoTo Exit_Left:
                            End If
                        End If
                    Next intCount
                End If
            End If

Exit_Left:
            'Right
            If (((numLocation(intLoop) + 1) Mod 15) <> 0) Then
                If (imgTile(numLocation(intLoop) + 1).Picture <> 0) Then
                    For intCount = (numLocation(intLoop)) To ((Int(numLocation(intLoop) / 15) * 15) + 15) Step 1
                        If (intCount <> numLocation(intLoop)) Then
                            If (imgTile(intCount).Picture <> 0) Then
                                strCrossWords(intLoop) = strCrossWords(intLoop) + strTileValue(intGridTiles(intCount))
                            Else
                                GoTo Exit_Right:
                            End If
                        End If

                        If intCount = 224 Then
                            GoTo Exit_Right
                        End If
                    Next intCount
                End If
            End If

Exit_Right:

        Next intLoop

    End If

    For intCount = 1 To numCounter
        If (WordCheck(strCrossWords(intCount)) <> 0) And (Len(strCrossWords(intCount)) > 1) Then    'If the crossword is more than two letters long and not a valid word
            lblMessages.FontSize = 18
            lblMessages.Caption = strCrossWords(intCount) & " is INVALID" & vbNewLine & "Try again"
            crossCheck = False
        ElseIf (Len(strCrossWords(intCount)) > 1) And StrComp(strCrossWords(intCount), strMainWord, vbTextCompare) <> 0 Then
            Call ComputeCrossWordScore(numLocation(intCount), strCrossWords(intCount))
            intPlayerScore(intTurn) = intPlayerScore(intTurn) + intScore
            lblMessages.Caption = lblMessages.Caption & strCrossWords(intCount) & " " & intScore & ", "


            If intBestWord < intScore Then
                intBestWord = intScore
                strGameSummary(1) = "Best word: " & strCrossWords(intCount) & " " & Str(intBestWord)
            End If

            intWordsPlayed(intTurn) = intWordsPlayed(intTurn) + 1
            frmGameOver.lstPlayer(intTurn).ListItems.Add(1).Text = Str(intWordsPlayed(intTurn))
            frmGameOver.lstPlayer(intTurn).ListItems.Item(1).ListSubItems.Add.Text = strCrossWords(intCount)
            frmGameOver.lstPlayer(intTurn).ListItems.Item(1).ListSubItems.Add.Text = intScore
            frmGameOver.lblPlayerScore(intTurn).Caption = intPlayerScore(intTurn)
        End If
    Next intCount

End Function

Private Sub ComputeCrossWordScore(ByRef locationNum As Integer, ByRef crossWord As String)    'For the crosswords
    Dim value As Integer

    intScore = 0

    For intLoop = 1 To Len(crossWord)    'Traverse from the first letter of the word to the last letter
        Select Case (UCase((Mid(crossWord, intLoop, 1))))    'Determine the value of the letter
        Case " "
            value = 0
        Case "E", "A", "I", "O", "N", "R", "T", "L", "S", "U"
            value = 1
        Case "D", "G"
            value = 2
        Case "B", "C", "M", "P"
            value = 3
        Case "F", "H", "V", "W", "Y"
            value = 4
        Case "K"
            value = 5
        Case "J", "X"
            value = 8
        Case "Q", "Z"
            value = 10
        End Select

        If intLoop = 1 Then    'Once
            Select Case locationNum    'Determine if the letter is on a triple letter or double letter score tile
            Case 20, 24, 76, 80, 84, 88, 136, 140, 144, 148, 200, 204
                value = value * 3    'If it is a triple letter score tile, multiply the letter value by 3
            Case 3, 11, 36, 38, 45, 52, 59, 92, 96, 98, 102, 108, 116, 122, 126, 128, 132, 165, 172, 179, 186, 188, 213, 221
                value = value * 2
            End Select
        End If

        If locationNum = intBlankTile(0) Or locationNum = intBlankTile(1) Then
            value = 0    'Blank tile
        End If

        intScore = intScore + value    'Add the letter score to the word score
    Next intLoop


    Select Case locationNum    'Determine if the letter is on a triple word or double word score tile
    Case 0, 7, 14, 105, 119, 210, 217, 224
        intScore = intScore * 3    'Triple word score tile multiplies the score by 3
    Case 16, 28, 32, 42, 48, 56, 64, 70, 112, 154, 160, 168, 176, 182, 192, 196, 208    'centre tile (112) is also double word score tile
        intScore = intScore * 2
    End Select
End Sub


Private Sub cmdSubmit_Click()    'Most important procedure in the game

    lblMessages.Caption = ""

    If frmScrabble.MouseIcon <> 0 Then
        GoTo End_Action
    End If

    If cmdPass.Enabled = False Then
        Continue (True)
        GoTo New_Turn:
    End If

    Dim temp(0 To 224) As Boolean
    Dim counter As Integer
    Dim location(1 To 15) As Integer
    Dim strWord As String

    'This for loop finds all letters played on the gameboard during the player turn when the player submits his/her move
    For intCount = 0 To 224
        If imgTile(intCount).Picture <> 0 Then
            If blnGridTilesPlayed(intCount) = False Then    'If gameboard tile played is not locked
                temp(intCount) = True
                counter = counter + 1
                location(counter) = intCount    'Record the location the grid location of the letter tile played
            End If
        End If
    Next intCount

    If counter = 0 Then    'means no tile has been placed
        GoTo End_Action:
    End If

    If blnGridTilesPlayed(112) = False Then    'If the first word played is not through the central tile
        If temp(112) = False Then
            lblMessages.FontSize = 16
            lblMessages.Caption = "The first word must" & vbNewLine & "pass through the star"
            GoTo End_Action:
        End If
    End If


    'If there are tiles above and below, or left and right of the letter, the letter and token is added twice in the process of traversing both directions
    Dim rowToken As Integer
    Dim columnToken As Integer

    strWord = strTileValue(intGridTiles(location(1)))    'Initally, word has only one letter

    If ((location(counter)) - (location(1)) < 14) Then
        'Left
        If ((location(1) Mod 15) <> 0) Then
            If (imgTile(location(1) - 1).Picture <> 0) Then
                For intCount = (location(1)) To (Int(location(1) / 15) * 15) Step -1
                    If (intCount <> location(1)) Then
                        If (imgTile(intCount).Picture <> 0) Then
                            strWord = strTileValue(intGridTiles(intCount)) + strWord
                            If temp(intCount) = True Then
                                rowToken = rowToken + 1
                            End If
                        Else
                            GoTo Exit_Left:
                        End If
                    End If
                Next intCount
            End If
        End If

Exit_Left:

        'Right
        If (((location(1) + 1) Mod 15) <> 0) Then
            If (imgTile(location(1) + 1).Picture <> 0) Then
                For intCount = (location(1)) To ((Int(location(1) / 15) * 15) + 15) Step 1
                    If (intCount <> location(1)) Then
                        If (imgTile(intCount).Picture <> 0) Then
                            strWord = strWord + strTileValue(intGridTiles(intCount))
                            If temp(intCount) = True Then
                                rowToken = rowToken + 1
                            End If
                        Else
                            GoTo Exit_Right:
                        End If
                    End If

                    If intCount = 224 Then
                        GoTo Exit_Right
                    End If
                Next intCount
            End If
        End If

Exit_Right:
    End If

    If WordCheck(strWord) = 0 Or (Len(strWord) = 1) Then  'Now check vertically (former condition: if rowtoken = 0 then)
        If ((location(counter)) - (location(1)) >= 14) Or (Len(strWord) = 1) Then
            strWord = strTileValue(intGridTiles(location(1)))

            'Up
            If (location(1) >= 15) Then
                If (imgTile(location(1) - 15).Picture <> 0) Then
                    For intCount = (location(1)) To (location(1) Mod 15) Step -15
                        If (intCount <> location(1)) Then
                            If (imgTile(intCount).Picture <> 0) Then
                                strWord = strTileValue(intGridTiles(intCount)) + strWord
                                If temp(intCount) = True Then
                                    columnToken = columnToken + 1
                                End If
                            Else
                                GoTo Exit_Up:
                            End If
                        End If
                    Next intCount
                End If
            End If

Exit_Up:

            'Down
            If (location(1) <= 209) Then
                If (imgTile(location(1) + 15).Picture <> 0) Then
                    For intCount = (location(1)) To ((location(1) Mod 15) + 210) Step 15
                        If (intCount <> location(1)) Then
                            If (imgTile(intCount).Picture <> 0) Then
                                strWord = strWord + strTileValue(intGridTiles(intCount))
                                If temp(intCount) = True Then
                                    columnToken = columnToken + 1
                                End If
                            Else
                                GoTo Exit_Down:
                            End If
                        End If
                    Next intCount
                End If
            End If

Exit_Down:
        End If

    Else
        lblMessages.FontSize = 18
        lblMessages.Caption = strWord & " is INVALID" & vbNewLine & "Try again"
        GoTo End_Action:
    End If

    Dim blnNotIsolated As Boolean

    '+1 for central token
    If (counter <> rowToken + 1) And (counter <> columnToken + 1) Then    'If all player letter tiles are not on a single row or single column
        lblMessages.FontSize = 16
        lblMessages.Caption = "Your tiles must be all in" & vbNewLine & "one row or column"
        GoTo End_Action:
    Else    'go through this massive for loop

        'Check if word formed is isolated or not
        For intCount = 1 To counter
            If rowToken > columnToken Then    'Horizontal word
                Select Case location(intCount)
                Case 0
                    If (imgTile(location(intCount) + 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If

                    If (imgTile(location(intCount) + 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 14
                    If (imgTile(location(intCount) + 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If

                    If (imgTile(location(intCount) - 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 1 To 13
                    If intCount = 1 Then    'if header
                        If (imgTile(location(intCount) - 1).Picture <> 0) Or (imgTile(location(intCount) + 15).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf intCount = counter Then    'if trailer
                        If (imgTile(location(intCount) + 1).Picture <> 0) Or (imgTile(location(intCount) + 15).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf (imgTile(location(intCount) + 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 210
                    If (imgTile(location(intCount) - 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If

                    If (imgTile(location(intCount) + 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 224
                    If (imgTile(location(intCount) - 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If

                    If (imgTile(location(intCount) - 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 211 To 223
                    If intCount = 1 Then    'if header
                        If (imgTile(location(intCount) - 1).Picture <> 0) Or (imgTile(location(intCount) - 15).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf intCount = counter Then    'if trailer
                        If (imgTile(location(intCount) + 1).Picture <> 0) Or (imgTile(location(intCount) - 15).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf (imgTile(location(intCount) - 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 15, 30, 45, 60, 75, 90, 105, 120, 135, 150, 165, 180, 195
                    If (imgTile(location(intCount) - 15).Picture <> 0) Or (imgTile(location(intCount) + 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 29, 44, 59, 74, 89, 104, 119, 134, 149, 179, 194, 209
                    If (imgTile(location(intCount) - 15).Picture <> 0) Or (imgTile(location(intCount) + 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case Else
                    If intCount = 1 Then    'if header
                        If (imgTile(location(intCount) - 15).Picture <> 0) Or (imgTile(location(intCount) + 15).Picture <> 0) Or (imgTile(location(intCount) - 1).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf intCount = counter Then    'if trailer
                        If (imgTile(location(intCount) - 15).Picture <> 0) Or (imgTile(location(intCount) + 15).Picture <> 0) Or (imgTile(location(intCount) + 1).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf (imgTile(location(intCount) - 15).Picture <> 0) Or (imgTile(location(intCount) + 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If

                    If blnGridTilesPlayed(location(intCount) + 1) = True Then
                        blnNotIsolated = True
                    End If
                End Select
            ElseIf rowToken < columnToken Then    'Vertical word
                Select Case location(intCount)
                Case 0
                    If (imgTile(location(intCount) + 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If

                    If (imgTile(location(intCount) + 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 14
                    If (imgTile(location(intCount) - 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If

                    If (imgTile(location(intCount) + 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 1 To 13
                    If (imgTile(location(intCount) - 1).Picture <> 0) Or (imgTile(location(intCount) + 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 210
                    If (imgTile(location(intCount) + 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If

                    If (imgTile(location(intCount) - 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 224
                    If (imgTile(location(intCount) - 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If

                    If (imgTile(location(intCount) - 15).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 211 To 223
                    If (imgTile(location(intCount) - 1).Picture <> 0) Or (imgTile(location(intCount) + 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 15, 30, 45, 60, 75, 90, 105, 120, 135, 150, 165, 180, 195
                    If intCount = 1 Then    'if header
                        If (imgTile(location(intCount) + 1).Picture <> 0) Or (imgTile(location(intCount) - 15).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                        If (imgTile(location(intCount) + 15).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf intCount = counter Then    'if trailer
                        If (imgTile(location(intCount) + 1).Picture <> 0) Or (imgTile(location(intCount) + 15).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf (imgTile(location(intCount) + 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If
                Case 29, 44, 59, 74, 89, 104, 119, 134, 149, 179, 194, 209
                    If intCount = 1 Then    'if header
                        If (imgTile(location(intCount) - 1).Picture <> 0) Or (imgTile(location(intCount) - 15).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                        If (imgTile(location(intCount) + 15).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf intCount = counter Then    'if trailer
                        If (imgTile(location(intCount) - 1).Picture <> 0) Or (imgTile(location(intCount) + 15).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf (imgTile(location(intCount) - 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If

                Case Else
                    If intCount = 1 Then    'if header
                        If (imgTile(location(intCount) - 15).Picture <> 0) Or (imgTile(location(intCount) + 1).Picture <> 0) Or (imgTile(location(intCount) - 1).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf intCount = counter Then    'if trailer
                        If (imgTile(location(intCount) + 15).Picture <> 0) Or (imgTile(location(intCount) + 1).Picture <> 0) Or (imgTile(location(intCount) - 1).Picture <> 0) Then
                            blnNotIsolated = True
                        End If
                    ElseIf (imgTile(location(intCount) - 1).Picture <> 0) Or (imgTile(location(intCount) + 1).Picture <> 0) Then
                        blnNotIsolated = True
                    End If

                    If blnGridTilesPlayed(location(intCount) + 15) = True Then
                        blnNotIsolated = True
                    End If
                End Select
            Else
                blnNotIsolated = True
            End If
        Next intCount

        If (blnNotIsolated = False) And (blnGridTilesPlayed(112) = True) Then    'If word is isolated
            lblMessages.FontSize = 16
            lblMessages.Caption = "Tiles must connect to existing tiles"
            GoTo End_Action:
        End If

    End If

    If Len(strWord) = 1 Then
        lblMessages.FontSize = 16
        lblMessages.Caption = "The word must be at least" & vbNewLine & "two letter tiles long"
        GoTo End_Action:
    End If

    If WordCheck(strWord) = 0 Then
        'At this point, the player's word is valid, crosscheck for other words before computing main word

        Call DisplayBonusScore(strWord, temp(), location())

        If crossCheck(rowToken, columnToken, counter, location(), strWord) = False Then    'This means that one of the crosswords is not valid
            GoTo End_Action
        End If

        Call ComputeWordScore(strWord, temp(), location())    'Player's word is valid, so calculate the word score by this method call
        intPlayerScore(intTurn) = intPlayerScore(intTurn) + intScore    'Method returns the calculated score which is added to the player score

Continue:
        lblMessages.Caption = lblMessages.Caption & strWord & " " & intScore    'Display the word played and score

        If intBestWord < intScore Then
            intBestWord = intScore
            strGameSummary(1) = "Best word: " & strWord & " " & Str(intBestWord)
        End If


        intWordsPlayed(intTurn) = intWordsPlayed(intTurn) + 1
        frmGameOver.lstPlayer(intTurn).ListItems.Add(1).Text = Str(intWordsPlayed(intTurn))
        frmGameOver.lstPlayer(intTurn).ListItems.Item(1).ListSubItems.Add.Text = strWord
        frmGameOver.lstPlayer(intTurn).ListItems.Item(1).ListSubItems.Add.Text = intScore
        frmGameOver.lblPlayerScore(intTurn).Caption = intPlayerScore(intTurn)

        If intTurn = 0 Then
            lblPlayer1Score.Caption = intPlayerScore(intTurn)    'Display the updated player's score
        ElseIf intTurn = 1 Then
            lblPlayer2Score.Caption = intPlayerScore(intTurn)
        End If

        For intCount = 0 To 224    'Lock the tiles played after submission
            If temp(intCount) = True Then
                blnGridTilesPlayed(intCount) = True    'Lock the tile played permanently
            End If
        Next intCount

        Call GetNewTiles    'Fill the player's tile rack with new letters from the letter bag

        Call Continue(False)
        cmdSubmit.Picture = LoadPicture(App.Path & "\game_files\game_screen\continue_black.gif")
        intSeconds = 60
        intMinutes = 2
        strSeconds = "00"
        strMinutes = "0"
        lblTimeElapsed.Caption = "0:00"
        tmrTimeElapsed.Enabled = False
        GoTo End_Action:

New_Turn:
        Call NewTurn    'Next player's turn if it is two-player mode
        GoTo End_Action:

    Else    'If player's word is not a valid word, reject it and display error
        lblMessages.FontSize = 18
        lblMessages.Caption = strWord & " is INVALID" & vbNewLine & "Try again"
        GoTo End_Action:
    End If

End_Action:        'Invalid moves or invalid wordplays end here
    strWord = ""
    intScore = 0

End Sub


Private Function WordCheck(ByVal wordplayed As String) As Integer
    WordCheck = 1    'Initially, word is invalid
    intword = FreeFile

    Open App.Path & "\game_files\dictionary\" + Left(wordplayed, 1) + ".txt" For Input As #intword    ' previous filepath "\game_files\dictionary.txt"
    Do
        Input #intword, strNext
        If StrComp((UCase(wordplayed)), (UCase(strNext)), vbTextCompare) = 0 Then
            WordCheck = StrComp((UCase(wordplayed)), (UCase(strNext)), vbTextCompare)
            Exit Do
        End If
    Loop Until EOF(intword)

    Close #intword
End Function

Private Sub Continue(ByVal blnFlag As Boolean)
    cmdDictionary.Enabled = blnFlag
    cmdExchange.Enabled = blnFlag
    cmdRecall.Enabled = blnFlag
    cmdShuffle.Enabled = blnFlag

    For intCount = 0 To 6
        imgLetterTile(intCount).Enabled = blnFlag
    Next intCount

    For intCount = 0 To 224
        imgTile(intCount).Enabled = blnFlag
    Next intCount

    frmScrabble.MouseIcon = Nothing
    cmdPass.Enabled = blnFlag
    cmdPause.Enabled = blnFlag
End Sub
Private Sub NewTurn()
    If intTurn = 0 Then    'change player turns
        intTurn = 1
        lblPlayer2Name.FontBold = True
        lblPlayer1Name.FontBold = False
    ElseIf intTurn = 1 Then
        intTurn = 0
        lblPlayer1Name.FontBold = True
        lblPlayer2Name.FontBold = False
    End If

    For intCount = 0 To 6
        Call IdentifyTile(intPlayerTiles(intTurn, intCount))
        imgLetterTile(intCount).Picture = LoadPicture(App.Path & strTilePath & ".gif")
    Next intCount

    blnExchangeDisabled = False    'Enable Exchange

    Call DisableExchange

    lblMessages.FontSize = 20
    lblMessages.Caption = strPlayer(intTurn) & "'s Turn"

    If intTurn = 1 And blnTwoPlayer = False Then    'If it is the computer's turn

        cmdPass.Visible = False
        cmdPause.Visible = False
        cmdSubmit.Visible = False
        cmdDictionary.Visible = False
        cmdExchange.Visible = False
        cmdRecall.Visible = False
        cmdShuffle.Visible = False
        imgBag.Visible = False
        lblTilesRemaining.Visible = False


        For intCount = 0 To 6
            imgLetterTile(intCount).Visible = False
        Next intCount

        lblMessages.Caption = lblMessages.Caption & vbNewLine & "Making a move...(" & intComputerTries & ")"
        swfComputing.Visible = True
        Call swfComputing.LoadMovie(0, App.Path + "\game_files\other\hourglass.swf")

    ElseIf intTurn = 0 And blnTwoPlayer = False Then    'If it is the player's turn

        cmdPass.Visible = True
        cmdPause.Visible = True
        cmdSubmit.Visible = True
        cmdDictionary.Visible = True
        cmdExchange.Visible = True
        cmdRecall.Visible = True
        cmdShuffle.Visible = True
        imgBag.Visible = True
        lblTilesRemaining.Visible = True

        For intCount = 0 To 6
            imgLetterTile(intCount).Visible = True
        Next intCount
    End If

    'Reset time variables and enable timer again
    intSeconds = 60
    intMinutes = 2
    strSeconds = "00"
    strMinutes = "0"
    lblTimeElapsed.Caption = "3:00"
    tmrTimeElapsed.Enabled = True
    
    'Endgame
    If intTotalTiles < 7 Then
        strWord = ""
        intScore = 0
        MsgBox "GAME OVER: There are less than seven tiles remaining in the tile bag", vbOKOnly, "Game Over"
        Call EndGame

        Unload Me
        frmDictionary.Hide
        frmGameOver.Show
    End If
End Sub

Private Sub firstTurn()
    Dim rndOrientation As Integer
    Dim strJumbledWord As String
    Dim position As Integer

    For intCount = 0 To 4
        strJumbledWord = strJumbledWord + strTileValue(intPlayerTiles(intTurn, intCount))
    Next intCount

    strJumbledWord = Replace$(strJumbledWord, Space(1), Space(0))

    FindCombinations vbNullString, strJumbledWord

    rndOrientation = Int(Rnd * 2)

    If rndOrientation = 0 Then
        For intCount = 112 To (112 + (Len(strMostValuable) - 1)) Step 1
            imgLetterTile(position).Picture = Nothing

            If (position + 1) <= Len(strMostValuable) Then
                position = position + 1
            End If

            imgTile(intCount).Picture = LoadPicture(App.Path & "\game_files\word_tiles\tile_" & (UCase((Mid(strMostValuable, position, 1)))) & ".gif")

            For Inner = 0 To 26
                If strTileValue(Inner) = (UCase((Mid(strMostValuable, position, 1)))) Then
                    intGridTiles(intCount) = Inner
                End If
            Next Inner

            blnGridTilesPlayed(intCount) = True

            Call UpdateComputerScore(intCount, (UCase((Mid(strMostValuable, position, 1)))))
        Next intCount
    ElseIf rndOrientation = 1 Then
        For intCount = 112 To (112 + (15 * (Len(strMostValuable) - 1))) Step 15
            imgLetterTile(position).Picture = Nothing

            If (position + 1) <= Len(strMostValuable) Then
                position = position + 1
            End If

            imgTile(intCount).Picture = LoadPicture(App.Path & "\game_files\word_tiles\tile_" & (UCase((Mid(strMostValuable, position, 1)))) & ".gif")

            For Inner = 0 To 26
                If strTileValue(Inner) = (UCase((Mid(strMostValuable, position, 1)))) Then
                    intGridTiles(intCount) = Inner
                End If
            Next Inner

            blnGridTilesPlayed(intCount) = True

            Call UpdateComputerScore(intCount, (UCase((Mid(strMostValuable, position, 1)))))
        Next intCount
    End If

    strWordComputer = strMostValuable

    intPlayerScore(intTurn) = intPlayerScore(intTurn) + intScore
    lblMessages.Caption = "Computer played:" & vbNewLine & strWordComputer & " " & intScore

    If intBestWord < intScore Then
        intBestWord = intScore
        strGameSummary(1) = "Best word: " & strWordComputer & " " & Str(intBestWord)
    End If

    strMostValuable = ""
    intMostValuable = 0
End Sub
Private Sub ComputerMove()

    If blnGridTilesPlayed(112) = False Then    'Means this is the first turn
        Call firstTurn
        GoTo End_Turn:
    End If

Recompute:

    intScore = 0

    For intCount = 0 To 224
        If (imgTile(intCount).Picture <> 0) And (blnGridTilesPlayed(intCount) = False) Then
            imgTile(intCount).Picture = Nothing
        End If
    Next intCount

    Call cmdShuffle_Click

    Dim rndPosition As Integer
    Dim rndOrientation As Integer
    Dim intSpace As Integer

    Do    'Pick a random tile on the board
        Randomize
        rndPosition = Int(Rnd * 225)
    Loop Until (blnGridTilesPlayed(rndPosition) = True)

    rndOrientation = 2    'undefined

    If ((rndPosition >= 15) And (rndPosition <= 209)) Then
        If (imgTile(rndPosition - 15).Picture = 0) And (imgTile(rndPosition + 15).Picture = 0) Then
            rndOrientation = 1    'vertical
        End If
    End If

    If (((rndPosition Mod 15) <> 0) And (((rndPosition + 1) Mod 15) <> 0)) Then
        If (imgTile(rndPosition - 1).Picture = 0 And imgTile(rndPosition + 1).Picture = 0) Then
            rndOrientation = 0    'horizontal
        End If
    End If

    If rndOrientation = 0 Then    'find horizontal space
        'Left
        If ((rndPosition Mod 15) <> 0) Then
            For intCount = (rndPosition - 1) To (Int(rndPosition / 15) * 15) Step -1
                If (imgTile(intCount).Picture = 0) Then
                    If (intCount + 15) > 224 Then
                        GoTo Exit_Left
                    End If
                    If (imgTile(intCount - 15).Picture <> 0) Or (imgTile(intCount + 15).Picture <> 0) Then
                        GoTo Exit_Left:
                    End If
                    intSpace = intSpace + 1
                Else
                    GoTo Exit_Left:
                End If
            Next intCount
        End If

Exit_Left:

        'Right
        If rndPosition = 224 Then
            GoTo Exit_Right
        End If

        If (((rndPosition + 1) Mod 15) <> 0) Then
            For intCount = (rndPosition + 1) To ((Int(rndPosition / 15) * 15) + 15) Step 1
                If (imgTile(intCount).Picture = 0) Then
                    If (intCount + 15) > 224 Then
                        GoTo Exit_Right
                    End If
                    If (imgTile(intCount - 15).Picture <> 0) Or (imgTile(intCount + 15).Picture <> 0) Then
                        GoTo Exit_Right:
                    End If
                    intSpace = intSpace + 1
                Else
                    GoTo Exit_Right:
                End If
            Next intCount
        End If

Exit_Right:

    ElseIf rndOrientation = 1 Then    'find vertical space
        'Up
        If (rndPosition >= 15) Then
            For intCount = (rndPosition - 15) To (rndPosition Mod 15) Step -15
                If (imgTile(intCount).Picture = 0) Then
                    If (intCount + 1) > 224 Then
                        GoTo Exit_Up:
                    End If
                    If (imgTile(intCount - 1).Picture <> 0) Or (imgTile(intCount + 1).Picture <> 0) Then
                        GoTo Exit_Up:
                    End If
                    intSpace = intSpace + 1
                Else
                    GoTo Exit_Up:
                End If
            Next intCount
        End If

Exit_Up:

        'Down
        If (rndPosition <= 209) Then
            For intCount = (rndPosition + 15) To ((rndPosition Mod 15) + 210) Step 15
                If (imgTile(intCount).Picture = 0) Then
                    If (intCount + 1) > 224 Then
                        GoTo Exit_Down:
                    End If
                    If (imgTile(intCount - 1).Picture <> 0) Or (imgTile(intCount + 1).Picture <> 0) Then
                        GoTo Exit_Down:
                    End If
                    intSpace = intSpace + 1
                Else
                    GoTo Exit_Down:
                End If
            Next intCount
        End If

Exit_Down:

    End If

    If intSpace > 3 Then
        intSpace = 3    'max 4 letters long excluding base letter
    ElseIf intSpace = 0 Then    'if there is absolutely no space around the letter to form a word, find another random letter by recomputing
        GoTo Recompute:
    End If

    intComputerTries = intComputerTries + 1
    If intComputerTries = 5 Then
        Call cmdExchange_Click
    End If

    lblMessages.Caption = strPlayer(intTurn) & "'s Turn" & vbNewLine & "Making a move...(" & intComputerTries & ")"

    If intComputerTries = 7 Then    'After seven unsuccessful tries, stop computer's turn
        For intCount = 0 To 224
            If (imgTile(intCount).Picture <> 0) And (blnGridTilesPlayed(intCount) = False) Then
                imgTile(intCount).Picture = Nothing
            End If
        Next intCount

        strMostValuable = ""
        intScore = 0
        lblMessages.Caption = "Computer played:" & vbNewLine & strMostValuable & " " & intScore
        GoTo End_Turn
    End If

    Dim strComputerWord As String

    strComputerWord = strTileValue(intGridTiles(rndPosition))
    intSpace = intSpace + 1

    Call BestValidWord(strComputerWord, intSpace)

    Dim pivot As Integer
    Dim pos As Integer

    For intCount = 1 To Len(strMostValuable)
        If StrComp((UCase((Mid(strMostValuable, intCount, 1)))), strTileValue(intGridTiles(rndPosition)), vbTextCompare) = 0 Then    'Find out grid position of pivot letter
            pivot = intCount
        End If
    Next intCount

    If rndOrientation = 0 Then

        For intCount = (rndPosition - (pivot - 1)) To (rndPosition) Step 1    'left to pivot letter
            If ((intCount + 15) > 224) Or ((intCount - 15) < 0) Then
                Call ComputerMove
            End If
            If (rndPosition <> intCount) And ((imgTile(intCount - 15).Picture <> 0) Or (imgTile(intCount + 15).Picture <> 0)) Then
                Call ComputerMove
            End If
            pos = pos + 1
            If ((UCase((Mid(strMostValuable, pos, 1))))) <> "" Then
                Call CompPlayTile(intCount, pos, rndPosition)
            End If
        Next intCount

        For intCount = (rndPosition + 1) To (rndPosition + (Len(strMostValuable) - (pivot - 1))) Step 1    'pivot letter to right
            If ((intCount + 15) > 224) Or ((intCount - 15) < 0) Then
                Call ComputerMove
            End If
            If (rndPosition <> intCount) And ((imgTile(intCount - 15).Picture <> 0) Or (imgTile(intCount + 15).Picture <> 0)) Then
                Call ComputerMove
            End If
            pos = pos + 1
            If ((UCase((Mid(strMostValuable, pos, 1))))) <> "" Then
                Call CompPlayTile(intCount, pos, rndPosition)
            End If
        Next intCount

    ElseIf rndOrientation = 1 Then

        For intCount = (rndPosition - (15 * (pivot - 1))) To (rndPosition) Step 15    'up to pivot letter
            If (intCount >= 223) Then
                Call ComputerMove
            Else
                If (rndPosition <> intCount) And ((imgTile(intCount - 1).Picture <> 0) Or (imgTile(intCount + 1).Picture <> 0)) Then
                    Call ComputerMove
                End If
            End If
            pos = pos + 1
            If ((UCase((Mid(strMostValuable, pos, 1))))) <> "" Then
                Call CompPlayTile(intCount, pos, rndPosition)
            End If
        Next intCount

        For intCount = (rndPosition + 15) To (rndPosition + (15 * (Len(strMostValuable) - (pivot - 1)))) Step 15    'pivot letter to down
            If (intCount >= 224) Then
                Call ComputerMove
            Else
                If (rndPosition <> intCount) And ((imgTile(intCount - 1).Picture <> 0) Or (imgTile(intCount + 1).Picture <> 0)) Then
                    Call ComputerMove
                End If
            End If
            pos = pos + 1
            If ((UCase((Mid(strMostValuable, pos, 1))))) <> "" Then
                Call CompPlayTile(intCount, pos, rndPosition)
            End If
        Next intCount

    End If

    strWordComputer = strMostValuable

    intPlayerScore(intTurn) = intPlayerScore(intTurn) + intScore
    lblMessages.Caption = "Computer played:" & vbNewLine & strWordComputer & " " & intScore

    If intBestWord < intScore Then
        intBestWord = intScore
        strGameSummary(1) = "Best word: " & strWordComputer & " " & Str(intBestWord)
    End If

    Call DiscardTiles(intSpace, rndPosition)

End_Turn:
    swfComputing.Visible = False

    'Update word play log and statistics
    intWordsPlayed(intTurn) = intWordsPlayed(intTurn) + 1
    frmGameOver.lstPlayer(intTurn).ListItems.Add(1).Text = Str(intWordsPlayed(intTurn))
    frmGameOver.lstPlayer(intTurn).ListItems.Item(1).ListSubItems.Add.Text = strWordComputer
    frmGameOver.lstPlayer(intTurn).ListItems.Item(1).ListSubItems.Add.Text = intScore
    frmGameOver.lblPlayerScore(intTurn).Caption = intPlayerScore(intTurn)


    strMostValuable = ""
    intMostValuable = 0
    intScore = 0
    intComputerTries = 0

    'Display the updated player's score
    If intTurn = 0 Then
        lblPlayer1Score.Caption = intPlayerScore(intTurn)
    ElseIf intTurn = 1 Then
        lblPlayer2Score.Caption = intPlayerScore(intTurn)
    End If

    cmdSubmit.Visible = True

    Call GetNewTiles

    Call Continue(False)
    cmdSubmit.Picture = LoadPicture(App.Path & "\game_files\game_screen\continue_black.gif")
    intSeconds = 60
    intMinutes = 2
    strSeconds = "00"
    strMinutes = "0"
    lblTimeElapsed.Caption = "0:00"
    tmrTimeElapsed.Enabled = False

End Sub

Private Sub DiscardTiles(ByRef numSpace As Integer, ByRef position As Integer)
'remove the pivot letter first from the word played
    For Inner = 1 To Len(strMostValuable)
        If ((UCase((Mid(strMostValuable, Inner, 1))))) = strTileValue(intGridTiles(position)) Then
            Mid(strMostValuable, Inner, 1) = " "
            strMostValuable = Replace$(strMostValuable, Space(1), Space(0))
        End If
    Next Inner

    'remove the rest of the letters containing in the word from the computer`s tile rack
    For intCount = 0 To (numSpace - 1)
        For Inner = 1 To Len(strMostValuable)
            If ((UCase((Mid(strMostValuable, Inner, 1))))) = strTileValue(intPlayerTiles(intTurn, intCount)) Then
                Mid(strMostValuable, Inner, 1) = " "
                strMostValuable = Replace$(strMostValuable, Space(1), Space(0))
                imgLetterTile(intCount).Picture = Nothing
            End If
        Next Inner

        If strTileValue(intPlayerTiles(intTurn, intCount)) = " " Then    'If there is a blank, dont play it
            imgLetterTile(intCount).Picture = Nothing
        End If
    Next intCount


    For intCount = 0 To 224
        If (imgTile(intCount).Picture <> 0) And (blnGridTilesPlayed(intCount) = False) Then
            blnGridTilesPlayed(intCount) = True    'lock the tiles played by computer permanently
        End If
    Next intCount
End Sub

Private Sub CompPlayTile(ByRef counter As Integer, ByRef position As Integer, ByRef location As Integer)
    If blnGridTilesPlayed(counter) = True And (counter <> location) Then
        Call ComputerMove
    End If

    imgTile(counter).Picture = LoadPicture(App.Path & "\game_files\word_tiles\tile_" & (UCase((Mid(strMostValuable, position, 1)))) & ".gif")

    For Inner = 0 To 26
        If strTileValue(Inner) = (UCase((Mid(strMostValuable, position, 1)))) Then
            intGridTiles(counter) = Inner
        End If
    Next Inner

    Call UpdateComputerScore(counter, (UCase((Mid(strMostValuable, position, 1)))))
End Sub


Private Sub UpdateComputerScore(ByRef location As Integer, ByRef letter As String)
    Dim value As Integer

    Select Case (UCase(letter))    'Determine the value of the letter
    Case " "
        value = 0
    Case "E", "A", "I", "O", "N", "R", "T", "L", "S", "U"
        value = 1
    Case "D", "G"
        value = 2
    Case "B", "C", "M", "P"
        value = 3
    Case "F", "H", "V", "W", "Y"
        value = 4
    Case "K"
        value = 5
    Case "J", "X"
        value = 8
    Case "Q", "Z"
        value = 10
    End Select

    Select Case location    'Determine if the letter is on a triple letter or double letter score tile
    Case 20, 24, 76, 80, 84, 88, 136, 140, 144, 148, 200, 204
        value = value * 3    'If it is a triple letter score tile, multiply the letter value by 3
    Case 3, 11, 36, 38, 45, 52, 59, 92, 96, 98, 102, 108, 116, 122, 126, 128, 132, 165, 172, 179, 186, 188, 213, 221
        value = value * 2
    Case Else
    End Select

    If location = intBlankTile(0) Or location = intBlankTile(1) Then
        value = 0    'Blank tile
    End If
    intScore = intScore + value

    Select Case location
    Case 0, 7, 14, 105, 119, 210, 217, 224
        intScore = intScore * 3    'Triple word score tile multiplies the score by 3
    Case 16, 28, 32, 42, 48, 56, 64, 70, 112, 154, 160, 168, 176, 182, 192, 196, 208    'centre tile (112) is also double word score tile
        intScore = intScore * 2
    End Select

End Sub

Private Sub BestValidWord(ByRef strJumbledWord As String, ByRef length As Integer)
'AI- Randomly select the first four jumbled letters from the rack (plus a random letter from board), so max 5 letter word
'Go for the maximum possible word
    Dim pivotLetter As String

    pivotLetter = strJumbledWord

    strPivot = pivotLetter

Recompute:

    strJumbledWord = pivotLetter

    For intCount = 0 To (length - 1)
        strJumbledWord = strJumbledWord + strTileValue(intPlayerTiles(intTurn, intCount))
    Next intCount

    strJumbledWord = Replace$(strJumbledWord, Space(1), Space(0))

    FindCombinations vbNullString, strJumbledWord
End Sub

Private Sub FindCombinations(ByVal Prefix$, ByVal msg$)
'finds all combinations of letters in msg$. Adds prefix to msg to make the whole word
    Dim i%, n%
    n = Len(msg)
    If n Then    'find sub letters
        For i = 1 To n
            FindCombinations Prefix & Mid$(msg$, i, 1), Left$(msg$, i - 1) & Mid$(msg$, i + 1)
        Next i
    Else    'we are now at bottom so output the word
    End If

    If Len(Prefix) > 1 Then

        If WordCheck(Prefix) = 0 Then
            'Check the value of the valid word
            If blnGridTilesPlayed(112) = False Then
                GoTo noPivot:
            End If

            Dim pivotPresent As Boolean

            For intCount = 1 To Len(Prefix)
                If strPivot = (UCase((Mid(Prefix, intCount, 1)))) Then
                    pivotPresent = True
                End If
            Next intCount

            If pivotPresent = False Then
                GoTo Skip
            End If
noPivot:
            Dim value As Integer

            For intCount = 1 To Len(Prefix)    'Traverse from the first letter of the word to the last letter
                Select Case (UCase((Mid(Prefix, intCount, 1))))    'Determine the value of the letter
                Case " "
                    value = value + 0
                Case "E", "A", "I", "O", "N", "R", "T", "L", "S", "U"
                    value = value + 1
                Case "D", "G"
                    value = value + 2
                Case "B", "C", "M", "P"
                    value = value + 3
                Case "F", "H", "V", "W", "Y"
                    value = value + 4
                Case "K"
                    value = value + 5
                Case "J", "X"
                    value = value + 8
                Case "Q", "Z"
                    value = value + 10
                End Select
            Next intCount

            If value > intMostValuable Then
                intMostValuable = value
                strMostValuable = Prefix
            End If
        End If
    End If

Skip:
End Sub


Private Sub DisplayBonusScore(ByRef word As String, ByRef blnTemp() As Boolean, ByRef intLocation() As Integer)    'Identify and display any bonus square tile scores
    Dim doublewordscore, triplewordscore As Integer
    Dim doubleletterscore, tripleletterscore As Integer

    For intCount = 1 To Len(word)    'Traverse from the first letter of the word to the last letter
        Select Case intLocation(intCount)    'Determine if the letter is on a triple letter or double letter score tile
        Case 20, 24, 76, 80, 84, 88, 136, 140, 144, 148, 200, 204
            tripleletterscore = tripleletterscore + 1
        Case 3, 11, 36, 38, 45, 52, 59, 92, 96, 98, 102, 108, 116, 122, 126, 128, 132, 165, 172, 179, 186, 188, 213, 221
            doubleletterscore = doubleletterscore + 1
        End Select
    Next intCount

    For intCount = 0 To 224    'Determine if the letter is on a triple word or double word score tile
        If blnTemp(intCount) = True Then
            Select Case intCount
            Case 0, 7, 14, 105, 119, 210, 217, 224
                triplewordscore = triplewordscore + 1
            Case 16, 28, 32, 42, 48, 56, 64, 70, 112, 154, 160, 168, 176, 182, 192, 196, 208    'centre tile (112) is also double word score tile
                doublewordscore = doublewordscore + 1
            End Select
        End If
    Next intCount

    'Setting message

    Select Case doublewordscore    'Display message if double word score is achieved
    Case 1
        lblMessages.Caption = lblMessages.Caption + "Double word score! "
    Case 2
        lblMessages.Caption = lblMessages.Caption + "Double word score! x 2 "
    End Select

    Select Case triplewordscore    'Display message if triple word score is achieved
    Case 1
        lblMessages.Caption = lblMessages.Caption + " Triple word score!"
    Case 2
        lblMessages.Caption = lblMessages.Caption + " Triple word score! x 2"
    Case 3
        lblMessages.Caption = lblMessages.Caption + " Triple word score! x 3"
    End Select

    If (doublewordscore > 0) Or (triplewordscore > 0) Then
        lblMessages.Caption = lblMessages.Caption & vbNewLine    'Adjust the message
        lblMessages.FontSize = 16
    End If

    Select Case doubleletterscore    'Display message if double letter score is achieved
    Case 1
        lblMessages.Caption = lblMessages.Caption + " Double letter score!"
    Case 2
        lblMessages.Caption = lblMessages.Caption + " Double letter score! x 2"
    Case 3
        lblMessages.Caption = lblMessages.Caption + " Double letter score! x 3"
    Case 4
        lblMessages.Caption = lblMessages.Caption + " Double letter score! x 4"
    End Select

    Select Case tripleletterscore    'Display message if triple letter score is achieved
    Case 1
        lblMessages.Caption = lblMessages.Caption + " Triple letter score!"
    Case 2
        lblMessages.Caption = lblMessages.Caption + " Triple letter score! x 2"
    Case 3
        lblMessages.Caption = lblMessages.Caption + " Triple letter score! x 3"
    Case 4
        lblMessages.Caption = lblMessages.Caption + " Triple letter score! x 4"
    End Select

    If (doubleletterscore > 0) Or (tripleletterscore > 0) Then    'Adjust message to bigger or smaller font based on size of message
        lblMessages.Caption = lblMessages.Caption & vbNewLine
        lblMessages.FontSize = 16
        If (doublewordscore > 0) Or (triplewordscore > 0) Then
            lblMessages.FontSize = 12
        Else
            lblMessages.FontSize = 20
        End If
    End If
End Sub

Private Sub ComputeWordScore(ByRef word As String, ByRef blnTemp() As Boolean, ByRef intLocation() As Integer)    'For the main word
    Dim value As Integer

    intScore = 0

    For intCount = 1 To Len(word)    'Traverse from the first letter of the word to the last letter
        Select Case (UCase((Mid(word, intCount, 1))))    'Determine the value of the letter
        Case " "
            value = 0
        Case "E", "A", "I", "O", "N", "R", "T", "L", "S", "U"
            value = 1
        Case "D", "G"
            value = 2
        Case "B", "C", "M", "P"
            value = 3
        Case "F", "H", "V", "W", "Y"
            value = 4
        Case "K"
            value = 5
        Case "J", "X"
            value = 8
        Case "Q", "Z"
            value = 10
        End Select

        Select Case intLocation(intCount)    'Determine if the letter is on a triple letter or double letter score tile
        Case 20, 24, 76, 80, 84, 88, 136, 140, 144, 148, 200, 204
            value = value * 3    'If it is a triple letter score tile, multiply the letter value by 3
        Case 3, 11, 36, 38, 45, 52, 59, 92, 96, 98, 102, 108, 116, 122, 126, 128, 132, 165, 172, 179, 186, 188, 213, 221
            value = value * 2
        End Select

        If intLocation(intCount) = intBlankTile(0) Or intLocation(intCount) = intBlankTile(1) Then
            value = 0    'Blank tile
        End If
        intScore = intScore + value    'Add the letter score to the word score
    Next intCount


    For intCount = 0 To 224    'Determine if the letter is on a triple word or double word score tile
        If blnTemp(intCount) = True Then
            Select Case intCount
            Case 0, 7, 14, 105, 119, 210, 217, 224
                intScore = intScore * 3    'Triple word score tile multiplies the score by 3
            Case 16, 28, 32, 42, 48, 56, 64, 70, 112, 154, 160, 168, 176, 182, 192, 196, 208    'centre tile (112) is also double word score tile
                intScore = intScore * 2
            End Select
        End If
    Next intCount
End Sub

Private Sub GetNewTiles()
    Dim random As Integer

    For intCount = 0 To 6
        If imgLetterTile(intCount).Picture = 0 Then    'Number of new tiles needed is dependent on number of empty tiles

            Do    'revised loop
                Randomize
                random = Int(Rnd * 100)
                intPlayerTiles(intTurn, intCount) = intTileBank(random)
            Loop Until (intTileBank(random) <> -1)    'Get a random letter tile as long as that letter has quantity remaining

            intTileBank(random) = -1
            intTilesInPlay(intTurn, intCount) = random

            intTilesRemaining(intPlayerTiles(intTurn, intCount)) = intTilesRemaining(intPlayerTiles(intTurn, intCount)) - 1    'Update the letter tile quantity remaining

        End If

        Call IdentifyTile(intPlayerTiles(intTurn, intCount))    'Identifies the letter based on letter number (example: A= 0, B= 1 etc.)

        imgLetterTile(intCount).Picture = LoadPicture(App.Path & strTilePath & ".gif")    'Send the new random tile to an empty tile position on the rack

    Next intCount

    Call DisableRecall
    Call DisableExchange

    Call UpdateTotalTiles
End Sub

Private Sub cmdSubmit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetButtons    'Reset all buttons to black color
    If cmdPass.Enabled = True Then
        cmdSubmit.Picture = LoadPicture(App.Path & "\game_files\game_screen\submit_red.bmp")    'Change Submit button to red if mouse over
    Else
        cmdSubmit.Picture = LoadPicture(App.Path & "\game_files\game_screen\continue_red.gif")
    End If
End Sub

Private Sub cmdPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetButtons
    cmdPause.Picture = LoadPicture(App.Path & "\game_files\game_screen\pause_red.bmp")
End Sub

Private Sub cmdPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetButtons
    cmdPass.Picture = LoadPicture(App.Path & "\game_files\game_screen\pass_red.bmp")
End Sub

Private Sub cmdPause_Click()
    If frmScrabble.MouseIcon = 0 Then    'if cursor is not holding a tile
        tmrTimeElapsed.Enabled = False    'Pause the time elapsed clock
        frmPauseMenu.Show vbModal    'Show the pause menu
    End If
End Sub

Private Sub Form_Load()
    NewGame    'Method call for a new game
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetButtons    'Reset buttons to black because cursor is on the form
End Sub

Private Sub imgBoardGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetButtons
End Sub

Private Sub imgLetterTile_Click(Index As Integer)    'This click operation is each of the seven player rack tiles
    If imgLetterTile(Index).Picture <> 0 Then    'if tile is already here
        If frmScrabble.MouseIcon = 0 Then    'if cursor is empty, pick up tile
            intTile = intPlayerTiles(intTurn, Index)
            Call IdentifyTile(intTile)    'Identify which letter image is to be loaded
            frmScrabble.MouseIcon = LoadPicture(App.Path & strTilePath & ".ico")
            imgLetterTile(Index).Picture = Nothing
        ElseIf frmScrabble.MouseIcon <> 0 Then    'if cursor already has tile, swap tiles
            intTempTile = intPlayerTiles(intTurn, Index)
            imgLetterTile(Index).Picture = LoadPicture(App.Path & strTilePath & ".gif")    'Load the picture of the tile placed
            intPlayerTiles(intTurn, Index) = intTile
            Call IdentifyTile(intTempTile)
            frmScrabble.MouseIcon = LoadPicture(App.Path & strTilePath & ".ico")    'Load the cursor picture because a tile is picked up
            intTile = intTempTile
        End If
    ElseIf imgLetterTile(Index).Picture = 0 Then    'if tile is not here
        If frmScrabble.MouseIcon <> 0 Then    'place tile
            intPlayerTiles(intTurn, Index) = intTile
            imgLetterTile(Index).Picture = LoadPicture(App.Path & strTilePath & ".gif")    'Load the picture of the tile placed
            frmScrabble.MouseIcon = Nothing    'Reset cursor picture to cursor arrow
            'Check if the Exchange button or Recall button should be disabled because the tiles were moved
            Call DisableExchange
            Call DisableRecall
        End If
    End If
End Sub

Private Sub imgTile_Click(Index As Integer)    'This click operation is for each of the 225 tiles on the gameboard grid
    If blnGridTilesPlayed(Index) = False Then
        If imgTile(Index).Picture <> 0 Then    'if tile is already here
            If frmScrabble.MouseIcon = 0 Then    'if cursor is empty, pick up tile
                If Index = intBlankTile(0) Or Index = intBlankTile(1) Then    'if the player is picking up the tile within a turn which is actually a blank tile
                    intTile = 26
                    strTilePath = "\game_files\word_tiles\tile_blank"

                    If intBlankTile(0) = Index Then
                        intBlankTile(0) = -1
                    Else
                        intBlankTile(1) = -1
                    End If
                Else
                    intTile = intGridTiles(Index)
                    Call IdentifyTile(intTile)
                End If
                frmScrabble.MouseIcon = LoadPicture(App.Path & strTilePath & ".ico")
                imgTile(Index).Picture = Nothing
            ElseIf frmScrabble.MouseIcon <> 0 Then    'if cursor already has tiles, swap tiles
                If Index = intBlankTile(0) Or Index = intBlankTile(1) Then
                    intTempTile = 26
                Else
                    intTempTile = intGridTiles(Index)
                End If

                If intTile = 26 Then    'if tile being placed down is a blank tile
                    frmChoose.Show vbModal
                    If intTile = 26 Then    'This means that the user did not choose a letter for the blank tile
                        GoTo SetBlank_Fail
                    End If
                    If intBlankTile(0) = -1 Then
                        intBlankTile(0) = Index
                    ElseIf intBlankTile(1) = -1 Then
                        intBlankTile(1) = Index
                    End If
                End If
                imgTile(Index).Picture = LoadPicture(App.Path & strTilePath & ".gif")    'Load the picture of the letter tile placed down

                If intTempTile = 26 Then    'if tile being picked up is a blank tile
                    If Index = intBlankTile(0) Then
                        intBlankTile(0) = -1
                    ElseIf Index = intBlankTile(1) Then
                        intBlankTile(1) = -1
                    End If
                End If


                intGridTiles(Index) = intTile

                Call IdentifyTile(intTempTile)
                frmScrabble.MouseIcon = LoadPicture(App.Path & strTilePath & ".ico")    'Load the picture of the letter tile picked up for the cursor
                intTile = intTempTile
            End If
        ElseIf imgTile(Index).Picture = 0 Then    'if tile is not here
            If frmScrabble.MouseIcon <> 0 Then    'place tile
                If intTile = 26 Then    'if tile being placed down is a blank tile
                    frmChoose.Show vbModal
                    If intTile = 26 Then    'This means that the user did not choose a letter for the blank tile
                        GoTo SetBlank_Fail
                    End If
                    If intBlankTile(0) = -1 Then
                        intBlankTile(0) = Index
                    Else
                        intBlankTile(1) = Index
                    End If
                End If
                intGridTiles(Index) = intTile
                imgTile(Index).Picture = LoadPicture(App.Path & strTilePath & ".gif")
                frmScrabble.MouseIcon = Nothing    'Reset the mouse cursor to the cursor arrow
            End If
        End If
    End If

SetBlank_Fail:
    'Check to see if the Exchange button or Recall button should be disabled because tiles were moved
    If blnGridTilesPlayed(Index) = False Then
        Call DisableExchange
        Call DisableRecall
    End If
End Sub

Private Sub lblTilesRemaining_Click()
    If frmScrabble.MouseIcon = 0 Then
        frmTileBag.Show vbModal
    End If
End Sub


Private Sub tmrTimeElapsed_Timer()

    intGameTimeElapsed = intGameTimeElapsed + 1    'Records the total game time elapsed

    intSeconds = intSeconds - 1

    If intSeconds < 0 And intMinutes <> 0 Then
        intMinutes = intMinutes - 1
        intSeconds = 59
    ElseIf intSeconds < 0 And intMinutes = 0 Then
        lblMessages.FontSize = 18
        lblMessages.Caption = "Out of time!" & vbNewLine & "End of " & strPlayer(intTurn) & "'s turn"
        frmScrabble.MouseIcon = Nothing
        Call RecallTiles
        intSeconds = 0
        tmrTimeElapsed.Enabled = False
        Call Continue(False)
        cmdSubmit.Picture = LoadPicture(App.Path & "\game_files\game_screen\continue_black.gif")
    End If

    strSeconds = intSeconds
    strMinutes = intMinutes

    If intSeconds < 10 Then    'Converts the seconds to a double digit string
        strSeconds = "0" & intSeconds
    End If

    lblTimeElapsed.Caption = strMinutes & ":" & strSeconds    'Display the time elapsed (in string form)

    If (lblTimeElapsed.Caption = "2:58") And (blnTwoPlayer = False) And (intTurn = 1) Then    'AI- If it is the computer's turn, make a move after two second elapse
        Call ComputerMove
    End If
End Sub

Public Sub NewGame()
    Dim player As Integer

    lblPlayer1Name.Caption = strPlayer(0)    'Display the player's name
    lblPlayer1Name.FontBold = True
    lblPlayer2Name.Caption = strPlayer(1)

    intPlayerScore(0) = 0
    intPlayerScore(1) = 0
    lblPlayer1Score.Caption = "0"
    lblPlayer2Score.Caption = "0"

    intWordsPlayed(0) = 0
    intWordsPlayed(1) = 0
    frmGameOver.lstPlayer(0).ListItems.Clear
    frmGameOver.lstPlayer(1).ListItems.Clear

    For intCount = 0 To 224    'Reset gameboard tile variables
        intGridTiles(intCount) = 0
        blnGridTilesPlayed(intCount) = 0    '0 is false, unlock all the gameboard grid tiles
        imgTile(intCount).Picture = Nothing
    Next intCount

    'Reset time elapsed variables
    intSeconds = 60
    intMinutes = 2
    strSeconds = "00"
    strMinutes = "0"
    lblTimeElapsed.Caption = "3:00"
    tmrTimeElapsed.Enabled = True


    intBlankTile(0) = -1
    intBlankTile(1) = -1

    Call GenerateTiles    'Generates seven random tiles for the player

    Dim increment As Integer
    Dim random As Integer

    For intCount = 0 To 26
        For Inner = 1 To intTilesRemaining(intCount)
            intTileBank(increment) = intCount
            increment = increment + 1
        Next Inner
    Next intCount

    For player = 1 To 0 Step -1    'Give tiles for player 2 first, then give tiles for player 1

        For intCount = 0 To 6
            Do
                Randomize
                random = Int(Rnd * 100)
                intPlayerTiles(player, intCount) = intTileBank(random)
            Loop Until (intTileBank(random) <> -1)    'Get a random letter tile as long as that letter has quantity remaining

            intTileBank(random) = -1
            intTilesInPlay(player, intCount) = random

            intTilesRemaining(intPlayerTiles(player, intCount)) = intTilesRemaining(intPlayerTiles(player, intCount)) - 1

            Call IdentifyTile(intPlayerTiles(player, intCount))
            imgLetterTile(intCount).Picture = LoadPicture(App.Path & strTilePath & ".gif")
        Next intCount

        cmdExchange.Enabled = True
        Call DisableRecall    'Disable the Recall button because all the seven tiles are already on the player rack
        Call UpdateTotalTiles

    Next player

    intTurn = 0    'Player 1's plays first

    lblMessages.FontSize = 18
    lblMessages.Caption = "Welcome to Scrabble!" & vbNewLine & strPlayer(intTurn) & "'s Turn"

End Sub

Private Sub ResetButtons()    'Change all command buttons to black
    cmdDictionary.Picture = LoadPicture(App.Path & "\game_files\game_screen\dictionary_black.bmp")
    cmdExchange.Picture = LoadPicture(App.Path & "\game_files\game_screen\exchange_black.bmp")
    cmdRecall.Picture = LoadPicture(App.Path & "\game_files\game_screen\recall_black.bmp")
    cmdShuffle.Picture = LoadPicture(App.Path & "\game_files\game_screen\shuffle_black.bmp")
    If cmdPass.Enabled = True Then
        cmdSubmit.Picture = LoadPicture(App.Path & "\game_files\game_screen\submit_black.bmp")
    Else
        cmdSubmit.Picture = LoadPicture(App.Path & "\game_files\game_screen\continue_black.gif")
    End If
    cmdPause.Picture = LoadPicture(App.Path & "\game_files\game_screen\pause_black.bmp")
    cmdPass.Picture = LoadPicture(App.Path & "\game_files\game_screen\pass_black.bmp")
End Sub

Private Sub GenerateTiles()    'Generates all of the Scrabble tiles (100 tiles) and assigns them a value and quantity
    For intCount = 0 To 26
        Select Case intCount
        Case 0
            intTilesRemaining(intCount) = 9
            intTileValue(intCount) = 1
            strTileValue(intCount) = "A"
        Case 1
            intTilesRemaining(intCount) = 2
            intTileValue(intCount) = 3
            strTileValue(intCount) = "B"
        Case 2
            intTilesRemaining(intCount) = 2
            intTileValue(intCount) = 3
            strTileValue(intCount) = "C"
        Case 3
            intTilesRemaining(intCount) = 4
            intTileValue(intCount) = 2
            strTileValue(intCount) = "D"
        Case 4
            intTilesRemaining(intCount) = 12
            intTileValue(intCount) = 1
            strTileValue(intCount) = "E"
        Case 5
            intTilesRemaining(intCount) = 2
            intTileValue(intCount) = 4
            strTileValue(intCount) = "F"
        Case 6
            intTilesRemaining(intCount) = 3
            intTileValue(intCount) = 2
            strTileValue(intCount) = "G"
        Case 7
            intTilesRemaining(intCount) = 2
            intTileValue(intCount) = 4
            strTileValue(intCount) = "H"
        Case 8
            intTilesRemaining(intCount) = 9
            intTileValue(intCount) = 1
            strTileValue(intCount) = "I"
        Case 9
            intTilesRemaining(intCount) = 1
            intTileValue(intCount) = 8
            strTileValue(intCount) = "J"
        Case 10
            intTilesRemaining(intCount) = 1
            intTileValue(intCount) = 5
            strTileValue(intCount) = "K"
        Case 11
            intTilesRemaining(intCount) = 4
            intTileValue(intCount) = 1
            strTileValue(intCount) = "L"
        Case 12
            intTilesRemaining(intCount) = 2
            intTileValue(intCount) = 3
            strTileValue(intCount) = "M"
        Case 13
            intTilesRemaining(intCount) = 6
            intTileValue(intCount) = 1
            strTileValue(intCount) = "N"
        Case 14
            intTilesRemaining(intCount) = 8
            intTileValue(intCount) = 1
            strTileValue(intCount) = "O"
        Case 15
            intTilesRemaining(intCount) = 2
            intTileValue(intCount) = 3
            strTileValue(intCount) = "P"
        Case 16
            intTilesRemaining(intCount) = 1
            intTileValue(intCount) = 10
            strTileValue(intCount) = "Q"
        Case 17
            intTilesRemaining(intCount) = 6
            intTileValue(intCount) = 1
            strTileValue(intCount) = "R"
        Case 18
            intTilesRemaining(intCount) = 4
            intTileValue(intCount) = 1
            strTileValue(intCount) = "S"
        Case 19
            intTilesRemaining(intCount) = 6
            intTileValue(intCount) = 1
            strTileValue(intCount) = "T"
        Case 20
            intTilesRemaining(intCount) = 4
            intTileValue(intCount) = 1
            strTileValue(intCount) = "U"
        Case 21
            intTilesRemaining(intCount) = 2
            intTileValue(intCount) = 4
            strTileValue(intCount) = "V"
        Case 22
            intTilesRemaining(intCount) = 2
            intTileValue(intCount) = 4
            strTileValue(intCount) = "W"
        Case 23
            intTilesRemaining(intCount) = 1
            intTileValue(intCount) = 8
            strTileValue(intCount) = "X"
        Case 24
            intTilesRemaining(intCount) = 2
            intTileValue(intCount) = 4
            strTileValue(intCount) = "Y"
        Case 25
            intTilesRemaining(intCount) = 1
            intTileValue(intCount) = 10
            strTileValue(intCount) = "Z"
        Case 26
            intTilesRemaining(intCount) = 2
            intTileValue(intCount) = 0
            strTileValue(intCount) = " "
        End Select
    Next intCount

    Call UpdateTotalTiles
End Sub

Private Sub IdentifyTile(ByRef num As Integer)    'Identify the picture file path of the letter based on the letter number (example: A= 0)
    Select Case num
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
    Case 26
        strTilePath = "\game_files\word_tiles\tile_blank"
    End Select
End Sub

Private Sub UpdateTotalTiles()    'Calculates the current total number of tiles remaining in the game
    intTotalTiles = 0
    For intCount = 0 To 26
        intTotalTiles = intTotalTiles + intTilesRemaining(intCount)
    Next intCount

    lblTilesRemaining.Caption = intTotalTiles
End Sub

