VERSION 5.00
Begin VB.Form FrmPlayers 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Players"
   ClientHeight    =   4305
   ClientLeft      =   2055
   ClientTop       =   3435
   ClientWidth     =   6600
   Icon            =   "FrmPlayers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LstPlayerNo 
      Height          =   1035
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.ListBox LstPlayers 
      Height          =   1035
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox TxtPlayerName 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   900
      Width           =   2115
   End
   Begin VB.Label LbCFinished 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Finished"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label LbCEditDB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Edit &DataBase"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2820
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label LbCEnterPlayer 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Enter Player"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1620
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label LblPlayerNumb 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   900
      Width           =   495
   End
   Begin VB.Image ImgChosenCounter 
      Height          =   480
      Left            =   3480
      Top             =   840
      Width           =   480
   End
   Begin VB.Image ImgCounter 
      Height          =   480
      Index           =   1
      Left            =   4320
      Top             =   300
      Width           =   480
   End
   Begin VB.Label LblPlayerNameLab 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Name:"
      Height          =   195
      Left            =   1920
      TabIndex        =   2
      Top             =   540
      Width           =   585
   End
   Begin VB.Label LblPlayerLab 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Player:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   765
   End
   Begin VB.Menu MnuGame 
      Caption         =   "&Game"
      Begin VB.Menu MnuEnterPlayer 
         Caption         =   "Enter &Player"
      End
      Begin VB.Menu MnuEditDB 
         Caption         =   "&Edit DataBase"
      End
      Begin VB.Menu MnuFinished 
         Caption         =   "&Finished"
      End
   End
End
Attribute VB_Name = "FrmPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Dim Ctrl As Control
Call SetCmdText(FrmPlayers)
FrmPlayers.BackColor = BrdColour
For Each Ctrl In FrmPlayers.Controls
If Ctrl.Name Like "LbC*" Then
    Ctrl.BackColor = BrdColour
    Ctrl.ForeColor = LbcForeCol 'TextColour
End If
Next Ctrl
End Sub

Private Sub Form_Load()
Call ModPlayers.CreateForm      'Put available counters on Players form
End Sub

Private Sub ImgCounter_Click(Index As Integer)
FrmPlayers.ImgChosenCounter.Picture = ImgCounter(Index).Picture
'Put chosen counter on board
CounterNumb = ImgCounter(Index).Index
End Sub

Private Sub LbCEditDB_Click()
FrmEditDB.Show
Call EditDB           'Go to Edit DataBase Options
End Sub

Private Sub LbCEnterPlayer_Click()
Call EnterPlyr          'Add player to DataBase
End Sub

Private Sub LbCfinished_Click()
Call ModPlayers.Finished
End Sub

Private Sub MnuEditDB_Click()
FrmEditDB.Show
Call EditDB           'Go to Edit DataBase Options
End Sub

Private Sub MnuEnterPlayer_Click()
Call EnterPlyr          'Add player to DataBase
End Sub

Private Sub MnuFinished_Click()
Call ModPlayers.Finished
End Sub
