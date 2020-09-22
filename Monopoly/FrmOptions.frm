VERSION 5.00
Begin VB.Form FrmOptions 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Options"
   ClientHeight    =   5040
   ClientLeft      =   2055
   ClientTop       =   3030
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CboVersion 
      Height          =   315
      ItemData        =   "FrmOptions.frx":0000
      Left            =   2048
      List            =   "FrmOptions.frx":0002
      TabIndex        =   4
      ToolTipText     =   "Select the Required Game Version"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label LbCSelectVers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Select Version"
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
      Left            =   2408
      TabIndex        =   5
      ToolTipText     =   "Click Here to Load the Version Selected Below"
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label LbCToGame 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Back To &Game"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4695
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label LbCPlayers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Enter Players"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Click Here to Enter Players"
      Top             =   840
      Width           =   975
   End
   Begin VB.Label LbCEditDB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Edit DataBase"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1095
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label LbCLoadGame 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Load Game"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "Click Here to Load a Saved Game"
      Top             =   840
      Width           =   975
   End
   Begin VB.Menu MnuGame 
      Caption         =   "&Game"
      Begin VB.Menu MnuBack 
         Caption         =   "&Back to Game"
      End
      Begin VB.Menu MnuSelectVers 
         Caption         =   "&Select Version"
      End
      Begin VB.Menu MnuLoadGame 
         Caption         =   "&Load Game"
      End
   End
   Begin VB.Menu MnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu MnuEditDB 
         Caption         =   "&Edit Data Base"
      End
      Begin VB.Menu MnuEnterPlayers 
         Caption         =   "Enter &Players"
      End
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Dim Ctrl As Control
Call SetCmdText(FrmOptions)
FrmOptions.BackColor = BrdColour
For Each Ctrl In FrmOptions.Controls
If Ctrl.Name Like "LbC*" And Ctrl.Name <> "LbcPlayers" Then
    Ctrl.BackColor = BrdColour
    Ctrl.ForeColor = LbcForeCol
End If
Next Ctrl
LbCPlayers.Enabled = False
LbCPlayers.ForeColor = &H8000000F
End Sub

Private Sub Form_Load()
Call CreateVersionList
End Sub

Private Sub LbCLoadGame_Click()
Unload FrmBoard
Call LoadGame
Call ModOptions.BackToGame
End Sub

Private Sub LbCSelectVers_Click()
Dim i As Integer
If CboVersion.Text = "" Then
    i = MsgBox("Please Select a Version" & vbCrLf & _
        "From The List Below", vbCritical, "Select a Version")
    Exit Sub
End If
VersName = CboVersion.Text
Call LoadVersion
LbCPlayers.Enabled = True
LbCPlayers.ForeColor = TextColour
End Sub

Private Sub MnuBack_Click()
Call ModOptions.BackToGame  'Go back to Game
End Sub

Private Sub MnuBoardCol_Click()
Call ModOptions.BoardColour         'Change Board Colour
End Sub

Private Sub MnuEditDB_Click()
FrmEditDB.Show
Call EditDB         'Go to Edit DataBase Options
End Sub

Private Sub MnuEnterPlayers_Click()
Me.Hide
FrmPlayers.Show             'Enter Players
End Sub

Private Sub LbCEditDB_Click()
FrmEditDB.Show
Call EditDB           'Go to Edit DataBase Options
End Sub

Private Sub LbCPlayers_Click()
Unload FrmPlayers
FrmPlayers.Show             'Enter Players
End Sub

Private Sub LbCToGame_Click()
Call ModOptions.BackToGame  'Go back to Game
End Sub

Private Sub MnuLoadGame_Click()
Call LoadGame
Call ModOptions.BackToGame
End Sub

Private Sub MnuTextProp_Click()
Call ModOptions.BoardText       'Change board text appearance
End Sub

Private Sub MnuSelectVers_Click()
VersName = CboVersion.Text
Call LoadVersion
End Sub
