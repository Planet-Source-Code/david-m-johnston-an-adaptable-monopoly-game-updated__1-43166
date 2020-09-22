VERSION 5.00
Begin VB.Form FrmAuction 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Auction"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6000
      Top             =   120
   End
   Begin VB.CommandButton CmdBid 
      Caption         =   "Bid"
      Height          =   855
      Index           =   0
      Left            =   240
      Picture         =   "FrmAuction.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label LblGoingFor 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "GoingFor"
      Height          =   375
      Left            =   2393
      TabIndex        =   8
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label LblGoingToLab 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Going to"
      Height          =   375
      Left            =   3173
      TabIndex        =   7
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label LblProp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label LblSecsLab 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Seconds"
      Height          =   375
      Left            =   4380
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LblSecs 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      Height          =   375
      Left            =   3900
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Label LblGoing 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Auction Closes in"
      Height          =   375
      Left            =   1380
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label LblGoingTo 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      Height          =   375
      Left            =   3900
      TabIndex        =   1
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label LblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2513
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "FrmAuction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBid_Click(Index As Integer)
Call ReSetAuction(Index + 1, False)
End Sub

Private Sub Form_Activate()
Dim Ctrl As Control
Call SetCmdText(FrmAuction)
FrmAuction.BackColor = BrdColour
For Each Ctrl In FrmAuction.Controls
If Ctrl.Name Like "Lb*" Then
    Ctrl.BackColor = BrdColour
    Ctrl.ForeColor = LbcForeCol
End If
Next Ctrl
End Sub

Private Sub Form_Load()
Dim i As Integer
Plyr.MoveFirst
Do Until Plyr.EOF
    i = Plyr.Fields("Number") - 1
    If i <> -1 And i <> 98 Then
        If i = 0 Then
            LblGoingTo.Caption = Plyr.Fields("Name")
            ElseIf i <> 0 Then
            Load CmdBid(i)
            CmdBid(i).Visible = True
            CmdBid(i).Left = CmdBid(i - 1).Left + CmdBid(i - 1).Width + 20
            CmdBid(i).Top = CmdBid(i - 1).Top
            If CmdBid(i).Left + CmdBid(i).Width > FrmAuction.Width Then
                CmdBid(i).Top = CmdBid(i - 1).Top + CmdBid(i - 1).Height + 20
                CmdBid(i).Left = 240
            End If
        End If
    CmdBid(i).Caption = Plyr.Fields("Name") & " BID"
    CmdBid(i).Picture = FrmBoard.ImgCounter(i + 1).Picture
    CmdBid(i).Width = TextWidth(CmdBid(i).Caption) + 400
    End If
    Plyr.MoveNext
Loop
End Sub

Private Sub Timer1_Timer()
With FrmAuction
.LblSecs.Caption = Val(.LblSecs.Caption) - 1
If Val(.LblSecs.Caption) < 5 Then
    .LblGoing.FontBold = True
    .LblGoing.FontSize = 14
    .LblSecs.Visible = False
    .LblSecsLab.Visible = False
    .LblGoing.Alignment = 0
    .LblGoing.Caption = "Going"
    If Val(.LblSecs.Caption) < 3 Then .LblGoing.Caption = "Going, Going"
    If .LblSecs.Caption = "0" Then
        .LblGoing.Caption = "Gone"
        Call CloseAuction(.LblGoingTo.Caption)
        .Hide
    End If
End If
End With
End Sub
