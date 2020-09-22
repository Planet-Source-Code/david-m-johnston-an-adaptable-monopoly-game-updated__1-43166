Attribute VB_Name = "ModPlayers"
Option Explicit

Public Sub CreateForm() 'Show available counters
Dim i As Integer
i = 0
FrmPlayers.LstPlayerNo.Clear
FrmPlayers.LstPlayers.Clear
PlyrAdd = 1
FrmPlayers.LblPlayerNumb.Caption = PlyrAdd

Counter.MoveFirst
Do Until Counter.EOF
    i = i + 1
    With FrmPlayers
    If i > 1 Then
        Load .ImgCounter(i)
        .ImgCounter(i).Left = .ImgCounter(i - 1).Left + .ImgCounter(i - 1).Width + 20
        .ImgCounter(i).Top = .ImgCounter(i - 1).Top
        If .ImgCounter(i).Left + .ImgCounter(i).Width > FrmPlayers.Width - 400 Then
            .ImgCounter(i).Top = .ImgCounter(i - 1).Top + .ImgCounter(i - 1).Height
            .ImgCounter(i).Left = 4320
        End If
    End If
    .ImgCounter(i).Picture = LoadPicture(App.Path & (Counter.Fields("FilePath")))
    .ImgCounter(i).Visible = True
    End With
    Counter.MoveNext
Loop
End Sub

Public Sub EnterPlyr()  'Enter a new player
Dim n As Integer
If FrmPlayers.TxtPlayerName.Text = "" Then
    n = MsgBox("Please Enter Players", vbCritical, "Players")
    Exit Sub

ElseIf TotPlayers > 5 Then
    n = MsgBox("Sorry you can't have more than 6 Players", vbInformation, "Players")
    Exit Sub
    
ElseIf FrmPlayers.ImgChosenCounter.Picture = 0 Then
    n = MsgBox("Please chose an icon", vbCritical, "Choose an Icon")
    Exit Sub

Else
Call EnterPlayer    'Add player to DataBase
FrmPlayers.TxtPlayerName.SetFocus
If TotPlayers > 0 Then
    Call DrawBoard
End If
    'Place players on "GO" square
PosX (1)
PosY (1)
Call ShowCounters
Plyr.Index = "Number"
CurPlayer = 1       'Player 1 starts
Plyr.Seek "=", CurPlayer
FrmBoard.LblInfo.Caption = Plyr.Fields("Name") & " To Go"
End If

End Sub

Public Sub EnterPlayer()
Dim i As Integer, n As Integer
If TotPlayers > 6 Then n = MsgBox("Sorry you can't have more than 6 Players", vbInformation, "Players")

If FrmPlayers.TxtPlayerName.Text = "" Then
    n = MsgBox("Please enter a name", vbCritical, "Enter a Name")
    Exit Sub

ElseIf FrmPlayers.ImgChosenCounter.Picture = 0 Then
    n = MsgBox("Please chose an icon", vbCritical, "Choose an Icon")
    Exit Sub

Else
Call DBAddPlayer(CounterNumb)   'Update DataBase
FrmBoard.CboViewPlayer.AddItem FrmPlayers.TxtPlayerName.Text

With FrmPlayers
.LstPlayerNo.AddItem PlyrAdd
.LstPlayers.AddItem FrmPlayers.TxtPlayerName.Text

.TxtPlayerName.Text = ""
.ImgCounter(CounterNumb).Visible = False
.ImgChosenCounter.Picture = LoadPicture("")
End With

Plyr.Index = "Number"
Plyr.Seek "=", PlyrAdd
If PlyrAdd > 1 Then Load FrmBoard.ImgCounter(PlyrAdd)
FrmBoard.ImgCounter(PlyrAdd).Picture = LoadPicture(App.Path & (Counter.Fields("FilePath")))
PlyrAdd = PlyrAdd + 1
TotPlayers = TotPlayers + 1
FrmPlayers.LblPlayerNumb.Caption = PlyrAdd
End If
End Sub

Public Sub Finished()  'finished entering players
Dim n As Integer
If Plyr.RecordCount < 1 Then    'No players entered
    n = MsgBox("Please enter players", vbCritical, "Players")
Exit Sub
End If
Unload FrmPlayers
Call StartProperty(StartProps)
FrmBoard.Show
ViewPlayer = 1
Call UpdateBoard    'Update board
FrmOptions.LbCPlayers.ForeColor = &H8000000F
FrmOptions.LbCPlayers.Enabled = False
FrmOptions.MnuEnterPlayers.Enabled = False
FrmOptions.LbCToGame.ForeColor = &H80000012
FrmOptions.LbCToGame.Enabled = True
FrmOptions.MnuBack.Enabled = True
End Sub
