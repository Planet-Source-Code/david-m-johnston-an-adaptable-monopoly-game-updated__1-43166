Attribute VB_Name = "ModDataBase"
Option Explicit

Private WrkJet As Workspace

Sub LoadDataBase()    'Load the DataBase
Dim StrFileName As String, Vers As String
Set DB = Nothing
On Error GoTo LoadError
Set WrkJet = CreateWorkspace("", "admin", "")
Set DB = WrkJet.OpenDatabase(App.Path & "\Data")
DBPath = App.Path & "\Data"
Exit Sub
LoadError:
MsgBox Err.Description, vbCritical, Err.Number
End
End Sub

Public Function SetRecordSets()     'Set RecordSet Variables

Set Prop = DB.OpenRecordset("Property")
Set PropSet = DB.OpenRecordset("PropertySet")
Set Plyr = DB.OpenRecordset("Player")
Set Counter = DB.OpenRecordset("Counter")
Set Chnce = DB.OpenRecordset("Chance")
Set CChest = DB.OpenRecordset("ComChest")
End Function

Public Sub DBAddPlayer(ByVal CounterNumb)       'Add a Player to DataBase
On Error GoTo AddError
Counter.Index = "Number"
Counter.Seek "=", CounterNumb

With FrmPlayers
Plyr.AddNew
Plyr.Fields("Number") = PlyrAdd
Plyr.Fields("Name") = .TxtPlayerName.Text
Plyr.Fields("CounterPath") = Counter.Fields("FilePath")
Plyr.Fields("Square") = 1
Plyr.Fields("Salary") = Salary
Plyr.Fields("Money") = PlyrStartMon
Plyr.Update
Exit Sub
End With

AddError:
    MsgBox Err.Description, vbExclamation, Err.Number
End Sub

Public Sub ReSetDB()         'Re-Set DataBase
On Error GoTo DeleteError

If Not (Prop.BOF And Prop.EOF) Then
Prop.MoveFirst
Do Until Prop.EOF       'Go through all Records in Properties Table
    Prop.Delete
    Prop.MoveNext
Loop
End If

If Not (PropSet.BOF And PropSet.EOF) Then 'PropSet.RecordCount <> 0 Then
PropSet.MoveFirst
Do Until PropSet.EOF
    PropSet.Delete
    PropSet.MoveNext
Loop
End If

If Not (Chnce.BOF And Chnce.EOF) Then 'Chnce.RecordCount <> 0 Then
Chnce.MoveFirst
Do Until Chnce.EOF      'Go through all Chance cards
    Chnce.Delete
    Chnce.MoveNext
Loop
End If

If Not (CChest.BOF And CChest.EOF) Then 'CChest.RecordCount <> 0 Then
CChest.MoveFirst
Do Until CChest.EOF      'Go through all Community Chest cards
    CChest.Delete
    CChest.MoveNext
Loop
End If

Plyr.MoveFirst
Do Until Plyr.EOF        'Go through all records in Players Table
    If Plyr.Fields("Number") = 99 Then
        Plyr.Edit
        Plyr.Fields("Money") = BnkStartMon
        Plyr.Fields("Salary") = 0
        Plyr.Update
    ElseIf Plyr.Fields("Number") = 0 Then
        Plyr.Edit
        Plyr.Fields("Money") = 0
        Plyr.Fields("Salary") = 0
        Plyr.Update
    Else
        Plyr.Delete
    End If
Plyr.MoveNext
Loop
FrmBoard.CboViewPlayer.Clear

Exit Sub
DeleteError:
    MsgBox Err.Description, vbExclamation, Err.Number
End Sub

Public Function GetPlayerMoney(ByVal Player) As Single
    'Receives Player number & Returns Money Owned
Plyr.Index = "Number"
Plyr.Seek "=", Player
GetPlayerMoney = Plyr.Fields("Money")
End Function

Public Function PlyrMoney(ByVal Player, ByVal Cash)
    'Receives Player Number & Money to be added (-ve amount if to be deducted)
        'and updates DataBase
Plyr.Index = "Number"
Plyr.Seek "=", Player
Plyr.Edit
Plyr.Fields("Money") = Plyr.Fields("Money") + Cash
Plyr.Update
End Function

Public Function PlayerSquare(ByVal Player) As Integer
    'Receives Player Number & Returns the square they are on
Plyr.Index = "Number"
Plyr.Seek "=", Player
PlayerSquare = Plyr.Fields("Square")
End Function
