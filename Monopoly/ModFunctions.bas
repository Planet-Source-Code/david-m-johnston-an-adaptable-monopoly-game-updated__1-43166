Attribute VB_Name = "ModFunctions"
Option Explicit

Public Function PosX(ByVal Sqre) As Integer
    'Receives the Square (Sqre) and sets XPos to appropriate X co-ordinate
        'SqBShort = length of short side of squares on bottom
Select Case Sqre
Case 1: XPos = FWidth - Corner
Case 2 To 10: XPos = FWidth - Corner - (Sqre - 1) * SqBShort
Case 11 To 21: XPos = 0
Case 22 To 31: XPos = Corner + ((Sqre - 22) * SqBShort)
Case 31 To 40: XPos = FWidth - Corner - LowRes
End Select

End Function

Public Function PosY(ByVal Sqre) As Integer
    'Receives the Square (Sqre) and sets YPos to appropriate Y co-ordinate
        'SqSShort = length of short side of squares on side
Select Case Sqre
Case 1 To 11: YPos = FHeight - Corner
Case 12 To 20: YPos = FHeight - Corner - (Sqre - 11) * SqSShort
Case 21 To 31: YPos = 0
Case 32 To 40: YPos = Corner + ((Sqre - 32) * SqSShort)
End Select

End Function

Public Function SelectedProperty()
    'Returns name of property selected for trade action
Dim n As Integer
SelectedProperty = FrmTrade.LstPlayerProp.Text
End Function

Public Function SetOwned(ByVal s) As Boolean
    'Determines if square (s) is part of a set owned by the same player
        'Reterns true if yes, False if no
Dim i, SetNo, OwnerNumb As Integer

Prop.Index = "Number"
Prop.Seek "=", s
OwnerNumb = Prop.Fields("OwnerNo")
SetNo = Prop.Fields("Set")
SetOwned = True
Prop.MoveFirst

Do Until Prop.EOF   'Check all properties
    If Prop.Fields("Set") = SetNo Then
        If Prop.Fields("OwnerNo") <> OwnerNumb Then
            SetOwned = False
        End If
    End If
Prop.MoveNext
Loop
End Function

Public Function HousesOnSet(ByVal s) As Boolean
    'Determines if square (s) is part of a set owned by the same player
        'and if any property in the set has houses/hotels
        'Reterns true if houses exist, False if none
Dim SetNo As Integer
Prop.Index = "Number"
Prop.Seek "=", s
SetNo = Prop.Fields("Set")
HousesOnSet = False
Prop.MoveFirst

Do Until Prop.EOF   'Check all properties
    If Prop.Fields("Set") = SetNo Then
        If Prop.Fields("HousesOwned") > 0 Then
            HousesOnSet = True
            Exit Do
        End If
    End If
Prop.MoveNext
Loop
End Function

Public Function EvenBuildSell(ByVal s As Integer, ByVal Build As Boolean) As Boolean
    'Determines if square (s) has more houses than any other
    'property in the set and reterns False if yes, True if no
Dim i, SetNo As Integer, OwnedHses As Integer
Prop.Index = "Number"
Prop.Seek "=", s
i = Prop.Fields("HousesOwned")
If i = 5 Then i = HotelNo
SetNo = Prop.Fields("Set")
EvenBuildSell = True
Prop.MoveFirst

Do Until Prop.EOF   'Check all properties
    If Prop.Fields("Set") = SetNo Then
        OwnedHses = Prop.Fields("HousesOwned")
        If OwnedHses = 5 Then OwnedHses = HotelNo
        If (Build And i > OwnedHses) Or _
        (Build = False And i < OwnedHses) Then
            EvenBuildSell = False
            Exit Do
        End If
    End If
Prop.MoveNext
Loop
End Function

Public Function MortgInSetSet(ByVal s) As Boolean
    'Determines if square (s) is part of a set owned by the same player
        'and if any property in that set is mortgaged
        'Reterns true if yes, False if none
Dim i, OwnerNumb, SetNo As Integer
Prop.Index = "Number"
Prop.Seek "=", s
SetNo = Prop.Fields("Set")
MortgInSetSet = False
Prop.MoveFirst

Do Until Prop.EOF   'Check all properties
    If Prop.Fields("Set") = SetNo Then
        If Prop.Fields("Mortgaged") = True Then
            MortgInSetSet = True
            Exit Do
        End If
    End If
Prop.MoveNext
Loop
End Function

Public Sub MovePlayer(ByVal FinalSquare)
    'Moves players' counter forwards & updates database
Dim OldSqur, Move, s, i, n As Integer
Dim Start

ViewPlayer = CurPlayer

If CurPlayer = 0 Then
    n = MsgBox("Please Click Options to Enter Players", vbCritical, "Players")
    Exit Sub
End If

Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
Plyr.Edit
OldSqur = Plyr.Fields("Square")
Plyr.Fields("Square") = FinalSquare
Plyr.Update

    If OldSqur < FinalSquare Then
    For s = OldSqur + 1 To FinalSquare
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + CounterPause / 1000 'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    Else: For s = OldSqur To 40
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + CounterPause / 1000 'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    For s = 1 To FinalSquare
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + CounterPause / 1000 'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    End If
End Sub

Public Sub MovePlayerBack(FinalSquare)
    'Moves players' counter backwards & updates database
Dim OldSqur, Move, s, i, n As Integer
Dim Start

ViewPlayer = CurPlayer

Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
Plyr.Edit
OldSqur = Plyr.Fields("Square")
Plyr.Fields("Square") = FinalSquare
Plyr.Update

    If OldSqur > FinalSquare Then
    For s = OldSqur To FinalSquare Step -1
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + CounterPause / 1000 + 0.1 'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    Else: For s = OldSqur To 1 Step -1
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + CounterPause / 1000 + 0.1 'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    For s = 40 To FinalSquare Step -1
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + CounterPause / 1000 + 0.1 'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    End If
End Sub

Public Sub PositionPlayer(ByVal s)
    'Moves counter to new square (s)
PosX (s)
PosY (s)
If s >= 1 And s <= 11 Or s >= 21 And s <= 31 Then
    FrmBoard.ImgCounter(CurPlayer).Move XPos + SqBShort / 3, YPos + Corner / 2
Else
    FrmBoard.ImgCounter(CurPlayer).Move XPos + Corner / 6, YPos + SqSShort / 6
End If
End Sub

Public Sub BuyHouse(ByVal Numb)
    'Update DataBase & board when a player buys a house
Dim HousesOwned, SetNo As Integer

Prop.Index = "Number"
Prop.Seek "=", Numb
PropSet.Index = "Number"
PropSet.Seek "=", Prop.Fields("Set")
SetNo = PropSet.Fields("Number")
Prop.Edit
If Prop.Fields("HousesOwned") = HotelNo - 1 Then
    Prop.Fields("HousesOwned") = 5
Else
    Prop.Fields("HousesOwned") = Prop.Fields("HousesOwned") + 1
End If
Prop.Update
Call DrawHouses(Numb, HousesOwned, SetNo)
    'Draw houses (HousesOwned) on Square (Numb)
Call PlyrMoney(CurPlayer, -PropSet.Fields("HousePrice"))
    'Reduce players' money
Call PlyrMoney(99, PropSet.Fields("HousePrice"))
    'Increase Banks' money

End Sub

Public Sub SellHouse(Numb)
    'Update DataBase & board when a player sells a house
Dim HousesOwned, SetNo As Integer

Prop.Index = "Number"
Prop.Seek "=", Numb
PropSet.Index = "Number"
PropSet.Seek "=", Prop.Fields("Set")
SetNo = PropSet.Fields("Number")
Prop.Edit
If Prop.Fields("HousesOwned") >= HotelNo Then
    Prop.Fields("HousesOwned") = Prop.Fields("HousesOwned") - 1 - (5 - HotelNo)
Else
    Prop.Fields("HousesOwned") = Prop.Fields("HousesOwned") - 1
End If
Prop.Update
Call DrawHouses(Numb, HousesOwned, SetNo)
    'Draw houses (HousesOwned) on Square (Numb)
Call PlyrMoney(CurPlayer, PropSet.Fields("HousePrice") / 2)
    'Increase players' money
Call PlyrMoney(99, -PropSet.Fields("HousePrice") / 2)
    'Reduce Banks' money
End Sub

Public Sub Stations(ByVal SetNum)
    'Set rent for stations owned by Player according to number owned
Dim RentOwed As Single: Dim Player, i, Count As Integer

Plyr.Index = "Number"
Plyr.MoveFirst
Do Until Plyr.EOF
    Player = Plyr.Fields("Number")
    Count = 0
    If Player <> 0 Then
        Prop.Index = "Number"
        Prop.MoveFirst
        Do Until Prop.EOF   'Count number owned by Player
            If Prop.Fields("Set") = SetNum And Prop.Fields("OwnerNo") = Player Then
                Count = Count + 1
            End If
        Prop.MoveNext
        Loop

        If Count > 0 Then
        Prop.MoveFirst
        Do Until Prop.EOF   'Update Rent
            If Prop.Fields("Set") = SetNum And Prop.Fields("OwnerNo") = Player Then
                Prop.Edit
                Prop.Fields("HousesOwned") = Count - 1
                Prop.Update
            End If
            Prop.MoveNext
        Loop
        End If
    End If
    Plyr.MoveNext
Loop
End Sub

Public Sub MissTurn(ByVal Numb)
    'Set number (Numb) of turns to be missed by curent player
Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
Plyr.Edit
Plyr.Fields("MissTurns") = Numb
Plyr.Update
End Sub

Public Sub TurnMissed(ByVal Miss)
    'Reduced number of turns to be missed by curent player by 1
Dim n As Integer
Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
Plyr.Edit
Plyr.Fields("MissTurns") = Miss - 1
Plyr.Update
n = MsgBox(Plyr.Fields("Name") & " to Miss " & Miss - 1 _
    & " more turns", vbInformation, "Miss a Turn")
End Sub

Public Sub LowMoney()
    'Current player can't afford rent owed
Dim n As Integer: Dim PlayerMoney, TotAssets As Single

Prop.Index = "Number"
PropSet.Index = "Number"
Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
PlayerMoney = GetPlayerMoney(CurPlayer)
TotAssets = PlayerMoney

Do Until Prop.EOF   'Check if player bankrupt - Can't raise enough money by selling assets
    If Prop.Fields("OwnerNo") = CurPlayer Then
        If Prop.Fields("HousesOwned") > 0 Then
            PropSet.Seek "=", Prop.Fields("Set")
            TotAssets = TotAssets + Prop.Fields("HousesOwned") * (PropSet.Fields("HousePrice") / 2)
        End If
        If Prop.Fields("Mortgaged") = True Then _
            TotAssets = TotAssets + ((Prop.Fields("Price") / 2) * 1.1)
    TotAssets = TotAssets + Prop.Fields("Price")
    End If
Prop.MoveNext
Loop

If TotAssets < AmountOwed Then  'Player is bankrupt
    n = MsgBox(Plyr.Fields("Name") & " is BANKRUPT", vbExclamation)
    Call RemovePlayer(CurPlayer)    'Player leaves game
    If TotPlayers < 2 Then
        Plyr.MoveFirst
        Plyr.MoveNext
        n = MsgBox(Plyr.Fields("Name") & " WINS", vbInformation, "Winner")
    End If
    Exit Sub
Else
LowMon = True
Call Trading    'Go back to trade options to sell more assets
End If
End Sub

Public Sub RemovePlayer(ByVal Player As Integer)
    'Bankrupt player leaves game
    
Dim n, s, Recipient, PropertySet As Integer

Plyr.Index = "Number"
Plyr.Seek "=", Player
Prop.Index = "Number"
PropSet.Index = "Number"
s = Plyr.Fields("Square")
Prop.Seek "=", s
Recipient = Prop.Fields("OwnerNo")
CChest.Index = "Number"
Chnce.Index = "Number"

CChest.MoveFirst
Do Until CChest.EOF
    If CChest.Fields("Action") = "Get Out of " & Jail Then
        CChest.Edit
        CChest.Fields("Owner") = Recipient
        CChest.Update
    End If
    CChest.MoveNext
Loop

Chnce.MoveFirst
Do Until Chnce.EOF
    If Chnce.Fields("Action") = "Get Out of " & Jail Then
        Chnce.Edit
        Chnce.Fields("Owner") = Recipient
        Chnce.Update
    End If
    Chnce.MoveNext
Loop

Prop.MoveFirst
Do Until Prop.EOF   'Property transferred to player/Bank who is owed money
    If Prop.Fields("OwnerNo") = Player Then
        PropertySet = Prop.Fields("Set")
        PropSet.Seek "=", PropertySet
        Prop.Edit
        If Recipient = 99 Then
            Prop.Fields("Mortgaged") = False
            Prop.Fields("HousesOwned") = 0
        End If
        Prop.Fields("OwnerNo") = Recipient
        Prop.Update
    End If
    Prop.MoveNext
Loop
Call PlyrMoney(Recipient, (Plyr.Fields("Money")))
    'Bankrupt palyers money transferred to owed player
Call Stations(9)
Call Stations(10)
Plyr.Index = "Number"
Plyr.Seek "=", (Player)
Plyr.Delete
FrmBoard.ImgCounter(Player).Visible = False
FrmBoard.CboViewPlayer.RemoveItem (Player - 1)
FrmBoard.CboViewPlayer.Refresh
Plyr.MoveNext
If Plyr.Fields("Number") = 99 Then
    Plyr.MoveFirst
    Plyr.MoveNext
End If
n = Plyr.Fields("Number")
SetCurPlayer (n)

CurPlayer = GetCurPlayer
ViewPlayer = CurPlayer
TotPlayers = TotPlayers - 1
Dice2 = 7
FrmBoard.LblInfo.Caption = Plyr.Fields("Name") & " To Go"
Call NextPlayer
Call UpdateBoard
End Sub

Public Sub EndGame()    'End Programme
If MsgBox("Are You Sure you want to quit?", 36, "Quit?") = 6 Then
    If MsgBox("Do you want to save your game?", 36, "Save Game?") = 6 Then
        Call SaveGame
        End
    End If
    End
End If
End Sub

Public Sub SaveGame()
Dim IntFile As Integer
Dim StrFileName As String
On Error GoTo ErrorCheck
IntFile = FreeFile
FrmBoard.CmdFiles.Filter = ("Saved Games|*.mon")
FrmBoard.CmdFiles.ShowSave
StrFileName = FrmBoard.CmdFiles.FileName
Open StrFileName For Output As #IntFile
Write #IntFile, VersName
Plyr.MoveFirst
Do Until Plyr.EOF
    Write #IntFile, Plyr.Fields("Number"), Plyr.Fields("Name"), Plyr.Fields("CounterPath") _
    , Plyr.Fields("Money"), Plyr.Fields("Salary"), Plyr.Fields("Square"), Plyr.Fields("MissTurns"), Plyr.Fields("CurPlayer")
    Plyr.MoveNext
Loop
Prop.MoveFirst
Do Until Prop.EOF
    Write #IntFile, Prop.Fields("Number"), Prop.Fields("OwnerNo") _
    , Prop.Fields("Mortgaged"), Prop.Fields("HousesOwned")
    Prop.MoveNext
Loop
Close IntFile

ErrorCheck:
Exit Sub
End Sub

Public Sub LoadGame()
Dim StrFileName As String, Name As String, CounterPath As String
Dim CurPlyr As Boolean, Mortgaged As Boolean
Dim i As Integer, Current As String, IntFile As Integer
Dim Number As String, Square As String, Money As String, Sal As String
Dim Miss As String, OwnerNo As String, HousesOwned As String
Call ReSetDB
On Error GoTo ErrorCheck
IntFile = FreeFile
FrmBoard.CmdFiles.Filter = ("Saved Games|*.mon")
FrmBoard.CmdFiles.ShowOpen
StrFileName = FrmBoard.CmdFiles.FileName

FrmBoard.CboViewPlayer.Clear
Open StrFileName For Input As #IntFile
Input #IntFile, VersName
Close IntFile
Call LoadVersion
Plyr.MoveFirst
Do Until Plyr.EOF
    Plyr.Delete
    Plyr.MoveNext
Loop
Call DrawBoard
FrmBoard.CboViewPlayer.Clear
Open StrFileName For Input As #IntFile
Input #IntFile, VersName
Do Until Val(Number) = 99
    Input #IntFile, Number, Name, CounterPath, Money, Sal, Square, Miss, CurPlyr
    Plyr.AddNew
    Plyr.Fields("Number") = Val(Number): Plyr.Fields("Name") = Name
    Plyr.Fields("CounterPath") = CounterPath: Plyr.Fields("Money") = Val(Money)
    Plyr.Fields("Salary") = Sal: Plyr.Fields("Square") = Val(Square)
    Plyr.Fields("MissTurns") = Val(Miss): Plyr.Fields("CurPlayer") = CurPlyr
    If Val(Number) > 0 And Val(Number) < 99 Then
        FrmBoard.CboViewPlayer.AddItem (Name)
        If Number > 1 Then Load FrmBoard.ImgCounter(Number)
        FrmBoard.ImgCounter(Number).Picture = LoadPicture(App.Path & CounterPath)
        CurPlayer = Val(Number)
        PositionPlayer (Val(Square))
        If Plyr.Fields("CurPlayer") = True Then
            Current = Val(Number)
            FrmBoard.LblInfo.Caption = Name & " To Go"
        End If
    End If
Plyr.Update
Loop
Prop.Index = "Number"
Do Until i = 99
    Input #IntFile, Number, OwnerNo, Mortgaged, HousesOwned
    Prop.Seek "=", Val(Number)
    Prop.Edit
    Prop.Fields("OwnerNo") = Val(OwnerNo): Prop.Fields("Mortgaged") = Mortgaged
    Prop.Fields("HousesOwned") = Val(HousesOwned)
    Prop.Update
    If Prop.Fields("Number") = 40 Then i = 99
Loop
Close IntFile
Call ModOptions.LoadDB
CurPlayer = Val(Current)
ViewPlayer = CurPlayer
FrmOptions.LbCPlayers.ForeColor = &H80000012
FrmOptions.LbCPlayers.Enabled = True
Call BankProperty
Call UpdateBoard
Call UpdateHouses
ErrorCheck:
Exit Sub
End Sub

Public Sub SaveVersion()
Dim IntFile As Integer
Dim Vers As String, StrFileName As String
VersName = InputBox("Please enter a name for this version", "Version Name", "New Vers " & Now)
On Error GoTo ErrorCheck
IntFile = FreeFile
StrFileName = App.Path & "\UserVers.dat"
Open StrFileName For Append As #IntFile
Write #IntFile,
Write #IntFile, "." & VersName
Write #IntFile, House, Hotel, Go, Jail, FParking, Bank, PropInfo, Rent, Utility, Station, _
CurrencySymb, CurSymbPos, ChanceNme, CommChestNme
Write #IntFile, BnkStartMon, PlyrStartMon, Salary, BrdColour, CounterPause, TextColour, _
JailFine, JailDoubles, BuildEven, SellEven, StartProps, FreePark, HotelNo
Write #IntFile, FontName, ForeCol, FontSize, FontBold, FontItalic, FontUline, FontStrThru

PropSet.MoveFirst
Do Until PropSet.EOF
    Write #IntFile, PropSet.Fields("Number"), PropSet.Fields("Colour"), PropSet.Fields("HousePrice") _
    , PropSet.Fields("HouseColour"), PropSet.Fields("HotelColour")
    PropSet.MoveNext
Loop

Prop.MoveLast
Write #IntFile, Prop.Fields("Number")
Prop.MoveFirst
Do Until Prop.EOF
    Write #IntFile, Prop.Fields("Number"), Prop.Fields("Name"), Prop.Fields("Set"), Prop.Fields("Price"), Prop.Fields("OwnerNo"), Prop.Fields("Rent"), Prop.Fields("Rent1"), Prop.Fields("Rent2"), Prop.Fields("Rent3"), Prop.Fields("Rent4"), Prop.Fields("Rent5")
    Prop.MoveNext
Loop

Chnce.MoveLast
Write #IntFile, Chnce.Fields("Number")
Chnce.MoveFirst
Do Until Chnce.EOF
    Write #IntFile, Chnce.Fields("Number"), Chnce.Fields("Text"), _
    Chnce.Fields("Action"), Chnce.Fields("Amount")
    Chnce.MoveNext
Loop

CChest.MoveLast
Write #IntFile, CChest.Fields("Number")
CChest.MoveFirst
Do Until CChest.EOF
    Write #IntFile, CChest.Fields("Number"), CChest.Fields("Text"), _
    CChest.Fields("Action"), CChest.Fields("Amount")
    CChest.MoveNext
Loop
Close IntFile

ErrorCheck:
Exit Sub
End Sub

Public Sub LoadVersion()
Dim IntFile As Integer, i As Integer, Numb1 As String
Dim Numb2 As String, Numb3 As String, Prce As String
Dim Numbs(9) As String, Count As String
Dim Vers As String, Str(11) As String
Dim StrFileName As String, CounterPath As String
IntFile = FreeFile
StrFileName = App.Path & "\Versions.dat"
Call ReSetDB
Do Until Vers = "." & VersName
Open StrFileName For Input As #IntFile
Do Until EOF(IntFile) Or Vers = VersName
    Input #IntFile, Vers
    If Vers = "." & VersName Then Exit Do
Loop
If EOF(IntFile) Then Close IntFile
StrFileName = App.Path & "\UserVers.dat"
Loop
Input #IntFile, House, Hotel, Go, Jail, FParking, Bank, PropInfo, Rent, Utility, Station, _
CurrencySymb, CurSymbPos, ChanceNme, CommChestNme
Input #IntFile, Str(1), Str(2), Str(3), BrdColour, Str(4), TextColour, _
Str(5), Str(6), Str(7), Str(8), Str(9), Str(10), Str(11)
Input #IntFile, FontName, ForeCol, FontSize, FontBold, FontItalic, FontUline, FontStrThru
BnkStartMon = Val(Str(1)): PlyrStartMon = Val(Str(2)): Salary = Val(Str(3)): CounterPause = Val(Str(4))
JailFine = Val(Str(5)): JailDoubles = Val(Str(6)): StartProps = Val(Str(9))
FreePark = Val(Str(10)): HotelNo = Val(Str(11))
If Str(7) = "YES" Then BuildEven = True Else BuildEven = False
If Str(8) = "YES" Then SellEven = True Else SellEven = False
Do Until Numb1 = "10"
    Input #IntFile, Numb1, Str(1), Prce, Str(2), Str(3)
    PropSet.AddNew
    PropSet.Fields("Number") = Val(Numb1): PropSet.Fields("Colour") = Str(1)
    PropSet.Fields("HousePrice") = Val(Prce): PropSet.Fields("HouseColour") = Str(2)
    PropSet.Fields("HotelColour") = Str(3)
    PropSet.Update
Loop

Input #IntFile, Count
Do Until Numb1 = Count
    Input #IntFile, Numb1, Str(1), Numbs(0), Numbs(1), Numbs(2), Numbs(3), _
    Numbs(4), Numbs(5), Numbs(6), Numbs(7), Numbs(8)
    Prop.AddNew
    Prop.Fields("Number") = Val(Numb1): Prop.Fields("Name") = Str(1)
    For i = 0 To 8
        Prop.Fields(i + 3) = Val(Numbs(i))
    Next i
    Prop.Fields("HousesOwned") = 0
    Prop.Update
Loop

Input #IntFile, Count
Do Until Numb1 = Count
    Input #IntFile, Numb1, Str(1), Str(2), Numb2
    Chnce.AddNew
    Chnce.Fields("Number") = Val(Numb1): Chnce.Fields("Text") = Str(1)
    Chnce.Fields("Action") = Str(2): Chnce.Fields("Amount") = Val(Numb2)
    Chnce.Fields("Owner") = 0: If Str(2) = "Get Out of Jail" Then Chnce.Fields("Owner") = 99
    Chnce.Update
Loop

Numb1 = 0
Input #IntFile, Count
Do Until Numb1 = Count
    Input #IntFile, Numb1, Str(1), Str(2), Numb2
    CChest.AddNew
    CChest.Fields("Number") = Val(Numb1): CChest.Fields("Text") = Str(1)
    CChest.Fields("Action") = Str(2): CChest.Fields("Amount") = Val(Numb2)
    CChest.Fields("Owner") = 0: If Str(2) = "Get Out of Jail" Then CChest.Fields("Owner") = 99
    CChest.Update
Loop
Close IntFile
If CurSymbPos = "Before" Then
    CurSymbBefore = CurrencySymb
    CurSymbAfter = ""
Else
    CurSymbAfter = CurrencySymb
    CurSymbBefore = ""
End If
Call ModOptions.LoadDB
StartTm = Now
End Sub

Public Sub StartProperty(ByVal Numb)
Dim i As Integer, SetNo As Integer, p As Integer

Prop.Index = "Number"
Plyr.MoveFirst
Plyr.MoveNext
p = Plyr.Fields("Number")
i = 0
Do Until i = Numb Or i * TotPlayers > 28
    Plyr.MoveFirst
    Plyr.MoveNext
    p = Plyr.Fields("Number")
    i = i + 1
    Do Until p = 99 Or i * TotPlayers > 28
        Randomize
        Prop.Seek "=", Random(40)
        SetNo = Prop.Fields("Set")
        If SetNo > 0 And Prop.Fields("OwnerNo") = 99 Then
            If Plyr.Fields("Money") < Prop.Fields("Price") Then Exit Sub
            Prop.Edit
            Prop.Fields("OwnerNo") = p
            Prop.Update
            Call PlyrMoney(p, -(Prop.Fields("Price")))
            Call PlyrMoney(99, Prop.Fields("Price"))
            If SetNo = 9 Then Call Stations(9)
            If SetNo = 10 Then Call Stations(10)
            Plyr.Seek "=", p
            Plyr.MoveNext
            p = Plyr.Fields("Number")
        End If
    Loop
Loop
End Sub

Public Sub Duration()   'Update Elapsed Time
Dim TotTm As Integer, hrs As Integer, mins As Integer, secs As Integer
TotTm = DateDiff("s", StartTm, Now)
secs = TotTm Mod 60
hrs = TotTm \ 3600
mins = TotTm \ 60 - hrs
FrmBoard.LblDuration.Caption = hrs & ":" & mins & ":" & secs
End Sub

Public Function Random(ByVal Numb) As Integer    'Produce random numbers
Dim n As Integer
Randomize
    Random = Int(Numb * Rnd + 1)
End Function

Public Sub EnterDice()
Dim Values As String

Values = InputBox("Please enter TWO 1 digit numbers", "Enter Dice Numbers")
Dice1 = Val(Left$(Values, 1))
Dice2 = Val(Right$(Values, 1))
End Sub
