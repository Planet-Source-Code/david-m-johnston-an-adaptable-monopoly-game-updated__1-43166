Attribute VB_Name = "ModGameFunctions"
Option Explicit

Public Sub NamePriceClicked(ByVal Indx)
    'Show Deed or, if Jail clicked, option to use Get out of jail card
Dim Nme As String
If Indx <> 11 Then
    Prop.Index = "Number"
    Prop.Seek "=", Indx
    Nme = Prop.Fields("Name")
    If Prop.Fields("OwnerNo") <> 0 Then Call Deed(Nme)  'Show Title Deed
Else: Call GetOutOfJail
End If
End Sub

Public Sub GetOutOfJail()
Dim n As Integer
Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
Chnce.Index = "Action"
Chnce.Seek "=", "Get Out of Jail"
CChest.Index = "Action"
CChest.Seek "=", "Get Out of Jail"
Plyr.Edit
If Chnce.Fields("Owner") = CurPlayer Then 'If current player owns Get Out of Jail card
    If MsgBox("Are you Sure you want to use your " & vbCrLf & _
        "Get Out of " & Jail & " Free Card?", 36, "") = 6 Then
        Chnce.Edit
        Chnce.Fields("Owner") = 99
        Chnce.Update
        Plyr.Fields("MissTurns") = 0
    End If
ElseIf CChest.Fields("Owner") = CurPlayer Then 'If current player owns Get Out of Jail card
        If MsgBox("Are you Sure you want to use your " & vbCrLf & _
        "Get Out of " & Jail & " Free Card?", 36, "") = 6 Then
        CChest.Edit
        CChest.Fields("Owner") = 99
        CChest.Update
        Plyr.Fields("MissTurns") = 0
        End If
Else: n = MsgBox("Sorry, You don't own a" & vbCrLf & "Get Out Of Jail Free Card", vbInformation, "Use Card")
End If
Plyr.Update
End Sub

Public Sub Deed(ByVal Name) 'Show Ttile Deed
Dim i As Integer, OwnersSet As Integer, OwnerNo As Integer, s As Integer
Dim Ctrl As Control: Dim HasSet As Boolean, Houses As Boolean, Motgd As Boolean

Chnce.MoveFirst
    Do Until Chnce.EOF
        If Chnce.Fields("Text") = Name Then
        FrmCard.LblText = Chnce.Fields("Text")
        FrmCard.Show
        Exit Sub
        Else
        Chnce.MoveNext
        End If
    Loop
If Chnce.EOF Then
    Do Until CChest.EOF
        If CChest.Fields("Text") = Name Then
        FrmCard.LblText = CChest.Fields("Text")
        FrmCard.Show
        Exit Sub
        Else
        CChest.MoveNext
        End If
    Loop
End If
Prop.Index = ("Name")
Prop.Seek "=", Name
s = Prop.Fields("Number")
HasSet = SetOwned(s)
Houses = HousesOnSet(s)
Motgd = MortgInSetSet(s)
PropSet.Index = "Number"
Prop.Index = "Name"
Prop.Seek "=", Name
If Prop.Fields("Set") = 0 Then Exit Sub
Plyr.Index = "Number"
Plyr.Seek "=", Prop.Fields("OwnerNo")
OwnerNo = Prop.Fields("OwnerNo")
PropSet.Seek "=", (Prop.Fields("Set"))
OwnersSet = Prop.Fields("Set")
With FrmProperty

.Caption = PropInfo
.LblName = Prop.Fields("Name")
If .TextWidth(.LblName.Caption) > .LblName.Width * 1.8 Then
    .LblName.FontSize = 11
Else
    .LblName.FontSize = 14
End If

    If OwnersSet <> 9 Then  'Not a Company
        For i = 0 To 6
            .LblRentHouses(i).Visible = True
            .LblRent(i).Visible = True
        Next i
        .LblHPriceLab.Visible = True
        .LblHPrice.Visible = True
        .LblEachLab.Visible = True
        .LblRent(0) = CurSymbBefore & Prop.Fields("Rent") & CurSymbAfter
        If OwnersSet < 9 And Houses = False And HasSet = True And OwnerNo <> 99 Then .LblRent(0) = CurSymbBefore & Prop.Fields("Rent") * 2 & CurSymbAfter
        .LblRent(1) = CurSymbBefore & Prop.Fields("Rent1") & CurSymbAfter
        .LblRent(2) = CurSymbBefore & Prop.Fields("Rent2") & CurSymbAfter
        .LblRent(3) = CurSymbBefore & Prop.Fields("Rent3") & CurSymbAfter
        .LblSet9.Visible = False
    
    If OwnersSet <> 10 Then 'Not a Station
        .LblName.ForeColor = &HFFFFFF
        .LblName.BackColor = Val(PropSet.Fields("Colour"))
        If OwnersSet < 9 And HasSet = True And OwnerNo <> 99 And Houses = False And Motgd = False Then
        .LblRentHouses(0).Caption = Rent & " - Site only (Set Owned)"
        Else: .LblRentHouses(0).Caption = Rent & " - Site only"
        End If
        .LblRentHouses(1).Caption = "      ""      With 1 " & House
        .LblRentHouses(2).Caption = "      ""          ""  2 " & House & "s"
        .LblRentHouses(3).Caption = "      ""          ""  3 " & House & "s"
        .LblRentHouses(4).Caption = "      ""          ""  4 " & House & "s"
        .LblRentHouses(5).Caption = "      ""          ""  " & Hotel
        .LblRent(4) = CurSymbBefore & Prop.Fields("Rent4") & CurSymbAfter
        .LblRent(5) = CurSymbBefore & Prop.Fields("Rent5") & CurSymbAfter
        .LblHPrice.Caption = CurSymbBefore & PropSet.Fields("HousePrice") & CurSymbAfter
        .LblSet.Visible = True
    Else
        .LblName.BackColor = &HFFFFFF
        .LblName.ForeColor = &H0&
        .LblRentHouses(0).Caption = "Rent"
        .LblRentHouses(1).Caption = "If 2 " & Station & "s are owned"
        .LblRentHouses(2).Caption = "If 3        ""       ""       """
        .LblRentHouses(3).Caption = "If 4        ""       ""       """
        .LblRentHouses(4).Visible = False
        .LblRentHouses(5).Visible = False
        .LblRent(4).Visible = False
        .LblRent(5).Visible = False
        .LblSet.Visible = False
        .LblHPriceLab.Visible = False
        .LblHPrice.Visible = False
        .LblEachLab.Visible = False
    End If
Else
    .LblName.BackColor = &HFFFFFF
    .LblName.ForeColor = &H0&
    For i = 0 To 5
        .LblRentHouses(i).Visible = False
        .LblRent(i).Visible = False
    Next i
    .LblSet.Visible = False
    .LblSet9.Caption = "If one " & Utility & " is owned rent is " & CurSymbBefore & Prop.Fields("Rent") & CurSymbAfter & _
    " times amount shown on dice 1. If both " & Utility & "s are owned " & Rent & " is " & CurSymbBefore & _
    Prop.Fields("Rent1") & CurSymbAfter & " times amount shown on dice 1."
    .LblSet9.Visible = True
    .LblHPriceLab.Visible = False
    .LblHPrice.Visible = False
    .LblEachLab.Visible = False
End If
.LblRent(6) = CurSymbBefore & Prop.Fields("Price") / 2 & CurSymbAfter
.LblOwner.Caption = Plyr.Fields("Name")

    If Prop.Fields("Mortgaged") = True Then 'Pink text on grey background
                                                'if property mortgaged
        .BackColor = &HE0E0E0
        For Each Ctrl In .Controls
            If Ctrl.Name Like "Lbl*" And Ctrl.Name <> "LblName" Then
                Ctrl.ForeColor = &H8080FF
            End If
        Next Ctrl
    Else
    .BackColor = &HFFFFFF
        For Each Ctrl In .Controls
            If Ctrl.Name Like "Lbl*" And Ctrl.Name <> "LblName" Then
                Ctrl.ForeColor = &H80000012
            End If
        Next Ctrl
        If OwnerNo <> 99 Then .LblRentHouses(Prop.Fields("HousesOwned")).ForeColor = vbRed
        If OwnerNo <> 99 Then .LblRent(Prop.Fields("HousesOwned")).ForeColor = vbRed
    End If
.Show
End With
End Sub

Public Sub Cards(ByVal Action, ByVal Amnt)
    'Receives Action & Amount from Chance or CommChest
    'Performs Action
        
Dim s As Integer, p As Integer, n As Integer: Dim Choice As String
Dim Amount As Single, SNo As Integer
If InStr(Amnt, Station) > 0 Then SNo = 10
If InStr(Amnt, Utility) > 0 Then SNo = 9
If Amnt <> Station And Amnt <> Utility Then Amount = Val(Amnt)
Plyr.Index = "Number"
Prop.Index = "Number"
s = PlayerSquare(CurPlayer) 'Get square current player is on

Select Case Action

Case "Receive From Bank"
    Call PlyrMoney(CurPlayer, Amount)
    Call PlyrMoney(99, -Amount)
    Call EndTurn

Case "Receive From All Players"
    For p = 1 To TotPlayers
        If p <> CurPlayer Then
            Call PlyrMoney(p, -Amount)
            Call PlyrMoney(CurPlayer, Amount)
            Call EndTurn
        End If
    Next p
    
Case "Pay To Bank"
    Call PlyrMoney(CurPlayer, -Amount)
    Call PlyrMoney(FreePark, Amount)
    Call EndTurn
    
Case "General Repairs"
Prop.MoveFirst
    Do Until Prop.EOF
        If Prop.Fields("OwnerNo") = CurPlayer And Prop.Fields("Set") < 9 Then
            If Prop.Fields("HousesOwned") = 5 Then      'Hotel
                Call PlyrMoney(CurPlayer, -Amount * 4)
                Call PlyrMoney(FreePark, Amount * 4)
            Else                                        'Houses
            Call PlyrMoney(CurPlayer, -Amount * Prop.Fields("HousesOwned"))
            Call PlyrMoney(FreePark, Amount * Prop.Fields("HousesOwned"))
            End If
        End If
    Prop.MoveNext
    Loop
    Call EndTurn

Case "Street Repairs"
    Prop.MoveFirst
    Do Until Prop.EOF
    If Prop.Fields("OwnerNo") = CurPlayer And Prop.Fields("Set") < 9 Then
        If Prop.Fields("HousesOwned") = 5 Then          'Hotel
            Call PlyrMoney(CurPlayer, -Amount * 3)
            Call PlyrMoney(FreePark, Amount * 3)
        Else                                            'Houses
            Call PlyrMoney(CurPlayer, -(Amount * Prop.Fields("HousesOwned")))
            Call PlyrMoney(FreePark, (Amount * Prop.Fields("HousesOwned")))
        End If
    End If
    Prop.MoveNext
    Loop
    Call EndTurn

Case "Increase Salary"
    Plyr.Seek "=", CurPlayer
    Plyr.Fields("Salary") = Plyr.Fields("Salary") + Amount
    Call EndTurn
    
Case "Decrease Salary"
    Plyr.Seek "=", CurPlayer
    Plyr.Fields("Salary") = Plyr.Fields("Salary") - Amount
    Call EndTurn
    
Case "Advance To"
    If InStr(Amnt, "Next") > 0 Then
        Prop.Seek "=", s
        Do Until Prop.Fields("set") = SNo
            Prop.MoveNext
        Loop
    Amount = Prop.Fields("Number")
    End If
    If Amount < s Then  'Player gets "Salary" for passing "GO"
        Plyr.Seek "=", CurPlayer
        Call PlyrMoney(CurPlayer, Plyr.Fields("Salary"))
        Call PlyrMoney(99, -Plyr.Fields("Salary"))
    End If
    s = Amount
    Call MovePlayer(s)
    
Case "Back To"
    If InStr(Amnt, "Last") > 0 Then
        Prop.Seek "=", s
        Do Until Prop.Fields("set") = SNo
            Prop.MoveLast
        Loop
    Amount = Prop.Fields("Number")
    End If
    s = Amount
    Call MovePlayerBack(s)
    
Case "Go Back"
    s = s - Amount
    If s < 1 Then s = s + 40
    Call MovePlayerBack(s)
   
Case "Go Forward"
    s = s + Amount
        If s > 40 Then  'Salary for passing "GO"
            Plyr.Seek "=", CurPlayer
            s = s - 40
            Call PlyrMoney(CurPlayer, Plyr.Fields("Salary"))
            Call PlyrMoney(99, -Plyr.Fields("Salary"))
        End If
    Call MovePlayer(s)
    
Case "Fine or " & ChanceNme
    Choice = InputBox("Please Type 'F' for Fine or 'C' for " & ChanceNme, "Fine or " & ChanceNme, "C")
        If Choice = "F" Or Choice = "f" Then
            Call PlyrMoney(CurPlayer, -Amount)
            Call PlyrMoney(FreePark, Amount)
            Call EndTurn
        ElseIf Choice = "C" Or Choice = "c" Then
            FrmBoard.LblComChest.Caption = ""
            Call Chance 'Chance Card
            Exit Sub
        Else
            n = MsgBox("Please type 'F' or 'C'", vbCritical, "Fine or " & ChanceNme)
            Call Cards(Action, Amount)
        End If
        
Case "Goto " & Jail
    Call GoToJail

Case "Miss Turns"
    Call MissTurn(Amount)
    Call EndTurn
End Select

End Sub

Public Sub GoToJail()
    Plyr.Seek "=", CurPlayer
    Plyr.Edit
    Plyr.Fields("Square") = 11
    Plyr.Update
    PositionPlayer (11)
    Call MissTurn(3)
    Dice2 = 7
    Call EndTurn
End Sub

Public Sub CommChest()
Dim Action As String, Amnt As String
Randomize
CChest.Index = "Number"
CChest.Seek "=", Random(16) 'Select Card at Random
Action = CChest.Fields("Action")
Amnt = CChest.Fields("Amount")

If Action = "Get Out of " & Jail Then
    If CChest.Fields("Owner") = 99 Then
        CChest.Edit
        CChest.Fields("Owner") = CurPlayer
        CChest.Update
        FrmBoard.LblComChest.Caption = CChest.Fields("text")
        Call EndTurn
        Exit Sub
    Else
    Call CommChest  'If Get Out of Jail card held by another player
                        'select a different card
    Exit Sub
    End If
End If

FrmBoard.LblComChest.Caption = CChest.Fields("text")
Call Cards(Action, Amnt)  'Perform action

End Sub

Public Sub Chance()
Dim Action As String, Amnt As String
Randomize
Chnce.Index = "Number"
Chnce.Seek "=", Random(16) 'Select Card at Random
Action = Chnce.Fields("Action")
Amnt = Chnce.Fields("Amount")

If Action = "Get Out of " & Jail Then
    If Chnce.Fields("Owner") = 99 Then
        Chnce.Edit
        Chnce.Fields("Owner") = CurPlayer
        Chnce.Update
        FrmBoard.LblChance.Caption = Chnce.Fields("text")
        Call EndTurn
        Exit Sub
    Else
    Call Chance  'If Get Out of Jail card held by another player
                        'select a different card
    Exit Sub
    End If
End If

FrmBoard.LblChance.Caption = Chnce.Fields("text")
Call Cards(Action, Amnt)  'Perform action

End Sub

Public Sub BuyProperty(ByVal s As Integer, ByVal Buyer As Integer, ByVal Price As Single)
'Player buys a property
Dim i As Integer, n As Integer, Count As Integer, SetNo As Integer
Prop.Index = "Number"
Prop.Seek "=", s
SetNo = Prop.Fields("Set")
    If GetPlayerMoney(Buyer) - Price < 0 Then    'Not enough money
        n = MsgBox("Sorry you only have " & CurSymbBefore & GetPlayerMoney(CurPlayer) & CurSymbAfter & vbLf & _
            "You can't afford " & Prop.Fields("Name"), vbCritical, "Insufficient Funds")
        Exit Sub
    End If
Call PlyrMoney(Buyer, -Price)    'Reduce players' money
Call PlyrMoney(99, Price)    'Increase Banks' money
Prop.Edit
Prop.Fields("OwnerNo") = Buyer
Prop.Update

If SetNo = 10 Then Call Stations(10)
If SetNo = 9 Then Call Stations(9)
End Sub

Public Sub Auction(ByVal Nme As Integer)
Dim s As Integer
Prop.Index = "Number"
Prop.Seek "=", Nme
AucPrice = Prop.Fields("Price")
FrmAuction.LblPrice.Caption = CurSymbBefore & AucPrice & CurSymbAfter
FrmAuction.LblProp.Caption = Prop.Fields("Name")
s = Prop.Fields("Number")
Call ReSetAuction(99, True)
FrmAuction.Show
End Sub

Public Sub ReSetAuction(ByVal PlyrNumb As Integer, ByVal First As Boolean)
Dim n As Integer
With FrmAuction
    .LblGoing.FontBold = False
    .LblGoing.FontSize = 8
    .LblGoing.Alignment = 1
    .LblSecs.Visible = True
    .LblSecsLab.Visible = True
    .LblGoing.Caption = "Auction Closes in"
    .LblSecs.Caption = "10"
End With
Plyr.Index = "Number"
Plyr.Seek "=", PlyrNumb
If GetPlayerMoney(PlyrNumb) < AucPrice Then
    n = MsgBox("Sorry you can't afford " & vbCrLf & _
    FrmAuction.LblProp.Caption, vbCritical, "")
    Exit Sub
End If
FrmAuction.LblGoingFor.Caption = "For " & CurSymbBefore & AucPrice & CurSymbAfter
If Not First Then AucPrice = AucPrice + 10
FrmAuction.LblPrice.Caption = CurSymbBefore & AucPrice & CurSymbAfter
FrmAuction.LblGoingTo.Caption = Plyr.Fields("Name")
End Sub

Public Sub CloseAuction(ByVal Winner As String)
Plyr.Index = "Name"
Plyr.Seek "=", Winner
Prop.Index = "Name"
Prop.Seek "=", FrmAuction.LblProp.Caption
MsgBox (Plyr.Fields("Name") & " Buys " & Prop.Fields("Name") & " For " & _
    CurSymbBefore & AucPrice - 10 & CurSymbAfter)
Call BuyProperty(Prop.Fields("Number"), Plyr.Fields("Number"), AucPrice - 10)
Call UpdateBoard
End Sub

Public Sub EndTurn()    'Complete players' turn
Dim s As Integer, PropOwner As Integer: Dim PropMort As Boolean
Prop.Index = "Number"
s = PlayerSquare(CurPlayer)
Prop.Seek "=", s
PropOwner = Prop.Fields("OwnerNo")
PropMort = Prop.Fields("Mortgaged")
If GetPlayerMoney(CurPlayer) < AmountOwed Then Call LowMoney
Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer

Select Case s   'Update player's/Banks' money
Case 5, 39
    Call PlyrMoney(CurPlayer, -AmountOwed)
    Call PlyrMoney(FreePark, AmountOwed)

Case Else
    If PropOwner <> 0 And PropOwner <> 99 And PropOwner <> CurPlayer And PropMort = False Then
    If Prop.Fields("Set") = 9 Then
        Call PlyrMoney(CurPlayer, -AmountOwed * Dice1)
        Call PlyrMoney(PropOwner, AmountOwed * Dice1)
    Else
        Call PlyrMoney(CurPlayer, -AmountOwed)
        Call PlyrMoney(PropOwner, AmountOwed)
    End If
    End If
End Select
AmountOwed = 0
End Sub

Public Function GetCurPlayer()
Plyr.Index = "Number"
Plyr.MoveFirst
GetCurPlayer = 0
Do Until Plyr.EOF
    If Plyr.Fields("CurPlayer") = True Then
        GetCurPlayer = Plyr.Fields("Number")
        Exit Do
    End If
Plyr.MoveNext
Loop
If GetCurPlayer = 0 Then GetCurPlayer = 1
End Function

Public Sub SetCurPlayer(ByVal PlayerNo)
Plyr.Index = "Number"
Plyr.MoveFirst
Do Until Plyr.EOF
Plyr.Edit
    If Plyr.Fields("Number") = PlayerNo Then
    Plyr.Fields("CurPlayer") = True
    Else
    Plyr.Fields("CurPlayer") = False
    End If
Plyr.Update
Plyr.MoveNext
Loop
End Sub
