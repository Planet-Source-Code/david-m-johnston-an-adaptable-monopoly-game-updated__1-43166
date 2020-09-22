Attribute VB_Name = "ModGame"
Option Explicit

Public Sub Turn()   'Player takes turn - Has clicked "Roll Dice"
Dim s As Integer, NewSqr As Integer, n As Integer, Miss As Integer
Dim PropOwner As Integer, HOwned As Integer
Dim PlayerMoney As Single
Dim OwnerName As String

Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
LowMon = False
AmountOwed = 0
ViewPlayer = CurPlayer
Prop.Index = "Number"
s = PlayerSquare(CurPlayer)
Miss = Plyr.Fields("MissTurns")

If Dice1 = Dice2 Then DoublesCount = DoublesCount + 1 Else DoublesCount = 0
If DoublesCount > JailDoubles Then
    n = MsgBox("You have rolled " & DoublesCount & " doubles" & vbLf & "Go To " & Jail, vbCritical, "Go To " & Jail)
    Call GoToJail
    DoublesCount = 0
    Call NextPlayer
    Exit Sub
End If
If Miss > 0 And Plyr.Fields("Square") <> 11 Then    'Player missing a turn
    Call TurnMissed(Miss)                               'but not in jail
    Call EndTurn
Else

Prop.Seek "=", s
s = s + Dice1 + Dice2   'New square = old square + dice
FrmBoard.LblDice1.Caption = Dice1
FrmBoard.LblDice2.Caption = Dice2
FrmBoard.LblChance.Caption = ""
FrmBoard.LblComChest.Caption = ""

If Plyr.Fields("Square") = 11 And Miss > 0 Then 'In jail & to miss turn
    If Dice1 <> Dice2 Then      'Didn't shake a double
        Call TurnMissed(Miss)   'Reduce turns to miss by 1
        Call NextPlayer
        Exit Sub
    End If
    n = MsgBox("You got a double, You can leave " & Jail, vbExclamation, "Leave " & Jail)
    Plyr.Edit
    Plyr.Fields("MissTurns") = 0
    Plyr.Update
End If

If s > 40 Then      'Get salary for passing "GO"
    s = s - 40
    Call PlyrMoney(CurPlayer, Plyr.Fields("Salary"))
    Call PlyrMoney(99, -Plyr.Fields("Salary"))
End If
Call MovePlayer(s)

Select Case s
Case 8, 23, 37, 3, 18, 34

        Select Case s
            Case 8, 23, 37: Call Chance
            Case 3, 18, 34: Call CommChest
        End Select
        NewSqr = PlayerSquare(CurPlayer)
        If NewSqr = s Then
            Call EndTurn
            If Dice1 <> Dice2 Then Call NextPlayer
            Exit Sub
        End If
        s = NewSqr
End Select

PlayerMoney = GetPlayerMoney(CurPlayer) 'Update PlayerMoney after Chance,Com. Chest Cards
Call UpdateBoard
Prop.Seek "=", s

Select Case s
Case 5        'Income Tax
    AmountOwed = Prop.Fields("Rent")
    If PlayerMoney < AmountOwed Then
        n = MsgBox("You can't afford this tax" & vbLf & _
        "You must sell some property to raise " & _
        CurSymbBefore & AmountOwed - PlayerMoney & CurSymbAfter, vbCritical, "Insufficient Funds")
        Call LowMoney
    Else
    Call EndTurn
    End If
    
Case 39       'Super Tax
    AmountOwed = Prop.Fields("Rent")
    If PlayerMoney < AmountOwed Then
        n = MsgBox("You can't afford this tax" & vbLf & _
        "You must sell some property to raise " & _
        CurSymbBefore & AmountOwed - PlayerMoney & CurSymbAfter, vbCritical, "Insufficient Funds")
        Call LowMoney
    Else
    Call EndTurn
    End If
    
Case 31     'Go to Jail
    s = 11
    Call GoToJail
    Call EndTurn
  
Case 1, 11
    If Dice1 <> Dice2 Then Call EndTurn
 
Case 21
If FreePark = 0 Then
    n = MsgBox("You've Landed on " & FParking & vbCrLf & "You Get " _
    & CurSymbBefore & GetPlayerMoney(0) & CurSymbAfter & "!", vbInformation, FParking)
    Call PlyrMoney(CurPlayer, GetPlayerMoney(0))
    Call PlyrMoney(0, -GetPlayerMoney(0))
End If

Case Else   'Set Rent Owed
    PropOwner = Prop.Fields("OwnerNo")
    Plyr.Seek "=", PropOwner
    HOwned = Prop.Fields("HousesOwned")
    AmountOwed = Prop.Fields(HOwned + 6)
    If Prop.Fields("Mortgaged") = True Then AmountOwed = 0
    If SetOwned(s) = True And HousesOnSet(s) = False And MortgInSetSet(s) = False Then AmountOwed = AmountOwed * 2
    Prop.Seek "=", s
    OwnerName = Plyr.Fields("Name")
    
    If PropOwner = 99 Then  'Unsold property
        If MsgBox("Would You Like to Buy " & vbLf & _
            Prop.Fields("Name") & vbLf & "For " & _
            CurSymbBefore & Prop.Fields("Price") & CurSymbAfter, 36, "Buy?") = 6 Then
        Call BuyProperty(s, CurPlayer, Prop.Fields("Price"))
        Call EndTurn
    Else
    Call Auction(s)
    End If
    
    ElseIf Prop.Fields("Set") = 9 And PropOwner <> CurPlayer Then   'Company owned by another player
        n = MsgBox("You have landed on " & Prop.Fields("Name") & vbLf & _
        "Which is owned by " & OwnerName & vbLf & _
        "Pay " & CurSymbBefore & AmountOwed & " * Dice 1" & CurSymbAfter & vbLf & _
        CurSymbBefore & AmountOwed * Dice1 & CurSymbAfter, vbExclamation, "Pay " & Rent)
        
        If PlayerMoney < AmountOwed Then    'Can't afford Rent
            n = MsgBox("You can't afford this " & Rent & vbLf & _
                "You must sell some property to raise " & _
                CurSymbBefore & AmountOwed - PlayerMoney & CurSymbAfter, vbCritical, "Insufficient Funds")
            Call LowMoney
        Else: Call EndTurn
        End If

    ElseIf PropOwner <> 0 And PropOwner <> 99 And _
        PropOwner <> CurPlayer And Prop.Fields("Mortgaged") = False Then
            'Pay Rent
        n = MsgBox("You have landed on " & Prop.Fields("Name") & vbLf & _
        "Which is owned by " & OwnerName & vbLf & _
        "Pay " & CurSymbBefore & AmountOwed & CurSymbAfter, vbExclamation, "Pay " & Rent)
        
        If PlayerMoney < AmountOwed Then    'Cant afford Rent
            n = MsgBox("You can't afford this " & Rent & vbLf & _
                "You must sell some property to raise " & CurSymbBefore & _
                AmountOwed - PlayerMoney & CurSymbAfter, vbCritical, "Insufficient Funds")
            Call LowMoney
        Else: Call EndTurn
        End If
    End If
End Select
End If
Call NextPlayer
End Sub

Public Sub NextPlayer() 'Move to next player
Dim n As Integer
If CurPlayer <> 0 Then
    Plyr.Index = "Number"
    Plyr.Seek "=", CurPlayer
    If Dice1 = Dice2 Or AmountOwed > GetPlayerMoney(CurPlayer) Then Exit Sub
End If
Plyr.MoveNext
If Plyr.Fields("Number") = 99 Then
    Plyr.MoveFirst
    Plyr.MoveNext
End If
CurPlayer = Plyr.Fields("Number")
SetCurPlayer (CurPlayer)
ViewPlayer = CurPlayer
Call UpdateBoard    'Show new players property & Money
Plyr.Seek "=", CurPlayer
FrmBoard.LblInfo.Caption = Plyr.Fields("Name") & " To Go"

If Plyr.Fields("Square") = 11 And Plyr.Fields("Missturns") > 0 Then
    If MsgBox("Would " & Plyr.Fields("Name") & " like to pay the fine of :" & _
    vbCrLf & CurSymbBefore & JailFine & CurSymbAfter & vbCrLf & "To get out of jail?" _
    & vbCrLf & vbCrLf & "(If you have a 'Get out of Jail Free' Card, " & _
    vbCrLf & "You can use it by clicking the 'Jail' square)", _
    vbYesNo, "Leave Jail") = 6 Then
        If GetPlayerMoney(CurPlayer) < JailFine Then
            n = MsgBox("You Can't afford the fine" & vbCrLf & _
                "You Must Stay In Jail", vbCritical, "Stay In Jail")
            Exit Sub
        Else
        Call PlyrMoney(CurPlayer, -JailFine)
        Call PlyrMoney(FreePark, JailFine)
        n = MsgBox("Fine Paid" & vbCrLf & _
            "You can move on your next turn", vbInformation, "Left Jail")
        Plyr.Seek "=", CurPlayer
        Plyr.Edit
        Plyr.Fields("MissTurns") = 0
        Plyr.Update
        Call NextPlayer
        End If
    End If
End If
End Sub
