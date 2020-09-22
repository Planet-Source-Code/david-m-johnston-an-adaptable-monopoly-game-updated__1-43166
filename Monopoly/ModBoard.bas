Attribute VB_Name = "ModDrawBoard"
Option Explicit

Public Sub ResCheck()
'Detects the resolution in use & Draws board to fill screen
'Moves controls acordingly

'Rescomp used to keep control proportions the same
Dim Ctrl As Object: Dim s As Integer, n As Integer
FWidth = Screen.Width   'Find width of screen
FHeight = Screen.Height - (FrmBoard.SBar.Height * 4)    'Used for board hieght
ResComp = Screen.Width / 12000
If Screen.Width <= 9600 Then 'Programme won't run in lower than 800 * 600
    n = MsgBox("Sorry this program can't run in this resolution", vbCritical, "Resolution")
    End
Else
LowRes = 600    'Used to increase width of side squares on board
                'to allow more space for property names
For s = 1 To 40
    FrmBoard.LblName(s).FontSize = 8
    FrmBoard.LblPrice(s).FontSize = 8
Next s
'Alter position & size of controls using rescomp so that board looks
    'the same in any resolution
    For Each Ctrl In FrmBoard.Controls
        If Ctrl.Name Like "Lb*" Or Ctrl.Name Like "Lst*" Then
            Ctrl.Move Ctrl.Left * ResComp, Ctrl.Top * ResComp, Ctrl.Width * ResComp, Ctrl.Height * ResComp
        End If
    Next Ctrl
FrmBoard.CboViewPlayer.Move FrmBoard.CboViewPlayer.Left * ResComp, FrmBoard.CboViewPlayer.Top * ResComp
End If
End Sub

Sub DrawBoard()
Dim SetCol As String: Dim s As Integer
'set sizes of squares according to resolution in use
Corner = (FWidth / 13) * 1.5
SqBShort = (FWidth - (Corner * 2)) / 9
SqSShort = (FHeight - (Corner * 2)) / 9

Prop.Index = "Number"
PropSet.Index = "Number"

For s = 1 To 40     'Go through all squares
PosX (s)    'Sets XPos (X Position) according to square being drawn
PosY (s)    'Sets YPos (Y Position) according to square being drawn
Prop.Seek "=", s    'Move to Property for Square being drawn
PropSet.Seek "=", Prop.Fields("Set")    'Move to property set for Property being drawn
SetCol = Val(PropSet.Fields("Colour"))  'Colour of property
FrmBoard.BackColor = BrdColour      'Set colour of board
FrmBoard.LblName(s).Caption = Prop.Fields("Name")   'Property Name

With FrmBoard
Select Case s       's = Square

Case 1, 11, 21, 31      'Corners
FrmBoard.Line (XPos, YPos)-Step(Corner, Corner), BrdColour, BF

Case 1 To 11, 21 To 31  'Top & Bottom
FrmBoard.Line (XPos, YPos)-Step(SqBShort, Corner), BrdColour, BF

Case 11 To 21, 31 To 40 'Sides
FrmBoard.Line (XPos, YPos)-Step(Corner + LowRes, SqSShort), BrdColour, BF

End Select

Select Case s
Case 2, 4, 7, 9, 10     'Sets on bottom of board
FrmBoard.Line (XPos, YPos + 20)-Step(SqBShort - 10, 200), SetCol, BF

Case 12, 14, 15, 17, 19, 20 'Sets on Left of board
FrmBoard.Line (Corner + LowRes - 220, YPos + 10)-Step(200, SqSShort - 20), SetCol, BF

Case 22, 24, 25, 27, 28, 30 'Sets at Top of board
FrmBoard.Line (XPos + 10, Corner - 220)-Step(SqBShort - 30, 200), SetCol, BF

Case 32, 33, 35, 38, 40     'Sets on right of board
FrmBoard.Line (XPos + 10, YPos + 10)-Step(210, SqSShort - 20), SetCol, BF

End Select

Select Case s

Case 1, 11, 21, 31 'Other Corners
If Len(.LblName(s).Caption) > 6 Then
    .LblName(s).Move XPos + 300, (YPos + (Corner / 3)), Corner - 300, Corner - (Corner / 3)
    .LblName(s).FontSize = 14
    Else
    .LblName(s).Move XPos + 300, (YPos + (Corner / 3)), Corner - 300, Corner / 3
    .LblName(s).FontSize = 25
End If
FrmBoard.Line (XPos, YPos)-Step(Corner, Corner), , B
    .LblPrice(s).Move XPos + 300, YPos + (Corner - 300), Corner - 300, 300

Case 2 To 10, 22 To 30  'Bottom & Top
    .LblName(s).Move XPos + 10, YPos + 300, SqBShort - 10, (Corner / 2)
    .LblPrice(s).Move XPos, (YPos + .LblName(s).Height), SqBShort, (Corner / 2)
    FrmBoard.Line (XPos, YPos)-Step(SqBShort, Corner), , B
    
Case 12 To 20   'Left
    .LblName(s).Move XPos, (YPos + 10), (Corner + LowRes - 200), (SqSShort / 3) * 2
    .LblPrice(s).Move XPos, (YPos + .LblName(s).Height), (Corner + LowRes - 200)
    FrmBoard.Line (XPos, YPos)-Step(Corner + LowRes, SqSShort), , B

Case 32 To 40   'Right
    .LblName(s).Move XPos, (YPos + 10), (Corner + LowRes + 200), (SqSShort / 3) * 2
    .LblPrice(s).Move XPos, (YPos + .LblName(s).Height), (Corner + LowRes + 200)
    FrmBoard.Line (XPos, YPos)-Step(Corner + LowRes, SqSShort), , B
End Select

If Prop.Fields("Price") <> "0" Then _
    .LblPrice(s).Caption = CurSymbBefore & Prop.Fields("Price") & CurSymbAfter  'Property Price

If Prop.Fields("Set") > 0 And Prop.Fields("Set") < 11 Then   'Set Colour
    .LblName(s).ToolTipText = "Click Here to View " & PropInfo
    .LblPrice(s).ToolTipText = "Click Here to View " & PropInfo
    .LstBankProp.ToolTipText = "Click Here to View " & PropInfo
    .LstPlayerProp.ToolTipText = "Click Here to View " & PropInfo
End If

If s = 11 Then 'Jail Square
    .LblName(11).ToolTipText = "Clik Here to use Get Out of " & Jail & " Free Card"
End If
.LblChance.Caption = ChanceNme
.LblComChest.Caption = CommChestNme
End With

Next s
Call ChangeCols
End Sub

Public Sub ChangeCols()     'Set Board & Text(Squares only) colours
Dim Ctrl As Control

FrmBoard.BackColor = BrdColour
For Each Ctrl In FrmBoard.Controls
If Ctrl.Name Like "Lb*" Then
    Ctrl.BackColor = BrdColour
    Ctrl.ForeColor = TextColour
End If
Next Ctrl

With FrmBoard   'GO Square doesn't change
.LblName(1).ForeColor = &HFF&
.LblName(1).FontBold = True
.LblChance.BackColor = &H80C0FF
.LblComChest.BackColor = &HFFD0FF
End With
End Sub

Public Sub UpdateBoard()    'Update Property Lists & Current Player
Dim Square As Integer
Plyr.Index = "Number"
Plyr.Seek "=", ViewPlayer
Prop.Index = "Number"
Square = 1

Prop.MoveFirst
Do Until Prop.EOF
If Prop.Fields("Set") <> 0 Then
Square = Prop.Fields("Number")
    If Prop.Fields("Mortgaged") = True Then 'Pink text on grey background
                                                'if property mortgaged
        FrmBoard.LblName(Square).ForeColor = &H8080FF
        FrmBoard.LblPrice(Square).ForeColor = &H8080FF
    Else
        FrmBoard.LblName(Square).ForeColor = TextColour
        FrmBoard.LblPrice(Square).ForeColor = TextColour
    End If
End If
Prop.MoveNext
Loop

Call BankProperty   'Update list of property in bank
PlayerProperty (ViewPlayer) 'Update list of property held by player being viewed
With FrmBoard
.LblOwner.Caption = Plyr.Fields("Name")
.CboViewPlayer.Text = Plyr.Fields("Name")
.LblMoney = CurSymbBefore & GetPlayerMoney(CurPlayer) & CurSymbAfter 'Money Owned by vurrent player
.LblBankMoney.Caption = CurSymbBefore & GetPlayerMoney(99) & CurSymbAfter  'Money in Bank
.LblPrice(21).Caption = CurSymbBefore & GetPlayerMoney(0) & CurSymbAfter
End With

End Sub

Public Sub ShowCounters()
Dim i As Integer

For i = 1 To TotPlayers
    FrmBoard.ImgCounter(i).Move XPos, YPos + Corner / 2
    FrmBoard.ImgCounter(i).Visible = True
    XPos = XPos + (Corner / TotPlayers)
Next i
End Sub

Public Sub UpdateHouses()   'Re-Draw Houses/Hotels
Dim i As Integer, Houses As Integer, PropSet As Integer

Prop.Index = "Number"
For i = 1 To 40
Prop.Seek "=", i
PropSet = Prop.Fields("Set")
Houses = Prop.Fields("HousesOwned")
If Houses > 0 And PropSet > 0 And PropSet < 9 Then Call DrawHouses(i, Houses, PropSet) 'Draw Houses
Next i
End Sub

Public Sub ClearHouses(ByVal Numb)
'Remove Houses/Hotels from Square (Numb) at position XPos,YPos
Dim SetCol As String

PropSet.Index = "Number"
PropSet.Seek "=", Prop.Fields("Set")
SetCol = Val(PropSet.Fields("Colour"))  'Set Colour
PosX (Numb)
PosY (Numb)
Select Case Numb
Case 2 To 10
    FrmBoard.Line (XPos, YPos + 20)-Step(SqBShort - 10, 200), SetCol, BF
    FrmBoard.Line (XPos, YPos)-Step(SqBShort, Corner), , B
Case 22 To 30:
    FrmBoard.Line (XPos + 10, Corner - 220)-Step(SqBShort - 30, 200), SetCol, BF
    FrmBoard.Line (XPos, YPos)-Step(SqBShort, Corner), , B
Case 12 To 20:
    FrmBoard.Line (XPos, YPos)-Step(Corner + LowRes, SqSShort), , B
    FrmBoard.Line (Corner + LowRes - 220, YPos + 10)-Step(200, SqSShort - 20), SetCol, BF
Case 32 To 40:
    FrmBoard.Line (XPos, YPos)-Step(Corner + LowRes, SqSShort), , B
    FrmBoard.Line (XPos + 10, YPos + 10)-Step(210, SqSShort - 20), SetCol, BF
End Select

End Sub

Public Sub DrawHouses(ByVal Numb, ByVal HousesOwned, ByVal SetNo)
'Draw HousesOwned Houses/Hotels on Square Numb
Dim i As Integer
Dim HouseColour, HotelColour As String

PropSet.Index = "Number"
PropSet.Seek "=", SetNo
HouseColour = PropSet.Fields("HouseColour")
HotelColour = PropSet.Fields("HotelColour")

Call ClearHouses(Numb)  'Remove Houses alredy on square Numb
PosX (Numb)
PosY (Numb)

For i = 1 To HousesOwned
Select Case Numb

Case 2 To 10    'Bottom
    If HousesOwned = 5 Then
        FrmBoard.Line ((XPos + ((SqBShort / 2) - SqBShort / 4)), YPos + 40)-Step(SqBShort / 2, 150), HotelColour, BF
        Exit Sub
    End If
    If i = 1 Then XPos = XPos + (SqBShort / 12) + 20
    If i > 0 And i < 5 Then FrmBoard.Line ((XPos + ((SqBShort / 6) + 40) * (i - 1)), YPos + 40)-Step(SqBShort / 6, 150), HouseColour, BF

Case 22 To 30   'Top
    If HousesOwned = 5 Then
        FrmBoard.Line ((XPos + ((SqBShort / 2) - SqBShort / 4)), YPos + Corner - 190)-Step(SqBShort / 2, 150), HotelColour, BF
        Exit Sub
    End If
    If i = 1 Then XPos = XPos + (SqBShort / 12) + 20
    If i > 0 And i < 5 Then FrmBoard.Line ((XPos + ((SqBShort / 6) + 40) * (i - 1)), YPos + Corner - 190)-Step(SqBShort / 6, 150), HouseColour, BF

Case 12 To 20   'Left
    If HousesOwned = 5 Then
        FrmBoard.Line ((XPos + LowRes + Corner - 190), YPos + (SqSShort / 2) - (SqSShort / 4))-Step(150, SqSShort / 2), HotelColour, BF
        Exit Sub
    End If
    If i = 1 Then YPos = YPos + (SqSShort / 12) + 20
    If i > 0 And i < 5 Then FrmBoard.Line ((XPos + LowRes + Corner - 200), YPos + ((SqSShort / 6) + 40) * (i - 1))-Step(150, SqSShort / 6), HouseColour, BF

Case 32 To 40   'Right
    If HousesOwned = 5 Then
    FrmBoard.Line ((XPos + 40), YPos + ((SqSShort / 2) - SqSShort / 4))-Step(150, SqSShort / 2), HotelColour, BF
    Exit Sub
    End If
    If i = 1 Then YPos = YPos + (SqSShort / 12) + 20
    If i > 0 And i < 5 Then FrmBoard.Line (XPos + 40, YPos + ((SqSShort / 6) + 40) * (i - 1))-Step(150, SqSShort / 6), HouseColour, BF

End Select

Next i
End Sub

Public Sub BankProperty()   'Clear & Re-Create list of Properties in Bank

With FrmBoard
.LstBankProp.Clear
Prop.MoveFirst
Do Until Prop.EOF   'Go throug all properties
    If Prop.Fields("OwnerNo") = 99 And Prop.Fields("Set") <> 0 Then _
        .LstBankProp.AddItem Prop.Fields("Name")
Prop.MoveNext
Loop
Chnce.MoveFirst
Do Until Chnce.EOF 'List all the current players' property
    If Chnce.Fields("Owner") = 99 Then _
        .LstBankProp.AddItem Chnce.Fields("Text")
Chnce.MoveNext
Loop
CChest.MoveFirst
Do Until CChest.EOF 'List all the current players' property
    If CChest.Fields("Owner") = 99 Then _
        .LstBankProp.AddItem CChest.Fields("Text")
CChest.MoveNext
Loop
End With
End Sub

Public Function PlayerProperty(ByVal Player)
'Clear & Re-Create list of all properties held by Player

With FrmBoard
.LstPlayerProp.Clear
Prop.MoveFirst
Do Until Prop.EOF   'Go throug all properties
    If Prop.Fields("OwnerNo") = Player Then _
    .LstPlayerProp.AddItem Prop.Fields("Name")
Prop.MoveNext
Loop
Chnce.MoveFirst
Do Until Chnce.EOF 'List all the current players' property
    If Chnce.Fields("Owner") = Player Then _
        .LstPlayerProp.AddItem Chnce.Fields("Text")
Chnce.MoveNext
Loop
CChest.MoveFirst
Do Until CChest.EOF 'List all the current players' property
    If CChest.Fields("Owner") = Player Then _
        .LstPlayerProp.AddItem CChest.Fields("Text")
CChest.MoveNext
Loop
End With
End Function
