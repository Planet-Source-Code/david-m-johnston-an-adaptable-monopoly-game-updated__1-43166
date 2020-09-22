Attribute VB_Name = "ModEditDB"
Option Explicit

Public Sub EditDB()
Call CreateLists    'Create list of properties on Edit DataBase form
End Sub

Public Sub CreateLists()
Dim i As Integer

Prop.Index = "Number"
Chnce.Index = "Number"
CChest.Index = "Number"
PropSet.Index = ("Number")

With FrmEditDB
For i = 1 To 12     'Clear property List boxes
    If i <> 3 And i <> 6 Then .LstProperty(i).Clear
Next i

Prop.MoveFirst
Do Until Prop.EOF   'Add all properties to property list boxes
    For i = 1 To 12
        If i <> 3 And i <> 6 Then
            .LstProperty(i).AddItem Prop.Fields(i - 1)
        End If
    Next i
Prop.MoveNext
Loop
.LblPropNo.Caption = .LstProperty(1).Text
.LstProperty(1).ListIndex = 0
Call PropertySelected(1)

For i = 0 To 3      'Clear Chance & Community Chest List Boxes
    .LstChance(i).Clear
    .LstCChest(i).Clear
Next i

Chnce.MoveFirst
Do Until Chnce.EOF  'Create Chance List Boxes
    For i = 0 To 3
       .LstChance(i).AddItem Chnce.Fields(i)
    Next i
    If Chnce.Fields("Action") = "Get Out of " & Jail Then
    .TxtJailFine.Text = JailFine
    End If
    Chnce.MoveNext
Loop
.LstChance(1).ListIndex = 0
Call ChanceSelected(1)

CChest.MoveFirst
Do Until CChest.EOF  'Create Community Chest List Boxes
    For i = 0 To 3
        .LstCChest(i).AddItem CChest.Fields(i)
    Next i
CChest.MoveNext
Loop
.LstCChest(1).ListIndex = 0
Call CChestSelected(1)

'Create action combo boxes for Chance & Community Chest Cards
For i = 0 To 1
.CboAction(i).Clear
.CboAmount(i).Clear
.CboAction(i).AddItem "Receive From " & Bank
.CboAction(i).AddItem "Receive From All Players"
.CboAction(i).AddItem "Pay To " & Bank
.CboAction(i).AddItem "General Repairs"
.CboAction(i).AddItem "Street Repairs"
.CboAction(i).AddItem "Increase Salary"
.CboAction(i).AddItem "Decrease Salary"
.CboAction(i).AddItem "Advance To"
.CboAction(i).AddItem "Back To"
.CboAction(i).AddItem "Go Back"
.CboAction(i).AddItem "Go Forward"
.CboAction(i).AddItem "Fine or " & ChanceNme
.CboAction(i).AddItem "Goto " & Jail
.CboAction(i).AddItem "Miss Turns"

.CboAmount(i).AddItem "Next " & Station
.CboAmount(i).AddItem "Last " & Station
.CboAmount(i).AddItem "Next " & Utility
.CboAmount(i).AddItem "Last " & Utility
Next i

PropSet.MoveFirst
Do Until PropSet.EOF    'Set colours in Set Colour options
    PropSet.MoveNext
    i = PropSet.Fields("Number")
    If i > 8 Then Exit Do
    .LblHousePriceLab(i - 1).Caption = House & " Price"
    .LblHouseColLab(i - 1).Caption = House & " Colour"
    .LblHotelColLab(i - 1).Caption = Hotel & " Colour"
    .LblSet(i - 1).BackColor = Val(PropSet.Fields("Colour"))
    .TxtHPrice(i - 1).Text = CurSymbBefore & PropSet.Fields("HousePrice") & CurSymbAfter
    .LblHouseCol(i - 1).BackColor = Val(PropSet.Fields("HouseColour"))
    .LblHotelCol(i - 1).BackColor = Val(PropSet.Fields("HotelColour"))
Loop
.LblApplyAll.Caption = "Apply " & House & " and " & Hotel & " colour to all Sets"

.TxtNamesName(0).Text = House
.TxtNamesName(1).Text = Hotel
.TxtNamesName(2).Text = Go
.TxtNamesName(3).Text = Jail
.TxtNamesName(4).Text = Bank
.TxtNamesName(5).Text = PropInfo
.TxtNamesName(6).Text = Rent
.TxtNamesName(7).Text = Utility
.TxtNamesName(8).Text = Station
.TxtNamesName(9).Text = CurSymbBefore & CurSymbAfter
.TxtNamesName(10).Text = FParking
.TxtNamesName(11).Text = ChanceNme
.TxtNamesName(12).Text = CommChestNme
.CboSymPos.Text = CurSymbPos

.TxtDoubles.Text = JailDoubles
.CboBuildEven.Text = "YES"
.CboSellEven.Text = "NO"
.TxtJailFine.Text = CurSymbBefore & JailFine & CurSymbAfter
.TxtBkStartMon.Text = CurSymbBefore & BnkStartMon & CurSymbAfter
.TxtPlyrStartMon.Text = CurSymbBefore & PlyrStartMon & CurSymbAfter
.TxtSalary.Text = CurSymbBefore & Salary & CurSymbAfter
.TxtStartProps.Text = StartProps
.CboFinesTo.AddItem Bank
.CboFinesTo.AddItem FParking
.CboFinesTo.ListIndex = 0
.CboHotelNo.Text = HotelNo

End With
End Sub

Public Sub PropertySelected(ByVal Clicked)
    'Select all fields for chosen property
Dim i As Integer
With FrmEditDB
For i = 1 To 12
    If i <> Clicked And i <> 3 And i <> 6 Then _
        .LstProperty(i).ListIndex = .LstProperty(Clicked).ListIndex
Next i

.LblPropNo.Caption = .LstProperty(1).Text
.TxtName.Text = .LstProperty(2).Text
.CboSet.Text = .LstProperty(4).Text
.TxtPrice.Text = CurSymbBefore & .LstProperty(5).Text & CurSymbAfter
For i = 7 To 12
    .TxtRent(i).Text = CurSymbBefore & .LstProperty(i).Text & CurSymbAfter
Next i
End With
End Sub

Public Sub ChanceSelected(ByVal Clicked)
    'Select all fields for chosen Chance card
Dim i As Integer
With FrmEditDB
For i = 0 To 3
    If i <> Clicked Then _
        .LstChance(i).ListIndex = .LstChance(Clicked).ListIndex
Next i
.LblCardNo(0).Caption = .LstChance(0).Text
.TxtText(0).Text = .LstChance(1).Text
.CboAction(0).Text = .LstChance(2).Text
.CboAmount(0).Text = .LstChance(3).Text
End With
End Sub

Public Sub CChestSelected(ByVal Clicked)
    'Select all fields for chosen Community Chest
Dim i As Integer
With FrmEditDB
For i = 0 To 3
    If i <> Clicked Then _
        .LstCChest(i).ListIndex = .LstCChest(Clicked).ListIndex
Next i
.LblCardNo(1).Caption = .LstCChest(0).Text
.TxtText(1).Text = .LstCChest(1).Text
.CboAction(1).Text = .LstCChest(2).Text
.CboAmount(1).Text = .LstCChest(3).Text
End With
End Sub

Public Sub UpdateDBProperty()
    'Update DataBase with new data
Dim i As Integer, n As Integer

Prop.Index = "Number"
With FrmEditDB
If .TxtName = "" Or .TxtPrice = "" Or .TxtRent(7) = "" Or _
    .TxtRent(8) = "" Or .TxtRent(9) = "" Or .TxtRent(10) = "" Or _
    .TxtRent(11) = "" Or .TxtRent(12) = "" Or .CboSet.Text = "" Then
    n = MsgBox("Please Enter a Value for all fields", vbCritical, "Empty Field")
    Exit Sub
End If
Prop.Seek "=", FrmEditDB.LstProperty(1).Text
Prop.Edit
Prop.Fields("Name") = .TxtName.Text
Prop.Fields("Set") = Value(.CboSet.Text)
If .CboSet.Text = "0" Then Prop.Fields("OwnerNo") = 0
If .CboSet.Text <> "0" Then Prop.Fields("OwnerNo") = 99

Prop.Fields("Price") = Value(.TxtPrice.Text)
For i = 7 To 12
Prop.Fields(i - 1) = Value(.TxtRent(i).Text)
Next i
Prop.Fields("Mortgaged") = False
Prop.Fields("HousesOwned") = 0
Prop.Update
End With
n = MsgBox("Property Updated", vbInformation, "Updated")
End Sub

Public Sub UpdateDBChance()
    'Update DataBase with new data
Dim Numb As Integer, n As Integer
With FrmEditDB
Numb = Value(.LstChance(0).Text)
Chnce.Index = "Number"
Chnce.Seek "=", Numb
Chnce.Edit
Chnce.Fields("Text") = .TxtText(0).Text
Chnce.Fields("Action") = .CboAction(0).Text
Chnce.Fields("Amount") = Value(.CboAmount(0).Text)
End With
Chnce.Update
n = MsgBox("Chance Cards Updated", vbInformation, "Updated")
End Sub

Public Sub UpdateDBCChest()
    'Update DataBase with new data
Dim Numb As Integer, n As Integer
With FrmEditDB
Numb = Value(.LstCChest(0).Text)
CChest.Index = "Number"
CChest.Seek "=", Numb
CChest.Edit
CChest.Fields("Text") = .TxtText(1).Text
CChest.Fields("Action") = .CboAction(1).Text
CChest.Fields("Amount") = Value(.CboAmount(1).Text)
End With
CChest.Update
n = MsgBox("Community Chest Cards Updated", vbInformation, "Updated")
End Sub

Public Sub UpdateDBCols()
    'Update DataBase with new data
PropSet.Index = "Number"
Dim i As Integer

With FrmEditDB
For i = 0 To 7
    PropSet.Seek "=", i + 1
    PropSet.Edit
    PropSet.Fields("Colour") = FrmEditDB.LblSet(i).BackColor
    PropSet.Fields("HousePrice") = Value(.TxtHPrice(i).Text)
    PropSet.Fields("HouseColour") = FrmEditDB.LblHouseCol(i).BackColor
    PropSet.Fields("HotelColour") = FrmEditDB.LblHotelCol(i).BackColor
    PropSet.Update
Next i
End With
End Sub

Public Sub UpdateNames()
With FrmEditDB
    House = .TxtNamesName(0).Text
    Hotel = .TxtNamesName(1).Text
    Go = .TxtNamesName(2).Text
    Jail = .TxtNamesName(3).Text
    Bank = .TxtNamesName(4).Text
    PropInfo = .TxtNamesName(5).Text
    Rent = .TxtNamesName(6).Text
    Utility = .TxtNamesName(7).Text
    Station = .TxtNamesName(8).Text
    CurrencySymb = .TxtNamesName(9).Text
    If .CboSymPos.Text = "Before" Then
        CurSymbBefore = CurrencySymb
        CurSymbAfter = ""
    Else
        CurSymbAfter = CurrencySymb
        CurSymbBefore = ""
    End If
    FParking = .TxtNamesName(10).Text
    ChanceNme = .TxtNamesName(11).Text
    CommChestNme = .TxtNamesName(12).Text
End With
End Sub

Public Sub UpdateRules()
With FrmEditDB
JailFine = Value(.TxtJailFine.Text)
JailDoubles = Value(.TxtDoubles.Text)
If .CboBuildEven = "YES" Then BuildEven = True Else BuildEven = False
If .CboSellEven = "YES" Then SellEven = True Else SellEven = False
BnkStartMon = Value(.TxtBkStartMon.Text)
PlyrStartMon = Value(.TxtPlyrStartMon.Text)
Salary = Value(.TxtSalary.Text)
StartProps = Value(.TxtStartProps.Text)
If .CboFinesTo.Text = Bank Then FreePark = 99
If .CboFinesTo.Text = FParking Then FreePark = 0
HotelNo = .CboHotelNo.Text
End With
End Sub

Public Function Value(ByVal StrAmnt As String) As Single
Dim X As Integer, n As Integer, Amnt As String
For n = 1 To Len(StrAmnt)
    X = Asc(Mid(StrAmnt, n, 1))
    If (X > 47 And X < 58) Or X = 14 Then Amnt = Amnt & Mid(StrAmnt, n, 1)
    Value = Val(Amnt)
Next n
End Function
