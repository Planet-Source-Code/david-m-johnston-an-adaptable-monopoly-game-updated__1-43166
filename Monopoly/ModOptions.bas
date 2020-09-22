Attribute VB_Name = "ModOptions"
Option Explicit

Public Sub LoadDB() 'Player selects DataBase to be loaded
Dim i As Integer

If CurSymbPos = "Before" Then
    CurSymbBefore = CurrencySymb
Else
    CurSymbAfter = CurrencySymb
End If
FrmEditDB.LbCBoardColour.ForeColor = &H80000012
FrmEditDB.LbCBoardColour.Enabled = True
FrmEditDB.LbCBoardText.ForeColor = &H80000012
FrmEditDB.LbCBoardText.Enabled = True
FrmOptions.LbCToGame.ForeColor = &H8000000F
FrmOptions.LbCToGame.Enabled = False
FrmOptions.MnuBack.Enabled = False
FrmOptions.LbCPlayers.ForeColor = &H80000012
FrmOptions.MnuEnterPlayers.Enabled = True
FrmOptions.LbCPlayers.Enabled = True

LowMon = False
CurPlayer = 1

Plyr.Index = "Number"
Plyr.MoveFirst
TotPlayers = 0

Do Until Plyr.EOF
    If Plyr.Fields("Number") <> 0 And Plyr.Fields("Number") <> 99 Then
        TotPlayers = TotPlayers + 1
        CurPlayer = (Plyr.Fields("Number"))
        FrmBoard.ImgCounter(Plyr.Fields("Number")).Visible = True
        FrmBoard.ImgCounter(Plyr.Fields("Number")).Picture = LoadPicture(App.Path & (Plyr.Fields("CounterPath")))
        Call PositionPlayer(Plyr.Fields("Square"))
    End If
Plyr.MoveNext
Loop

'Call BankProperty   'Create new list of property in Bank
Plyr.Index = "Number"
If TotPlayers > 1 Then
    CurPlayer = GetCurPlayer
    ViewPlayer = CurPlayer
    FrmOptions.Hide
    Plyr.Seek "=", CurPlayer
    FrmBoard.LblInfo.Caption = Plyr.Fields("Name") & " To Go"
Else
    Plyr.Seek "=", 99
    Plyr.Edit
    Plyr.Fields("Money") = BnkStartMon
    Plyr.Update
End If

End Sub

Public Sub BoardColour()    'Change board colour

FrmEditDB.CD1.CancelError = True
On Error GoTo ErrHandler
FrmEditDB.CD1.Flags = cdlCCRGBInit
FrmEditDB.CD1.ShowColor
BrdColour = FrmEditDB.CD1.Color
Call DrawBoard
Exit Sub

ErrHandler:
Exit Sub
End Sub

Public Sub BoardText()  'Change text settings
Dim Ctrl As Object: Dim i As Integer

With FrmEditDB
    .CD1.CancelError = True
    On Error GoTo ErrHandler
    .CD1.Flags = cdlCFBoth Or cdlCFEffects
    .CD1.ShowFont

FontName = .CD1.FontName
ForeCol = .CD1.Color
FontSize = .CD1.FontSize
FontBold = .CD1.FontBold
FontItalic = .CD1.FontItalic
FontUline = .CD1.FontUnderline
FontStrThru = .CD1.FontStrikethru

For Each Ctrl In FrmBoard.Controls
    If Ctrl.Name Like "LblName*" Or Ctrl.Name Like "LblPrice*" Then
        Ctrl.Font.Name = .CD1.FontName
        Ctrl.ForeColor = .CD1.Color
        Ctrl.Font.Size = .CD1.FontSize
        Ctrl.Font.Bold = .CD1.FontBold
        Ctrl.Font.Italic = .CD1.FontItalic
        Ctrl.Font.Underline = .CD1.FontUnderline
        Ctrl.Font.Strikethrough = .CD1.FontStrikethru
    End If
Next Ctrl
TextColour = .CD1.Color
Exit Sub
End With
ErrHandler:
Exit Sub
End Sub

Public Sub ButtonText()
Dim Ctrl As Object: Dim i As Integer

With FrmEditDB
    .CcdCmdText.CancelError = True
    On Error GoTo ErrHandler
    .CcdCmdText.Flags = cdlCFBoth Or cdlCFEffects
    .CcdCmdText.ShowFont

LbcFontName = .CcdCmdText.FontName
LbcForeCol = .CcdCmdText.Color
LbcFontSize = .CcdCmdText.FontSize
LbcFontBold = .CcdCmdText.FontBold
LbcFontItalic = .CcdCmdText.FontItalic
LbcFontUline = .CcdCmdText.FontUnderline
LbcFontStrThru = .CcdCmdText.FontStrikethru

Exit Sub
End With
ErrHandler:
Exit Sub
End Sub

Public Sub SetCmdText(ByVal FrmName As Form)
Dim Ctrl As Control, FormNme As Form

With FrmEditDB
For Each Ctrl In FrmName
    If Ctrl.Name Like "LbC*" Then
        Ctrl.Font.Name = LbcFontName
        Ctrl.ForeColor = LbcForeCol
        Ctrl.Font.Size = LbcFontSize
        Ctrl.Font.Bold = LbcFontBold
        Ctrl.Font.Italic = LbcFontItalic
        Ctrl.Font.Underline = LbcFontUline
        Ctrl.Font.Strikethrough = LbcFontStrThru
    End If
Next Ctrl
End With
End Sub

Public Sub BackToGame() 'Go back to game
Dim n As Integer

Plyr.Index = "Number"
If Plyr.RecordCount < 2 Then    'Not enough players entered
    n = MsgBox("Please enter player details", vbCritical, "Options")
Exit Sub
End If
Call UpdateHouses   'Re-create houses/hotels
Unload FrmOptions
FrmBoard.Show
End Sub
