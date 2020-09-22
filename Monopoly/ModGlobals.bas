Attribute VB_Name = "ModGlobals"
'***************************************************************************
'Programme:  Monopoly game
'
'Files:          Monopoly.vbp,
'                FrmAuction.frm, FrmBoard.frm , FrmCard.frm, FrmEditDB.frm
'                FrmOptions.frm, FrmPlayers.frm , FrmProperty.frm, FrmTrade.frm
'
'                ModDrawBoard.bas , ModDataBase.bas, ModEditDB.bas, ModFunctions.bas
'                ModGame.bas , ModGameFunctions.bas, ModGlobals.bas
'                ModOptions.bas , ModPlayers.bas, ModTrade.bas,
'
'                Versions.dat, UserVers.dat, Data.mdb
'
'                icon files
'
'Function:       To be a flexible game of monopoly that will allow the user to make
'                alterations to both the content of the game and the appearance of the
'                games' user interface.
'
'Description:    The game allows 2 to 6 players who take turns to click 'roll dice'.
'                Then the player moves a number of squares depending or the randomly
'                generated number on the dice.  The player will then be asked to
'                pay rent, buy property as appropriate.  In addition players may also
'                receive a chance or community chest card and, through the 'Trade' form
'                can buy & sell houses as well as sell property to other players or the bank.
'
'Author:         David Johnston
'
'Environments:   MS Visual Basic 6.0, Pentium III 500MHZ, 128mb RAM, Windows 98SE.
'                Pentium 100MHZ, 64mb RAM, Windows 98SE
'
'Notes:          In most instances the user is prevented from entering invalid data through
'                the use of error messages in the form of message boxes.
'                Message boxes are also used to provide the user with information.
'                This programme won't run at resolutions of less than 800 x 600.
'
'Revisions:      1.00   12/3/2002 Version 1
'                2.00 21/4/2002 Version 2
'                20/12/2002 Version 3 - Greater error checking. Fewer bugs.
'                                       Improved User Interface
'                13/2/2003 Version 4 -  More versions. Better save game function.  Auction property.
'                                       More customization options.
'
'
'***************************************************************************

Public DB As Database
Public Prop As Recordset, PropSet As Recordset
Public Plyr As Recordset, Vers As Recordset
Public Counter As Recordset, Chnce As Recordset, CChest As Recordset
Public DBPath As String, BrdColour As String, TextColour As String
Public FontName As String, ForeCol As String, FontSize As String
Public FontBold As String, FontItalic As String, FontUline As String
Public FontStrThru As String
Public LbcFontName As String, LbcForeCol As String, LbcFontSize As String
Public LbcFontBold As String, LbcFontItalic As String, LbcFontUline As String
Public LbcFontStrThru As String
Public LowRes As Integer, FHeight As Integer, FWidth As Integer
Public SqBShort As Integer, SqSShort As Integer, Corner As Integer
Public XPos As Integer, YPos As Integer
Public ViewPlayer As Integer, CurPlayer As Integer, TotPlayers As Integer
Public PlyrAdd As Integer, CounterNumb As Integer
Public Dice1 As Integer, Dice2 As Integer, DoublesCount As Integer
Public ResComp As Single, CounterPause As Single, AucPrice As Single
Public LowMon As Boolean
Public StartTm As Date
Public AmountOwed As Single
Public VersName As String, House As String, Hotel As String
Public Go As String, Jail As String, Bank As String, FParking As String
Public ChanceNme As String, CommChestNme As String
Public PropInfo As String, Rent As String, Utility As String, HotelNo As Integer
Public Station As String, CurrencySymb As String, CurSymbPos As String
Public CurSymbBefore As String, CurSymbAfter As String
Public BnkStartMon As Single, PlyrStartMon As Single, Salary As Single
Public JailFine As Single, JailDoubles As Integer, StartProps As Integer
Public BuildEven As Boolean, SellEven As Boolean, FreePark As Integer

Sub Main()  'First procedure to run
Dim CountersPath As String
Set DB = Nothing
StartTm = Now
CountersPath = (App.Path & "\Counters")
DoublesCount = 0

LbcFontName = "Hobo": LbcForeCol = vbBlack: LbcFontSize = 12
LbcFontBold = False: LbcFontItalic = False: LbcFontUline = False
LbcFontStrThru = False

VersName = "London UK"
Call ModDataBase.LoadDataBase     'Load the DataBase
Call SetRecordSets
Call ReSetDB
Call LoadVersion
If ResComp = 0 Then Call ResCheck
FrmOptions.Show     'Show Options Screen
End Sub

Public Sub CreateVersionList()
Dim IntFile As Integer, i As Integer, Vers As String
FrmOptions.CboVersion.Clear
IntFile = FreeFile
StrFileName = App.Path & "\Versions.dat"
For i = 0 To 1
If i = 1 Then StrFileName = App.Path & "\UserVers.dat"
Open StrFileName For Input As #IntFile
Do Until EOF(IntFile)
    Input #IntFile, Vers
    If Left(Vers, 1) = "." Then FrmOptions.CboVersion.AddItem (Right(Vers, Len(Vers) - 1))
Loop
Close IntFile
Next i
FrmOptions.CboVersion.Text = FrmOptions.CboVersion.List(0)

End Sub

Public Sub NewGame()    'Re-set the game
FrmPlayers.LstPlayerNo.Clear
FrmPlayers.LstPlayers.Clear
FrmPlayers.LblPlayerNumb.Caption = ""
Unload FrmBoard
TotPlayers = 0
PlyrAdd = 1
CurSymbBefore = ""
CurSymbAfter = ""
Set DB = Nothing
Call Main
End Sub
