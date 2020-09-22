VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmEditDB 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Edit DataBase"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   7605
      Left            =   68
      TabIndex        =   0
      Top             =   120
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   13414
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabHeight       =   520
      BackColor       =   12648384
      TabCaption(0)   =   "Property"
      TabPicture(0)   =   "FrmEditDB.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CboSet"
      Tab(0).Control(1)=   "TxtRent(7)"
      Tab(0).Control(2)=   "TxtPrice"
      Tab(0).Control(3)=   "LstProperty(12)"
      Tab(0).Control(4)=   "LstProperty(10)"
      Tab(0).Control(5)=   "LstProperty(8)"
      Tab(0).Control(6)=   "LstProperty(7)"
      Tab(0).Control(7)=   "LstProperty(5)"
      Tab(0).Control(8)=   "LstProperty(4)"
      Tab(0).Control(9)=   "LstProperty(2)"
      Tab(0).Control(10)=   "LstProperty(1)"
      Tab(0).Control(11)=   "TxtName"
      Tab(0).Control(12)=   "TxtRent(8)"
      Tab(0).Control(13)=   "TxtRent(9)"
      Tab(0).Control(14)=   "TxtRent(10)"
      Tab(0).Control(15)=   "TxtRent(11)"
      Tab(0).Control(16)=   "TxtRent(12)"
      Tab(0).Control(17)=   "LstProperty(9)"
      Tab(0).Control(18)=   "LstProperty(11)"
      Tab(0).Control(19)=   "LblPropNo"
      Tab(0).Control(20)=   "LbCAddRec"
      Tab(0).Control(21)=   "Label1"
      Tab(0).Control(22)=   "Label2"
      Tab(0).Control(23)=   "Label5"
      Tab(0).Control(24)=   "Label6"
      Tab(0).Control(25)=   "Label7"
      Tab(0).Control(26)=   "Label8"
      Tab(0).Control(27)=   "Label9"
      Tab(0).Control(28)=   "Label10"
      Tab(0).Control(29)=   "Label3"
      Tab(0).Control(30)=   "Label4"
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Cards"
      TabPicture(1)   =   "FrmEditDB.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraChance"
      Tab(1).Control(1)=   "FraCChest"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Sets"
      TabPicture(2)   =   "FrmEditDB.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LblSetInfoLab"
      Tab(2).Control(1)=   "LbCUpdateCols"
      Tab(2).Control(2)=   "LblHousePriceLab(7)"
      Tab(2).Control(3)=   "LblHousePriceLab(1)"
      Tab(2).Control(4)=   "LblHousePriceLab(2)"
      Tab(2).Control(5)=   "LblHousePriceLab(3)"
      Tab(2).Control(6)=   "LblHousePriceLab(4)"
      Tab(2).Control(7)=   "LblHousePriceLab(5)"
      Tab(2).Control(8)=   "LblHousePriceLab(6)"
      Tab(2).Control(9)=   "LblHousePriceLab(0)"
      Tab(2).Control(10)=   "LblSet8Lab"
      Tab(2).Control(11)=   "LblSet7Lab"
      Tab(2).Control(12)=   "LblSet6Lab"
      Tab(2).Control(13)=   "LblSet5Lab"
      Tab(2).Control(14)=   "LblSet4"
      Tab(2).Control(15)=   "LblSet3Lab"
      Tab(2).Control(16)=   "LblSet2Lab"
      Tab(2).Control(17)=   "LblSet1Lab"
      Tab(2).Control(18)=   "LblSet(7)"
      Tab(2).Control(19)=   "LblSet(6)"
      Tab(2).Control(20)=   "LblSet(5)"
      Tab(2).Control(21)=   "LblSet(4)"
      Tab(2).Control(22)=   "LblSet(3)"
      Tab(2).Control(23)=   "LblSet(2)"
      Tab(2).Control(24)=   "LblSet(1)"
      Tab(2).Control(25)=   "LblSet(0)"
      Tab(2).Control(26)=   "LblHouseColLab(0)"
      Tab(2).Control(27)=   "LblHotelColLab(0)"
      Tab(2).Control(28)=   "LblHouseCol(0)"
      Tab(2).Control(29)=   "LblHotelCol(0)"
      Tab(2).Control(30)=   "LblHouseColLab(1)"
      Tab(2).Control(31)=   "LblHotelColLab(1)"
      Tab(2).Control(32)=   "LblHouseCol(1)"
      Tab(2).Control(33)=   "LblHotelCol(1)"
      Tab(2).Control(34)=   "LblHouseColLab(2)"
      Tab(2).Control(35)=   "LblHotelColLab(2)"
      Tab(2).Control(36)=   "LblHouseCol(2)"
      Tab(2).Control(37)=   "LblHotelCol(2)"
      Tab(2).Control(38)=   "LblHouseColLab(3)"
      Tab(2).Control(39)=   "LblHotelColLab(3)"
      Tab(2).Control(40)=   "LblHouseCol(3)"
      Tab(2).Control(41)=   "LblHotelCol(3)"
      Tab(2).Control(42)=   "LblHouseColLab(4)"
      Tab(2).Control(43)=   "LblHotelColLab(4)"
      Tab(2).Control(44)=   "LblHouseCol(4)"
      Tab(2).Control(45)=   "LblHotelCol(4)"
      Tab(2).Control(46)=   "LblHouseColLab(5)"
      Tab(2).Control(47)=   "LblHotelColLab(5)"
      Tab(2).Control(48)=   "LblHouseCol(5)"
      Tab(2).Control(49)=   "LblHotelCol(5)"
      Tab(2).Control(50)=   "LblHouseColLab(6)"
      Tab(2).Control(51)=   "LblHotelColLab(6)"
      Tab(2).Control(52)=   "LblHouseCol(6)"
      Tab(2).Control(53)=   "LblHotelCol(6)"
      Tab(2).Control(54)=   "LblHouseColLab(7)"
      Tab(2).Control(55)=   "LblHotelColLab(7)"
      Tab(2).Control(56)=   "LblHouseCol(7)"
      Tab(2).Control(57)=   "LblHotelCol(7)"
      Tab(2).Control(58)=   "LblApplyAll"
      Tab(2).Control(59)=   "CD2"
      Tab(2).Control(60)=   "TxtHPrice(7)"
      Tab(2).Control(61)=   "TxtHPrice(6)"
      Tab(2).Control(62)=   "TxtHPrice(5)"
      Tab(2).Control(63)=   "TxtHPrice(4)"
      Tab(2).Control(64)=   "TxtHPrice(3)"
      Tab(2).Control(65)=   "TxtHPrice(2)"
      Tab(2).Control(66)=   "TxtHPrice(1)"
      Tab(2).Control(67)=   "TxtHPrice(0)"
      Tab(2).ControlCount=   68
      TabCaption(3)   =   "Names"
      TabPicture(3)   =   "FrmEditDB.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "LblNameLab(8)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "LblNameLab(7)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "LblNameLab(6)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "LblNameLab(5)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "LblNameLab(4)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "LblNameLab(3)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "LblNameLab(2)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "LblNameLab(1)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "LblNameLab(0)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "LbCUpdateNames"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "LblNamesLab"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "LblCurSymbLab(9)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "LbCBoardText"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "LbCBoardColour"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "LblNameLab(12)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "LblNameLab(13)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "LblFParkingLab(9)"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "LblChanceLab(0)"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "LblCommChestLab(1)"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "LbCButtonText"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "CD1"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "TxtNamesName(0)"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "TxtNamesName(1)"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "TxtNamesName(2)"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "TxtNamesName(3)"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "TxtNamesName(4)"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "TxtNamesName(6)"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "TxtNamesName(7)"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "TxtNamesName(8)"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "TxtNamesName(5)"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "TxtNamesName(9)"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "CboSymPos"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "TxtNamesName(10)"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "TxtNamesName(11)"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "TxtNamesName(12)"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).Control(35)=   "CcdCmdText"
      Tab(3).Control(35).Enabled=   0   'False
      Tab(3).ControlCount=   36
      TabCaption(4)   =   "Rules"
      TabPicture(4)   =   "FrmEditDB.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "CboHotelNo"
      Tab(4).Control(1)=   "CboFinesTo"
      Tab(4).Control(2)=   "TxtStartProps"
      Tab(4).Control(3)=   "CboSellEven"
      Tab(4).Control(4)=   "TxtBkStartMon"
      Tab(4).Control(5)=   "TxtPlyrStartMon"
      Tab(4).Control(6)=   "TxtSalary"
      Tab(4).Control(7)=   "TxtDoubles"
      Tab(4).Control(8)=   "CboBuildEven"
      Tab(4).Control(9)=   "TxtJailFine"
      Tab(4).Control(10)=   "HScrCPause"
      Tab(4).Control(11)=   "LblHotelNoLab"
      Tab(4).Control(12)=   "LblPayFinesToLab"
      Tab(4).Control(13)=   "LblPropsLab"
      Tab(4).Control(14)=   "LblStartPropsLab"
      Tab(4).Control(15)=   "LblRuleLab(3)"
      Tab(4).Control(16)=   "LblBkStartMonLab"
      Tab(4).Control(17)=   "LblPlyrStartMonLab"
      Tab(4).Control(18)=   "LblSalaryLab"
      Tab(4).Control(19)=   "LblRuleLab(2)"
      Tab(4).Control(20)=   "LblRuleLab(1)"
      Tab(4).Control(21)=   "LblRuleLab(0)"
      Tab(4).Control(22)=   "LbCRules"
      Tab(4).Control(23)=   "LblNameLab(11)"
      Tab(4).ControlCount=   24
      Begin MSComDlg.CommonDialog CcdCmdText 
         Left            =   2280
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         FontName        =   "Hobo"
         FontSize        =   12
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   12
         Left            =   8280
         TabIndex        =   174
         Text            =   "Text1"
         Top             =   3132
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   11
         Left            =   8280
         TabIndex        =   172
         Text            =   "Text1"
         Top             =   2580
         Width           =   1575
      End
      Begin VB.ComboBox CboHotelNo 
         Height          =   315
         ItemData        =   "FrmEditDB.frx":008C
         Left            =   -73200
         List            =   "FrmEditDB.frx":009F
         TabIndex        =   170
         Top             =   4110
         Width           =   855
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   10
         Left            =   4800
         TabIndex        =   168
         Text            =   "Text1"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.ComboBox CboFinesTo 
         Height          =   315
         ItemData        =   "FrmEditDB.frx":00B2
         Left            =   -69240
         List            =   "FrmEditDB.frx":00B4
         TabIndex        =   167
         Top             =   4110
         Width           =   1455
      End
      Begin VB.TextBox TxtStartProps 
         Height          =   285
         Left            =   -69240
         TabIndex        =   164
         Text            =   "Text1"
         Top             =   2685
         Width           =   855
      End
      Begin VB.ComboBox CboSellEven 
         Height          =   315
         ItemData        =   "FrmEditDB.frx":00B6
         Left            =   -73200
         List            =   "FrmEditDB.frx":00C0
         TabIndex        =   162
         Top             =   2670
         Width           =   855
      End
      Begin VB.TextBox TxtBkStartMon 
         Height          =   285
         Left            =   -69240
         TabIndex        =   160
         Text            =   "Text1"
         Top             =   1485
         Width           =   855
      End
      Begin VB.TextBox TxtPlyrStartMon 
         Height          =   285
         Left            =   -69240
         TabIndex        =   159
         Text            =   "Text1"
         Top             =   2085
         Width           =   855
      End
      Begin VB.TextBox TxtSalary 
         Height          =   285
         Left            =   -69240
         TabIndex        =   158
         Text            =   "Text1"
         Top             =   3405
         Width           =   855
      End
      Begin VB.TextBox TxtDoubles 
         Height          =   285
         Left            =   -73200
         TabIndex        =   154
         Text            =   "Text1"
         Top             =   3405
         Width           =   855
      End
      Begin VB.ComboBox CboBuildEven 
         Height          =   315
         ItemData        =   "FrmEditDB.frx":00CD
         Left            =   -73200
         List            =   "FrmEditDB.frx":00D7
         TabIndex        =   152
         Top             =   2070
         Width           =   855
      End
      Begin VB.TextBox TxtJailFine 
         Height          =   285
         Left            =   -73200
         TabIndex        =   149
         Text            =   "Text1"
         Top             =   1485
         Width           =   855
      End
      Begin VB.HScrollBar HScrCPause 
         Height          =   285
         LargeChange     =   50
         Left            =   -69960
         Max             =   5
         Min             =   400
         SmallChange     =   5
         TabIndex        =   146
         Top             =   5205
         Value           =   50
         Width           =   1575
      End
      Begin VB.ComboBox CboSymPos 
         Height          =   315
         ItemData        =   "FrmEditDB.frx":00E4
         Left            =   4800
         List            =   "FrmEditDB.frx":00EE
         TabIndex        =   106
         Text            =   "Before"
         ToolTipText     =   "Select position for symbol"
         Top             =   5700
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   9
         Left            =   1560
         TabIndex        =   105
         Text            =   "Text1"
         Top             =   5700
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   103
         Text            =   "Text1"
         Top             =   4236
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   102
         Text            =   "Text1"
         Top             =   4788
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   7
         Left            =   4800
         TabIndex        =   99
         Text            =   "Text1"
         Top             =   4236
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   6
         Left            =   4800
         TabIndex        =   98
         Text            =   "Text1"
         Top             =   3132
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   97
         Text            =   "Text1"
         Top             =   3684
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   3
         Left            =   4800
         TabIndex        =   96
         Text            =   "Text1"
         Top             =   3684
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   2
         Left            =   4800
         TabIndex        =   95
         Text            =   "Text1"
         Top             =   2580
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   94
         Text            =   "Text1"
         Top             =   3132
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   93
         Text            =   "Text1"
         Top             =   2580
         Width           =   1575
      End
      Begin VB.ComboBox CboSet 
         Height          =   315
         ItemData        =   "FrmEditDB.frx":0101
         Left            =   -70440
         List            =   "FrmEditDB.frx":0126
         TabIndex        =   45
         Text            =   "Set No"
         Top             =   6360
         Width           =   600
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   7
         Left            =   -68850
         TabIndex        =   44
         Text            =   "Rent"
         Top             =   6360
         Width           =   765
      End
      Begin VB.TextBox TxtPrice 
         Height          =   315
         Left            =   -69720
         TabIndex        =   43
         Text            =   "Price"
         Top             =   6360
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   4545
         Index           =   12
         Left            =   -64350
         TabIndex        =   42
         Top             =   1530
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   4545
         Index           =   10
         Left            =   -66150
         TabIndex        =   41
         Top             =   1530
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   4545
         Index           =   8
         Left            =   -67950
         TabIndex        =   40
         Top             =   1530
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   4545
         Index           =   7
         Left            =   -68850
         TabIndex        =   39
         Top             =   1530
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   4545
         Index           =   5
         Left            =   -69750
         TabIndex        =   38
         Top             =   1530
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   4545
         Index           =   4
         Left            =   -70440
         TabIndex        =   37
         Top             =   1530
         Width           =   600
      End
      Begin VB.ListBox LstProperty 
         Height          =   4545
         Index           =   2
         Left            =   -74040
         TabIndex        =   36
         Top             =   1530
         Width           =   3465
      End
      Begin VB.ListBox LstProperty 
         Height          =   4545
         Index           =   1
         Left            =   -74700
         TabIndex        =   35
         Top             =   1530
         Width           =   615
      End
      Begin VB.TextBox TxtName 
         Height          =   315
         Left            =   -74010
         TabIndex        =   34
         Text            =   "Property Name"
         Top             =   6360
         Width           =   3525
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   8
         Left            =   -67920
         TabIndex        =   33
         Text            =   "Rent 1 House"
         Top             =   6360
         Width           =   765
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   9
         Left            =   -66960
         TabIndex        =   32
         Text            =   "Rent 2 Houses"
         Top             =   6360
         Width           =   765
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   10
         Left            =   -66120
         TabIndex        =   31
         Text            =   "Rent 3 Houses"
         Top             =   6360
         Width           =   765
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   11
         Left            =   -65280
         TabIndex        =   30
         Text            =   "Rent 4 Houses"
         Top             =   6360
         Width           =   765
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   12
         Left            =   -64320
         TabIndex        =   29
         Text            =   "Rent Hotel"
         Top             =   6360
         Width           =   765
      End
      Begin VB.Frame FraChance 
         Caption         =   "Chance"
         Height          =   2895
         Left            =   -74760
         TabIndex        =   20
         Top             =   1080
         Width           =   10935
         Begin VB.ComboBox CboAmount 
            Height          =   315
            Index           =   0
            Left            =   9360
            TabIndex        =   144
            Top             =   2040
            Width           =   1455
         End
         Begin VB.ListBox LstChance 
            Height          =   1620
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   615
         End
         Begin VB.ListBox LstChance 
            Height          =   1620
            Index           =   1
            Left            =   840
            TabIndex        =   25
            Top             =   240
            Width           =   6375
         End
         Begin VB.TextBox TxtText 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   24
            Text            =   "Card Text"
            Top             =   2040
            Width           =   6375
         End
         Begin VB.ComboBox CboAction 
            Height          =   315
            Index           =   0
            Left            =   7320
            TabIndex        =   23
            Text            =   "Action"
            Top             =   2040
            Width           =   1935
         End
         Begin VB.ListBox LstChance 
            Height          =   1620
            Index           =   2
            Left            =   7320
            TabIndex        =   22
            Top             =   240
            Width           =   1935
         End
         Begin VB.ListBox LstChance 
            Height          =   1620
            Index           =   3
            Left            =   9360
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label LblCardNo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Card No"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label LbCUpdateChance 
            Alignment       =   2  'Center
            Caption         =   "Change &Card"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5220
            TabIndex        =   27
            Top             =   2400
            Width           =   1455
         End
      End
      Begin VB.Frame FraCChest 
         Caption         =   "Community Chest"
         Height          =   2895
         Left            =   -74760
         TabIndex        =   11
         Top             =   4200
         Width           =   10935
         Begin VB.ComboBox CboAmount 
            Height          =   315
            Index           =   1
            Left            =   9360
            TabIndex        =   145
            Top             =   2040
            Width           =   1455
         End
         Begin VB.ListBox LstCChest 
            Height          =   1620
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   615
         End
         Begin VB.ListBox LstCChest 
            Height          =   1620
            Index           =   1
            Left            =   840
            TabIndex        =   16
            Top             =   240
            Width           =   6375
         End
         Begin VB.TextBox TxtText 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   15
            Text            =   "Card Text"
            Top             =   2040
            Width           =   6375
         End
         Begin VB.ComboBox CboAction 
            Height          =   315
            Index           =   1
            Left            =   7320
            TabIndex        =   14
            Text            =   "Action"
            Top             =   2040
            Width           =   1935
         End
         Begin VB.ListBox LstCChest 
            Height          =   1620
            Index           =   2
            Left            =   7320
            TabIndex        =   13
            Top             =   240
            Width           =   1935
         End
         Begin VB.ListBox LstCChest 
            Height          =   1620
            Index           =   3
            Left            =   9360
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label LblCardNo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Card No"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label LbCUpdateChest 
            Alignment       =   2  'Center
            Caption         =   "Change &Card"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5220
            TabIndex        =   18
            Top             =   2400
            Width           =   1455
         End
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -72907
         TabIndex        =   10
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   2325
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -72907
         TabIndex        =   9
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   5085
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -70387
         TabIndex        =   8
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   2325
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -70387
         TabIndex        =   7
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   5085
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   -67627
         TabIndex        =   6
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   2325
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   -67627
         TabIndex        =   5
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   5085
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   -64867
         TabIndex        =   4
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   2325
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   -64747
         TabIndex        =   3
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   5085
         Width           =   735
      End
      Begin VB.ListBox LstProperty 
         Height          =   4545
         Index           =   9
         Left            =   -67050
         TabIndex        =   2
         Top             =   1530
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   4545
         Index           =   11
         Left            =   -65250
         TabIndex        =   1
         Top             =   1530
         Width           =   765
      End
      Begin MSComDlg.CommonDialog CD2 
         Left            =   -74760
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   1320
         Top             =   1260
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         InitDir         =   "app.path"
      End
      Begin VB.Label LbCButtonText 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Command &Button Text"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9960
         TabIndex        =   178
         Top             =   5700
         Width           =   975
      End
      Begin VB.Label LblCommChestLab 
         Caption         =   "Community Chest"
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   175
         Top             =   3087
         Width           =   1095
      End
      Begin VB.Label LblChanceLab 
         Caption         =   "Chance"
         Height          =   375
         Index           =   0
         Left            =   7200
         TabIndex        =   173
         Top             =   2535
         Width           =   1095
      End
      Begin VB.Label LblHotelNoLab 
         Caption         =   "Hotel = Houses"
         Height          =   375
         Left            =   -74280
         TabIndex        =   171
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label LblFParkingLab 
         Caption         =   "Free Parking"
         Height          =   375
         Index           =   9
         Left            =   3720
         TabIndex        =   169
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label LblPayFinesToLab 
         Caption         =   "Pay Fines To"
         Height          =   375
         Left            =   -70440
         TabIndex        =   166
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label LblPropsLab 
         Alignment       =   1  'Right Justify
         Caption         =   "Properties"
         Height          =   375
         Left            =   -68400
         TabIndex        =   165
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label LblStartPropsLab 
         Caption         =   "Players Start With"
         Height          =   375
         Left            =   -70440
         TabIndex        =   163
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label LblRuleLab 
         Caption         =   "Sell Evenly"
         Height          =   375
         Index           =   3
         Left            =   -74280
         TabIndex        =   161
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label LblBkStartMonLab 
         Caption         =   "Bank Start Money"
         Height          =   375
         Left            =   -70440
         TabIndex        =   157
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label LblPlyrStartMonLab 
         Caption         =   "Player Start Money"
         Height          =   375
         Left            =   -70440
         TabIndex        =   156
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label LblSalaryLab 
         Caption         =   "Salary"
         Height          =   375
         Left            =   -70440
         TabIndex        =   155
         Top             =   3420
         Width           =   735
      End
      Begin VB.Label LblRuleLab 
         Caption         =   "Doubles Before Jail"
         Height          =   375
         Index           =   2
         Left            =   -74280
         TabIndex        =   153
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label LblRuleLab 
         Caption         =   "Build Evenly"
         Height          =   375
         Index           =   1
         Left            =   -74280
         TabIndex        =   151
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label LblRuleLab 
         Caption         =   "Get Out of Jail Fine"
         Height          =   375
         Index           =   0
         Left            =   -74280
         TabIndex        =   150
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label LbCRules 
         Alignment       =   2  'Center
         Caption         =   "&ApplyChanges"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69840
         TabIndex        =   148
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Counter Speed"
         Height          =   375
         Index           =   11
         Left            =   -71160
         TabIndex        =   147
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Value"
         Height          =   375
         Index           =   13
         Left            =   6600
         TabIndex        =   143
         Top             =   5700
         Width           =   735
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Positioned"
         Height          =   375
         Index           =   12
         Left            =   3720
         TabIndex        =   142
         Top             =   5700
         Width           =   735
      End
      Begin VB.Label LbCBoardColour 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Board &Colour"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7440
         TabIndex        =   141
         Top             =   5820
         Width           =   975
      End
      Begin VB.Label LbCBoardText 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Board Text"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8760
         TabIndex        =   140
         Top             =   5820
         Width           =   975
      End
      Begin VB.Label LblApplyAll 
         Caption         =   "Apply house and hotel colour to all &Sets"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   139
         Top             =   3720
         Width           =   10095
      End
      Begin VB.Label LblHotelCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   -64380
         TabIndex        =   138
         Top             =   5985
         Width           =   375
      End
      Begin VB.Label LblHouseCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   -64387
         TabIndex        =   137
         Top             =   5535
         Width           =   375
      End
      Begin VB.Label LblHotelColLab 
         Caption         =   "Hotel Colour:"
         Height          =   375
         Index           =   7
         Left            =   -65940
         TabIndex        =   136
         Top             =   5940
         Width           =   1575
      End
      Begin VB.Label LblHouseColLab 
         Caption         =   "House Colour:"
         Height          =   375
         Index           =   7
         Left            =   -65940
         TabIndex        =   135
         Top             =   5490
         Width           =   1575
      End
      Begin VB.Label LblHotelCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   -64507
         TabIndex        =   134
         Top             =   3225
         Width           =   375
      End
      Begin VB.Label LblHouseCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   -64500
         TabIndex        =   133
         Top             =   2775
         Width           =   375
      End
      Begin VB.Label LblHotelColLab 
         Caption         =   "Hotel Colour:"
         Height          =   375
         Index           =   6
         Left            =   -66060
         TabIndex        =   132
         Top             =   3180
         Width           =   1575
      End
      Begin VB.Label LblHouseColLab 
         Caption         =   "House Colour:"
         Height          =   375
         Index           =   6
         Left            =   -66060
         TabIndex        =   131
         Top             =   2730
         Width           =   1575
      End
      Begin VB.Label LblHotelCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   -67260
         TabIndex        =   130
         Top             =   5985
         Width           =   375
      End
      Begin VB.Label LblHouseCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   -67260
         TabIndex        =   129
         Top             =   5535
         Width           =   375
      End
      Begin VB.Label LblHotelColLab 
         Caption         =   "Hotel Colour:"
         Height          =   375
         Index           =   5
         Left            =   -68820
         TabIndex        =   128
         Top             =   5940
         Width           =   1575
      End
      Begin VB.Label LblHouseColLab 
         Caption         =   "House Colour:"
         Height          =   375
         Index           =   5
         Left            =   -68820
         TabIndex        =   127
         Top             =   5490
         Width           =   1575
      End
      Begin VB.Label LblHotelCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   -67260
         TabIndex        =   126
         Top             =   3225
         Width           =   375
      End
      Begin VB.Label LblHouseCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   -67260
         TabIndex        =   125
         Top             =   2775
         Width           =   375
      End
      Begin VB.Label LblHotelColLab 
         Caption         =   "Hotel Colour:"
         Height          =   375
         Index           =   4
         Left            =   -68820
         TabIndex        =   124
         Top             =   3180
         Width           =   1575
      End
      Begin VB.Label LblHouseColLab 
         Caption         =   "House Colour:"
         Height          =   375
         Index           =   4
         Left            =   -68820
         TabIndex        =   123
         Top             =   2730
         Width           =   1575
      End
      Begin VB.Label LblHotelCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   -70020
         TabIndex        =   122
         Top             =   5985
         Width           =   375
      End
      Begin VB.Label LblHouseCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   -70020
         TabIndex        =   121
         Top             =   5535
         Width           =   375
      End
      Begin VB.Label LblHotelColLab 
         Caption         =   "Hotel Colour:"
         Height          =   375
         Index           =   3
         Left            =   -71580
         TabIndex        =   120
         Top             =   5940
         Width           =   1575
      End
      Begin VB.Label LblHouseColLab 
         Caption         =   "House Colour:"
         Height          =   375
         Index           =   3
         Left            =   -71580
         TabIndex        =   119
         Top             =   5490
         Width           =   1575
      End
      Begin VB.Label LblHotelCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   -70020
         TabIndex        =   118
         Top             =   3225
         Width           =   375
      End
      Begin VB.Label LblHouseCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   -70020
         TabIndex        =   117
         Top             =   2775
         Width           =   375
      End
      Begin VB.Label LblHotelColLab 
         Caption         =   "Hotel Colour:"
         Height          =   375
         Index           =   2
         Left            =   -71580
         TabIndex        =   116
         Top             =   3180
         Width           =   1575
      End
      Begin VB.Label LblHouseColLab 
         Caption         =   "House Colour:"
         Height          =   375
         Index           =   2
         Left            =   -71580
         TabIndex        =   115
         Top             =   2730
         Width           =   1575
      End
      Begin VB.Label LblHotelCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   -72540
         TabIndex        =   114
         Top             =   5985
         Width           =   375
      End
      Begin VB.Label LblHouseCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   -72540
         TabIndex        =   113
         Top             =   5535
         Width           =   375
      End
      Begin VB.Label LblHotelColLab 
         Caption         =   "Hotel Colour:"
         Height          =   375
         Index           =   1
         Left            =   -74100
         TabIndex        =   112
         Top             =   5940
         Width           =   1575
      End
      Begin VB.Label LblHouseColLab 
         Caption         =   "House Colour:"
         Height          =   375
         Index           =   1
         Left            =   -74100
         TabIndex        =   111
         Top             =   5490
         Width           =   1575
      End
      Begin VB.Label LblHotelCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   -72540
         TabIndex        =   110
         Top             =   3225
         Width           =   375
      End
      Begin VB.Label LblHouseCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   -72540
         TabIndex        =   109
         Top             =   2775
         Width           =   375
      End
      Begin VB.Label LblHotelColLab 
         Caption         =   "Hotel Colour:"
         Height          =   375
         Index           =   0
         Left            =   -74100
         TabIndex        =   108
         Top             =   3180
         Width           =   1575
      End
      Begin VB.Label LblHouseColLab 
         Caption         =   "House Colour:"
         Height          =   375
         Index           =   0
         Left            =   -74100
         TabIndex        =   107
         Top             =   2730
         Width           =   1575
      End
      Begin VB.Label LblCurSymbLab 
         Caption         =   "Curency Symbol"
         Height          =   375
         Index           =   9
         Left            =   480
         TabIndex        =   104
         Top             =   5700
         Width           =   735
      End
      Begin VB.Label LblNamesLab 
         Caption         =   "Enter new names for the following in the boxes provided"
         Height          =   255
         Left            =   3893
         TabIndex        =   101
         Top             =   1740
         Width           =   4095
      End
      Begin VB.Label LbCUpdateNames 
         Alignment       =   2  'Center
         Caption         =   "&ApplyChanges"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5213
         TabIndex        =   100
         Top             =   6900
         Width           =   1455
      End
      Begin VB.Label LblPropNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PropNo"
         Height          =   315
         Left            =   -74700
         TabIndex        =   92
         Top             =   6360
         Width           =   615
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   -74100
         TabIndex        =   91
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   -74100
         TabIndex        =   90
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   -71580
         TabIndex        =   89
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   -71580
         TabIndex        =   88
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   -68820
         TabIndex        =   87
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   -68820
         TabIndex        =   86
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   -66060
         TabIndex        =   85
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   -65940
         TabIndex        =   84
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label LblSet1Lab 
         Caption         =   "Set 1"
         Height          =   255
         Left            =   -73500
         TabIndex        =   83
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label LblSet2Lab 
         Caption         =   "Set 2"
         Height          =   255
         Left            =   -73500
         TabIndex        =   82
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label LblSet3Lab 
         Caption         =   "Set 3"
         Height          =   255
         Left            =   -70980
         TabIndex        =   81
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label LblSet4 
         Caption         =   "Set 4"
         Height          =   255
         Left            =   -71100
         TabIndex        =   80
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label LblSet5Lab 
         Caption         =   "Set 5"
         Height          =   255
         Left            =   -68220
         TabIndex        =   79
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label LblSet6Lab 
         Caption         =   "Set 6"
         Height          =   255
         Left            =   -68220
         TabIndex        =   78
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label LblSet7Lab 
         Caption         =   "Set 7"
         Height          =   255
         Left            =   -65220
         TabIndex        =   77
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label LblSet8Lab 
         Caption         =   "Set 8"
         Height          =   255
         Left            =   -65340
         TabIndex        =   76
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   375
         Index           =   0
         Left            =   -74100
         TabIndex        =   75
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   375
         Index           =   6
         Left            =   -66060
         TabIndex        =   74
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   375
         Index           =   5
         Left            =   -68820
         TabIndex        =   73
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   375
         Index           =   4
         Left            =   -68820
         TabIndex        =   72
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   375
         Index           =   3
         Left            =   -71580
         TabIndex        =   71
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   375
         Index           =   2
         Left            =   -71580
         TabIndex        =   70
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   375
         Index           =   1
         Left            =   -74100
         TabIndex        =   69
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   375
         Index           =   7
         Left            =   -65940
         TabIndex        =   68
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label LbCAddRec 
         Alignment       =   2  'Center
         Caption         =   "&Apply Changes"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69907
         TabIndex        =   67
         Top             =   6960
         Width           =   1695
      End
      Begin VB.Label LbCUpdateCols 
         Alignment       =   2  'Center
         Caption         =   "&Apply Changes"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70027
         TabIndex        =   66
         Top             =   6960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Property Number"
         Height          =   375
         Left            =   -74640
         TabIndex        =   65
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Property Name"
         Height          =   375
         Left            =   -73680
         TabIndex        =   64
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Rent"
         Height          =   375
         Left            =   -68760
         TabIndex        =   63
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Rent 1 House"
         Height          =   375
         Left            =   -67920
         TabIndex        =   62
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "Rent 2 Houses"
         Height          =   375
         Left            =   -66960
         TabIndex        =   61
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label8 
         Caption         =   "Rent 3 Houses"
         Height          =   375
         Left            =   -66120
         TabIndex        =   60
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label9 
         Caption         =   "Rent 4 Houses"
         Height          =   375
         Left            =   -65160
         TabIndex        =   59
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label10 
         Caption         =   "Rent Hotel"
         Height          =   375
         Left            =   -64320
         TabIndex        =   58
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Set"
         Height          =   375
         Left            =   -70440
         TabIndex        =   57
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Price"
         Height          =   375
         Left            =   -69720
         TabIndex        =   56
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label LblSetInfoLab 
         Caption         =   $"FrmEditDB.frx":014C
         Height          =   615
         Left            =   -72300
         TabIndex        =   55
         Top             =   960
         Width           =   6375
      End
      Begin VB.Label LblNameLab 
         Caption         =   "House"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   54
         Top             =   2580
         Width           =   735
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Station"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   53
         Top             =   4788
         Width           =   735
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Hotel"
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   52
         Top             =   3132
         Width           =   735
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Go"
         Height          =   375
         Index           =   3
         Left            =   3720
         TabIndex        =   51
         Top             =   2580
         Width           =   735
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Jail"
         Height          =   375
         Index           =   4
         Left            =   3720
         TabIndex        =   50
         Top             =   3684
         Width           =   735
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Bank"
         Height          =   375
         Index           =   5
         Left            =   480
         TabIndex        =   49
         Top             =   3684
         Width           =   735
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Deed"
         Height          =   375
         Index           =   6
         Left            =   480
         TabIndex        =   48
         Top             =   4236
         Width           =   735
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Rent"
         Height          =   375
         Index           =   7
         Left            =   3720
         TabIndex        =   47
         Top             =   3132
         Width           =   735
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Utility"
         Height          =   375
         Index           =   8
         Left            =   3720
         TabIndex        =   46
         Top             =   4236
         Width           =   735
      End
   End
   Begin VB.Label LbCFinished 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Finished"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4215
      TabIndex        =   177
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label LbCToFile 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Save as New &Version"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6135
      TabIndex        =   176
      Top             =   7800
      Width           =   2535
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuSaveVers 
         Caption         =   "Save &Version"
      End
      Begin VB.Menu MnuFinished 
         Caption         =   "&Finished"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuProperty 
         Caption         =   "&Property"
      End
      Begin VB.Menu MnuCards 
         Caption         =   "&Cards"
      End
      Begin VB.Menu MnuSets 
         Caption         =   "&Sets"
      End
      Begin VB.Menu MnuNames 
         Caption         =   "&Names"
      End
   End
End
Attribute VB_Name = "FrmEditDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboAmount_Click(Index As Integer)
Dim i As Integer
If (CboAction(Index).Text <> "Advance To" And CboAction(Index).Text <> "Back To") And _
    (Mid(CboAmount(Index), 1, 1) = "N" Or Mid(CboAmount(Index), 1, 1) = "L") Then
    i = MsgBox("Sorry, This value can't be used with the selected action", vbInformation, "")
End If
End Sub

Private Sub Form_Activate()
Call SetCmdText(FrmEditDB)
FrmEditDB.LbCFinished.BackColor = BrdColour
FrmEditDB.BackColor = BrdColour
End Sub

Private Sub HScrCPause_Change()
CounterPause = HScrCPause.Value
End Sub

Private Sub LbCAddRec_Click()
Call UpdateDBProperty           'Update DataBase
Call CreateLists                'Update property list on Board
End Sub

Private Sub LbCBoardColour_Click()
Call ModOptions.BoardColour      'Change Board Colour
End Sub

Private Sub LbCBoardText_Click()
Call ModOptions.BoardText      'Change board text appearance
End Sub

Private Sub LbCButtonText_Click()
Call ModOptions.ButtonText
End Sub

Private Sub LbCfinished_Click()
Call UpdateDBCols
Call UpdateNames
Call UpdateRules
Call DrawBoard          'Re-draw (Update) Board
Call UpdateHouses       'Re-draw House/Hotels
Unload FrmEditDB
End Sub

Private Sub LbCRules_Click()
Dim n As Integer
Call UpdateRules
n = MsgBox("Rules Updated", vbInformation, "Updated")
End Sub

Private Sub LbCUpdateChance_Click()  'Update DataBase
Call UpdateDBChance
Call CreateLists
End Sub

Private Sub LbCUpdateChest_Click()  'Update DataBase
Call UpdateDBCChest
Call CreateLists
End Sub

Private Sub LbCUpdateCols_Click()
Dim n As Integer
Call UpdateDBCols           'Update DataBase with new Colours
n = MsgBox("Set Info Updated", vbInformation, "Updated")
End Sub

Private Sub LbCUpdateNames_Click()
Dim n As Integer
Call UpdateNames
n = MsgBox("Names Updated", vbInformation, "Updated")
End Sub

Private Sub LblApplyAll_Click()
Dim i As Integer

For i = 1 To 7
LblHouseCol(i).BackColor = LblHouseCol(0).BackColor
LblHotelCol(i).BackColor = LblHotelCol(0).BackColor
Next i

End Sub

Private Sub LblSet_Click(Index As Integer)  'Allow user to Choose Colour
CD2.CancelError = True
On Error GoTo ErrHandler
CD2.Flags = cdlCCRGBInit
CD2.ShowColor
LblSet(Index).BackColor = CD2.Color
Exit Sub

ErrHandler:
MsgBox "Error"
Exit Sub

End Sub

Private Sub LblHouseCol_Click(Index As Integer)  'Allow user to Choose Colour
CD2.CancelError = True
On Error GoTo ErrHandler
CD2.Flags = cdlCCRGBInit
CD2.ShowColor
LblHouseCol(Index).BackColor = CD2.Color
Exit Sub

ErrHandler:
MsgBox "Error"
Exit Sub

End Sub

Private Sub LblHotelCol_Click(Index As Integer)  'Allow user to Choose Colour
CD2.CancelError = True
On Error GoTo ErrHandler
CD2.Flags = cdlCCRGBInit
CD2.ShowColor
LblHotelCol(Index).BackColor = CD2.Color
Exit Sub

ErrHandler:
MsgBox "Error"
Exit Sub

End Sub

Private Sub LbCToFile_Click()
Call UpdateDBCols
Call UpdateNames
Call UpdateRules
Call SaveVersion
MsgBox "File Saved"
Call DrawBoard          'Re-draw (Update) Board
Call UpdateHouses       'Re-draw House/Hotels
Unload FrmEditDB
End Sub

Private Sub LstCChest_Click(Index As Integer)
Dim Clicked As Integer
Clicked = LstCChest(Index).Index
Call CChestSelected(Clicked)    'Select same card in all other fields
End Sub

Private Sub LstChance_Click(Index As Integer)
Dim Clicked As Integer
Clicked = LstChance(Index).Index
Call ChanceSelected(Clicked)    'Select same card in all other fields
End Sub

Private Sub LstProperty_Click(Index As Integer)
Dim Clicked As Integer
Clicked = LstProperty(Index).Index
Call PropertySelected(Clicked)  'Select same property in all other fields
End Sub

Private Sub MnuCards_Click()
SSTab1.Tab = 1
End Sub

Private Sub MnuFinished_Click()
Call DrawBoard          'Re-draw (Update) Board
Call UpdateHouses       'Re-draw House/Hotels
Me.Hide
End Sub

Private Sub MnuNames_Click()
SSTab1.Tab = 3
End Sub

Private Sub MnuProperty_Click()
SSTab1.Tab = 0
End Sub

Private Sub MnuSaveVers_Click()
Call SaveVersion
MsgBox "File Saved"
Call DrawBoard          'Re-draw (Update) Board
Call UpdateHouses       'Re-draw House/Hotels
Me.Hide
End Sub

Private Sub MnuSets_Click()
SSTab1.Tab = 2
End Sub

Private Sub TxtPrice_Click()
FrmEditDB.TxtPrice.Text = ""
End Sub

Private Sub TxtRent_Click(Index As Integer)
FrmEditDB.TxtRent(Index).Text = ""
End Sub
