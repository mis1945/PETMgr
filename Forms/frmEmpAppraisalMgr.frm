VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0919B38C-FEB7-4581-AC15-BCF315A77232}#1.0#0"; "xxxControl.ocx"
Begin VB.Form frmEmpAppraisalMgr 
   BorderStyle     =   0  'None
   Caption         =   "Employee Appraisal - Manager"
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6105
      Left            =   225
      TabIndex        =   22
      Tag             =   "wt0;fb0"
      Top             =   2760
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10769
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   -2147483634
      TabCaption(0)   =   "Legend"
      TabPicture(0)   =   "frmEmpAppraisalMgr.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3(12)"
      Tab(0).Control(1)=   "Label3(13)"
      Tab(0).Control(2)=   "Label3(14)"
      Tab(0).Control(3)=   "Label3(15)"
      Tab(0).Control(4)=   "Label3(16)"
      Tab(0).Control(5)=   "Label3(17)"
      Tab(0).Control(6)=   "Label3(18)"
      Tab(0).Control(7)=   "Label3(19)"
      Tab(0).Control(8)=   "Label3(20)"
      Tab(0).Control(9)=   "Label3(21)"
      Tab(0).Control(10)=   "Label3(22)"
      Tab(0).Control(11)=   "Label3(35)"
      Tab(0).Control(12)=   "Label3(36)"
      Tab(0).Control(13)=   "Label3(37)"
      Tab(0).Control(14)=   "Label3(38)"
      Tab(0).Control(15)=   "Label3(39)"
      Tab(0).Control(16)=   "Label3(40)"
      Tab(0).Control(17)=   "xrFrame3(13)"
      Tab(0).Control(18)=   "xrFrame3(12)"
      Tab(0).Control(19)=   "xrFrame3(11)"
      Tab(0).Control(20)=   "xrFrame3(10)"
      Tab(0).Control(21)=   "xrFrame3(9)"
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "I. - IV."
      TabPicture(1)   =   "frmEmpAppraisalMgr.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(11)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3(10)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3(9)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3(8)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label3(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label3(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label3(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label3(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label4"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Shape3(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Shape4(2)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label5"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Shape3(3)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Shape4(3)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label3(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Shape4(5)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Shape3(5)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label7"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label3(5)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label3(6)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label3(7)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label3(41)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Shape4(6)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Shape3(6)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label8"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label3(42)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label3(45)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label3(44)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label3(43)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "xrFrame3(4)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "xrFrame3(8)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "xrFrame3(7)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "xrFrame3(3)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "xrFrame3(2)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "xrFrame3(0)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "xrFrame3(1)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "xrFrame3(6)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "xrFrame3(5)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).ControlCount=   38
      TabCaption(2)   =   "V. - VII."
      TabPicture(2)   =   "frmEmpAppraisalMgr.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3(27)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label3(29)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Shape3(7)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Shape4(7)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label3(30)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Shape4(8)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Shape3(8)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label10"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Shape4(9)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Shape3(9)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label11"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label3(46)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label3(48)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label3(50)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label3(52)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label3(53)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label3(47)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label3(23)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Label3(24)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Label3(25)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "xrFrame3(21)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "xrFrame3(20)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "xrFrame3(23)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "xrFrame3(19)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "xrFrame3(18)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "xrFrame3(17)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "xrFrame3(16)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "xrFrame3(15)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "xrFrame3(14)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).ControlCount=   30
      TabCaption(3)   =   "VIII. - IX."
      TabPicture(3)   =   "frmEmpAppraisalMgr.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label3(34)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label3(33)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label3(32)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label3(31)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label3(26)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Shape4(4)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Shape3(4)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label6"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label3(28)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label3(49)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label3(51)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label3(54)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label3(55)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "xrFrame3(24)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "xrFrame3(22)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Text1"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Text3"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "Text2"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).ControlCount=   18
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   1140
         TabIndex        =   78
         Top             =   3645
         Width           =   9300
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   1140
         TabIndex        =   77
         Top             =   4710
         Width           =   9300
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   1140
         TabIndex        =   76
         Top             =   2565
         Width           =   9300
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   5
         Left            =   -67155
         Top             =   2355
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   5
            Left            =   -90
            TabIndex        =   24
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   6
         Left            =   -67155
         Top             =   2745
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   6
            Left            =   -90
            TabIndex        =   25
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   1
         Left            =   -67155
         Top             =   1350
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   1
            Left            =   -90
            TabIndex        =   41
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   9
         Left            =   -71100
         Top             =   735
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   9
            Left            =   -90
            TabIndex        =   46
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            Rating          =   10
            Enabled         =   0   'False
            StarCount       =   10
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   10
         Left            =   -71085
         Top             =   1395
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   10
            Left            =   -90
            TabIndex        =   47
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            Rating          =   8
            Enabled         =   0   'False
            StarCount       =   10
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   11
         Left            =   -71085
         Top             =   2055
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   11
            Left            =   -90
            TabIndex        =   48
            Top             =   -90
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            Rating          =   6
            Enabled         =   0   'False
            StarCount       =   10
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   12
         Left            =   -71085
         Top             =   2715
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   12
            Left            =   -90
            TabIndex        =   49
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            Rating          =   4
            Enabled         =   0   'False
            StarCount       =   10
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   13
         Left            =   -71070
         Top             =   3375
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   13
            Left            =   -90
            TabIndex        =   50
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            Rating          =   2
            Enabled         =   0   'False
            StarCount       =   10
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   0
         Left            =   -67155
         Top             =   885
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   0
            Left            =   -90
            TabIndex        =   59
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   2
         Left            =   -67155
         Top             =   3585
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   2
            Left            =   -90
            TabIndex        =   61
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   3
         Left            =   -67155
         Top             =   3975
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   3
            Left            =   -90
            TabIndex        =   62
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   7
         Left            =   -67155
         Top             =   5175
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   7
            Left            =   -90
            TabIndex        =   68
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   8
         Left            =   -67155
         Top             =   5535
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   8
            Left            =   -90
            TabIndex        =   72
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   4
         Left            =   -67155
         Top             =   4815
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   4
            Left            =   -90
            TabIndex        =   73
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   14
         Left            =   -67155
         Top             =   2265
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   14
            Left            =   -90
            TabIndex        =   83
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   15
         Left            =   -67155
         Top             =   2625
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   15
            Left            =   -90
            TabIndex        =   84
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   16
         Left            =   -67155
         Top             =   1245
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   16
            Left            =   -90
            TabIndex        =   85
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   17
         Left            =   -67155
         Top             =   885
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   17
            Left            =   -90
            TabIndex        =   86
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   18
         Left            =   -67155
         Top             =   3945
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   18
            Left            =   -90
            TabIndex        =   87
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   19
         Left            =   -67155
         Top             =   4305
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   19
            Left            =   -90
            TabIndex        =   88
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   23
         Left            =   -67155
         Top             =   2985
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   23
            Left            =   -90
            TabIndex        =   100
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   20
         Left            =   -67155
         Top             =   4665
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   20
            Left            =   -90
            TabIndex        =   103
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   21
         Left            =   -67155
         Top             =   5025
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   21
            Left            =   -90
            TabIndex        =   104
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   22
         Left            =   7845
         Top             =   1245
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   22
            Left            =   -90
            TabIndex        =   107
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   24
         Left            =   7845
         Top             =   885
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   24
            Left            =   -90
            TabIndex        =   108
            Top             =   -105
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(i.e. promotion, salary increase, and or trainings, and others) "
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   55
         Left            =   1125
         TabIndex        =   115
         Top             =   4440
         Width           =   5085
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(i.e. insubordination, violations of company rules, tardiness, absentism, and others)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   54
         Left            =   1125
         TabIndex        =   114
         Top             =   3375
         Width           =   6885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(i.e. improvements of system or procedures, continuous achievement  of quota or target, and others)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   51
         Left            =   1125
         TabIndex        =   113
         Top             =   2295
         Width           =   8370
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Hair well kept and grooms properly"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   49
         Left            =   585
         TabIndex        =   112
         Top             =   1320
         Width           =   3105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIII. Personal Appearance/Grooming"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   270
         TabIndex        =   111
         Top             =   540
         Width           =   4035
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rating: 10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   4410
         TabIndex        =   110
         Tag             =   "eb0;et0"
         Top             =   510
         Width           =   1710
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   4
         Left            =   4350
         Top             =   450
         Width           =   1830
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   4
         Left            =   4380
         Top             =   480
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a. Wears appropriate company uniform"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   26
         Left            =   585
         TabIndex        =   109
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "d. Demonstrates appropriate problem solving skills"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   25
         Left            =   -74415
         TabIndex        =   106
         Top             =   5100
         Width           =   4215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c. Works well in group problem - solving situations"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   24
         Left            =   -74415
         TabIndex        =   105
         Top             =   4740
         Width           =   4140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Develops/finds solution"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   23
         Left            =   -74415
         TabIndex        =   102
         Top             =   4380
         Width           =   2115
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c. Works in an organized manner"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   47
         Left            =   -74415
         TabIndex        =   101
         Top             =   3060
         Width           =   2730
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VI. Planning and Organization"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   -74730
         TabIndex        =   99
         Top             =   1890
         Width           =   3240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a. Prioritizes and plans work activities"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   52
         Left            =   -74415
         TabIndex        =   98
         Top             =   2340
         Width           =   3135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Uses time effectively and efficiently"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   50
         Left            =   -74415
         TabIndex        =   97
         Top             =   2700
         Width           =   3150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Includes appropriate people in decision - making process"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   48
         Left            =   -74415
         TabIndex        =   96
         Top             =   1320
         Width           =   4965
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V. Judgement"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   46
         Left            =   -74730
         TabIndex        =   95
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rating: 10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   -71325
         TabIndex        =   94
         Tag             =   "eb0;et0"
         Top             =   510
         Width           =   1710
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   9
         Left            =   -71385
         Top             =   450
         Width           =   1830
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   9
         Left            =   -71355
         Top             =   480
         Width           =   1770
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rating: 15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   -71325
         TabIndex        =   93
         Tag             =   "eb0;et0"
         Top             =   1860
         Width           =   1710
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   8
         Left            =   -71385
         Top             =   1800
         Width           =   1830
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   8
         Left            =   -71355
         Top             =   1830
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a. Supports and explains reasoning for decision"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   30
         Left            =   -74415
         TabIndex        =   92
         Top             =   960
         Width           =   3960
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   7
         Left            =   -71355
         Top             =   3510
         Width           =   1770
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   7
         Left            =   -71385
         Top             =   3480
         Width           =   1830
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rating: 20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   -71325
         TabIndex        =   91
         Tag             =   "eb0;et0"
         Top             =   3540
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a. Gathers and analyzes information skillfully"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   29
         Left            =   -74415
         TabIndex        =   90
         Top             =   4020
         Width           =   3720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VII. Problem Solving"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   -74730
         TabIndex        =   89
         Top             =   3570
         Width           =   2205
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IX. Knowledge and Skills:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   270
         TabIndex        =   82
         Top             =   1755
         Width           =   2745
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a.  Employee""s Most Notable Accomplishments: "
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   32
         Left            =   585
         TabIndex        =   81
         Top             =   2100
         Width           =   3960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Employee""s Policy Violations or Negative Comments"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   33
         Left            =   585
         TabIndex        =   80
         Top             =   3180
         Width           =   4515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c. Evaluator's Recommendations"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   34
         Left            =   585
         TabIndex        =   79
         Top             =   4245
         Width           =   2730
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c. Uses resources (manpower, equipment, technology) effectively"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   43
         Left            =   -74415
         TabIndex        =   75
         Top             =   5610
         Width           =   5460
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a. Competent in required job skills and knowledge"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   44
         Left            =   -74415
         TabIndex        =   74
         Top             =   4897
         Width           =   4125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IV. Job Knowledge"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   -74730
         TabIndex        =   71
         Top             =   4440
         Width           =   2040
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Exhibits ability tom learn and apply new skills"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   42
         Left            =   -74415
         TabIndex        =   70
         Top             =   5250
         Width           =   3915
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rating: 15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   -71910
         TabIndex        =   69
         Tag             =   "eb0;et0"
         Top             =   4410
         Width           =   1710
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   6
         Left            =   -71970
         Top             =   4350
         Width           =   1830
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   6
         Left            =   -71940
         Top             =   4380
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "III. Dependability"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   -74730
         TabIndex        =   67
         Top             =   3210
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a. Follows instructions, responds to management direction;"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   -74415
         TabIndex        =   66
         Top             =   3585
         Width           =   4920
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "takes responsibility for own actions"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   -74040
         TabIndex        =   65
         Top             =   3795
         Width           =   2940
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Keeps commitments and meets attendance and punctuality guidelines"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   -74415
         TabIndex        =   64
         Top             =   4050
         Width           =   6045
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rating: 10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   -71910
         TabIndex        =   63
         Tag             =   "eb0;et0"
         Top             =   3180
         Width           =   1710
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   5
         Left            =   -71970
         Top             =   3120
         Width           =   1830
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   5
         Left            =   -71940
         Top             =   3150
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a. Uses appropriate communication tools to convey ideas"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   -74415
         TabIndex        =   60
         Top             =   885
         Width           =   4770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Highly Effective (76-85)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   40
         Left            =   -74265
         TabIndex        =   58
         Top             =   4740
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Needs Improvement (31-40)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   39
         Left            =   -74265
         TabIndex        =   57
         Top             =   5250
         Width           =   2310
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not Effective (20-30)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   38
         Left            =   -74265
         TabIndex        =   56
         Top             =   5505
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Performing (60-75)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   37
         Left            =   -74265
         TabIndex        =   55
         Top             =   4995
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exceptional (86-100)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   36
         Left            =   -74265
         TabIndex        =   54
         Top             =   4470
         Width           =   1725
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rating Equivalent"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   35
         Left            =   -74535
         TabIndex        =   53
         Top             =   4155
         Width           =   1710
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   3
         Left            =   -71940
         Top             =   1920
         Width           =   1770
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   3
         Left            =   -71970
         Top             =   1890
         Width           =   1830
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rating: 10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   -71910
         TabIndex        =   52
         Tag             =   "eb0;et0"
         Top             =   1950
         Width           =   1710
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   2
         Left            =   -71940
         Top             =   480
         Width           =   1770
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   2
         Left            =   -71970
         Top             =   450
         Width           =   1830
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rating: 10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   -71910
         TabIndex        =   51
         Tag             =   "eb0;et0"
         Top             =   510
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I. Communication"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74730
         TabIndex        =   45
         Top             =   540
         Width           =   1905
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(i.e. written/verbal communication)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   -74040
         TabIndex        =   44
         Top             =   1095
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Keeps subordinates/associates adequately informed"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   -74415
         TabIndex        =   43
         Top             =   1350
         Width           =   4575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(i.e. inter office communication; mktg bulletin; policies and procedures)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   -74040
         TabIndex        =   42
         Top             =   1605
         Width           =   5925
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Performance does not meet position requirements; immediate attention to improvement is required."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   22
         Left            =   -73665
         TabIndex        =   40
         Top             =   3765
         Width           =   8295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Performance meets some position requirements, objectives and expectations"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   21
         Left            =   -73665
         TabIndex        =   39
         Top             =   3120
         Width           =   6450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Performance meets position requirements"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   20
         Left            =   -73665
         TabIndex        =   38
         Top             =   2460
         Width           =   3525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Performance exceeds normal job requirements"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   19
         Left            =   -73665
         TabIndex        =   37
         Top             =   1800
         Width           =   3945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Performance is Superior on a consistent and systematic basis"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   18
         Left            =   -73665
         TabIndex        =   36
         Top             =   1140
         Width           =   5145
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Summary of Performance"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   17
         Left            =   -74535
         TabIndex        =   35
         Top             =   465
         Width           =   2475
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O = Outstanding (9-10 Stars)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   16
         Left            =   -74220
         TabIndex        =   34
         Top             =   810
         Width           =   2355
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E = Exceeds Expectations (7-8 Stars)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   15
         Left            =   -74220
         TabIndex        =   33
         Top             =   1470
         Width           =   3030
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M = Meets Expectations (5-6 Stars)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   14
         Left            =   -74220
         TabIndex        =   32
         Top             =   2130
         Width           =   2850
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NI = Needs Improvement (3-4 Stars)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   13
         Left            =   -74220
         TabIndex        =   31
         Top             =   2790
         Width           =   2925
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "U = Unsatisfactory (1-2 Stars)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   12
         Left            =   -74220
         TabIndex        =   30
         Top             =   3450
         Width           =   2430
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Works cooperatively in group situations and works actively to resolve conflicts"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   8
         Left            =   -74415
         TabIndex        =   29
         Top             =   2820
         Width           =   6705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(i.e. offers assistance and support to co - workers; exhibit consideration)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   9
         Left            =   -74040
         TabIndex        =   28
         Top             =   2565
         Width           =   6030
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a. Establish and maintain effective working relationship w/ associates"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   10
         Left            =   -74415
         TabIndex        =   27
         Top             =   2355
         Width           =   5760
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "II. Cooperation"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -74730
         TabIndex        =   26
         Top             =   1980
         Width           =   1635
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   10650
      TabIndex        =   12
      Top             =   9210
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmpAppraisalMgr.frx":0070
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   9090
      TabIndex        =   10
      Top             =   9210
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Save"
      AccessKey       =   "S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmpAppraisalMgr.frx":07EA
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   9870
      TabIndex        =   11
      Top             =   9210
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "Searc&h"
      AccessKey       =   "h"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmpAppraisalMgr.frx":0F64
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   570
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   1005
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   8160
         MaxLength       =   50
         TabIndex        =   20
         Top             =   90
         Width           =   2000
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   1
         Top             =   90
         Width           =   4815
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Employed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   6675
         TabIndex        =   21
         Top             =   150
         Width           =   1335
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   105
         TabIndex        =   0
         Top             =   150
         Width           =   1440
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   7830
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1155
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   13811
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   8175
         MaxLength       =   50
         TabIndex        =   18
         Text            =   "MIS Associate"
         Top             =   1095
         Width           =   2970
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   8175
         MaxLength       =   50
         TabIndex        =   16
         Text            =   "Feb. 26, 2013"
         Top             =   705
         Width           =   2000
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   86
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "GMC Dagupan - Honda"
         Top             =   1095
         Width           =   4815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   89
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "Dela Cruz, Juan Santos"
         Top             =   705
         Width           =   4815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "M00113-000001"
         Top             =   105
         Width           =   2000
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Total Rating: 100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   8685
         TabIndex        =   23
         Tag             =   "eb0;et0"
         Top             =   135
         Width           =   2400
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   1
         Left            =   8625
         Top             =   75
         Width           =   2520
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   1
         Left            =   8655
         Top             =   105
         Width           =   2460
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   6705
         TabIndex        =   19
         Top             =   1155
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Employed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   6705
         TabIndex        =   17
         Top             =   765
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   225
         X2              =   10980
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   225
         X2              =   10980
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   105
         TabIndex        =   7
         Top             =   1155
         Width           =   615
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   0
         Left            =   3945
         Top             =   105
         Width           =   2460
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   0
         Left            =   3915
         Top             =   75
         Width           =   2520
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Regular"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3975
         TabIndex        =   4
         Tag             =   "eb0;et0"
         Top             =   135
         Width           =   2400
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1725
         Tag             =   "et0;ht2"
         Top             =   150
         Width           =   1995
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   765
         Width           =   1440
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   105
         TabIndex        =   2
         Top             =   165
         Width           =   1110
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   9870
      TabIndex        =   14
      Top             =   9210
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Browse"
      AccessKey       =   "B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmpAppraisalMgr.frx":16DE
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   9090
      TabIndex        =   9
      Top             =   9210
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&New"
      AccessKey       =   "N"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmpAppraisalMgr.frx":1E58
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   10650
      TabIndex        =   15
      Top             =   9210
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Close"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmpAppraisalMgr.frx":25D2
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   9090
      TabIndex        =   13
      Top             =   9210
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Update"
      AccessKey       =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmpAppraisalMgr.frx":2D4C
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmEmpAppraisalMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Employee Appraisal Form
'        - Used for Employee Evaluation
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' MAC [02-21-12 12:00 PM]
'     Start creating this form.
'        -this form requires xxxControl.ctl for the Star Rating
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº

Option Explicit

Private Const pxeMODULENAME = "frmEmpAppraisal"
Private oDriver As clsFormDriver
Private oSkin As clsFormSkin

Private Sub Form_Load()
10   Dim lsOldProc As String

20   lsOldProc = "Form_Load"
30   On Error GoTo errProc

40   CenterChildForm mdiMain, Me

50   Set oDriver = New clsFormDriver
60   Set oDriver.AppDriver = oApp
70   Set oDriver.MainForm = Me

80   Set oSkin = New clsFormSkin
90   Set oSkin.AppDriver = oApp
100   Set oSkin.Form = Me
110   oSkin.ApplySkin xeFormMaintenance
   
120   Call InitForm

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitForm()
   Dim lsOldProc As String
   Dim loTxt As TextBox
   Dim loStar As StarRating
   
   lsOldProc = "initForm"
   On Error GoTo errProc
   
'   For Each loTxt In txtField
'      loTxt = ""
'   Next
   
'   For Each loTxt In txtSearch
'      loTxt = ""
'   Next
   
   For Each loStar In strRate
      loStar.FixedColorHover = Yellow
      loStar.FixedColorRated = Yellow
   Next
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
10       With oApp
20          .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
30          If bEnd Then
40             .xShowError
50             End
60          Else
70             With Err
80                .Raise .Number, .Source, .Description
90             End With
100         End If
110      End With
End Sub

