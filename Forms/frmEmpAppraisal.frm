VERSION 5.00
Object = "{0919B38C-FEB7-4581-AC15-BCF315A77232}#1.0#0"; "xxxControl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmEmpAppraisal 
   BorderStyle     =   0  'None
   Caption         =   "Employee Appraisal - Rank and File"
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      TabIndex        =   23
      Tag             =   "wt0;fb0"
      Top             =   2760
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10769
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   -2147483634
      TabCaption(0)   =   "Legend"
      TabPicture(0)   =   "frmEmpAppraisal.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "xrFrame3(9)"
      Tab(0).Control(1)=   "xrFrame3(10)"
      Tab(0).Control(2)=   "xrFrame3(11)"
      Tab(0).Control(3)=   "xrFrame3(12)"
      Tab(0).Control(4)=   "xrFrame3(13)"
      Tab(0).Control(5)=   "Label3(17)"
      Tab(0).Control(6)=   "Label3(35)"
      Tab(0).Control(7)=   "Label3(40)"
      Tab(0).Control(8)=   "Label3(39)"
      Tab(0).Control(9)=   "Label3(38)"
      Tab(0).Control(10)=   "Label3(37)"
      Tab(0).Control(11)=   "Label3(36)"
      Tab(0).Control(12)=   "Label3(22)"
      Tab(0).Control(13)=   "Label3(21)"
      Tab(0).Control(14)=   "Label3(20)"
      Tab(0).Control(15)=   "Label3(19)"
      Tab(0).Control(16)=   "Label3(18)"
      Tab(0).Control(17)=   "Label3(16)"
      Tab(0).Control(18)=   "Label3(15)"
      Tab(0).Control(19)=   "Label3(14)"
      Tab(0).Control(20)=   "Label3(13)"
      Tab(0).Control(21)=   "Label3(12)"
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "I."
      TabPicture(1)   =   "frmEmpAppraisal.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3(5)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3(4)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label3(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label3(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Shape3(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Shape4(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "xrFrame3(4)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "xrFrame3(3)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "xrFrame3(2)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "xrFrame3(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "xrFrame3(0)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "II."
      TabPicture(2)   =   "frmEmpAppraisal.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "xrFrame3(5)"
      Tab(2).Control(1)=   "xrFrame3(6)"
      Tab(2).Control(2)=   "xrFrame3(7)"
      Tab(2).Control(3)=   "xrFrame3(8)"
      Tab(2).Control(4)=   "Shape4(3)"
      Tab(2).Control(5)=   "Shape3(3)"
      Tab(2).Control(6)=   "Label5"
      Tab(2).Control(7)=   "Label3(7)"
      Tab(2).Control(8)=   "Label3(8)"
      Tab(2).Control(9)=   "Label3(9)"
      Tab(2).Control(10)=   "Label3(10)"
      Tab(2).Control(11)=   "Label3(11)"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "III."
      TabPicture(3)   =   "frmEmpAppraisal.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "xrFrame3(14)"
      Tab(3).Control(1)=   "xrFrame3(15)"
      Tab(3).Control(2)=   "xrFrame3(16)"
      Tab(3).Control(3)=   "xrFrame3(17)"
      Tab(3).Control(4)=   "Label3(23)"
      Tab(3).Control(5)=   "Label3(24)"
      Tab(3).Control(6)=   "Label3(25)"
      Tab(3).Control(7)=   "Label3(26)"
      Tab(3).Control(8)=   "Label3(27)"
      Tab(3).Control(9)=   "Label3(28)"
      Tab(3).Control(10)=   "Label3(29)"
      Tab(3).Control(11)=   "Label3(30)"
      Tab(3).Control(12)=   "Label6"
      Tab(3).Control(13)=   "Shape3(4)"
      Tab(3).Control(14)=   "Shape4(4)"
      Tab(3).ControlCount=   15
      TabCaption(4)   =   "IV."
      TabPicture(4)   =   "frmEmpAppraisal.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text2"
      Tab(4).Control(1)=   "Text3"
      Tab(4).Control(2)=   "Text1"
      Tab(4).Control(3)=   "Label3(31)"
      Tab(4).Control(4)=   "Label3(32)"
      Tab(4).Control(5)=   "Label3(33)"
      Tab(4).Control(6)=   "Label3(34)"
      Tab(4).ControlCount=   7
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   -73860
         TabIndex        =   58
         Top             =   2445
         Width           =   9300
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   -73860
         TabIndex        =   57
         Top             =   3375
         Width           =   9300
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   -73860
         TabIndex        =   56
         Top             =   1440
         Width           =   9300
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   0
         Left            =   825
         Top             =   1560
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   0
            Left            =   -90
            TabIndex        =   34
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   1
         Left            =   825
         Top             =   2415
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   1
            Left            =   -90
            TabIndex        =   35
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   2
         Left            =   825
         Top             =   3270
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   2
            Left            =   -90
            TabIndex        =   36
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   3
         Left            =   825
         Top             =   4140
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   3
            Left            =   -90
            TabIndex        =   37
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   4
         Left            =   825
         Top             =   4995
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   4
            Left            =   -90
            TabIndex        =   38
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   9
         Left            =   -71430
         Top             =   990
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   9
            Left            =   -90
            TabIndex        =   45
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
         Left            =   -71415
         Top             =   1650
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   10
            Left            =   -90
            TabIndex        =   46
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
         Left            =   -71415
         Top             =   2310
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   11
            Left            =   -90
            TabIndex        =   47
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
         Left            =   -71415
         Top             =   2970
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   12
            Left            =   -90
            TabIndex        =   48
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
         Left            =   -71400
         Top             =   3630
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   13
            Left            =   -90
            TabIndex        =   49
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
         Index           =   5
         Left            =   -74160
         Top             =   1560
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   5
            Left            =   -90
            TabIndex        =   63
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   6
         Left            =   -74160
         Top             =   2415
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   6
            Left            =   -90
            TabIndex        =   64
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   7
         Left            =   -74160
         Top             =   3285
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   7
            Left            =   -90
            TabIndex        =   65
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   8
         Left            =   -74160
         Top             =   4140
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   8
            Left            =   -90
            TabIndex        =   66
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   14
         Left            =   -74220
         Top             =   1695
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   14
            Left            =   -90
            TabIndex        =   73
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   15
         Left            =   -74220
         Top             =   2550
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   15
            Left            =   -90
            TabIndex        =   74
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   16
         Left            =   -74220
         Top             =   3615
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   16
            Left            =   -90
            TabIndex        =   75
            Top             =   -90
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   345
         Index           =   17
         Left            =   -74220
         Top             =   4680
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         BackColor       =   -2147483644
         ClipControls    =   0   'False
         Begin xxxControl.StarRating strRate 
            Height          =   495
            Index           =   17
            Left            =   -90
            TabIndex        =   76
            Top             =   -105
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   873
            StarCount       =   10
            FixedColorHover =   0
            FixedColorRated =   1
         End
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
         Left            =   -74865
         TabIndex        =   87
         Top             =   720
         Width           =   2475
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
         Left            =   -67275
         TabIndex        =   86
         Top             =   720
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "III. Interpersonal Relations:"
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
         Index           =   23
         Left            =   -74730
         TabIndex        =   85
         Top             =   720
         Width           =   3000
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a. Employee cooperate w/ others on team assignments, actively participating"
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
         TabIndex        =   84
         Top             =   1155
         Width           =   6390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " and performing assigned functions consistent with team goals."
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
         Left            =   -74265
         TabIndex        =   83
         Top             =   1365
         Width           =   5265
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Employee assist other staff as needed w/out complaining"
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
         Left            =   -74415
         TabIndex        =   82
         Top             =   2220
         Width           =   4920
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c. Employee promote harmony among co - workers, keeping personal problems"
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
         Index           =   27
         Left            =   -74415
         TabIndex        =   81
         Top             =   3075
         Width           =   6585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "from affecting the performance of work"
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
         Index           =   28
         Left            =   -74220
         TabIndex        =   80
         Top             =   3285
         Width           =   3300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "d. Adheres to the established work hours for arrival and departure from work,"
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
         TabIndex        =   79
         Top             =   4140
         Width           =   6450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lunch and break periods"
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
         Left            =   -74220
         TabIndex        =   78
         Top             =   4350
         Width           =   2025
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rating: 40"
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
         Left            =   -71535
         TabIndex        =   77
         Tag             =   "eb0;et0"
         Top             =   690
         Width           =   1710
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   4
         Left            =   -71595
         Top             =   630
         Width           =   1830
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   4
         Left            =   -71565
         Top             =   660
         Width           =   1770
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   3
         Left            =   -72705
         Top             =   660
         Width           =   1770
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   3
         Left            =   -72735
         Top             =   630
         Width           =   1830
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rating: 40"
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
         Left            =   -72675
         TabIndex        =   72
         Tag             =   "eb0;et0"
         Top             =   690
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "d. Do not display a demeaning attitude or behavior in the workplace towards another employee/customer."
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
         TabIndex        =   71
         Top             =   3810
         Width           =   8775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c. Have Positive Disposition towards work."
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
         TabIndex        =   70
         Top             =   2940
         Width           =   3480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Serves Customers/Clients in a Friendly Manner."
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
         Left            =   -74415
         TabIndex        =   69
         Top             =   2085
         Width           =   4110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a. Follows Company Policies and Regulations."
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
         TabIndex        =   68
         Top             =   1230
         Width           =   3780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "II. Work Attitude:"
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
         TabIndex        =   67
         Top             =   720
         Width           =   1860
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IV. Knowledge and Skills:"
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
         Left            =   -74730
         TabIndex        =   62
         Top             =   720
         Width           =   2760
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a.  Positive Comments towards the Employee "
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
         Left            =   -74415
         TabIndex        =   61
         Top             =   1110
         Width           =   3750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b. Negative Comments towards the Employee:"
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
         Left            =   -74415
         TabIndex        =   60
         Top             =   2115
         Width           =   3795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c. Recommendation:"
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
         Left            =   -74415
         TabIndex        =   59
         Top             =   3045
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Highly Effective (91-116)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   40
         Left            =   -66975
         TabIndex        =   55
         Top             =   1395
         Width           =   2040
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Needs Improvement (39-64)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   39
         Left            =   -66975
         TabIndex        =   54
         Top             =   2055
         Width           =   2310
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not Effective (13-38)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   38
         Left            =   -66975
         TabIndex        =   53
         Top             =   2385
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Performing (65-90)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   37
         Left            =   -66975
         TabIndex        =   52
         Top             =   1725
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exceptional (117-130)"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   36
         Left            =   -66975
         TabIndex        =   51
         Top             =   1065
         Width           =   1830
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   2
         Left            =   3060
         Top             =   660
         Width           =   1770
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   2
         Left            =   3030
         Top             =   630
         Width           =   1830
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Rating: 50"
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
         Left            =   3090
         TabIndex        =   50
         Tag             =   "eb0;et0"
         Top             =   690
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I. Knowledge and Skills:"
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
         Left            =   270
         TabIndex        =   44
         Top             =   720
         Width           =   2610
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a.  Able to demonstrate sufficient knowledge in completing work assignments."
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
         Left            =   585
         TabIndex        =   43
         Top             =   1230
         Width           =   6465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b.  Able to do the job with minimal supervision."
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
         Left            =   585
         TabIndex        =   42
         Top             =   2085
         Width           =   3825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c. Does the required assigned job effectively and efficiently."
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
         Left            =   585
         TabIndex        =   41
         Top             =   2940
         Width           =   4980
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "d. Finishes assigned task/work on the required time frame."
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
         Left            =   585
         TabIndex        =   40
         Top             =   3795
         Width           =   4845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e. Handles customer concerns efficiently."
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
         Left            =   585
         TabIndex        =   39
         Top             =   4665
         Width           =   3465
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
         Left            =   -73995
         TabIndex        =   33
         Top             =   4020
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
         Left            =   -73995
         TabIndex        =   32
         Top             =   3375
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
         Left            =   -73995
         TabIndex        =   31
         Top             =   2715
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
         Left            =   -73995
         TabIndex        =   30
         Top             =   2055
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
         Left            =   -73995
         TabIndex        =   29
         Top             =   1395
         Width           =   5145
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
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   16
         Left            =   -74550
         TabIndex        =   28
         Top             =   1065
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
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   15
         Left            =   -74550
         TabIndex        =   27
         Top             =   1725
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
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   14
         Left            =   -74550
         TabIndex        =   26
         Top             =   2385
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
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   13
         Left            =   -74550
         TabIndex        =   25
         Top             =   3045
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
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   12
         Left            =   -74550
         TabIndex        =   24
         Top             =   3705
         Width           =   2430
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   10650
      TabIndex        =   19
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
      Picture         =   "frmEmpAppraisal.frx":008C
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   9090
      TabIndex        =   17
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
      Picture         =   "frmEmpAppraisal.frx":0806
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   9870
      TabIndex        =   18
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
      Picture         =   "frmEmpAppraisal.frx":0F80
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         Index           =   83
         Left            =   8175
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
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
         Index           =   82
         Left            =   8175
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
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
         Index           =   81
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
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
         Index           =   80
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Dela Cruz, Juan Santos"
         Top             =   705
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
         Index           =   0
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "M00113-000001"
         Top             =   105
         Width           =   2000
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Total Rating: 110"
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
         TabIndex        =   15
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
         TabIndex        =   12
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
         TabIndex        =   10
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
         TabIndex        =   8
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
         Caption         =   "Agency"
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
         TabIndex        =   14
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
         TabIndex        =   6
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
         TabIndex        =   4
         Top             =   165
         Width           =   1110
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   9870
      TabIndex        =   21
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
      Picture         =   "frmEmpAppraisal.frx":16FA
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   9090
      TabIndex        =   16
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
      Picture         =   "frmEmpAppraisal.frx":1E74
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   10650
      TabIndex        =   22
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
      Picture         =   "frmEmpAppraisal.frx":25EE
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   9090
      TabIndex        =   20
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
      Picture         =   "frmEmpAppraisal.frx":2D68
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmEmpAppraisal"
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

Private Const pxeModuleName = "frmEmpAppraisal"
Private oDriver As clsFormDriver
Private oSkin As clsFormSkin
Private WithEvents oTrans As clsEmpAppraisal
Attribute oTrans.VB_VarHelpID = -1

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
   
   Set oTrans = New clsEmpAppraisal
   Set oTrans.AppDriver = oApp
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
   
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
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
   For Each loTxt In txtSearch
      loTxt = ""
   Next
   
   For Each loStar In strRate
      loStar.FixedColorHover = Yellow
      loStar.FixedColorRated = Yellow
   Next
   
   Label1.Caption = "Total Rating : -"
   Label2.Caption = "UNKNOWN"
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
10       With oApp
20          .xLogError Err.Number, Err.Description, pxeModuleName, lsProcName, Erl
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

Private Sub oTrans_MasterRetrieve(ByVal Index As Variant)
   With oTrans
      Select Case Index
         Case 80, 81, 83
            txtField(Index) = IFNull(.Master(Index))
         Case 82
            txtField(Index) = strLongDate(.Master(Index))
         Case 84
            Select Case LCase(.Master(Index))
               Case "a"
                  Label2.Caption = "Agency"
               Case "p"
                  Label2.Caption = "Probationary"
               Case "r"
                  Label2.Caption = "Regular"
               Case Else
                  Label2.Caption = "UNKNOWN"
            End Select
      End Select
   End With
End Sub

Private Sub strRate_RatingSelected(Index As Integer, ByVal Rating As Integer)
   strRate(Index).StarCount = Rating
   
   Select Case Index
      Case 0
   End Select
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case Index
      Case 0
      Case 1
         If KeyCode = vbKeyF3 Then
            If oTrans.SearchMaster(80, txtSearch(Index)) Then
               txtSearch(Index) = ""
            Else
               txtSearch(Index).SetFocus
            End If
         ElseIf KeyCode = vbKeyReturn Then
            oTrans.Master(80) = txtSearch(Index)
         End If
   End Select
End Sub
