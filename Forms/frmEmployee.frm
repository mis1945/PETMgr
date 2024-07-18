VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmEmployee 
   BorderStyle     =   0  'None
   Caption         =   "Employee Information"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Visible         =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   10605
      TabIndex        =   139
      Top             =   8925
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
      Picture         =   "frmEmployee.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   9045
      TabIndex        =   137
      Top             =   8925
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
      Picture         =   "frmEmployee.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   9825
      TabIndex        =   138
      Top             =   8925
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
      Picture         =   "frmEmployee.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   510
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   900
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   1
         Top             =   75
         Width           =   2000
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   6345
         MaxLength       =   50
         TabIndex        =   3
         Top             =   75
         Width           =   4815
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Control No"
         Height          =   195
         Index           =   24
         Left            =   90
         TabIndex        =   0
         Top             =   135
         Width           =   750
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         Height          =   195
         Index           =   25
         Left            =   4605
         TabIndex        =   2
         Top             =   135
         Width           =   1155
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1665
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1080
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2937
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   35
         Left            =   7500
         TabIndex        =   156
         Text            =   "8106"
         Top             =   1215
         Width           =   2000
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   155
         Text            =   "M00111-000021"
         Top             =   105
         Width           =   2000
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   89
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   154
         Text            =   "Sayson, Josh Ramsej C. Sayson"
         Top             =   555
         Width           =   4815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   86
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   153
         Top             =   885
         Width           =   4815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   13
         Left            =   7500
         TabIndex        =   152
         Text            =   "8106"
         Top             =   885
         Width           =   2000
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   12
         Left            =   7500
         TabIndex        =   151
         Text            =   "001-00283-06"
         Top             =   555
         Width           =   2000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   195
         Index           =   61
         Left            =   6690
         TabIndex        =   157
         Top             =   1275
         Width           =   690
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1695
         Tag             =   "et0;ht2"
         Top             =   195
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl. No"
         Height          =   195
         Index           =   48
         Left            =   6690
         TabIndex        =   8
         Top             =   615
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         Height          =   195
         Index           =   47
         Left            =   6690
         TabIndex        =   9
         Top             =   945
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   7
         Top             =   945
         Width           =   510
      End
      Begin VB.Shape Shape4 
         Height          =   315
         Index           =   0
         Left            =   7005
         Top             =   120
         Width           =   2460
      End
      Begin VB.Shape Shape3 
         Height          =   375
         Index           =   0
         Left            =   6975
         Top             =   90
         Width           =   2520
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "INACTIVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7035
         TabIndex        =   5
         Tag             =   "eb0;et0"
         Top             =   150
         Width           =   2400
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1380
         Left            =   9705
         Top             =   75
         Width           =   1380
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   6
         Top             =   615
         Width           =   1155
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         Height          =   195
         Index           =   21
         Left            =   105
         TabIndex        =   4
         Top             =   165
         Width           =   900
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   9825
      TabIndex        =   141
      Top             =   8925
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
      Picture         =   "frmEmployee.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   8265
      TabIndex        =   136
      Top             =   8925
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
      Picture         =   "frmEmployee.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   10605
      TabIndex        =   142
      Top             =   8925
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
      Picture         =   "frmEmployee.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   9045
      TabIndex        =   140
      Top             =   8925
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
      Picture         =   "frmEmployee.frx":2CDC
      PicturePos      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   105
      TabIndex        =   10
      Tag             =   "wt0;fb0"
      Top             =   2760
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10398
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Personal Info"
      TabPicture(0)   =   "frmEmployee.frx":3456
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "xrFrame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Contact Info"
      TabPicture(1)   =   "frmEmployee.frx":3472
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "xrFrame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Employment Information"
      TabPicture(2)   =   "frmEmployee.frx":348E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "xrFrame5"
      Tab(2).Control(1)=   "xrFrame6"
      Tab(2).ControlCount=   2
      Begin xrControl.xrFrame xrFrame6 
         Height          =   420
         Left            =   -74955
         Top             =   5385
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   741
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Account#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6975
            TabIndex        =   143
            Tag             =   "wt0;fb0"
            Top             =   45
            Width           =   4140
         End
      End
      Begin xrControl.xrFrame xrFrame5 
         Height          =   5535
         Left            =   -74970
         Tag             =   "wt0;fb0"
         Top             =   315
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   9763
         BackColor       =   12632256
         ClipControls    =   0   'False
         BorderStyle     =   4
         Begin VB.CheckBox chkField 
            Caption         =   "Selfie Log"
            Height          =   315
            Index           =   41
            Left            =   9675
            TabIndex        =   158
            Tag             =   "wt0;fb0"
            Top             =   4710
            Width           =   1005
         End
         Begin VB.ComboBox cmbSecType 
            Height          =   315
            ItemData        =   "frmEmployee.frx":34AA
            Left            =   7410
            List            =   "frmEmployee.frx":34B4
            Style           =   2  'Dropdown List
            TabIndex        =   149
            Top             =   4710
            Width           =   1770
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   17
            Left            =   6885
            Locked          =   -1  'True
            TabIndex        =   111
            TabStop         =   0   'False
            Text            =   "January 12, 2011"
            Top             =   1575
            Width           =   3000
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Credit Investigator"
            Height          =   315
            Index           =   34
            Left            =   4080
            TabIndex        =   147
            Tag             =   "wt0;fb0"
            Top             =   4710
            Width           =   1575
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Mechanic"
            Height          =   315
            Index           =   33
            Left            =   2700
            TabIndex        =   146
            Tag             =   "wt0;fb0"
            Top             =   4710
            Width           =   1005
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Manager"
            Height          =   315
            Index           =   32
            Left            =   1395
            TabIndex        =   145
            Tag             =   "wt0;fb0"
            Top             =   4710
            Width           =   930
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Collector"
            Height          =   315
            Index           =   31
            Left            =   105
            TabIndex        =   144
            Tag             =   "wt0;fb0"
            Top             =   4710
            Width           =   915
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Company Subsidized?"
            Height          =   315
            Index           =   28
            Left            =   6885
            TabIndex        =   127
            Tag             =   "wt0;fb0"
            Top             =   3375
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   91
            Left            =   6885
            TabIndex        =   135
            Top             =   4230
            Width           =   3000
         End
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   9
            ItemData        =   "frmEmployee.frx":34CF
            Left            =   1620
            List            =   "frmEmployee.frx":34E5
            Style           =   2  'Dropdown List
            TabIndex        =   115
            Top             =   2400
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   26
            Left            =   6885
            TabIndex        =   133
            Top             =   3900
            Width           =   1920
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   90
            Left            =   6885
            TabIndex        =   125
            Text            =   "National Capital Region"
            Top             =   2745
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   80
            Left            =   1620
            TabIndex        =   92
            Text            =   "Senior Programmer 1"
            Top             =   90
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   81
            Left            =   1620
            TabIndex        =   94
            Text            =   "IT Department"
            Top             =   420
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   87
            Left            =   6885
            TabIndex        =   99
            Text            =   "GMC Dagupan - Honda"
            Top             =   90
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   88
            Left            =   6885
            TabIndex        =   101
            Text            =   "LGK Guanzon"
            Top             =   420
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   15
            Left            =   1620
            TabIndex        =   107
            Text            =   "January 12, 2011"
            Top             =   1575
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1620
            TabIndex        =   105
            Text            =   "December 31, 2012"
            Top             =   1245
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   16
            Left            =   6885
            TabIndex        =   109
            Text            =   "January 12, 2011"
            Top             =   1245
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   22
            Left            =   1605
            TabIndex        =   131
            Text            =   "0"
            Top             =   4230
            Width           =   1920
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   23
            Left            =   1605
            TabIndex        =   129
            Text            =   "0"
            Top             =   3870
            Width           =   1920
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   82
            Left            =   1620
            TabIndex        =   96
            Text            =   "IT Department"
            Top             =   750
            Width           =   2100
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   84
            Left            =   6885
            TabIndex        =   103
            Text            =   "National Capital Region"
            Top             =   750
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   83
            Left            =   1620
            TabIndex        =   117
            Top             =   2745
            Width           =   1920
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   85
            Left            =   6885
            TabIndex        =   123
            Text            =   "0"
            Top             =   2415
            Width           =   3000
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Deduct Gov't Contribution"
            Height          =   315
            Index           =   19
            Left            =   6885
            TabIndex        =   126
            Tag             =   "wt0;fb0"
            Top             =   3060
            Width           =   3000
         End
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   8
            ItemData        =   "frmEmployee.frx":3541
            Left            =   1605
            List            =   "frmEmployee.frx":354E
            Style           =   2  'Dropdown List
            TabIndex        =   113
            Top             =   2070
            Width           =   3000
         End
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   20
            ItemData        =   "frmEmployee.frx":3571
            Left            =   1620
            List            =   "frmEmployee.frx":357B
            Style           =   2  'Dropdown List
            TabIndex        =   119
            Top             =   3060
            Width           =   3000
         End
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   21
            ItemData        =   "frmEmployee.frx":3591
            Left            =   6885
            List            =   "frmEmployee.frx":359E
            Style           =   2  'Dropdown List
            TabIndex        =   121
            Top             =   2070
            Width           =   3000
         End
         Begin xrControl.xrButton cmdRestDay 
            Height          =   315
            Index           =   3
            Left            =   3765
            TabIndex        =   97
            Top             =   750
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            Caption         =   "&Shedule"
            AccessKey       =   "S"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   6060
            X2              =   6060
            Y1              =   4725
            Y2              =   5040
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   6075
            X2              =   6075
            Y1              =   4725
            Y2              =   5040
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SEC TYPE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   62
            Left            =   6465
            TabIndex        =   148
            Top             =   4770
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resigned/Terminated"
            Height          =   195
            Index           =   60
            Left            =   5040
            TabIndex        =   110
            Top             =   1635
            Width           =   1545
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   15
            X2              =   11175
            Y1              =   4635
            Y2              =   4635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
            Height          =   315
            Index           =   55
            Left            =   5070
            TabIndex        =   134
            Top             =   4230
            Width           =   375
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   30
            X2              =   11190
            Y1              =   3795
            Y2              =   3795
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   45
            X2              =   11205
            Y1              =   1980
            Y2              =   1980
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   30
            X2              =   11190
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Account#"
            Height          =   315
            Index           =   54
            Left            =   5070
            TabIndex        =   132
            Top             =   3900
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Region"
            Height          =   195
            Index           =   53
            Left            =   5040
            TabIndex        =   124
            Top             =   2805
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job Description"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   91
            Top             =   150
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   93
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Branch"
            Height          =   195
            Index           =   11
            Left            =   5040
            TabIndex        =   98
            Top             =   150
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg Branch"
            Height          =   195
            Index           =   13
            Left            =   5040
            TabIndex        =   100
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Started"
            Height          =   195
            Index           =   34
            Left            =   105
            TabIndex        =   106
            Top             =   1635
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Hired"
            Height          =   195
            Index           =   35
            Left            =   105
            TabIndex        =   104
            Top             =   1305
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Regular"
            Height          =   195
            Index           =   36
            Left            =   5040
            TabIndex        =   108
            Top             =   1305
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Basic Pay"
            Height          =   195
            Index           =   38
            Left            =   105
            TabIndex        =   130
            Top             =   4290
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg Salary"
            Height          =   195
            Index           =   39
            Left            =   105
            TabIndex        =   128
            Top             =   3960
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule/Shift"
            Height          =   195
            Index           =   37
            Left            =   105
            TabIndex        =   95
            Top             =   810
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Status"
            Height          =   195
            Index           =   40
            Left            =   105
            TabIndex        =   112
            Top             =   2130
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Level"
            Height          =   195
            Index           =   41
            Left            =   105
            TabIndex        =   114
            Top             =   2475
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Level"
            Height          =   195
            Index           =   42
            Left            =   105
            TabIndex        =   116
            Top             =   2805
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Region/Area"
            Height          =   195
            Index           =   43
            Left            =   5040
            TabIndex        =   102
            Top             =   810
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Type"
            Height          =   195
            Index           =   44
            Left            =   105
            TabIndex        =   118
            Top             =   3150
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Computation Rate"
            Height          =   195
            Index           =   45
            Left            =   5040
            TabIndex        =   120
            Top             =   2130
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exemption Status"
            Height          =   195
            Index           =   46
            Left            =   5040
            TabIndex        =   122
            Top             =   2475
            Width           =   1230
         End
      End
      Begin xrControl.xrFrame xrFrame4 
         Height          =   6210
         Left            =   -74985
         Tag             =   "wt0;fb0"
         Top             =   315
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   10954
         BackColor       =   12632256
         ClipControls    =   0   'False
         BorderStyle     =   4
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   14
            Left            =   6855
            TabIndex        =   77
            Text            =   "masayson@guanzongroup.com.ph"
            Top             =   1080
            Width           =   4275
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   85
            Left            =   1620
            MultiLine       =   -1  'True
            TabIndex        =   69
            Text            =   "frmEmployee.frx":35BA
            Top             =   1080
            Width           =   3885
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   80
            Left            =   1620
            TabIndex        =   67
            Text            =   "City of San Carlos"
            Top             =   750
            Width           =   3885
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   12
            Left            =   6855
            TabIndex        =   71
            Text            =   "(075) 522-1097"
            Top             =   420
            Width           =   2070
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   13
            Left            =   6855
            TabIndex        =   74
            Text            =   "+639163103800"
            Top             =   750
            Width           =   2070
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   9
            Left            =   1620
            TabIndex        =   63
            Text            =   "45"
            Top             =   90
            Width           =   975
         End
         Begin VB.TextBox txtContact 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1620
            TabIndex        =   80
            Top             =   1890
            Width           =   3885
         End
         Begin VB.TextBox txtContact 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1620
            TabIndex        =   86
            Top             =   2880
            Width           =   2070
         End
         Begin VB.TextBox txtContact 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1620
            TabIndex        =   84
            Top             =   2550
            Width           =   2070
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   10
            Left            =   1620
            MultiLine       =   -1  'True
            TabIndex        =   65
            Text            =   "frmEmployee.frx":35DD
            Top             =   420
            Width           =   3885
         End
         Begin VB.TextBox txtContact 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1620
            MultiLine       =   -1  'True
            TabIndex        =   82
            Top             =   2220
            Width           =   3885
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   21
            Left            =   8130
            TabIndex        =   90
            Top             =   2220
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   20
            Left            =   8130
            TabIndex        =   88
            Top             =   1890
            Width           =   3000
         End
         Begin xrControl.xrButton cmdContact 
            Height          =   315
            Index           =   0
            Left            =   9015
            TabIndex        =   72
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            Caption         =   "&Phone"
            AccessKey       =   "P"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin xrControl.xrButton cmdContact 
            Height          =   315
            Index           =   1
            Left            =   9015
            TabIndex        =   75
            Top             =   750
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            Caption         =   "&Mobile"
            AccessKey       =   "M"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email Add"
            Height          =   195
            Index           =   28
            Left            =   5775
            TabIndex        =   76
            Top             =   1140
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Barangay"
            ForeColor       =   &H000040C0&
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   68
            Top             =   1140
            Width           =   675
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   45
            X2              =   11205
            Y1              =   1470
            Y2              =   1470
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Town"
            Height          =   195
            Index           =   20
            Left            =   105
            TabIndex        =   66
            Top             =   810
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Index           =   23
            Left            =   105
            TabIndex        =   64
            Top             =   480
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone #"
            Height          =   195
            Index           =   26
            Left            =   5775
            TabIndex        =   70
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile #"
            Height          =   195
            Index           =   27
            Left            =   5775
            TabIndex        =   73
            Top             =   810
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "House No"
            Height          =   195
            Index           =   29
            Left            =   105
            TabIndex        =   62
            Top             =   150
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Index           =   24
            Left            =   105
            TabIndex        =   79
            Top             =   1950
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relation"
            Height          =   195
            Index           =   30
            Left            =   105
            TabIndex        =   85
            Top             =   2940
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Index           =   31
            Left            =   105
            TabIndex        =   81
            Top             =   2280
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact No"
            Height          =   195
            Index           =   32
            Left            =   105
            TabIndex        =   83
            Top             =   2610
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CONTACT PERSON(in case of emergency):"
            Height          =   195
            Index           =   33
            Left            =   105
            TabIndex        =   78
            Top             =   1590
            Width           =   3135
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mother's Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   51
            Left            =   6450
            TabIndex        =   89
            Top             =   2280
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Father's Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   52
            Left            =   6450
            TabIndex        =   87
            Top             =   1950
            Width           =   1020
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   6210
         Left            =   15
         Tag             =   "wt0;fb0"
         Top             =   315
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   10954
         BackColor       =   12632256
         ClipControls    =   0   'False
         BorderStyle     =   4
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   29
            Left            =   6885
            TabIndex        =   59
            Top             =   4260
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   30
            Left            =   6885
            TabIndex        =   61
            Top             =   4590
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   30
            Left            =   1620
            TabIndex        =   57
            Top             =   4590
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   15
            Left            =   1620
            TabIndex        =   55
            Top             =   4260
            Width           =   3000
         End
         Begin VB.ComboBox cmbPhysical 
            Height          =   315
            Index           =   5
            ItemData        =   "frmEmployee.frx":35EE
            Left            =   1620
            List            =   "frmEmployee.frx":3607
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   3075
            Width           =   3000
         End
         Begin VB.ComboBox cmbPersonal 
            Height          =   315
            Index           =   31
            ItemData        =   "frmEmployee.frx":3688
            Left            =   6885
            List            =   "frmEmployee.frx":36A7
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   3930
            Width           =   2400
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   83
            Left            =   1620
            TabIndex        =   51
            Text            =   "Jehovah's Witness"
            Top             =   3930
            Width           =   3000
         End
         Begin VB.CheckBox chkPhysical 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Gambling"
            Height          =   285
            Index           =   10
            Left            =   9165
            TabIndex        =   49
            Tag             =   "wt0;fb0"
            Top             =   3435
            Width           =   1185
         End
         Begin VB.CheckBox chkPhysical 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Smoking"
            Height          =   285
            Index           =   9
            Left            =   8040
            TabIndex        =   48
            Tag             =   "wt0;fb0"
            Top             =   3435
            Width           =   1185
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   6885
            TabIndex        =   46
            Top             =   3075
            Width           =   3000
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   6
            Left            =   6885
            TabIndex        =   44
            Top             =   2745
            Width           =   3000
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1620
            TabIndex        =   40
            Top             =   3420
            Width           =   975
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   6885
            TabIndex        =   42
            Top             =   2415
            Width           =   3000
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1620
            TabIndex        =   36
            Top             =   2745
            Width           =   3000
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1620
            TabIndex        =   34
            Top             =   2415
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   82
            Left            =   6885
            TabIndex        =   32
            Text            =   "Sayson, Marlon Agbuya"
            Top             =   1905
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   29
            Left            =   6885
            TabIndex        =   30
            Text            =   "Dela Cruz"
            ToolTipText     =   "Apleyido ng Nanay ng Dalaga Pa!!!"
            Top             =   1575
            Width           =   3000
         End
         Begin VB.ComboBox cmbPersonal 
            Height          =   315
            Index           =   5
            ItemData        =   "frmEmployee.frx":3753
            Left            =   1620
            List            =   "frmEmployee.frx":3769
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1905
            Width           =   3000
         End
         Begin VB.ComboBox cmbPersonal 
            Height          =   315
            Index           =   4
            ItemData        =   "frmEmployee.frx":37C3
            Left            =   1620
            List            =   "frmEmployee.frx":37CD
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1575
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1620
            TabIndex        =   12
            Text            =   "Sayson"
            Top             =   90
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1620
            TabIndex        =   14
            Text            =   "Gabriel Markus"
            Top             =   420
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1620
            TabIndex        =   16
            Text            =   "Caramat"
            Top             =   750
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   84
            Left            =   6885
            TabIndex        =   20
            Text            =   "Filipino"
            Top             =   90
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   6885
            TabIndex        =   22
            Text            =   "January 25, 1999"
            Top             =   420
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   81
            Left            =   6885
            TabIndex        =   24
            Text            =   "San Carlos City, Pangasinan"
            Top             =   750
            Width           =   4275
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   36
            Left            =   1620
            TabIndex        =   18
            Text            =   "Jr."
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkPhysical 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Drinking "
            Height          =   285
            Index           =   8
            Left            =   6885
            TabIndex        =   47
            Tag             =   "wt0;fb0"
            Top             =   3435
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Philhealth No"
            Height          =   195
            Index           =   59
            Left            =   5070
            TabIndex        =   58
            Top             =   4320
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HDMF/Pag-Ibig No"
            Height          =   195
            Index           =   58
            Left            =   5070
            TabIndex        =   60
            Top             =   4650
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SSS No"
            Height          =   195
            Index           =   56
            Left            =   105
            TabIndex        =   56
            Top             =   4650
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIN No"
            Height          =   195
            Index           =   57
            Left            =   105
            TabIndex        =   54
            Top             =   4320
            Width           =   525
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   45
            X2              =   11205
            Y1              =   3825
            Y2              =   3825
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   45
            X2              =   11205
            Y1              =   2310
            Y2              =   2310
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Educ Attnmnt"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   50
            Left            =   5070
            TabIndex        =   52
            Top             =   3990
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Religion"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   49
            Left            =   105
            TabIndex        =   50
            Top             =   3990
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Body Built"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   19
            Left            =   105
            TabIndex        =   37
            Top             =   3135
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hair Color"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   18
            Left            =   5070
            TabIndex        =   45
            Top             =   3135
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Eye Color"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   17
            Left            =   5070
            TabIndex        =   43
            Top             =   2805
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Blood Type"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   16
            Left            =   105
            TabIndex        =   39
            Top             =   3480
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complexion"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   15
            Left            =   5070
            TabIndex        =   41
            Top             =   2475
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Weight (kg)"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   14
            Left            =   105
            TabIndex        =   35
            Top             =   2805
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height (cm)"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   12
            Left            =   105
            TabIndex        =   33
            Top             =   2475
            Width           =   810
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   30
            X2              =   11190
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spouse "
            Height          =   195
            Index           =   9
            Left            =   5070
            TabIndex        =   31
            Top             =   1965
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mother's Maiden Nm"
            ForeColor       =   &H000040C0&
            Height          =   195
            Index           =   8
            Left            =   5070
            TabIndex        =   29
            Top             =   1635
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Civil Status"
            Height          =   195
            Index           =   4
            Left            =   105
            TabIndex        =   27
            Top             =   1965
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            Height          =   195
            Index           =   22
            Left            =   105
            TabIndex        =   25
            Top             =   1635
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Citizenship"
            Height          =   195
            Index           =   5
            Left            =   5070
            TabIndex        =   19
            Top             =   150
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            Height          =   195
            Index           =   6
            Left            =   5070
            TabIndex        =   21
            Top             =   480
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Place"
            Height          =   195
            Index           =   7
            Left            =   5070
            TabIndex        =   23
            Top             =   810
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name"
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   15
            Top             =   810
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   11
            Top             =   150
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            Height          =   195
            Index           =   10
            Left            =   105
            TabIndex        =   13
            Top             =   480
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Suffix (Jr., II, III)"
            ForeColor       =   &H000040C0&
            Height          =   195
            Index           =   21
            Left            =   105
            TabIndex        =   17
            Top             =   1140
            Width           =   1095
         End
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   7
      Left            =   7485
      TabIndex        =   150
      Top             =   8925
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&201"
      AccessKey       =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmployee.frx":37DF
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmEmployee"
Private Const pxeEmployee = 0
Private Const pxePersonal = 1
Private Const pxePhysical = 2
Private Const pxeUnknownx = -1

Private Const pxeTextBox = 0
Private Const pxeCheckBox = 1
Private Const pxeComboBox = 2

Private WithEvents oRecord As clsEmployee
Attribute oRecord.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean
Private pnFocus As Integer
Private pnIndex As Integer
Private pnObjct As Integer
Private secIndex As Integer

Private paOthers(4) As String
Private Const pxeJavaPath As String = "D:\GGC_Java_Systems\"
Private Const pxeImageLocx As String = "d:/GGC_Java_Systems/temp/empid/download/"
Private Const pxeImageExtn As String = ".jpg"

Dim lsImgName As String
Dim loImageViewer As frmImage

Private Sub LoadMaster()
   Call loadPersonal
   Call loadPhysical
   Call loadEmployee
   Call loadContact
   Call loadHasMovement
   If (loImageViewer.Visible = True) Then
   Unload loImageViewer

   End If
   Call loadEmployeeImage
   FormOpen (xeModeReady)
End Sub

Private Sub loadPhysical()
   Dim loTxt As TextBox
   Dim locmb As ComboBox
   Dim lochk As CheckBox

   For Each loTxt In txtPhysical
      loTxt = oRecord.physicalinfo(loTxt.Index)
   Next

   For Each locmb In cmbPhysical
      If IsNumeric(oRecord.physicalinfo(locmb.Index)) Then
         locmb.ListIndex = IIf(oRecord.physicalinfo(locmb.Index) = "", -1, oRecord.physicalinfo(locmb.Index))
      Else
         locmb.ListIndex = FindIndex(oRecord.physicalinfo(locmb.Index), locmb)
      End If
   Next

   For Each lochk In chkPhysical
      lochk.Value = IFNull(oRecord.physicalinfo(lochk.Index), 0)
   Next

End Sub

Private Sub loadPersonal()
   Dim loTxt As TextBox
   Dim locmb As ComboBox
   Dim lochk As CheckBox

   For Each loTxt In txtPersonal
      Select Case loTxt.Index
      Case 7
         If IsDate(oRecord.personalinfo(loTxt.Index)) Then
            loTxt = Format(oRecord.personalinfo(loTxt.Index), "Mmmm DD, YYYY")
         Else
            loTxt = ""
         End If
      Case Else
         loTxt = oRecord.personalinfo(loTxt.Index)
      End Select
   Next

   For Each locmb In cmbPersonal
      If IsNumeric(oRecord.personalinfo(locmb.Index)) Then
         locmb.ListIndex = oRecord.personalinfo(locmb.Index)
      Else
         locmb.ListIndex = FindIndex(oRecord.personalinfo(locmb.Index), locmb)
      End If
   Next

End Sub

Private Sub loadContact()
   Dim loTxt As TextBox

   For Each loTxt In txtContact
      loTxt = IFNull(oRecord.contactinfo(loTxt.Index))
   Next

End Sub

Private Sub loadEmployee()
   Dim loTxt As TextBox
   Dim locmb As ComboBox
   Dim lochk As CheckBox

   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 0
         loTxt = Format(oRecord.Master(loTxt.Index), "@@@@@@-@@@@@@")
      Case 12
         loTxt = Format(oRecord.Master(loTxt.Index), "@@@-@@@@@-@@")
         txtSearch(0) = loTxt
      Case 4, 15, 16
         If IsDate(oRecord.Master(loTxt.Index)) Then
            loTxt = Format(oRecord.Master(loTxt.Index), "Mmmm DD, YYYY")
         Else
            loTxt = ""
         End If
      Case 17
         If oRecord.Master("cRecdStat") = "0" Then
            If IsDate(oRecord.Master(loTxt.Index)) Then
               loTxt = Format(oRecord.Master(loTxt.Index), "Mmmm DD, YYYY")
            Else
               loTxt = ""
            End If
         Else
            loTxt = ""
         End If
      Case 22
         loTxt = Format(oRecord.Master(loTxt.Index), "#,##0.00")
         'TO BE USED LATER ON WITH DISABLING SALARY UPDATE
         loTxt.Tag = oRecord.Master(loTxt.Index)

         If oApp.UserLevel <= xeSupervisor And oRecord.Master("sEmpLevID") >= "1" Then
            loTxt.Text = "0.00"
         End If
      Case 23
         loTxt = Format(IFNull(oRecord.Master(loTxt.Index), 0#), "#,##0.00")
         'TO BE USED LATER ON WITH DISABLING SALARY UPDATE
         loTxt.Tag = IFNull(oRecord.Master(loTxt.Index), 0#)
      Case 35
         If oApp.ProductID = "PetMgr" Then
            loTxt.Text = IFNull(oRecord.Master(loTxt.Index), "")
         Else
            loTxt.Text = "****"
         End If
      Case 83 'Salary Level
         'kalyptus - 2017.01.10 09:25AM
         'Allow supervisor
         loTxt = IFNull(oRecord.Master(loTxt.Index))
         loTxt.Tag = IFNull(oRecord.Master(loTxt.Index))
         If oApp.UserLevel <= xeSupervisor And oRecord.Master("sEmpLevID") >= "1" Then
            loTxt.Text = ""
         End If
      Case Else
         loTxt = IFNull(oRecord.Master(loTxt.Index))
         Select Case loTxt.Index
         Case 80, 81, 86, 87
            loTxt.Tag = IFNull(oRecord.Master(loTxt.Index))
         Case 89
            txtSearch(1) = loTxt
         End Select
      End Select
   Next

  If oRecord.Master("cSecTypex") <> "1" Then
   cmbSecType.ListIndex = 0
  Else
   cmbSecType.ListIndex = 1
  End If

   For Each locmb In cmbField
      If IsNumeric(oRecord.Master(locmb.Index)) Then
         locmb.ListIndex = oRecord.Master(locmb.Index)
      Else
         locmb.ListIndex = FindIndex(IFNull(oRecord.Master(locmb.Index), ""), locmb)
      End If
   Next

   For Each lochk In chkField
      lochk.Value = Val(IFNull(oRecord.Master(lochk.Index), 0))
   Next

   Label2 = IIf(oRecord.Master("cRecdStat") = "1", "ACTIVE", "INACTIVE")
End Sub
Private Sub loadEmployeeImage()
    lsImgName = ""

         With oRecord
                            lsImgName = pxeImageLocx & oRecord.Master("sEmployID") & pxeImageExtn

                If Dir(lsImgName) <> "" Then
                Image1.Stretch = True
                Image1.Picture = LoadPicture(lsImgName)

            Else
                Image1.Picture = Nothing
            End If

            'If .EmployeeID = "M00113001232" Then
            '    PictureBox1.BackgroundImage = My.Resources.MacarlingFed
            'Else
            'End If
        End With
End Sub

Private Sub Image1_Click()
   Dim lnResult As Long
'   lsImgName = pxeImageLocx & oRecord.Master("sEmployID") & pxeImageExtn
   If (loImageViewer.Visible = True) Then
       Unload frmImage
   End If
   If (Not txtField(0).Text = "") Then
           If (Image1.Picture = 0) Then
              If MsgBox("Do you want to Download this Employee Image?", vbQuestion + vbYesNo, "Employee Entry Warning") = vbYes Then
                If (Dir(pxeJavaPath & "DLEmpImage.bat") <> "") Then
                    lnResult = (RMJExecute(pxeJavaPath & "DLEmpImage.bat " & oRecord.Master("sEmployID")))
                    If (lnResult = 0) Then
                        If Dir(lsImgName) <> "" Then
                            Image1.Stretch = True
                            loImageViewer.ImageSource (lsImgName)
                            Image1.Picture = LoadPicture(lsImgName)

                            loImageViewer.Show vbModal
                        End If
                    End If

                    If (lnResult = 1) Then
                        MsgBox "Image Does'nt Exist!! Please Inform MIS Department for uploading image!!", vbInformation, "Notice"
                    End If
                 Else 'path check
                     MsgBox "File Path Does'nt Exist" & pxeJavaPath & "DLEmpImage.bat" & "Please Inform MIS Dept !!", vbInformation, "Notice"
                End If
             End If 'messageboxyesno

          Else

               loImageViewer.ImageSource (lsImgName)
               loImageViewer.Show vbModal
          End If
          End If

End Sub

Private Function FindIndex(ByVal Value As String, ByVal combo As ComboBox) As Integer
   Dim lnCtr As Integer

   FindIndex = -1
   If Len(Trim(Value)) > 0 Then
      For lnCtr = 0 To combo.ListCount - 1
         If Left(combo.List(lnCtr), 1) = Value Then
            FindIndex = lnCtr
            Exit For
         End If
      Next
   End If
End Function

Private Sub chkField_GotFocus(Index As Integer)
   pnIndex = Index
   pnObjct = pxeCheckBox
End Sub

Private Sub chkField_Validate(Index As Integer, Cancel As Boolean)
   oRecord.Master(Index) = chkField(Index).Value
End Sub

Private Sub chkPhysical_GotFocus(Index As Integer)
   pnIndex = Index
   pnObjct = pxeCheckBox
End Sub

Private Sub chkPhysical_Validate(Index As Integer, Cancel As Boolean)
   oRecord.physicalinfo(Index) = chkPhysical(Index).Value
End Sub

Private Sub cmbField_GotFocus(Index As Integer)
   pnIndex = Index
   pnObjct = pxeComboBox
End Sub

Private Sub cmbField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 8  ' Employee Status/Type
      oRecord.Master(Index) = IIf(cmbField(Index).ListIndex >= 0, Left(cmbField(Index).List(cmbField(Index).ListIndex), 1), "")
   Case 9  ' Employee Level
      'kalyptus - 2017.01.10 09:56am
      'Do not allow a supervisor level user to update the employee level of a supervisor level and up employees
      If oApp.UserLevel <= xeSupervisor And oRecord.Master("sEmpLevID") >= "1" Then
         Exit Sub
      End If
      oRecord.Master(Index) = cmbField(Index).ListIndex
   Case 20 ' Salary Type
      oRecord.Master(Index) = IIf(cmbField(Index).ListIndex >= 0, Left(cmbField(Index).List(cmbField(Index).ListIndex), 1), "")
   Case 21 ' Computation Rate
      oRecord.Master(Index) = cmbField(Index).ListIndex
   End Select
End Sub

Private Sub cmbPersonal_GotFocus(Index As Integer)
   pnIndex = Index
   pnObjct = pxeComboBox
End Sub

Private Sub cmbPersonal_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 4  ' Gender
      oRecord.personalinfo(Index) = IIf(cmbPersonal(Index).ListIndex = -1, "", cmbPersonal(Index).ListIndex)
   Case 5  ' Civil Status
      oRecord.personalinfo(Index) = IIf(cmbPersonal(Index).ListIndex = -1, "", cmbPersonal(Index).ListIndex)
   Case 31 ' Educational Attainment
      oRecord.personalinfo(Index) = IIf(cmbPersonal(Index).ListIndex = -1, "", cmbPersonal(Index).ListIndex)
   End Select
End Sub

Private Sub cmbPhysical_GotFocus(Index As Integer)
   pnIndex = Index
   pnObjct = pxeComboBox
End Sub

Private Sub cmbPhysical_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 5 'Body Built
      oRecord.physicalinfo(Index) = IIf(cmbPhysical(Index).ListIndex = -1, "", cmbPhysical(Index).ListIndex)
   End Select
End Sub

Private Sub cmbSecType_Click()
   secIndex = IIf(oRecord.Master("cSecTypex") = "1", "1", "0")

   If cmbSecType.ListIndex <> secIndex Then
      If oRecord.isApprovalOk = False Then cmbSecType.ListIndex = secIndex
   End If
End Sub

Private Sub cmbSecType_Validate(Cancel As Boolean)
 oRecord.Master("cSecTypex") = cmbSecType.ListIndex
End Sub

Private Sub cmdButton_Click(Index As Integer)

   Select Case Index
   Case 0 ' Close
      Unload Me
   Case 1 ' Browse
      Call oRecord.SearchTransaction("", False)
      LoadMaster
   Case 2 ' Update
      If oRecord.EditMode = xeModeReady Or oRecord.EditMode = xeModeUpdate Then
         InitForm xeModeUpdate
      End If
   Case 3 ' New
      Call oRecord.NewTransaction
      LoadMaster
      If Not oRecord.EditMode = xeModeUnknown Then
         InitForm xeModeUpdate
      End If
   Case 4 ' Save
      If oRecord.SaveTransaction Then
         MsgBox "Employee Info save successfully", vbInformation, "Employee Master Entry"
         InitForm xeModeReady
         clearFields
      End If
   Case 5 ' Search
      If pnIndex >= 80 Then
         Select Case pnFocus
         Case pxeEmployee
            oRecord.Master(pnIndex) = IIf(Trim(txtField(Index)) = "", "*", txtField(Index))
         Case pxePersonal
            oRecord.personalinfo(pnIndex) = IIf(Trim(txtPersonal(Index)) = "", "*", txtPersonal(Index))
         Case pxePhysical
            oRecord.physicalinfo(pnIndex) = IIf(Trim(txtPhysical(Index)) = "", "*", txtPhysical(Index))
         End Select
      End If
   Case 6 'cancel
      If MsgBox("Do you really want to cancel the entry?", vbCritical + vbYesNo, "Employee Entry Warning") = vbYes Then
         If oRecord.EditMode = xeModeReady Or oRecord.EditMode = xeModeUpdate Then
            oRecord.OpenTransaction oRecord.Master("sEmployID")
         Else
            oRecord.NewTransaction
         End If

         MsgBox "Employee entry cancelled successfully", vbInformation, "Employee Master Entry"

         LoadMaster
         Call InitForm(xeModeReady)
      End If
   Case 7
      Load frm201Master
      Call frm201Master.LoadMaster(oRecord.Master("sEmployID"))
      frm201Master.Show 1
   End Select
End Sub

Private Sub cmdRestDay_Click(Index As Integer)
   Dim loFrm As frmEmployeeShift
   Set loFrm = New frmEmployeeShift

   loFrm.EmployeeID = oRecord.Master("sEmployID")
   loFrm.Show 1
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      Call InitForm(xeModeReady)
      bLoaded = True
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oRecord = New clsEmployee
   Set oRecord.AppDriver = oApp
   oRecord.Branch = oApp.BranchCode
   oRecord.InitTransaction
   Set loImageViewer = New frmImage
   Call LoadMaster

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me

   oSkin.ApplySkin xeFormMaintenance

   FormOpen (xeModeReady)
   cmbSecType.Visible = LCase(oApp.ProductID) = "petmgr"

   Call clearFields
   SSTab1.Tab = 0
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oRecord = Nothing
   Set oSkin = Nothing
   If loImageViewer.Visible = True Then
   Unload loImageViewer
   End If
End Sub

Private Sub Label3_Click()
'   If Label3.Caption <> "" Then
'      If oRecord.Master("sEmployID") <> "" Then
'         frmEmployeeMovementReg.Tag = "Label3_Click"
'         Call frmEmployeeMovementReg.LoadMovement(Label3.Tag)
'         frmEmployeeMovementReg.Show
'      End If
'   End If
End Sub

Private Sub oRecord_ClientRetrieved(ByVal Index As Variant)
   txtPersonal(Index) = oRecord.personalinfo(Index)
End Sub

Private Sub oRecord_MasterRetrieved(ByVal Index As Variant)
   Select Case Index
   Case 89
      txtField(0) = Format(oRecord.Master("sEmployID"), "@@@@@@-@@@@@@")
      loadPersonal
   Case 36
      cmbSecType.ListIndex = oRecord.Master("cSecTypex")
   End Select

   txtField(Index) = oRecord.Master(Index)
End Sub

Private Sub oRecord_PhysicalRetrieved(ByVal Index As Variant)
   txtPhysical(Index) = oRecord.physicalinfo(Index)
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
   Dim lbCancel As Boolean
   'Is Save button visible
   If cmdButton(4).Visible Then
      With SSTab1
         Select Case pnFocus
         Case pxePersonal
            Select Case pnObjct
            Case pxeCheckBox
            Case pxeComboBox
               Call cmbPersonal_Validate(pnIndex, lbCancel)
            Case pxeTextBox
               Call txtPersonal_Validate(pnIndex, lbCancel)
            End Select
         Case pxePhysical
            Select Case pnObjct
            Case pxeCheckBox
               Call chkPhysical_Validate(pnIndex, lbCancel)
            Case pxeComboBox
               Call cmbPhysical_Validate(pnIndex, lbCancel)
            Case pxeTextBox
               Call cmbPhysical_Validate(pnIndex, lbCancel)
            End Select
         Case pxeEmployee
            Select Case pnObjct
            Case pxeCheckBox
               Call chkField_Validate(pnIndex, lbCancel)
            Case pxeComboBox
               Call cmbField_Validate(pnIndex, lbCancel)
            Case pxeTextBox
               Call txtField_Validate(pnIndex, lbCancel)
            End Select
         End Select

         If .Tab = 0 Then
            txtPersonal(1).SetFocus
         ElseIf .Tab = 1 Then
            txtPersonal(9).SetFocus
         ElseIf .Tab = 2 Then
            txtField(80).SetFocus
         End If
      End With
   End If
End Sub

Private Sub txtContact_GotFocus(Index As Integer)
   With txtContact(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtContact_LostFocus(Index As Integer)
   With txtContact(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtContact_Validate(Index As Integer, Cancel As Boolean)
   oRecord.contactinfo(Index) = txtContact(Index)
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
   Case 12 'Control No
      txtField(Index) = IFNull(oRecord.Master(Index))
   Case 4, 15, 16
      If IsDate(oRecord.Master(Index)) Then
         txtField(Index) = Format(oRecord.Master(Index), "MM/DD/YYYY")
      End If
   Case 22
      txtField(Index) = Format(oRecord.Master(Index), "##0.00")
      If Val(txtField(Index).Tag) > 0 Then
         txtField(Index).Locked = True
      Else
         txtField(Index).Locked = False
      End If
   Case 23
      txtField(Index) = Format(oRecord.Master(Index), "##0.00")
'      If Val(txtField(Index).Tag) > 0 Then
''         If Not (LCase(oApp.ProductID) = "petmgr" And oApp.UserLevel >= xeManager) Then
'            txtField(Index).Locked = True
''         End If
'      Else
'         txtField(Index).Locked = False
'      End If
   Case 80
      If SSTab1.Tab <> 2 Then SSTab1.Tab = 2
   Case 80, 81, 83, 86, 87
      If Trim(txtField(Index).Tag) <> "" Then
         txtField(Index).Locked = True
      Else
         txtField(Index).Locked = False
      End If
   End Select

   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
   pnFocus = pxeEmployee
   pnObjct = pxeTextBox
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3
      If Index >= 80 Then
         'kalyptus - 2017.01.10 09:56am
         'Do not allow a supervisor level user to update the employee level of a supervisor level and up employees
         If Index = 83 Then
            If oApp.UserLevel <= xeSupervisor And oRecord.Master("sEmpLevID") >= "1" Then
               Exit Sub
            End If
         End If

         If oRecord.SearchMaster(Index, txtField(Index).Text) Then
            SetNextFocus
         End If
      End If
   Case vbKeyReturn
      'kalyptus - 2017.01.10 09:56am
      'Do not allow a supervisor level user to update the employee level of a supervisor level and up employees
      If Index >= 80 Then
         If Index = 83 Then
            If oApp.UserLevel <= xeSupervisor And oRecord.Master("sEmpLevID") >= "1" Then
               Exit Sub
            End If
         End If

         If txtField(Index) <> "" Then
            Call oRecord.SearchMaster(Index, txtField(Index).Text)
         Else
            oRecord.Master(Index) = txtField(Index).Text
         End If
      End If
   End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)

   If Index = 0 Or Index = 12 Then
      txtField(Index) = Replace(txtField(Index).Text, "-", "")
   End If

   Select Case Index
   Case 22
      'kalyptus - 2017.01.10 09:56am
      'Do not allow a supervisor level user to update the employee level of a supervisor level and up employees

      If oApp.UserLevel <= xeSupervisor And oRecord.Master("sEmpLevID") >= "1" Then
         Exit Sub
      End If

      If Val(txtField(Index).Tag) = 0 Then
         oRecord.Master(Index) = txtField(Index).Text
      End If
   Case 23
      If Val(txtField(Index).Tag) = 0 Then
         oRecord.Master(Index) = txtField(Index).Text
      Else
         If LCase(oApp.ProductID) = "petmgr" And oApp.UserLevel >= xeManager Then
            oRecord.Master(Index) = txtField(Index).Text
         End If
      End If
   Case 83
      'kalyptus - 2017.01.10 09:56am
      'Do not allow a supervisor level user to update the employee level of a supervisor level and up employees
      If oApp.UserLevel <= xeSupervisor And oRecord.Master("sEmpLevID") >= "1" Then
         Exit Sub
      End If

      oRecord.Master(Index) = txtField(Index).Text

   Case Else
      oRecord.Master(Index) = txtField(Index).Text
   End Select

   Select Case Index
   Case 0
      txtField(Index) = Format(oRecord.Master(Index), "@@@@@@-@@@@@@")
   Case 12
      txtField(Index) = Format(oRecord.Master(Index), "@@@-@@@@@-@@")
   Case 4, 15, 16
      txtField(Index) = Format(oRecord.Master(Index), "Mmmm DD, YYYY")
   Case 22, 23
      txtField(Index) = Format(oRecord.Master(Index), "#,##0.00")
   End Select
End Sub

Private Sub txtPersonal_GotFocus(Index As Integer)

   Select Case Index
   Case 1
      If SSTab1.Tab <> 0 Then SSTab1.Tab = 0
   Case 7
      If IsDate(oRecord.personalinfo(Index)) Then
         txtPersonal(Index) = Format(oRecord.personalinfo(Index), "MM/DD/YYYY")
      End If
   Case 9
      If SSTab1.Tab <> 1 Then SSTab1.Tab = 1
   Case 29
      txtPersonal(Index).Locked = Not (cmbPersonal(4).ListIndex = 1 And cmbPersonal(5).ListIndex <> 0)
   End Select

   With txtPersonal(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
   pnFocus = pxePersonal
   pnObjct = pxeTextBox
End Sub

Private Sub txtPersonal_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3
      If Index >= 80 Then
         If oRecord.SearchPersonal(Index, txtPersonal(Index).Text) Then
            SetNextFocus
         End If
      End If
   Case vbKeyReturn
      If Index >= 80 Then
         If txtPersonal(Index) <> "" Then
            Call oRecord.SearchPersonal(Index, txtPersonal(Index).Text)
         Else
            oRecord.personalinfo(Index) = txtPersonal(Index).Text
         End If
      End If
   End Select
End Sub

Private Sub txtPersonal_LostFocus(Index As Integer)
   With txtPersonal(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtPersonal_Validate(Index As Integer, Cancel As Boolean)
   oRecord.personalinfo(Index) = Replace(txtPersonal(Index), vbCrLf, "")

   Select Case Index
   Case 7
      txtPersonal(Index) = Format(oRecord.personalinfo(Index), "MM/DD/YYYY")
   Case Else
      txtPersonal(Index) = oRecord.personalinfo(Index)
   End Select
End Sub

Private Sub txtPhysical_GotFocus(Index As Integer)
   With txtPhysical(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
   pnFocus = pxePhysical
   pnObjct = pxeTextBox

End Sub

Private Sub txtPhysical_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3
      If Index >= 80 Then
         If oRecord.SearchPhysical(Index, txtPhysical(Index).Text) Then
            SetNextFocus
         End If
      End If
   Case vbKeyReturn
      If Index >= 80 Then
         If txtPhysical(Index) <> "" Then
            Call oRecord.SearchPhysical(Index, txtPhysical(Index).Text)
         Else
            oRecord.physicalinfo(Index) = txtPhysical(Index).Text
         End If
      End If
   End Select
End Sub

Private Sub txtPhysical_LostFocus(Index As Integer)
   With txtPhysical(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtPhysical_Validate(Index As Integer, Cancel As Boolean)
   oRecord.physicalinfo(Index) = txtPhysical(Index)
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
   With txtSearch(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
   pnFocus = pxeUnknownx
   pnObjct = pxeUnknownx

End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Call oRecord.SearchTransaction(txtSearch(Index), IIf(Index = 0, True, False))
      LoadMaster
   End If
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtSearch_Validate(Index As Integer, Cancel As Boolean)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_Validate"
'   Debug.Print pxeMODULENAME & "." & lsOldProc
''   On Error GoTo errProc
'
'   With txtSearch(Index)
'         If .Text = "" Then
'            oRecord.InitTransaction
'            LoadMaster
'            Exit Sub
'         Else
'            If oRecord.SearchTransaction(.Text, IIf(Index = 0, True, False)) Then
'               LoadMaster
'            End If
'         End If
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index _
'                       & ", " & Cancel & " )", True
End Sub

Private Sub InitForm(ByVal fnEditMode As Integer)
   If oApp.ProductID = "AssInv" Then
      cmdButton(0).Visible = True
      cmdButton(1).Visible = True

      cmdButton(2).Visible = False
      cmdButton(3).Visible = False
      cmdButton(4).Visible = False
      cmdButton(5).Visible = False
      cmdButton(6).Visible = False
      cmdButton(7).Visible = False

      xrFrame1.Enabled = False      'for loading image
      xrFrame2.Enabled = True       'Search frame will be the only frame to be enabled....
      xrFrame3.Enabled = False
      xrFrame4.Enabled = False
      xrFrame5.Enabled = False

      Exit Sub
   End If

   cmdButton(0).Visible = (fnEditMode = xeModeReady)
   cmdButton(1).Visible = (fnEditMode = xeModeReady)
   cmdButton(2).Visible = (fnEditMode = xeModeReady)
   cmdButton(3).Visible = (fnEditMode = xeModeReady)
   cmdButton(7).Visible = (fnEditMode = xeModeReady)


   cmdButton(4).Visible = Not (fnEditMode = xeModeReady)
   cmdButton(5).Visible = Not (fnEditMode = xeModeReady)
   cmdButton(6).Visible = Not (fnEditMode = xeModeReady)

   xrFrame1.Enabled = Not (fnEditMode = xeModeReady)
   xrFrame3.Enabled = Not (fnEditMode = xeModeReady)
   xrFrame4.Enabled = Not (fnEditMode = xeModeReady)
   xrFrame5.Enabled = Not (fnEditMode = xeModeReady)

   If fnEditMode = xeModeReady Then
      txtSearch(0).SetFocus
   Else
      txtField(89).SetFocus
   End If

   SSTab1.Tab = 0
End Sub

    Private Sub FormOpen(ByVal fnEditMode As Integer)
If Not txtSearch(1).Text = "" Then
   xrFrame1.Enabled = (fnEditMode = xeModeReady)


   txtField(89).Locked = (fnEditMode = xeModeReady)
   txtField(86).Locked = (fnEditMode = xeModeReady)
   txtField(12).Locked = (fnEditMode = xeModeReady)
   txtField(13).Locked = (fnEditMode = xeModeReady)
   txtField(35).Locked = (fnEditMode = xeModeReady)
   End If
End Sub

Private Sub clearFields()
   Dim loTxt As TextBox
   Dim locmb As ComboBox
   Dim lochk As CheckBox
'   secIndex = 0

   For Each loTxt In txtField
      loTxt = ""
   Next
   For Each loTxt In txtPersonal
      loTxt = ""
   Next
   For Each loTxt In txtPhysical
      loTxt = ""
   Next
   For Each loTxt In txtSearch
      loTxt = ""
   Next
   For Each loTxt In txtContact
      loTxt = ""
   Next
   For Each locmb In cmbPersonal
      locmb.ListIndex = -1
   Next
   For Each locmb In cmbPhysical
      locmb.ListIndex = -1
   Next
   For Each locmb In cmbField
      locmb.ListIndex = -1
   Next
   For Each lochk In chkField
      lochk.Value = 0
   Next
   For Each lochk In chkPhysical
      lochk.Value = 0
   Next

         'Clear movement indicator
   Label3.Caption = ""
   Label3.Tag = ""
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oApp
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      If bEnd Then
         .xShowError
         End
      Else
         With Err
            .Raise .Number, .Source, .Description
         End With
      End If
   End With
End Sub

Private Sub loadHasMovement()
   Dim lsSQL As String
   Dim loRS As Recordset

   lsSQL = "SELECT sTransNox" & _
                ", dEffectve" & _
          " FROM Employee_Movement" & _
          " WHERE sEmployID = " & strParm(oRecord.Master("sEmployID")) & _
            " AND cTranStat = " & strParm(xeStateClosed) & _
          " ORDER BY dEffectve DESC"
   Set loRS = oApp.Connection.Execute(lsSQL, , adCmdText)

   If loRS.EOF Then
      Label3.Tag = ""
      Label3.Caption = ""
   Else
      Label3.Tag = loRS("sTransNox")
       Label3.Caption = "RECORD WILL BE UPDATED ON " & Format(loRS("dEffectve"), "MM/DD/YYYY")
    End If
End Sub


