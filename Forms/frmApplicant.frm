VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmApplicant 
   BorderStyle     =   0  'None
   Caption         =   "Employment Application"
   ClientHeight    =   9615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Visible         =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   10605
      TabIndex        =   130
      Top             =   8835
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
      Picture         =   "frmApplicant.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   9045
      TabIndex        =   128
      Top             =   8835
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
      Picture         =   "frmApplicant.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   9825
      TabIndex        =   129
      Top             =   8835
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
      Picture         =   "frmApplicant.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1650
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2910
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   144
         Top             =   1215
         Width           =   4815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   142
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
         TabIndex        =   141
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
         TabIndex        =   140
         Top             =   885
         Width           =   4815
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position Desired"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   145
         Top             =   1275
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "NEW APPLICANT"
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
         TabIndex        =   143
         Tag             =   "eb0;et0"
         Top             =   150
         Width           =   2400
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
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   2
         Top             =   945
         Width           =   570
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
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1380
         Left            =   9690
         Top             =   75
         Width           =   1380
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applicant Name"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   615
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applicant No."
         Height          =   195
         Index           =   21
         Left            =   105
         TabIndex        =   0
         Top             =   165
         Width           =   960
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   9825
      TabIndex        =   132
      Top             =   8835
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
      Picture         =   "frmApplicant.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   8265
      TabIndex        =   127
      Top             =   8835
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
      Picture         =   "frmApplicant.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   10605
      TabIndex        =   133
      Top             =   8835
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
      Picture         =   "frmApplicant.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   9045
      TabIndex        =   131
      Top             =   8835
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
      Picture         =   "frmApplicant.frx":2CDC
      PicturePos      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6390
      Left            =   90
      TabIndex        =   3
      Tag             =   "wt0;fb0"
      Top             =   2235
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11271
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Personal Info"
      TabPicture(0)   =   "frmApplicant.frx":3456
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "xrFrame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Educational Background"
      TabPicture(1)   =   "frmApplicant.frx":3472
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "xrFrame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Employment Information"
      TabPicture(2)   =   "frmApplicant.frx":348E
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
            TabIndex        =   134
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
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   17
            Left            =   6885
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Text            =   "January 12, 2011"
            Top             =   1575
            Width           =   3000
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Credit Investigator"
            Height          =   315
            Index           =   34
            Left            =   5790
            TabIndex        =   138
            Tag             =   "wt0;fb0"
            Top             =   4725
            Width           =   2235
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Mechanic"
            Height          =   315
            Index           =   33
            Left            =   3975
            TabIndex        =   137
            Tag             =   "wt0;fb0"
            Top             =   4725
            Width           =   1695
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Manager"
            Height          =   315
            Index           =   32
            Left            =   2130
            TabIndex        =   136
            Tag             =   "wt0;fb0"
            Top             =   4725
            Width           =   1695
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Collector"
            Height          =   315
            Index           =   31
            Left            =   225
            TabIndex        =   135
            Tag             =   "wt0;fb0"
            Top             =   4725
            Width           =   1695
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Company Subsidized?"
            Height          =   315
            Index           =   28
            Left            =   6885
            TabIndex        =   118
            Tag             =   "wt0;fb0"
            Top             =   3375
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   91
            Left            =   6885
            TabIndex        =   126
            Top             =   4230
            Width           =   3000
         End
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   9
            ItemData        =   "frmApplicant.frx":34AA
            Left            =   1620
            List            =   "frmApplicant.frx":34BD
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   2400
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   26
            Left            =   6885
            TabIndex        =   124
            Top             =   3900
            Width           =   1920
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   90
            Left            =   6885
            TabIndex        =   116
            Text            =   "National Capital Region"
            Top             =   2745
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   80
            Left            =   1620
            TabIndex        =   83
            Text            =   "Senior Programmer 1"
            Top             =   90
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   81
            Left            =   1620
            TabIndex        =   85
            Text            =   "IT Department"
            Top             =   420
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   87
            Left            =   6885
            TabIndex        =   90
            Text            =   "GMC Dagupan - Honda"
            Top             =   90
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   88
            Left            =   6885
            TabIndex        =   92
            Text            =   "LGK Guanzon"
            Top             =   420
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   15
            Left            =   1620
            TabIndex        =   98
            Text            =   "January 12, 2011"
            Top             =   1575
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1620
            TabIndex        =   96
            Text            =   "December 31, 2012"
            Top             =   1245
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   16
            Left            =   6885
            TabIndex        =   100
            Text            =   "January 12, 2011"
            Top             =   1245
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   22
            Left            =   1605
            TabIndex        =   122
            Text            =   "0"
            Top             =   4230
            Width           =   1920
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   23
            Left            =   1605
            TabIndex        =   120
            Text            =   "0"
            Top             =   3900
            Width           =   1920
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   82
            Left            =   1620
            TabIndex        =   87
            Text            =   "IT Department"
            Top             =   750
            Width           =   2100
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   84
            Left            =   6885
            TabIndex        =   94
            Text            =   "National Capital Region"
            Top             =   750
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   83
            Left            =   1620
            TabIndex        =   108
            Top             =   2745
            Width           =   1920
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   85
            Left            =   6885
            TabIndex        =   114
            Text            =   "0"
            Top             =   2415
            Width           =   3000
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Deduct Gov't Contribution"
            Height          =   315
            Index           =   19
            Left            =   6885
            TabIndex        =   117
            Tag             =   "wt0;fb0"
            Top             =   3060
            Width           =   3000
         End
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   8
            ItemData        =   "frmApplicant.frx":3505
            Left            =   1620
            List            =   "frmApplicant.frx":3512
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   2070
            Width           =   3000
         End
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   20
            ItemData        =   "frmApplicant.frx":3535
            Left            =   1620
            List            =   "frmApplicant.frx":353F
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   3060
            Width           =   3000
         End
         Begin VB.ComboBox cmbField 
            Height          =   315
            Index           =   21
            ItemData        =   "frmApplicant.frx":3555
            Left            =   6885
            List            =   "frmApplicant.frx":3562
            Style           =   2  'Dropdown List
            TabIndex        =   112
            Top             =   2070
            Width           =   3000
         End
         Begin xrControl.xrButton cmdRestDay 
            Height          =   315
            Index           =   3
            Left            =   3765
            TabIndex        =   88
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resigned/Terminated"
            Height          =   195
            Index           =   60
            Left            =   5040
            TabIndex        =   101
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
            TabIndex        =   125
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
            TabIndex        =   123
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
            TabIndex        =   115
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
            TabIndex        =   82
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
            TabIndex        =   84
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
            TabIndex        =   89
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
            TabIndex        =   91
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
            TabIndex        =   97
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
            TabIndex        =   95
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
            TabIndex        =   99
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
            TabIndex        =   121
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
            TabIndex        =   119
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
            TabIndex        =   86
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
            TabIndex        =   103
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
            TabIndex        =   105
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
            TabIndex        =   107
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
            TabIndex        =   93
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
            TabIndex        =   109
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
            TabIndex        =   111
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
            TabIndex        =   113
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
            TabIndex        =   68
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
            TabIndex        =   60
            Text            =   "frmApplicant.frx":357E
            Top             =   1080
            Width           =   3885
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   80
            Left            =   1620
            TabIndex        =   58
            Text            =   "City of San Carlos"
            Top             =   750
            Width           =   3885
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   12
            Left            =   6855
            TabIndex        =   62
            Text            =   "(075) 522-1097"
            Top             =   420
            Width           =   2070
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   13
            Left            =   6855
            TabIndex        =   65
            Text            =   "+639163103800"
            Top             =   750
            Width           =   2070
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   9
            Left            =   1620
            TabIndex        =   54
            Text            =   "45"
            Top             =   90
            Width           =   975
         End
         Begin VB.TextBox txtContact 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1620
            TabIndex        =   71
            Top             =   1890
            Width           =   3885
         End
         Begin VB.TextBox txtContact 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1620
            TabIndex        =   77
            Top             =   2880
            Width           =   2070
         End
         Begin VB.TextBox txtContact 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1620
            TabIndex        =   75
            Top             =   2550
            Width           =   2070
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   10
            Left            =   1620
            MultiLine       =   -1  'True
            TabIndex        =   56
            Text            =   "frmApplicant.frx":35A1
            Top             =   420
            Width           =   3885
         End
         Begin VB.TextBox txtContact 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1620
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   2220
            Width           =   3885
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   21
            Left            =   8130
            TabIndex        =   81
            Top             =   2220
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   20
            Left            =   8130
            TabIndex        =   79
            Top             =   1890
            Width           =   3000
         End
         Begin xrControl.xrButton cmdContact 
            Height          =   315
            Index           =   0
            Left            =   9015
            TabIndex        =   63
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
            TabIndex        =   66
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
            TabIndex        =   67
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
            TabIndex        =   59
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
            TabIndex        =   57
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
            TabIndex        =   55
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
            TabIndex        =   61
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
            TabIndex        =   64
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
            TabIndex        =   53
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
            TabIndex        =   70
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
            TabIndex        =   76
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
            TabIndex        =   72
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
            TabIndex        =   74
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
            TabIndex        =   69
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
            TabIndex        =   80
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
            TabIndex        =   78
            Top             =   1950
            Width           =   1020
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   6030
         Left            =   15
         Tag             =   "wt0;fb0"
         Top             =   315
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   10636
         BackColor       =   12632256
         ClipControls    =   0   'False
         BorderStyle     =   4
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   11
            Left            =   6885
            TabIndex        =   158
            Text            =   "Sayson, Marlon Agbuya"
            Top             =   1905
            Width           =   945
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   8
            Left            =   6885
            TabIndex        =   156
            Text            =   "Sayson, Marlon Agbuya"
            Top             =   2895
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   6
            Left            =   6885
            TabIndex        =   154
            Text            =   "Sayson, Marlon Agbuya"
            Top             =   2565
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   5
            Left            =   6885
            TabIndex        =   152
            Text            =   "Sayson, Marlon Agbuya"
            Top             =   2235
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   1620
            TabIndex        =   150
            Text            =   "Sayson, Marlon Agbuya"
            Top             =   2895
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   1620
            TabIndex        =   148
            Text            =   "Sayson, Marlon Agbuya"
            Top             =   2565
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   6885
            TabIndex        =   146
            Top             =   5610
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   29
            Left            =   6885
            TabIndex        =   50
            Top             =   4950
            Width           =   3000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   30
            Left            =   6885
            TabIndex        =   52
            Top             =   5280
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   30
            Left            =   1620
            TabIndex        =   48
            Top             =   5610
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   15
            Left            =   1620
            TabIndex        =   46
            Top             =   5280
            Width           =   3000
         End
         Begin VB.ComboBox cmbPhysical 
            Height          =   315
            Index           =   5
            ItemData        =   "frmApplicant.frx":35B2
            Left            =   1620
            List            =   "frmApplicant.frx":35CB
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   4095
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   83
            Left            =   1620
            TabIndex        =   44
            Text            =   "Jehovah's Witness"
            Top             =   4950
            Width           =   3000
         End
         Begin VB.CheckBox chkPhysical 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Gambling"
            Height          =   285
            Index           =   10
            Left            =   9270
            TabIndex        =   42
            Tag             =   "wt0;fb0"
            Top             =   4455
            Width           =   1185
         End
         Begin VB.CheckBox chkPhysical 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Smoking"
            Height          =   285
            Index           =   9
            Left            =   8145
            TabIndex        =   41
            Tag             =   "wt0;fb0"
            Top             =   4455
            Width           =   1185
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   6885
            TabIndex        =   39
            Top             =   4095
            Width           =   3000
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   6
            Left            =   6885
            TabIndex        =   37
            Top             =   3765
            Width           =   3000
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1620
            TabIndex        =   33
            Top             =   4440
            Width           =   975
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   6885
            TabIndex        =   35
            Top             =   3435
            Width           =   3000
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1620
            TabIndex        =   29
            Top             =   3765
            Width           =   3000
         End
         Begin VB.TextBox txtPhysical 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1620
            TabIndex        =   27
            Top             =   3435
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   82
            Left            =   1620
            TabIndex        =   25
            Text            =   "Sayson, Marlon Agbuya"
            Top             =   2235
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   29
            Left            =   6885
            TabIndex        =   23
            Text            =   "Dela Cruz"
            ToolTipText     =   "Apleyido ng Nanay ng Dalaga Pa!!!"
            Top             =   1575
            Width           =   3000
         End
         Begin VB.ComboBox cmbPersonal 
            Height          =   315
            Index           =   5
            ItemData        =   "frmApplicant.frx":364C
            Left            =   1620
            List            =   "frmApplicant.frx":365C
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1905
            Width           =   3000
         End
         Begin VB.ComboBox cmbPersonal 
            Height          =   315
            Index           =   4
            ItemData        =   "frmApplicant.frx":3685
            Left            =   1620
            List            =   "frmApplicant.frx":368F
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1575
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1620
            TabIndex        =   5
            Text            =   "Sayson"
            Top             =   90
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1620
            TabIndex        =   7
            Text            =   "Gabriel Markus"
            Top             =   420
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1620
            TabIndex        =   9
            Text            =   "Caramat"
            Top             =   750
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   84
            Left            =   6885
            TabIndex        =   13
            Text            =   "Filipino"
            Top             =   90
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   6885
            TabIndex        =   15
            Text            =   "January 25, 1999"
            Top             =   420
            Width           =   3000
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   81
            Left            =   6885
            TabIndex        =   17
            Text            =   "San Carlos City, Pangasinan"
            Top             =   750
            Width           =   4275
         End
         Begin VB.TextBox txtPersonal 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   36
            Left            =   1620
            TabIndex        =   11
            Text            =   "Jr."
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkPhysical 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Drinking "
            Height          =   285
            Index           =   8
            Left            =   7185
            TabIndex        =   40
            Tag             =   "wt0;fb0"
            Top             =   4455
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Children"
            Height          =   195
            Index           =   64
            Left            =   5475
            TabIndex        =   159
            Top             =   1965
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            Height          =   195
            Index           =   63
            Left            =   5700
            TabIndex        =   157
            Top             =   2955
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            Height          =   195
            Index           =   62
            Left            =   5700
            TabIndex        =   155
            Top             =   2625
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            Height          =   195
            Index           =   61
            Left            =   5700
            TabIndex        =   153
            Top             =   2295
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mother's Name"
            Height          =   195
            Index           =   50
            Left            =   105
            TabIndex        =   151
            Top             =   2955
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Father's Name"
            Height          =   195
            Index           =   48
            Left            =   105
            TabIndex        =   149
            Top             =   2625
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Driver's License No"
            Height          =   195
            Index           =   47
            Left            =   5145
            TabIndex        =   147
            Top             =   5670
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Philhealth No"
            Height          =   195
            Index           =   59
            Left            =   5580
            TabIndex        =   49
            Top             =   5010
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HDMF/Pag-Ibig No"
            Height          =   195
            Index           =   58
            Left            =   5145
            TabIndex        =   51
            Top             =   5340
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SSS No"
            Height          =   195
            Index           =   56
            Left            =   105
            TabIndex        =   47
            Top             =   5670
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIN No"
            Height          =   195
            Index           =   57
            Left            =   105
            TabIndex        =   45
            Top             =   5340
            Width           =   525
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   30
            X2              =   11190
            Y1              =   4845
            Y2              =   4845
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   30
            X2              =   11190
            Y1              =   3315
            Y2              =   3315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Religion"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   49
            Left            =   105
            TabIndex        =   43
            Top             =   5010
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
            TabIndex        =   30
            Top             =   4155
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hair Color"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   18
            Left            =   5835
            TabIndex        =   38
            Top             =   4155
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Eye Color"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   17
            Left            =   5850
            TabIndex        =   36
            Top             =   3825
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
            TabIndex        =   32
            Top             =   4500
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complexion"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   15
            Left            =   5715
            TabIndex        =   34
            Top             =   3495
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
            TabIndex        =   28
            Top             =   3825
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
            TabIndex        =   26
            Top             =   3495
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
            Caption         =   "Spouse Name"
            Height          =   195
            Index           =   9
            Left            =   105
            TabIndex        =   24
            Top             =   2295
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mother's Maiden Nm"
            ForeColor       =   &H000040C0&
            Height          =   195
            Index           =   8
            Left            =   5070
            TabIndex        =   22
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
            TabIndex        =   20
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
            TabIndex        =   18
            Top             =   1635
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Citizenship"
            Height          =   195
            Index           =   5
            Left            =   5775
            TabIndex        =   12
            Top             =   150
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            Height          =   195
            Index           =   6
            Left            =   5820
            TabIndex        =   14
            Top             =   480
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Place"
            Height          =   195
            Index           =   7
            Left            =   5760
            TabIndex        =   16
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
            TabIndex        =   8
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
            TabIndex        =   4
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
            TabIndex        =   6
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
            TabIndex        =   10
            Top             =   1140
            Width           =   1095
         End
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   7
      Left            =   7485
      TabIndex        =   139
      Top             =   8835
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
      Picture         =   "frmApplicant.frx":36A1
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmApplicant"
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

Private paOthers(4) As String

Private Sub LoadMaster()
   Call loadPersonal
   Call loadPhysical
   Call loadEmployee
   Call loadContact
   Call loadHasMovement
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
'         txtSearch(0) = loTxt
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
      Case 22, 23
         loTxt = Format(oRecord.Master(loTxt.Index), "#,##0.00")
         'TO BE USED LATER ON WITH DISABLING SALARY UPDATE
         loTxt.Tag = oRecord.Master(loTxt.Index)
      Case 35
         If oApp.ProductID = "PetMgr" Then
            loTxt.Text = oRecord.Master(loTxt.Index)
         Else
            loTxt.Text = "****"
         End If
      Case Else
         loTxt = IFNull(oRecord.Master(loTxt.Index))
         Select Case loTxt.Index
         Case 80, 81, 83, 86, 87
            loTxt.Tag = IFNull(oRecord.Master(loTxt.Index))
         Case 89
'            txtSearch(1) = loTxt
         End Select
      End Select
   Next

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

   Label2 = "NEW APPLICANT" ' IIf(oRecord.Master("cRecdStat") = "1", "ACTIVE", "INACTIVE")
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
         ClearFields
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
   Case 6
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
   Call LoadMaster

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormMaintenance

   Call ClearFields
   SSTab1.Tab = 0
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oRecord = Nothing
   Set oSkin = Nothing
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
         If oRecord.SearchMaster(Index, txtField(Index).Text) Then
            SetNextFocus
         End If
      End If
   Case vbKeyReturn
      If Index >= 80 Then
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

'   If Index = 0 Or Index = 12 Then
'      txtField(Index) = Replace(txtField(Index).Text, "-", "")
'   End If
'
'   Select Case Index
'   Case 22
'      If Val(txtField(Index).Tag) = 0 Then
'         oRecord.Master(Index) = txtField(Index).Text
'      End If
'   Case 23
'      If Val(txtField(Index).Tag) = 0 Then
'         oRecord.Master(Index) = txtField(Index).Text
'      Else
'         If LCase(oApp.ProductID) = "petmgr" And oApp.UserLevel >= xeManager Then
'            oRecord.Master(Index) = txtField(Index).Text
'         End If
'      End If
'   Case Else
'      oRecord.Master(Index) = txtField(Index).Text
'   End Select
'
'   Select Case Index
'   Case 0
'      txtField(Index) = Format(oRecord.Master(Index), "@@@@@@-@@@@@@")
'   Case 12
'      txtField(Index) = Format(oRecord.Master(Index), "@@@-@@@@@-@@")
'   Case 4, 15, 16
'      txtField(Index) = Format(oRecord.Master(Index), "Mmmm DD, YYYY")
'   Case 22, 23
'      txtField(Index) = Format(oRecord.Master(Index), "#,##0.00")
'   End Select
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
'   oRecord.personalinfo(Index) = Replace(txtPersonal(Index), vbCrLf, "")
'
'   Select Case Index
'   Case 7
'      txtPersonal(Index) = Format(oRecord.personalinfo(Index), "MM/DD/YYYY")
'   Case Else
'      txtPersonal(Index) = oRecord.personalinfo(Index)
'   End Select
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
      
      xrFrame1.Enabled = False
'      xrFrame2.Enabled = True       'Search frame will be the only frame to be enabled....
      xrFrame3.Enabled = False
      xrFrame4.Enabled = False
      xrFrame5.Enabled = False
      
      Exit Sub
   End If
   
   cmdButton(0).Visible = False '(fnEditMode = xeModeReady)
   cmdButton(1).Visible = False '(fnEditMode = xeModeReady)
   cmdButton(2).Visible = False '(fnEditMode = xeModeReady)
   cmdButton(3).Visible = False '(fnEditMode = xeModeReady)
   cmdButton(7).Visible = False '(fnEditMode = xeModeReady)
   
'   xrFrame2.Enabled = (fnEditMode = xeModeReady)

   cmdButton(4).Visible = True 'Not (fnEditMode = xeModeReady)
   cmdButton(5).Visible = True 'Not (fnEditMode = xeModeReady)
   cmdButton(6).Visible = True 'Not (fnEditMode = xeModeReady)

   xrFrame1.Enabled = True 'Not (fnEditMode = xeModeReady)
   xrFrame3.Enabled = True 'Not (fnEditMode = xeModeReady)
   xrFrame4.Enabled = True 'Not (fnEditMode = xeModeReady)
   xrFrame5.Enabled = True 'Not (fnEditMode = xeModeReady)

   If fnEditMode = xeModeReady Then
'      txtSearch(0).SetFocus
   Else
      txtField(89).SetFocus
   End If

   SSTab1.Tab = 0
   txtField(0) = "M00116000001"
End Sub
Private Sub ClearFields()
   Dim loTxt As TextBox
   Dim locmb As ComboBox
   Dim lochk As CheckBox

   For Each loTxt In txtField
      loTxt = ""
   Next
   For Each loTxt In txtPersonal
      loTxt = ""
   Next
   For Each loTxt In txtPhysical
      loTxt = ""
   Next
'   For Each loTxt In txtSearch
'      loTxt = ""
'   Next
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


