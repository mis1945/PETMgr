VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmEmp13thMonthRelease 
   BorderStyle     =   0  'None
   Caption         =   "Release 13th Pay and Bonus"
   ClientHeight    =   8310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   14010
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1005
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   1773
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.CheckBox chkField 
         Caption         =   "Exclude Release"
         Height          =   345
         Index           =   1
         Left            =   7890
         TabIndex        =   32
         Tag             =   "wt0;fb0"
         Top             =   510
         Width           =   1575
      End
      Begin VB.CheckBox chkField 
         Caption         =   "Include Resigned Employees"
         Height          =   345
         Index           =   0
         Left            =   9660
         TabIndex        =   40
         Tag             =   "wt0;fb0"
         Top             =   510
         Width           =   2385
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
         Index           =   0
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   37
         Top             =   510
         Width           =   5730
      End
      Begin VB.ComboBox cmbSearch 
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
         ItemData        =   "frmEmp13thMonthRelease.frx":0000
         Left            =   9660
         List            =   "frmEmp13thMonthRelease.frx":000D
         TabIndex        =   31
         Top             =   105
         Width           =   2415
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
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   2
         Top             =   105
         Width           =   5730
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   17
         Left            =   495
         TabIndex        =   38
         Top             =   570
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   7920
         TabIndex        =   0
         Top             =   165
         Width           =   1305
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   25
         Left            =   855
         TabIndex        =   1
         Top             =   165
         Width           =   705
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   150
      TabIndex        =   29
      Top             =   4440
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmEmp13thMonthRelease.frx":003F
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   150
      TabIndex        =   28
      Top             =   5700
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmEmp13thMonthRelease.frx":07B9
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   5880
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   1605
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   10372
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   3615
         MaxLength       =   50
         TabIndex        =   43
         Text            =   "10,000.00"
         Top             =   4935
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "GMC Dagupan - Main"
         Top             =   1380
         Width           =   3930
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   3615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Apr 25, 2011"
         Top             =   1770
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   3615
         MaxLength       =   50
         TabIndex        =   18
         Text            =   "10,000.00"
         Top             =   4545
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   3615
         MaxLength       =   50
         TabIndex        =   16
         Text            =   "10,000.00"
         Top             =   4155
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   3615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "10,000.00"
         Top             =   3765
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "7, 200.00"
         Top             =   3150
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2760
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Junior Programmer"
         Top             =   990
         Width           =   3930
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Apr 25, 2011"
         Top             =   1770
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Cuison, Michael T."
         Top             =   600
         Width           =   3930
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "INACTIVE"
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
         Left            =   2910
         TabIndex        =   47
         Tag             =   "eb0;et0"
         Top             =   180
         Width           =   2400
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   0
         Left            =   2850
         Top             =   120
         Width           =   2520
      End
      Begin VB.Shape Shape4 
         Height          =   360
         Index           =   0
         Left            =   2880
         Top             =   150
         Width           =   2460
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LESS: Deduction"
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
         Index           =   5
         Left            =   1950
         TabIndex        =   44
         Top             =   4995
         Width           =   1515
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
         Index           =   16
         Left            =   105
         TabIndex        =   36
         Top             =   1440
         Width           =   615
      End
      Begin VB.Line Line2 
         X1              =   3450
         X2              =   3315
         Y1              =   1800
         Y2              =   2115
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "30,000.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3615
         TabIndex        =   20
         Top             =   5370
         Width           =   1755
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   2820
         TabIndex        =   19
         Top             =   5475
         Width           =   645
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CB"
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
         Index           =   9
         Left            =   3195
         TabIndex        =   17
         Top             =   4605
         Width           =   270
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus"
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
         Index           =   8
         Left            =   2910
         TabIndex        =   15
         Top             =   4215
         Width           =   555
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "13th Month Pay"
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
         Index           =   7
         Left            =   2085
         TabIndex        =   13
         Top             =   3825
         Width           =   1380
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Salary"
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
         Left            =   2220
         TabIndex        =   11
         Top             =   3210
         Width           =   1245
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Work Months Completed"
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
         Left            =   1305
         TabIndex        =   9
         Top             =   2820
         Width           =   2160
      End
      Begin VB.Line Line1 
         X1              =   450
         X2              =   4965
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hired/Resigned"
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
         Left            =   105
         TabIndex        =   7
         Top             =   1845
         Width           =   1320
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Descript"
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
         TabIndex        =   5
         Top             =   1050
         Width           =   1080
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
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
         TabIndex        =   3
         Top             =   660
         Width           =   870
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   645
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   7545
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   1138
      BackColor       =   12632256
      Enabled         =   0   'False
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
         Index           =   5
         Left            =   7935
         MaxLength       =   50
         TabIndex        =   45
         Top             =   135
         Width           =   1620
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
         Index           =   4
         Left            =   5505
         MaxLength       =   50
         TabIndex        =   25
         Top             =   135
         Width           =   1620
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
         Index           =   3
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   24
         Top             =   135
         Width           =   1620
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
         Index           =   6
         Left            =   10575
         MaxLength       =   50
         TabIndex        =   27
         Top             =   120
         Width           =   1620
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
         Index           =   2
         Left            =   930
         MaxLength       =   50
         TabIndex        =   22
         Top             =   120
         Width           =   1620
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dedtn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   18
         Left            =   7305
         TabIndex        =   46
         Top             =   210
         Width           =   495
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   4875
         TabIndex        =   30
         Top             =   210
         Width           =   255
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   2610
         TabIndex        =   23
         Top             =   210
         Width           =   540
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "13th Pay && Bonuses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   13
         Left            =   9645
         TabIndex        =   26
         Top             =   60
         Width           =   960
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "13th Pay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   45
         TabIndex        =   21
         Top             =   195
         Width           =   705
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   150
      TabIndex        =   33
      Top             =   5070
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Retreive"
      AccessKey       =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmp13thMonthRelease.frx":0F33
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   150
      TabIndex        =   39
      Top             =   3810
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Release"
      AccessKey       =   "Release"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmp13thMonthRelease.frx":16AD
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   150
      TabIndex        =   41
      Top             =   1290
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "BDO"
      AccessKey       =   "BDO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmp13thMonthRelease.frx":1E27
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   150
      TabIndex        =   42
      Top             =   1920
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "MB"
      AccessKey       =   "MB"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmp13thMonthRelease.frx":2701
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   150
      TabIndex        =   48
      Top             =   3195
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Others"
      AccessKey       =   "Others"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmp13thMonthRelease.frx":2FDB
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   150
      TabIndex        =   49
      Top             =   660
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "No Account"
      AccessKey       =   "No Account"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmp13thMonthRelease.frx":3755
   End
   Begin MSFlexGridLib.MSFlexGrid MSHFlexGrid1 
      Height          =   5910
      Left            =   7155
      TabIndex        =   50
      Top             =   1635
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   10425
      _Version        =   393216
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   150
      TabIndex        =   51
      Top             =   2550
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "SB"
      AccessKey       =   "SB"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmp13thMonthRelease.frx":3ECF
   End
End
Attribute VB_Name = "frmEmp13thMonthRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmEmp13thMonthRelease"
Private Const pxeRobinsons = "00XX044"
Private Const pxeMetrobank = "00XX006"
Private Const pxeChinabank = "00XX003"
Private Const pxeBDObank As String = "00XX024"
Private Const pxeSecBank As String = "00XX022"

Private oSkin As clsFormSkin
Private WithEvents oTrans As clsYearEndBonus
Attribute oTrans.VB_VarHelpID = -1
Private bLoaded As Boolean
Private p_orsDetail As Recordset

Private p_oPayMiscx As clsPayMisc

Dim pnIndex As Integer
Dim pnRow As Integer
Dim poCtrl As Variant

Private Sub chkField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 0
      oTrans.WithResigned = chkField(Index).Value = 1
   Case 1
      oTrans.ExcludeReleased = chkField(Index).Value = 1
   End Select
End Sub

Private Sub cmbSearch_Change(Index As Integer)
   oTrans.EmpLevel = cmbSearch(Index).ListIndex
End Sub

Private Sub cmbSearch_GotFocus(Index As Integer)
   Set poCtrl = cmbSearch(Index)
End Sub

Private Sub cmbSearch_Validate(Index As Integer, Cancel As Boolean)
   oTrans.EmpLevel = cmbSearch(Index).ListIndex
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRow As Integer
   Dim lnRep As Integer
   Dim loFmx As frmChinaCMS

   lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc

   Select Case Index
   Case 0 'retrieve
      Call oTrans.LoadDetail
      Call LoadDetail
   Case 1 'search
      If poCtrl.Name = "txtSearch" Then
         Select Case poCtrl.Index
         Case 1
            Call oTrans.getBranch(txtSearch(1), False, True)
            txtSearch(1) = oTrans.Branch
         Case 0
            Call oTrans.getDepartment(txtSearch(0), False, True)
            txtSearch(0) = oTrans.Department
         End Select
      End If
   Case 2
      Unload Me
   Case 3
      If Trim(txtField(2).Text) <> "" Then
         Call oTrans.ReleaseBonus(oTrans.Detail(MSHFlexGrid1.Row - 1, "sEmployID"))
      End If
   Case 4
'      If Trim(MSHFlexGrid1.TextMatrix(1, 1)) <> "" Then
'         Call exportCMSRobinson
'         If MsgBox("Issue Release of 13th Month Pay and Bonus for employees" & vbCrLf & _
'                   "with Robinsons account?" & vbCrLf & vbCrLf & _
'                   "Please click [Yes] to continue...", vbYesNo, "Confirm") = vbYes Then
'            If ReleaseAccount(pxeRobinsons) Then
'               Call oTrans.LoadDetail
'               Call LoadDetail
'            End If
'         End If
'      End If
      If Trim(MSHFlexGrid1.TextMatrix(1, 1)) <> "" Then
         'Call ExportBDO(MSHFlexGrid1)
         Call ExportBDO
         If MsgBox("Issue Release of 13th Month Pay and Bonus for employees" & vbCrLf & _
                   "with BDO account?" & vbCrLf & vbCrLf & _
                   "Please click [Yes] to continue...", vbYesNo, "Confirm") = vbYes Then
            If ReleaseAccount(pxeBDObank) Then
               Call oTrans.LoadDetail
               Call LoadDetail
            End If
         End If
      End If
   Case 5
      If Trim(MSHFlexGrid1.TextMatrix(1, 1)) <> "" Then
         'Call ExportMB(MSHFlexGrid1)
         Call ExportMB
         If MsgBox("Issue Release of 13th Month Pay and Bonus for employees" & vbCrLf & _
                   "with Metrobank account?" & vbCrLf & vbCrLf & _
                   "Please click [Yes] to continue...", vbYesNo, "Confirm") = vbYes Then
            If ReleaseAccount(pxeMetrobank) Then
               Call oTrans.LoadDetail
               Call LoadDetail
            End If
         End If
      End If
   Case 6
      If Trim(MSHFlexGrid1.TextMatrix(1, 1)) <> "" Then
        Call initChina
        Set loFmx = New frmChinaCMS
        loFmx.txtField(0) = oTrans.Master("sPayPerID")
        loFmx.txtField(1) = Format(oTrans.Master("dCovergTo"), "Mmmm DD, YYYY")
        Call ExportCNB
        Call ExportChina
        Set loFmx.Data = p_orsDetail
        loFmx.Show 1

         If MsgBox("Issue Release of 13th Month Pay and Bonus for employees" & vbCrLf & _
                   "with Chinabank account?" & vbCrLf & vbCrLf & _
                   "Please click [Yes] to continue...", vbYesNo, "Confirm") = vbYes Then
            If ReleaseAccount(pxeChinabank) Then
               Call oTrans.LoadDetail
               Call LoadDetail
            End If
         End If
      End If
   Case 7
      If Trim(MSHFlexGrid1.TextMatrix(1, 1)) <> "" Then
         Call ExportNoAccount
         'Call ExportNoAccount(MSHFlexGrid1)
         If MsgBox("Issue Release of 13th Month Pay and Bonus for employees" & vbCrLf & _
                   "with NO ACCOUNT?" & vbCrLf & vbCrLf & _
                   "Please click [Yes] to continue...", vbYesNo, "Confirm") = vbYes Then
            If ReleaseAccount("") Then
               Call oTrans.LoadDetail
               Call LoadDetail
            End If
         End If
      End If
   Case 8 'SB
      If Trim(MSHFlexGrid1.TextMatrix(1, 1)) <> "" Then
         'Call ExportSB(MSHFlexGrid1)
         Call ExportSB
         If MsgBox("Issue Release of 13th Month Pay and Bonus for employees" & vbCrLf & _
                   "with Security Bank account?" & vbCrLf & vbCrLf & _
                   "Please click [Yes] to continue...", vbYesNo, "Confirm") = vbYes Then
            If ReleaseAccount(pxeSecBank) Then
               Call oTrans.LoadDetail
               Call LoadDetail
            End If
         End If
      End If
   
   End Select
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
'   'On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      If LCase(oApp.ProductID) <> "petmgr" Then
         txtSearch(0).Text = ""
         txtSearch(1).Text = oApp.BranchName

         txtSearch(0).Locked = True
         txtSearch(1).Locked = True
         cmbSearch(0).Locked = True

'         cmdButton(4).Visible = False
'         cmdButton(5).Visible = False
'         cmdButton(6).Visible = False
      End If
'
'      If oApp.BranchCode = "N001" Then
'         cmdButton(5).Visible = False
'         cmdButton(6).Visible = True
'      Else
'         cmdButton(5).Visible = True
'         cmdButton(6).Visible = False
'      End If

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
'   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsYearEndBonus
   Set oTrans.AppDriver = oApp
   oTrans.TransStatus = 2

   oTrans.InitTransaction
   oTrans.SearchTransaction "%"

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

   Set p_oPayMiscx = New clsPayMisc
   Set p_oPayMiscx.AppDriver = oApp

   Call InitGrid
   Call ClearFields(-1)

   Call oTrans.LoadDetail
   Call LoadDetail

   If oApp.ProductID <> "PetMgr" Then
      xrFrame2.Enabled = False
'      chkField(1).Visible = False
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub ClearFields(ByVal fnRow As Integer)
   Dim loTxt As TextBox
   Dim lnTotal As Currency

   txtSearch(1) = oTrans.Branch
   txtSearch(0) = oTrans.Department
   cmbSearch(0).ListIndex = oTrans.EmpLevel

   If fnRow < 0 Then
      For Each loTxt In txtField
         loTxt = ""
      Next

      lblTotal = ""
   Else
      For Each loTxt In txtField
         Select Case loTxt.Index
         Case 6, 7, 9, 10, 11
            loTxt = Format(oTrans.Detail(fnRow, loTxt.Index), "#,##0.00")
         Case 19, 20
            loTxt = IFNull(oTrans.Detail(fnRow, loTxt.Index), "")
         Case Else
            loTxt = IFNull(oTrans.Detail(fnRow, loTxt.Index), "")
         End Select
      Next

      lblTotal = Format(CCur(txtField(7)) + CCur(txtField(9)) + CCur(txtField(10)) - CCur(txtField(11)), "#,##0.00")
   End If

   Select Case oTrans.Detail(fnRow, "cReleased")
   Case 1
      Label2.Caption = "CONFIRMED"
   Case 2
      Label2.Caption = "RELEASED"
   Case Else
      Label2.Caption = "OPEN"
   End Select

End Sub

Private Sub MSHFlexGrid1_Click()
   xrFrame3.Enabled = False
   Call ClearFields(MSHFlexGrid1.Row - 1)
End Sub

Private Sub MSHFlexGrid1_SelChange()
   With MSHFlexGrid1
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSHFlexGrid1
      .Cols = 12
      .Rows = 2

      .Clear
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next

      .Row = 1

      'column title
      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = "Employee"
      .TextMatrix(0, 2) = "Designation"
      .TextMatrix(0, 3) = "Months"
      .TextMatrix(0, 4) = "13th Pay"
      .TextMatrix(0, 5) = "Bonus"
      .TextMatrix(0, 6) = "XMas"
      .TextMatrix(0, 7) = "Deduction"
      .TextMatrix(0, 8) = "Total"
      .TextMatrix(0, 9) = "Account"
      .TextMatrix(0, 10) = "BnkID"
      .TextMatrix(0, 11) = "Branch Name"

      'column width
      .ColWidth(0) = 590
      .ColWidth(1) = 2800
      .ColWidth(2) = 2300
      .ColWidth(3) = 800
      .ColWidth(4) = 1300
      .ColWidth(5) = 1300
      .ColWidth(6) = 1300
      .ColWidth(7) = 1300
      .ColWidth(8) = 1300
      .ColWidth(9) = 0
      .ColWidth(10) = 0
      .ColWidth(11) = 0

      'column allinment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      .ColAlignment(6) = flexAlignRightCenter
      .ColAlignment(7) = flexAlignRightCenter
      .ColAlignment(8) = flexAlignRightCenter

      'set location
      .Row = 1
      .Col = 0
      .ColSel = .Cols - 1

      pnRow = 0
   End With
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   Dim ln13thMnth As Currency
   Dim lnBonusxx1 As Currency
   Dim lnPartyxx1 As Currency
   Dim lnDeductnx As Currency

   With MSHFlexGrid1
      .Rows = 2
      If oTrans.ItemCount = 0 Then
         .TextMatrix(1, 0) = 1
         .TextMatrix(1, 1) = ""
         .TextMatrix(1, 2) = ""
         .TextMatrix(1, 3) = 0
         .TextMatrix(1, 4) = "0.00"
         .TextMatrix(1, 5) = "0.00"
         .TextMatrix(1, 6) = "0.00"
         .TextMatrix(1, 7) = "0.00"
         .TextMatrix(1, 8) = "0.00"
         .TextMatrix(1, 9) = ""
      Else
         .Rows = oTrans.ItemCount + 1
         For lnCtr = 0 To oTrans.ItemCount - 1
            DoEvents
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "xFullName")
            .TextMatrix(lnCtr + 1, 2) = IFNull(oTrans.Detail(lnCtr, "sPositnNm"))
            .TextMatrix(lnCtr + 1, 3) = Format(oTrans.Detail(lnCtr, "nAttend01"), "#0.00")
            .TextMatrix(lnCtr + 1, 4) = Format(oTrans.Detail(lnCtr, "n13thMnth"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 5) = Format(oTrans.Detail(lnCtr, "nBonusxx1"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 6) = Format(oTrans.Detail(lnCtr, "nPartyxx1"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 7) = Format(oTrans.Detail(lnCtr, "nDeductnx"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 8) = Format(CCur(.TextMatrix(lnCtr + 1, 4)) + CCur(.TextMatrix(lnCtr + 1, 5)) + CCur(.TextMatrix(lnCtr + 1, 6)) - CCur(.TextMatrix(lnCtr + 1, 7)), "#,##0.00")
            .TextMatrix(lnCtr + 1, 9) = IFNull(oTrans.Detail(lnCtr, "sBnkActno"))
            .TextMatrix(lnCtr + 1, 10) = IFNull(oTrans.Detail(lnCtr, "sBankIDxx"))
            .TextMatrix(lnCtr + 1, 11) = IFNull(oTrans.Detail(lnCtr, "sBranchNm"))

            ln13thMnth = ln13thMnth + CCur(.TextMatrix(lnCtr + 1, 4))
            lnBonusxx1 = lnBonusxx1 + CCur(.TextMatrix(lnCtr + 1, 5))
            lnPartyxx1 = lnPartyxx1 + CCur(.TextMatrix(lnCtr + 1, 6))
            lnDeductnx = lnDeductnx + CCur(.TextMatrix(lnCtr + 1, 7))
         Next
      End If
   End With

   txtSearch(2) = Format(ln13thMnth, "#,##0.00")
   txtSearch(3) = Format(lnBonusxx1, "#,##0.00")
   txtSearch(4) = Format(lnPartyxx1, "#,##0.00")
   txtSearch(5) = Format(lnDeductnx, "#,##0.00")
   txtSearch(6) = Format(ln13thMnth + lnBonusxx1 + lnPartyxx1 - lnDeductnx, "#,##0.00")

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

Private Sub txtSearch_GotFocus(Index As Integer)
   Set poCtrl = txtSearch(Index)
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 1 Then
      Select Case KeyCode
      Case vbKeyF3
         Call oTrans.getBranch(txtSearch(Index), False, True)
         txtSearch(Index) = oTrans.Branch
      Case vbKeyReturn
         If txtSearch(Index) <> "" Then
            Call oTrans.getBranch(txtSearch(Index), False, True)
         Else
            Call oTrans.getBranch(txtSearch(Index), False, False)
         End If
         txtSearch(Index) = oTrans.Branch
      End Select
   ElseIf Index = 0 Then
      Select Case KeyCode
      Case vbKeyF3
         Call oTrans.getDepartment(txtSearch(Index), False, True)
         txtSearch(Index) = oTrans.Department
      Case vbKeyReturn
         If txtSearch(Index) <> "" Then
            Call oTrans.getDepartment(txtSearch(Index), False, True)
         Else
            Call oTrans.getDepartment(txtSearch(Index), False, False)
         End If
         txtSearch(Index) = oTrans.Department
      End Select
   End If
End Sub

Private Sub txtSearch_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 1
      Call oTrans.getBranch(txtSearch(Index), False, False)
      txtSearch(Index) = oTrans.Branch
   Case 0
      Call oTrans.getDepartment(txtSearch(Index), False, False)
      txtSearch(Index) = oTrans.Department
   End Select
End Sub

Private Sub exportCMSRobinson()
   Dim loCls As clsRobinsonCMS
   Dim lnCtr As Integer
   Dim lnDta As Integer
   Dim loFrm As frmRobinsonsCMS
   Set loCls = New clsRobinsonCMS
   Set loCls.AppDriver = oApp
   loCls.isBonus = True
   loCls.InitTransaction
   loCls.NewTransaction
   loCls.Master("dTransact") = oApp.ServerDate
   loCls.Master("sPIRRefNo") = oTrans.Master("sPeriodID")

   For lnCtr = 1 To MSHFlexGrid1.Rows - 1
      If MSHFlexGrid1.TextMatrix(lnCtr, 10) = pxeRobinsons And CCur(MSHFlexGrid1.TextMatrix(lnCtr, 8)) > 0 Then
         loCls.addDetail
         loCls.Detail(loCls.ItemCount - 1, "sPayeAcct") = MSHFlexGrid1.TextMatrix(lnCtr, 9)
         loCls.Detail(loCls.ItemCount - 1, "sPayeName") = MSHFlexGrid1.TextMatrix(lnCtr, 1)
         loCls.Detail(loCls.ItemCount - 1, "nTranAmtx") = CCur(MSHFlexGrid1.TextMatrix(lnCtr, 8))
      End If
   Next

   Set loFrm = New frmRobinsonsCMS
   Set loFrm.CMS = loCls

   loFrm.Show 1
   If loFrm.IsOkey Then
      Call loCls.SaveTransaction(True)
      loCls.exportCMS
      MsgBox "Export of 13th Pay and Bonus successfully done!", vbInformation + vbOKOnly, "Success"
   End If

   Unload loFrm
End Sub

'Private Function ExportMB(MSHFlexGrid1 As MSHFlexGrid) As Boolean
Private Function ExportMB() As Boolean

   Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook
   
   Dim lsSQL As String
   Dim lors As Recordset
   
   Dim XLastName As String
   Dim XFrstName As String
   Dim XMiddName As String
   Dim XAcctNmbr As String
   Dim XNetPayxx As String
   
   Dim sFileName As String
   
   Dim lnLineCtr As Integer
   Dim lnItemCtr As Integer
   
   XLastName = "A"
   XFrstName = "B"
   XMiddName = "C"
   XAcctNmbr = "D"
   XNetPayxx = "E"
   
   'lnLineCtr = 6
   lnLineCtr = 1
   
   sFileName = "MetroBank Payroll"
    
   Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & sFileName & ".XLS")
   
   Set xlsheet = xlwbook.Sheets.Item(1)
   xlsheet.Range(XLastName & (lnLineCtr + 1) & ":" & XNetPayxx & 65536).ClearContents
    
   With MSHFlexGrid1
   
     For lnItemCtr = 1 To MSHFlexGrid1.Rows - 1
     If .TextMatrix(lnItemCtr, 10) = "00XX006" And CCur(.TextMatrix(lnItemCtr, 8)) > 0 And Trim(.TextMatrix(lnItemCtr, 9)) <> "" Then
            lnLineCtr = lnLineCtr + 1
'         If .TextMatrix(lnItemCtr, 10) = pxeMetrobank And CCur(.TextMatrix(lnItemCtr, 8)) > 0 And Trim(.TextMatrix(lnItemCtr, 9)) <> "" Then
'            lnLineCtr = lnLineCtr + 1
            'she 2020-12-21 modify for the template
            lsSQL = "SELECT b.sLastName, b.sFrstName, b.sMiddName" & _
                    " FROM Employee_Master001 a, Client_Master b" & _
                    " WHERE a.sEmployID = b.sClientID" & _
                    " AND b.sCompnyNm LIKE " & strParm(.TextMatrix(lnItemCtr, 1) + "%")
            Set lors = New Recordset
            lors.Open lsSQL, oApp.Connection, , , adCmdText
            If lors.EOF Then
                xlsheet.Range(XLastName & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
                xlsheet.Range(XFrstName & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
                xlsheet.Range(XMiddName & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
            Else
                xlsheet.Range(XLastName & (lnLineCtr)).Value = lors("sLastName")
                xlsheet.Range(XFrstName & (lnLineCtr)).Value = lors("sFrstName")
                xlsheet.Range(XMiddName & (lnLineCtr)).Value = lors("sMiddName")
            End If
            xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = Mid(.TextMatrix(lnItemCtr, 9), 4) ' .TextMatrix(lnItemCtr, 17)
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = CCur(.TextMatrix(lnItemCtr, 8)) 'CCur(.TextMatrix(lnItemCtr, 16)) - CCur(.TextMatrix(lnItemCtr, 19))
         End If
     Next
'     MsgBox xlsheet.Range("E" & Rows.Count).End(xlUp).Row
         
     'kalyptus-2013.08.19 02:03pm
     'delete remaining items in our template if ever template was updated...
     lnItemCtr = xlsheet.Range("E" & xlsheet.Rows.Count).End(xlUp).Row
     If lnItemCtr > lnLineCtr Then
         xlsheet.Range("A" & lnLineCtr + 1 & ":E" & lnItemCtr).Clear
     End If
   
   End With
    
   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & oTrans.Master(0) & ").XLS"
   xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
   xl.Quit
    
   Set xlwbook = Nothing
   Set xl = Nothing

'   Dim xl As New Excel.Application
'   Dim xlsheet As Excel.Worksheet
'   Dim xlwbook As Excel.Workbook
'
'   Dim lsSQL As String
'   Dim lors As Recordset
'
'   Dim XEmployID As String
'   Dim XEmployNm As String
'   Dim XBranchCD As String
'   Dim XAcctNmbr As String
'   Dim XNetPayxx As String
'
'   Dim sFileName As String
'
'   Dim lnLineCtr As Integer
'   Dim lnItemCtr As Integer
'
'   XEmployID = "A"
'   XEmployNm = "B"
'   XBranchCD = "C"
'   XAcctNmbr = "D"
'   XNetPayxx = "E"
'
'
'   lnLineCtr = 6
'
'   sFileName = "MetroBank Payroll"
'
'   Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & sFileName & ".XLS")
'   Set xlsheet = xlwbook.Sheets.Item(1)
'
'   'With MSHFlexGrid1
'   With MSHFlexGrid1
'     For lnItemCtr = 1 To MSHFlexGrid1.Rows - 1
'         If .TextMatrix(lnItemCtr, 10) = pxeMetrobank And CCur(.TextMatrix(lnItemCtr, 8)) > 0 And Trim(.TextMatrix(lnItemCtr, 9)) <> "" Then
'            lnLineCtr = lnLineCtr + 1
'            xlsheet.Range(XEmployID & (lnLineCtr)).Value = lnLineCtr - 6
''            xlsheet.Range(XEmployNm & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
'            xlsheet.Range(XBranchCD & (lnLineCtr)).Value = Left(.TextMatrix(lnItemCtr, 9), 3)
'            xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = Mid(.TextMatrix(lnItemCtr, 9), 4)
'            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = CCur(.TextMatrix(lnItemCtr, 8))
'         End If
'     Next
''     MsgBox xlsheet.Range("E" & Rows.Count).End(xlUp).Row
'
'     'kalyptus-2013.08.19 02:03pm
'     'delete remaining items in our template if ever template was updated...
'     lnItemCtr = xlsheet.Range("E" & xlsheet.Rows.Count).End(xlUp).Row
'     If lnItemCtr > lnLineCtr Then
'         xlsheet.Range("A" & lnLineCtr + 1 & ":E" & lnItemCtr).Clear
'     End If
'
'   End With
'
'   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & oTrans.Master(0) & ").XLS"
'   xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
'   xl.Quit
'
'   Set xlwbook = Nothing
'   Set xl = Nothing
End Function

'kalyptus - 2014.05.20 10:00am
'Indicate update of the release
Private Function ReleaseAccount(ByVal fsBankCode As String) As Boolean
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lsOldProc As String

   lsOldProc = "ReleaseAccount"
'   'On Error GoTo errProc

   If Not p_oPayMiscx.Makanengneng Then
      MsgBox "This user is not authorized to release the 13th month pay and bonus(es)!" & _
             "For inquiry please call Guanzon MIS SEG/SSG...", vbCritical + vbOKOnly, "Warning"
      GoTo endProc
   End If

'   If oTrans.Master("cRecdStat") = "0" Then
'      MsgBox "13th Month Pay was not yet processed!" & _
'             "Can't release 13th month pay and bonus(es)...", vbCritical + vbOKOnly, "Warning"
'      GoTo endProc
   If oTrans.Master("cRecdStat") = "0" Then
      MsgBox "13th Month Pay was not yet finalized!" & _
             "Can't release 13th month pay and bonus(es)...", vbCritical + vbOKOnly, "Warning"
      GoTo endProc
   End If

   oApp.BeginTrans
   With MSHFlexGrid1
      For lnCtr = 0 To oTrans.ItemCount - 1
         If fsBankCode = "" Then
            If .TextMatrix(lnCtr + 1, 10) = fsBankCode And _
               CCur(.TextMatrix(lnCtr + 1, 8)) > 0 Then

               If oTrans.Detail(lnCtr, "cReleased") = "0" Then
                  MsgBox "Bonus was not yet finalized for " & .TextMatrix(lnCtr + 1, 1) & "!" & vbCrLf & _
                         "Can't release 13th month pay and bonus(es)...", vbCritical + vbOKOnly, "Warning"
                  GoTo endProcWithRoll
               ElseIf oTrans.Detail(lnCtr, "cReleased") = "2" Then
                  MsgBox "13th month pay and bonus(es) was previously released for " & .TextMatrix(lnCtr + 1, 1) & "!" & vbCrLf & _
                         "For inquiry please call Guanzon MIS SEG/SSG...", vbCritical + vbOKOnly, "Warning"
                  GoTo endProcWithRoll
               End If

               'indicate that 13th Month Pay/Bonus was released...
               lsSQL = "UPDATE Payroll_Annual_Total" & _
                      " SET cReleased = '2'" & _
                      " WHERE sPeriodID = " & strParm(oTrans.Master("sPeriodID")) & _
                        " AND sEmployID = " & strParm(oTrans.Detail(lnCtr, "sEmployID"))
                oApp.Execute lsSQL, "Payroll_Annual_Total", oTrans.Detail(lnCtr, "sBranchCD")
            End If
         Else
            If .TextMatrix(lnCtr + 1, 10) = fsBankCode And _
               CCur(.TextMatrix(lnCtr + 1, 8)) > 0 And _
               Trim(.TextMatrix(lnCtr + 1, 9)) <> "" Then

               If oTrans.Detail(lnCtr, "cReleased") = "0" Then
                  MsgBox "Bonus was not yet finalized for " & .TextMatrix(lnCtr + 1, 1) & "!" & vbCrLf & _
                         "Can't release 13th month pay and bonus(es)...", vbCritical + vbOKOnly, "Warning"
                  GoTo endProcWithRoll
               ElseIf oTrans.Detail(lnCtr, "cReleased") = "2" Then
                  MsgBox "13th month pay and bonus(es) was previously released for " & .TextMatrix(lnCtr + 1, 1) & "!" & vbCrLf & _
                         "For inquiry please call Guanzon MIS SEG/SSG...", vbCritical + vbOKOnly, "Warning"
                  GoTo endProcWithRoll
               End If

               'indicate that 13th Month Pay/Bonus was released...
               lsSQL = "UPDATE Payroll_Annual_Total" & _
                      " SET cReleased = '2'" & _
                      " WHERE sPeriodID = " & strParm(oTrans.Master("sPeriodID")) & _
                        " AND sEmployID = " & strParm(oTrans.Detail(lnCtr, "sEmployID"))
                oApp.Execute lsSQL, "Payroll_Annual_Total", oTrans.Detail(lnCtr, "sBranchCD")
            End If
         End If
      Next

   End With
   oApp.CommitTrans

   ReleaseAccount = True
endProc:
   Exit Function
endProcWithRoll:
   oApp.RollbackTrans
   GoTo endProc
errProc:
   oApp.RollbackTrans
   ShowError lsOldProc & "( " & " )", True
End Function

Private Sub initChina()
   Set p_orsDetail = New Recordset
   With p_orsDetail
      .Fields.Append "sBankIDxx", adChar, 8
      .Fields.Append "sEmployNm", adVarChar, 100
      .Fields.Append "sBnkActNo", adVarChar, 50
      .Fields.Append "nNetPayxx", adCurrency
      .Open
8   End With
End Sub

Private Sub ExportChina()
   Dim lnItemCtr As Integer
   With MSHFlexGrid1
     For lnItemCtr = 1 To .Rows - 1
         If .TextMatrix(lnItemCtr, 10) = pxeChinabank And CCur(.TextMatrix(lnItemCtr, 8)) > 0 And Trim(.TextMatrix(lnItemCtr, 9)) <> "" Then
            p_orsDetail.AddNew
            p_orsDetail("sBankIDxx") = .TextMatrix(lnItemCtr, 10)
            p_orsDetail("sEmployNm") = .TextMatrix(lnItemCtr, 1)
            p_orsDetail("sBnkActNo") = .TextMatrix(lnItemCtr, 9)
            p_orsDetail("nNetPayxx") = CCur(.TextMatrix(lnItemCtr, 8))
         End If
     Next
  End With
End Sub

'Private Function ExportNoAccount(MSHFlexGrid1 As MSHFlexGrid) As Boolean
Private Function ExportNoAccount() As Boolean
   Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook

   Dim lsSQL As String
   Dim lors As Recordset

   Dim XEmployID As String
   Dim XEmployNm As String
   Dim XBranchCD As String
   Dim XAcctNmbr As String
   Dim XNetPayxx As String

   Dim sFileName As String

   Dim lnLineCtr As Integer
   Dim lnItemCtr As Integer

   XEmployID = "A"
   XEmployNm = "B"
   XBranchCD = "C"
   XAcctNmbr = "D"
   XNetPayxx = "E"

   lnLineCtr = 6

   sFileName = "MetroBank Payroll"

   Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & sFileName & ".XLS")
   Set xlsheet = xlwbook.Sheets.Item(1)

   With MSHFlexGrid1
   'With MSHFlexGrid1

     For lnItemCtr = 1 To MSHFlexGrid1.Rows - 1
         If .TextMatrix(lnItemCtr, 10) = "" Then
            If IsNumeric(.TextMatrix(lnItemCtr, 8)) Then
               If CCur(.TextMatrix(lnItemCtr, 8)) > 0 Then
                  lnLineCtr = lnLineCtr + 1
                  xlsheet.Range(XEmployID & (lnLineCtr)).Value = lnLineCtr - 6
                  xlsheet.Range(XEmployNm & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
                  xlsheet.Range(XBranchCD & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 11)
                  xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = Mid(.TextMatrix(lnItemCtr, 9), 4)
                  xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = CCur(.TextMatrix(lnItemCtr, 8))
               End If
            End If
         End If
     Next
'     MsgBox xlsheet.Range("E" & Rows.Count).End(xlUp).Row

     'kalyptus-2013.08.19 02:03pm
     'delete remaining items in our template if ever template was updated...
     lnItemCtr = xlsheet.Range("E" & xlsheet.Rows.Count).End(xlUp).Row
     If lnItemCtr > lnLineCtr Then
         xlsheet.Range("A" & lnLineCtr + 1 & ":E" & lnItemCtr).Clear
     End If

   End With

   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & oTrans.Master(0) & "-NOACCOUNT).XLS"
   xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
   xl.Quit

   Set xlwbook = Nothing
   Set xl = Nothing
End Function

'Private Function ExportBDO(MSHFlexGrid1 As MSHFlexGrid) As Boolean
Private Function ExportBDO() As Boolean
   Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook

   Dim lsSQL As String
   Dim lors As Recordset

   Dim XEmployNm As String
   Dim XRemarksx As String
   Dim XAcctNmbr As String
   Dim XNetPayxx As String

   Dim sFileName As String

   Dim lnLineCtr As Integer
   Dim lnItemCtr As Integer

   XAcctNmbr = "A"
   XNetPayxx = "B"
   XEmployNm = "C"
   XRemarksx = "D"

   lnLineCtr = 6

   sFileName = "BDO ATM Payroll Converter for BOB"

   Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & sFileName & ".XLS")

   Set xlsheet = xlwbook.Sheets.Item(1)

   'With MSHFlexGrid1
   With MSHFlexGrid1
     For lnItemCtr = 1 To MSHFlexGrid1.Rows - 1
         If .TextMatrix(lnItemCtr, 10) = pxeBDObank And CCur(.TextMatrix(lnItemCtr, 8)) > 0 And Trim(.TextMatrix(lnItemCtr, 9)) <> "" Then
            lnLineCtr = lnLineCtr + 1
            xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = "'" & .TextMatrix(lnItemCtr, 9)
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = CCur(.TextMatrix(lnItemCtr, 8))
            xlsheet.Range(XEmployNm & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
            xlsheet.Range(XRemarksx & (lnLineCtr)).Value = ""
         End If
     Next
   End With

   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & oTrans.Master(0) & ").XLS"
   xl.ActiveWorkbook.Close
   xl.Quit

   Set xlwbook = Nothing
   Set xl = Nothing
End Function

'Private Function ExportSB(MSHFlexGrid1 As MSHFlexGrid) As Boolean
Private Function ExportSB() As Boolean
   Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook

   Dim lsSQL As String
   Dim lors As Recordset

   Dim XEmployID As String
   Dim XEmployNm As String
   Dim XBranchCD As String
   Dim XAcctNmbr As String
   Dim XNetPayxx As String

   Dim sFileName As String

   Dim lnLineCtr As Integer
   Dim lnItemCtr As Integer

   XEmployNm = "A"
   XAcctNmbr = "B"
   XNetPayxx = "C"

   lnLineCtr = 8

   sFileName = "SBC_payroll_ver_1.7"

   Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & sFileName & ".XLS")

   Set xlsheet = xlwbook.Sheets.Item(1)
   xlsheet.Range(XEmployNm & (lnLineCtr + 1) & ":" & XNetPayxx & 65536).ClearContents

   'With MSHFlexGrid1
   With MSHFlexGrid1
     For lnItemCtr = 1 To MSHFlexGrid1.Rows - 1
         If .TextMatrix(lnItemCtr, 10) = pxeSecBank And CCur(.TextMatrix(lnItemCtr, 8)) > 0 And Trim(.TextMatrix(lnItemCtr, 9)) <> "" Then
            lnLineCtr = lnLineCtr + 1
            xlsheet.Range(XEmployNm & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
            xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 9)
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = CCur(.TextMatrix(lnItemCtr, 8))
         End If
     Next

     'kalyptus-2013.08.19 02:03pm
     'delete remaining items in our template if ever template was updated...
     lnItemCtr = xlsheet.Range("C" & xlsheet.Rows.Count).End(xlUp).Row
     If lnItemCtr > lnLineCtr Then
         xlsheet.Range("A" & lnLineCtr + 1 & ":C" & lnItemCtr).Clear
     End If

   End With

   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & oTrans.Master(0) & ").XLS"
   xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
   xl.Quit

   Set xlwbook = Nothing
   Set xl = Nothing
End Function

Private Function ExportCNB()

    Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook
   
   Dim lsSQL As String
   Dim lors As Recordset
   
   Dim XLastName As String
   Dim XFrstName As String
   Dim XMiddName As String
   Dim XAcctNmbr As String
   Dim XNetPayxx As String
   Dim xMobileNo As String
   Dim xEmailAdd As String
   
   Dim sFileName As String
   
   Dim lnLineCtr As Integer
   Dim lnItemCtr As Integer
   
   XAcctNmbr = "A"
   XLastName = "B"
   XFrstName = "C"
   XMiddName = "D"
   XNetPayxx = "E"
   xMobileNo = "F"
   xEmailAdd = "G"
   
   'lnLineCtr = 6
   lnLineCtr = 1
   
   sFileName = "CNBSpecial"
    
   Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & sFileName & ".XLS")
   
   Set xlsheet = xlwbook.Sheets.Item(1)
   xlsheet.Range(XLastName & (lnLineCtr + 1) & ":" & XNetPayxx & 65536).ClearContents
    
   With MSHFlexGrid1
     For lnItemCtr = 1 To MSHFlexGrid1.Rows - 1
         If .TextMatrix(lnItemCtr, 10) = "00XX003" And CCur(.TextMatrix(lnItemCtr, 8)) > 0 And Trim(.TextMatrix(lnItemCtr, 9)) <> "" Then
            lnLineCtr = lnLineCtr + 1
            'she 2020-10-02 modify for the template
            lsSQL = "SELECT b.sLastName, b.sFrstName, b.sMiddName, b.sEmailAdd, b.sMobileNo" & _
                    " FROM Employee_Master001 a, Client_Master b" & _
                    " WHERE a.sEmployID = b.sClientID" & _
                    " AND b.sCompnyNm LIKE " & strParm(.TextMatrix(lnItemCtr, 1) + "%")
            Set lors = New Recordset
            lors.Open lsSQL, oApp.Connection, , , adCmdText
            
            If lors.EOF Then
                xlsheet.Range(XLastName & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
                xlsheet.Range(XFrstName & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
                xlsheet.Range(XMiddName & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
                xlsheet.Range(xMobileNo & (lnLineCtr)).Value = ""
                xlsheet.Range(xEmailAdd & (lnLineCtr)).Value = ""
            Else
                xlsheet.Range(XLastName & (lnLineCtr)).Value = lors("sLastName")
                xlsheet.Range(XFrstName & (lnLineCtr)).Value = lors("sFrstName")
                xlsheet.Range(XMiddName & (lnLineCtr)).Value = lors("sMiddName")
                xlsheet.Range(xMobileNo & (lnLineCtr)).Value = lors("sMobileNo")
                xlsheet.Range(xEmailAdd & (lnLineCtr)).Value = lors("sEmailAdd")
            End If
            
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = CCur(.TextMatrix(lnItemCtr, 8))
            xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 9)
            
            
            
         End If
     Next
'     MsgBox xlsheet.Range("E" & Rows.Count).End(xlUp).Row
         
     'kalyptus-2013.08.19 02:03pm
     'delete remaining items in our template if ever template was updated...
     lnItemCtr = xlsheet.Range("E" & xlsheet.Rows.Count).End(xlUp).Row
     If lnItemCtr > lnLineCtr Then
         xlsheet.Range("A" & lnLineCtr + 1 & ":E" & lnItemCtr).Clear
     End If
   
   End With
    
   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & Format(oApp.ServerDate, "YYYYMMDD") & ").XLS"
   xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
   xl.Quit
    
   Set xlwbook = Nothing
   Set xl = Nothing
End Function

