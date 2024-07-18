VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmEmp13thMonth 
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
   Begin MSFlexGridLib.MSFlexGrid MSHFlexGrid1 
      Height          =   5910
      Left            =   7155
      TabIndex        =   34
      Top             =   1590
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   10425
      _Version        =   393216
   End
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
         TabIndex        =   39
         Top             =   510
         Width           =   5325
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
         ItemData        =   "frmEmp13thMonth.frx":0000
         Left            =   9165
         List            =   "frmEmp13thMonth.frx":000D
         TabIndex        =   33
         Top             =   105
         Width           =   3000
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
         Width           =   5325
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
         TabIndex        =   40
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
         Left            =   7590
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
      Left            =   135
      TabIndex        =   31
      Top             =   2250
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
      Picture         =   "frmEmp13thMonth.frx":003F
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   135
      TabIndex        =   30
      Top             =   4140
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
      Picture         =   "frmEmp13thMonth.frx":07B9
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   5880
      Left            =   1605
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
         Enabled         =   0   'False
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   46
         Text            =   "10,000.00"
         Top             =   4950
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "GMC Dagupan - Main"
         Top             =   1395
         Width           =   3930
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "Apr 25, 2011"
         Top             =   1785
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
         Index           =   4
         Left            =   3615
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "0"
         Top             =   2385
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
         TabIndex        =   20
         Text            =   "10,000.00"
         Top             =   4560
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
         TabIndex        =   18
         Text            =   "10,000.00"
         Top             =   4170
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "10,000.00"
         Top             =   3780
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "7, 200.00"
         Top             =   3165
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2775
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Top             =   1005
         Width           =   3930
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Top             =   1785
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Top             =   615
         Width           =   3930
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
         Index           =   6
         Left            =   1950
         TabIndex        =   47
         Top             =   5010
         Width           =   1515
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
         TabIndex        =   45
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
         TabIndex        =   38
         Top             =   1455
         Width           =   615
      End
      Begin VB.Line Line2 
         X1              =   3450
         X2              =   3315
         Y1              =   1815
         Y2              =   2130
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MVP Rank*"
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
         Left            =   2520
         TabIndex        =   9
         Top             =   2445
         Width           =   1020
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
         TabIndex        =   22
         Top             =   5385
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
         TabIndex        =   21
         Top             =   5490
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
         TabIndex        =   19
         Top             =   4620
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
         TabIndex        =   17
         Top             =   4230
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
         TabIndex        =   15
         Top             =   3840
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
         TabIndex        =   13
         Top             =   3225
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
         TabIndex        =   11
         Top             =   2835
         Width           =   2160
      End
      Begin VB.Line Line1 
         X1              =   450
         X2              =   4965
         Y1              =   3675
         Y2              =   3675
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
         Top             =   1860
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
         Top             =   1065
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
         Top             =   675
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
         Index           =   4
         Left            =   6750
         MaxLength       =   50
         TabIndex        =   27
         Top             =   135
         Width           =   2040
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
         Left            =   3900
         MaxLength       =   50
         TabIndex        =   26
         Top             =   135
         Width           =   2040
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
         Index           =   5
         Left            =   10095
         MaxLength       =   50
         TabIndex        =   29
         Top             =   120
         Width           =   2040
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
         Left            =   1005
         MaxLength       =   50
         TabIndex        =   24
         Top             =   120
         Width           =   2040
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
         Left            =   6120
         TabIndex        =   32
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
         Left            =   3270
         TabIndex        =   25
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
         Left            =   9075
         TabIndex        =   28
         Top             =   60
         Width           =   1005
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
         Left            =   120
         TabIndex        =   23
         Top             =   195
         Width           =   705
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   135
      TabIndex        =   35
      Top             =   2880
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
      Picture         =   "frmEmp13thMonth.frx":0F33
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   135
      TabIndex        =   41
      Top             =   3510
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmEmp13thMonth.frx":16AD
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   135
      TabIndex        =   42
      Top             =   4140
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmEmp13thMonth.frx":1E27
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   135
      TabIndex        =   43
      Top             =   3510
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmEmp13thMonth.frx":25A1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   135
      TabIndex        =   44
      Top             =   1620
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Confirm"
      AccessKey       =   "Confirm"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmp13thMonth.frx":2D1B
   End
End
Attribute VB_Name = "frmEmp13thMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmEmp13thMonth"

Private oSkin As clsFormSkin
Private WithEvents oTrans As clsYearEndBonus
Attribute oTrans.VB_VarHelpID = -1
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim pnRow As Integer
Dim poCtrl As Variant

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

   lsOldProc = "cmdButton_Click"
'   ''On Error GoTo errProc

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
   Case 2 'Close
      Unload Me
   Case 3 'Update
      If oTrans.UpdateTransaction Then
         Call InitForm(xeModeUpdate)
      End If
   Case 4 'Cancel Update
      If MsgBox("Do you really want to undo changes?", vbCritical + vbYesNo, "Confirmation") = vbYes Then
         If oTrans.UnSaveTransaction Then
            Call LoadDetail
            Call InitForm(xeModeReady)
         End If
      End If
   Case 5 'Save
      If oTrans.SaveTransaction Then
         MsgBox "Updates save successfully!", vbOKOnly, "Confirmation"
         oTrans.LoadDetail
         Call LoadDetail
         Call InitForm(xeModeReady)
      Else
         MsgBox "Unable to save changes!" & vbCrLf & _
                "Please check your entry and try again...", vbOKOnly, "Confirmation"
      End If
   Case 6 'Confirm
      If oTrans.CloseTransaction Then
         MsgBox "Bonus(es) was confirmed successfully!", vbOKOnly, "Confirmation"
         oTrans.LoadDetail
         Call LoadDetail
         Call InitForm(xeModeReady)
      Else
         MsgBox "Unable to confirm bonus(es)!" & vbCrLf & _
                "Please check your entry and try again...", vbOKOnly, "Confirmation"
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
'   ''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      If LCase(oApp.ProductID) <> "petmgr" Then
         txtSearch(0).Text = ""
         txtSearch(1).Text = oApp.BranchName
         
         txtSearch(0).Locked = True
         txtSearch(1).Locked = True
         cmbSearch(0).Locked = True
         
      End If
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
'   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsYearEndBonus
   Set oTrans.AppDriver = oApp
   oTrans.TransStatus = 10

   oTrans.InitTransaction
   oTrans.SearchTransaction ""
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

   Call InitGrid
   Call ClearFields(-1)
   Call InitForm(xeModeReady)
   
   Call oTrans.LoadDetail
   Call LoadDetail

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
   If oTrans.EmpLevel >= 0 Then cmbSearch(0).ListIndex = oTrans.EmpLevel
      
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
   If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> "" Then
      Call ClearFields(MSHFlexGrid1.Row - 1)
   Else
      Call ClearFields(-1)
   End If
End Sub

Private Sub MSHFlexGrid1_SelChange()
   With MSHFlexGrid1
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub oTrans_DetailRetrieve(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 4
      txtField(Index) = Value
   Case 9, 10, 11
      txtField(Index) = Format(Value, "#,##0.00")
      lblTotal = Format(CCur(txtField(7)) + CCur(txtField(9)) + CCur(txtField(10)) - CCur(txtField(11)), "#,##0.00")
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
      Set poCtrl = txtField(Index)
   End With

   pnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub InitForm(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(0).Visible = Not lbShow
   cmdButton(1).Visible = Not lbShow
   cmdButton(2).Visible = Not lbShow
   cmdButton(3).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   
   cmdButton(4).Visible = lbShow
   cmdButton(5).Visible = lbShow
   
   txtField(4).Enabled = lbShow
   txtField(9).Enabled = lbShow
   txtField(10).Enabled = lbShow
   
   xrFrame2.Enabled = Not lbShow
   xrFrame3.Enabled = lbShow
   
   If lbShow Then
      txtField(4).SetFocus
   End If
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSHFlexGrid1
      .Cols = 8
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
      .TextMatrix(0, 2) = "Months"
      .TextMatrix(0, 3) = "13th Pay"
      .TextMatrix(0, 4) = "Bonus"
      .TextMatrix(0, 5) = "XMas"
      
      'column width
      .ColWidth(0) = 590
      .ColWidth(1) = 2500
      .ColWidth(2) = 800
      .ColWidth(3) = 1300
      .ColWidth(4) = 1300
      .ColWidth(5) = 1300
      .ColWidth(6) = 0
      .ColWidth(7) = 0

      'column allinment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignRightCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      
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
   
   With MSHFlexGrid1
      .Rows = 2
      If oTrans.ItemCount = 0 Then
         .TextMatrix(1, 0) = 1
         .TextMatrix(1, 1) = ""
         .TextMatrix(1, 2) = 0
         .TextMatrix(1, 3) = "0.00"
         .TextMatrix(1, 4) = "0.00"
         .TextMatrix(1, 5) = "0.00"
         .TextMatrix(1, 6) = "0.00"
         .TextMatrix(1, 7) = "0.00"
      Else
         .Rows = oTrans.ItemCount + 1
         For lnCtr = 0 To oTrans.ItemCount - 1
            DoEvents
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "xFullName")
            .TextMatrix(lnCtr + 1, 2) = Format(oTrans.Detail(lnCtr, "nAttend01"), "#0.00")
            .TextMatrix(lnCtr + 1, 3) = Format(oTrans.Detail(lnCtr, "n13thMnth"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 4) = Format(oTrans.Detail(lnCtr, "nBonusxx1"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 5) = Format(oTrans.Detail(lnCtr, "nPartyxx1"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 6) = Format(oTrans.Detail(lnCtr, "nDeductnx"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 7) = Format(CCur(.TextMatrix(lnCtr + 1, 3)) + CCur(.TextMatrix(lnCtr + 1, 4)) + CCur(.TextMatrix(lnCtr + 1, 5)) - CCur(.TextMatrix(lnCtr + 1, 6)), "#,##0.00")
         
         
            ln13thMnth = ln13thMnth + CCur(.TextMatrix(lnCtr + 1, 3))
            lnBonusxx1 = lnBonusxx1 + CCur(.TextMatrix(lnCtr + 1, 4))
            lnPartyxx1 = lnPartyxx1 + CCur(.TextMatrix(lnCtr + 1, 5))
         Next
      End If
   End With
   
   txtSearch(2) = Format(ln13thMnth, "#,##0.00")
   txtSearch(3) = Format(lnBonusxx1, "#,##0.00")
   txtSearch(4) = Format(lnPartyxx1, "#,##0.00")
   txtSearch(5) = Format(ln13thMnth + lnBonusxx1 + lnPartyxx1, "#,##0.00")

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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With MSHFlexGrid1
      Select Case Index
      Case 4
         oTrans.Detail(.Row - 1, Index) = txtField(Index)
      Case 9, 10, 11
         oTrans.Detail(.Row - 1, Index) = txtField(Index)
      End Select
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
