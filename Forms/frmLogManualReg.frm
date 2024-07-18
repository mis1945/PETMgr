VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLogManualReg 
   BorderStyle     =   0  'None
   Caption         =   "Log Manual"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   13440
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmLogManualReg.frx":0000
   End
   Begin xrControl.xrButton cmdDetail 
      Height          =   600
      Index           =   1
      Left            =   13455
      TabIndex        =   27
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&ADD"
      AccessKey       =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmLogManualReg.frx":077A
   End
   Begin xrControl.xrButton cmdDetail 
      Height          =   600
      Index           =   0
      Left            =   13455
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&DEL"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmLogManualReg.frx":180C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   13440
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2445
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
      Picture         =   "frmLogManualReg.frx":1F86
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   13440
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3075
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
      Picture         =   "frmLogManualReg.frx":2700
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   13440
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2445
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
      Picture         =   "frmLogManualReg.frx":2E7A
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   4
      Left            =   13440
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3075
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
      Picture         =   "frmLogManualReg.frx":35F4
   End
   Begin xrControl.xrButton cmdDetail 
      Height          =   600
      Index           =   2
      Left            =   13440
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   555
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
      Picture         =   "frmLogManualReg.frx":3D6E
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   645
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   1138
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   6450
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   105
         Width           =   2355
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   1620
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   105
         Width           =   2415
      End
      Begin VB.Shape Shape4 
         Height          =   405
         Index           =   0
         Left            =   10455
         Top             =   105
         Width           =   2475
      End
      Begin VB.Shape Shape3 
         Height          =   345
         Index           =   0
         Left            =   10485
         Top             =   135
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10515
         TabIndex        =   4
         Tag             =   "eb0;et0"
         Top             =   165
         Width           =   2355
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
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
         Index           =   15
         Left            =   4665
         TabIndex        =   2
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         Index           =   14
         Left            =   105
         TabIndex        =   0
         Top             =   180
         Width           =   1365
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   8280
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1215
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   14605
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin xrControl.xrFrame xrFrame1 
         Height          =   1590
         Left            =   105
         Tag             =   "wt0;fb0"
         Top             =   1335
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   2805
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   16
            Top             =   585
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   6
            Left            =   10275
            MaxLength       =   50
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1035
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   5
            Left            =   10275
            MaxLength       =   50
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   585
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   4
            Left            =   5985
            MaxLength       =   50
            TabIndex        =   22
            Top             =   1035
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   3
            Left            =   5985
            MaxLength       =   50
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   585
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1050
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   14
            Top             =   105
            Width           =   5000
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime Out"
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
            Left            =   9030
            TabIndex        =   25
            Top             =   1125
            Width           =   1140
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime In"
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
            Left            =   9030
            TabIndex        =   23
            Top             =   675
            Width           =   975
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Out PM"
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
            Left            =   4680
            TabIndex        =   21
            Top             =   1125
            Width           =   1155
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time In PM"
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
            Left            =   4695
            TabIndex        =   19
            Top             =   675
            Width           =   990
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
            TabIndex        =   13
            Top             =   195
            Width           =   870
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time In AM"
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
            TabIndex        =   15
            Top             =   675
            Width           =   990
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Out AM"
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
            Left            =   105
            TabIndex        =   17
            Top             =   1125
            Width           =   1155
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   1215
         Index           =   0
         Left            =   105
         Tag             =   "wt0;fb0"
         Top             =   105
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   2143
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   120
            Width           =   2415
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   3
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   675
            Width           =   2415
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   7695
            MaxLength       =   50
            TabIndex        =   12
            Top             =   675
            Width           =   5000
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   7695
            MaxLength       =   50
            TabIndex        =   10
            Top             =   120
            Width           =   2355
         End
         Begin VB.ComboBox cmbField 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   7
            ItemData        =   "frmLogManualReg.frx":44E8
            Left            =   7695
            List            =   "frmLogManualReg.frx":44FB
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   675
            Width           =   2925
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Name"
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
            Left            =   6165
            TabIndex        =   11
            Top             =   765
            Width           =   1185
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction No."
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
            Index           =   8
            Left            =   105
            TabIndex        =   5
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referrence No."
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
            Left            =   105
            TabIndex        =   7
            Top             =   765
            Width           =   1290
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tran. Date"
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
            Index           =   10
            Left            =   6165
            TabIndex        =   9
            Top             =   210
            Width           =   900
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   420
            Left            =   1725
            Tag             =   "et0;ht2"
            Top             =   210
            Width           =   2415
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5175
         Left            =   105
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2955
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   9128
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         WordWrap        =   -1  'True
         Enabled         =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLogManualReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeModuleName = "frmLogManualReg"
Private Const pxBRANCHCODES = "M001»H001»N001»PHO1»PHO2"

Private WithEvents oTrans As clsLogManual
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim pnRow As Integer
Dim pnActiveRow As Integer

Dim pbCtrlPress As Boolean
Dim pbFormLoad As Boolean
Dim pbDetailGotFocus As Boolean
Dim pbUpdateMode As Boolean
Dim pbCopy2All As Boolean
Dim pnLastSelc As Integer
Dim pbMoveUpxx As Boolean
Dim pbMoveDown As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   Select Case Index
      Case 0   'save
         Call oTrans.DeleteDetail(oTrans.ItemCount)
         If oTrans.SaveTransaction Then
            MsgBox "Transaction was successfuly updated.", vbInformation, "Notice"
            Call ClearFields
            Call InitForm(0)
            Call InitGrid
            pbUpdateMode = False
            GoTo endWithFocus
         End If
      Case 2   'update
         If oTrans.UpdateTransaction Then
            Call InitForm(1)
'            oTrans.AddDetail
            
            'Reposition Row
            pnRow = oTrans.ItemCount - 1
            pnActiveRow = pnRow + 1
            
            With MSFlexGrid1
               .Row = 2
               pnActiveRow = 2
               .Col = 1
               .ColSel = .Cols - 1
            End With
            pbUpdateMode = True
            GoTo endWithFocus
         Else
            MsgBox "Updating Closed/Cancelled/Posted Transaction is strictly prohibited!!!", vbCritical, "Warning"
         End If
      Case 3   'browse
         If pnIndex > 1 Then pnIndex = 1
         
         If oTrans.SearchTransaction(txtOthers(pnIndex), True) Then
            Call ClearFields
            Call InitGrid
            Call LoadMaster
            Call LoadDetail
            Call detailFieldChange
         End If
         GoTo endWithFocus
         
      Case 4   'cancel
         lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
   
         If lnRep = vbYes Then
            Call ClearFields
            Call InitForm(0)
            Call InitGrid
            pbUpdateMode = False
            GoTo endWithFocus
         End If
         
      Case 5   'close
         Unload Me
   End Select
   
endProc:
   Exit Sub
endWithFocus:
   If xrFrame1.Enabled = False Then
      txtSearch(0).SetFocus
   Else
      txtOthers(0).SetFocus
   End If
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub cmdDetail_Click(Index As Integer)
   Dim lnCtr As Integer
   Dim lnCtr1 As Integer
   Dim lnRow As Integer

   If Not pbUpdateMode Then GoTo endWithFocus
   
   Select Case Index
      Case 0   'delete
         If oTrans.ItemCount = 0 And oTrans.Detail(0, 0) = "" Then GoTo endProc
         
         If pnActiveRow < 2 Then Exit Sub 'MAC(07.02.12)
         If oTrans.DeleteDetail(pnActiveRow - 2) Then
            pnActiveRow = pnActiveRow - 1
            If pnActiveRow < 2 Then pnActiveRow = 2
            LoadDetail
            ClearOthers
            Call detailFieldChange
            GoTo endWithFocus
         End If
      Case 1   'add
         If oTrans.ItemCount = 0 Then
            oTrans.InitTransaction
            pnRow = 0
         End If
         
'         MoveToLastRec
         'Causes an error
         'kalyptus-2012.04.18
         If oTrans.Detail(pnRow, "xfullname") = "" Then GoTo endProc
         'If oTrans.Detail(pnRow - 2, 0) = "" Then GoTo endProc
            
         With MSFlexGrid1
         
            .Rows = .Rows + 1
            .Row = .Rows - 2
            pnRow = .Row
            
            .TextMatrix(pnRow + 1, 0) = pnRow + 1
            
            For lnCtr = 0 To 6
               .TextMatrix(pnRow + 1, lnCtr + 1) = IFNull(oTrans.Detail(pnRow - 1, lnCtr), "")
            Next
            
            Call oTrans.addDetail
            ClearOthers
            
            If MSFlexGrid1.Rows > 13 Then
               .ColWidth(1) = 3050
'               .TopRow = pnRow
            Else
               .ColWidth(1) = 3300
            End If
               
            .TextMatrix(pnRow + 1, 0) = pnRow
            .Row = pnRow + 1
            pnActiveRow = .Row
                        

            Call detailFieldChange
         End With
         GoTo endWithFocus
      Case 2
         If oTrans.loadEmployee Then
            lnRow = oTrans.ItemCount
            
            With MSFlexGrid1
               MSFlexGrid1.Rows = lnRow + 2
'               If MSFlexGrid1.Rows > 13 Then .ColWidth(1) = 2750
               
               If MSFlexGrid1.Rows > 13 Then
                  .ColWidth(1) = 3050
               Else
                  .ColWidth(1) = 3300
               End If
               
               For lnCtr = 1 To lnRow
                  MSFlexGrid1.TextMatrix(lnCtr + 1, 0) = lnCtr
                  For lnCtr1 = 0 To 6
                     MSFlexGrid1.TextMatrix(lnCtr + 1, lnCtr1 + 1) = IFNull(oTrans.Detail(lnCtr - 1, lnCtr1), "")
                  Next
               Next
            End With
            
            Call detailFieldChange
            Call flexFocus
         Else
            MSFlexGrid1.Rows = 3
         End If
         GoTo endWithFocus
   End Select
   
endProc:
   Exit Sub
endWithFocus:
   If xrFrame1.Enabled Then
      txtOthers(0).SetFocus
   ElseIf xrFrame3(1).Enabled Then
      txtSearch(0).SetFocus
   End If
   GoTo endProc
End Sub


Private Sub cmdDetail_GotFocus(Index As Integer)
   Select Case Index
      Case 0, 1
         MSFlexGrid1_Click
   End Select
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
      
      If InStr(1, pxBRANCHCODES, oApp.BranchCode) = 0 Then
'      If oApp.BranchCode <> "M001" Then
         txtField(1).Text = oApp.BranchName
         cmbField(7).Visible = False
         txtField(1).Visible = True
         txtField(1).TabStop = False
         txtField(1).Locked = True
      Else
         lblField(1) = "Branch Name"
         cmbField(7).Visible = False
         txtField(1).Visible = True
      End If
   End If

   If Not pbFormLoad Then pbFormLoad = True
   pnActiveRow = 2
   pnRow = 2
   
'   txtSearch(0).SetFocus
   pbUpdateMode = False
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   With MSFlexGrid1
      Select Case KeyCode
      Case vbKeyReturn
         If Not pbMoveDown Then Exit Sub
         If GetFocus = .hWnd Then Exit Sub
         SetNextFocus
      Case vbKeyDown
         If pbCtrlPress Then

            If pnActiveRow = .Rows - 1 Then Exit Sub
            
            If Not (pnActiveRow = 0) Then
               pnActiveRow = pnActiveRow + 1
            End If
            
            .Row = pnActiveRow
            
            If Not .RowIsVisible(.Row + IIf(.Row = .Rows - 1, 0, 1)) Then .TopRow = .Row - 1

            Call detailFieldChange
         Else
            If Not pbMoveDown Then Exit Sub
            If Not (GotFocus = cmbField(7).hWnd) Then
               SetNextFocus
            End If
         End If
      Case vbKeyUp
         If pbCtrlPress Then
            If .Row > 2 Then
               If Not (pnActiveRow = 1) Then
                  pnActiveRow = pnActiveRow - 1
               End If
               
               .Row = pnActiveRow
   
               If Not .RowIsVisible(.Row) Then .TopRow = .TopRow - 1

               Call detailFieldChange
            End If
         Else
            If Not pbMoveUpxx Then Exit Sub
            If Not (GotFocus = cmbField(7).hWnd) Then
                  SetPreviousFocus
            End If
         End If
      Case vbKeyControl
         pbCtrlPress = True
      End Select
   End With
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
'   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   Set oTrans = New clsLogManual
   Set oTrans.AppDriver = oApp
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Call ClearFields
   Call InitGrid
   Call InitForm(0)
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing

   pnActiveRow = 0
   pbFormLoad = False
   pbCtrlPress = False
   pbUpdateMode = False
   bLoaded = False
   
End Sub

Private Sub detailFieldChange()
   Dim lnCtr As Integer
   Dim lnRow As Integer
   
   lnRow = pnActiveRow
   SetGridRowColor (lnRow)
   pnActiveRow = MSFlexGrid1.Row
   lnRow = pnActiveRow
   
   With oTrans
      For lnCtr = 0 To 6
         txtOthers(lnCtr) = IFNull(.Detail(lnRow - 2, lnCtr), "")
      Next
   End With
   
   If txtOthers(0).Enabled Then txtOthers(0).SetFocus
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Cols = 8
      .Rows = 2
      .MergeCells = flexMergeFree
      
      .Clear
      
      .Row = 0
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .MergeRow(0) = True
      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = ""
      .TextMatrix(0, 2) = "1ST QUARTER"
      .TextMatrix(0, 3) = "1ST QUARTER"
      .TextMatrix(0, 4) = "2ND QUARTER"
      .TextMatrix(0, 5) = "2ND QUARTER"
      .TextMatrix(0, 6) = "OVERTIME"
      .TextMatrix(0, 7) = "OVERTIME"

      .Row = 1
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      'column title
      .TextMatrix(1, 0) = "No"
      .TextMatrix(1, 1) = "Employee Name"
      .TextMatrix(1, 2) = "In"
      .TextMatrix(1, 3) = "Out"
      .TextMatrix(1, 4) = "In"
      .TextMatrix(1, 5) = "Out"
      .TextMatrix(1, 6) = "In"
      .TextMatrix(1, 7) = "Out"
      
      .RowHeightMin = 338
      
      'column width
      .ColWidth(0) = 500
      .ColWidth(1) = 3300
      .ColWidth(2) = 1500
      .ColWidth(3) = 1500
      .ColWidth(4) = 1500
      .ColWidth(5) = 1500
      .ColWidth(6) = 1500
      .ColWidth(7) = 1500
      
      'column allinment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignCenterCenter
      .ColAlignment(4) = flexAlignCenterCenter
      .ColAlignment(5) = flexAlignCenterCenter
      .ColAlignment(6) = flexAlignCenterCenter
      .ColAlignment(7) = flexAlignCenterCenter
            
      .Rows = 3
      .TextMatrix(2, 0) = "1"
      
      .Row = 2
      pnLastSelc = .Row
      SetGridRowColor (.Row)
   End With
End Sub

Private Sub InitForm(lnStat As Integer)
   Dim lbShow As Boolean
   Dim lnCtr As Integer

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(3).Visible = Not lbShow
   cmdButton(2).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow

   cmdButton(0).Visible = lbShow
   cmdButton(4).Visible = lbShow
   
   txtSearch(0).Enabled = Not lbShow
   txtSearch(1).Enabled = Not lbShow
   txtField(1).Enabled = lbShow
   txtField(2).Enabled = lbShow
   cmbField(7).Enabled = lbShow
   
   If lbShow Then
      xrFrame1.Enabled = lbShow
      txtOthers(1).Enabled = oTrans.IsActiveAMInxxxx
      txtOthers(2).Enabled = oTrans.IsActiveAMOutxxx
      txtOthers(3).Enabled = oTrans.IsActivePMInxxxx
      txtOthers(4).Enabled = oTrans.IsActivePMOutxxx
      txtOthers(5).Enabled = oTrans.IsActiveOTimeInx
      txtOthers(6).Enabled = oTrans.IsActiveOTimeOut
   Else
      For lnCtr = 1 To 6
         txtOthers(lnCtr).Enabled = False
      Next
      xrFrame1.Enabled = lbShow
   End If
End Sub

Private Sub LoadMaster()
   Dim lnCtr As Integer
   
   For lnCtr = 0 To 3
      If lnCtr = 1 Then
         If Len(oTrans.Master(7)) = 4 Then
            txtField(lnCtr) = oTrans.Master(lnCtr)
            txtField(1).Visible = True
            cmbField(7).Visible = False
            lblField(1) = "Branch Name"
         Else
            cmbField(7).ListIndex = oTrans.Master(7)
            txtField(1).Visible = False
            cmbField(7).Visible = True
            lblField(1) = "Employee Level"
         End If
      Else
         txtField(lnCtr) = oTrans.Master(lnCtr)
      End If
   Next
   
   txtSearch(0) = txtField(0)
   txtSearch(1) = txtField(3)
   txtField(2) = Format(txtField(2), "MMMM DD, YYYY")

   If oTrans.Master("cTranStat") = "4" Then
      Label2.Caption = "APPLIED"
   Else
      Label2.Caption = TransStat(CInt(oTrans.Master("cTranStat")))
   End If
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCtr As Integer

   If oTrans.ItemCount = 0 Then oTrans.InitTransaction
   lnRow = oTrans.ItemCount

   With MSFlexGrid1
      .Rows = 3
      .Rows = lnRow + 2
      
      If MSFlexGrid1.Rows > 13 Then
         .ColWidth(1) = 3050
      Else
         .ColWidth(1) = 3300
      End If
      
      For lnCtr = 0 To lnRow - 1
         .TextMatrix(lnCtr + 2, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 2, 1) = IFNull(oTrans.Detail(lnCtr, 0), "")
         .TextMatrix(lnCtr + 2, 2) = IFNull(oTrans.Detail(lnCtr, 1), "")
         .TextMatrix(lnCtr + 2, 3) = IFNull(oTrans.Detail(lnCtr, 2), "")
         .TextMatrix(lnCtr + 2, 4) = IFNull(oTrans.Detail(lnCtr, 3), "")
         .TextMatrix(lnCtr + 2, 5) = IFNull(oTrans.Detail(lnCtr, 4), "")
         .TextMatrix(lnCtr + 2, 6) = IFNull(oTrans.Detail(lnCtr, 5), "")
         .TextMatrix(lnCtr + 2, 7) = IFNull(oTrans.Detail(lnCtr, 6), "")
      Next
      
'      .Row = 2
'      pnActiveRow = .Row
'      pnRow = pnActiveRow - 2
      
      Call detailFieldChange 'set info into textbox
   End With
   
End Sub

Private Sub MSFlexGrid1_Click()
   Dim lnCtr As Integer
   
   With oTrans
      pnActiveRow = MSFlexGrid1.Row
      pnRow = pnActiveRow - 2

      
'      If pnActiveRow > oTrans.ItemCount Then GoTo endProc
      
      For lnCtr = 0 To 6
         txtOthers(lnCtr) = Format(IFNull(.Detail(pnRow, lnCtr), ""), "HH:MM AM/PM")
      Next
   End With
   
   Call detailFieldChange
endProc:
   If xrFrame1.Enabled = False Then
      txtSearch(0).SetFocus
   Else
      txtOthers(0).SetFocus
   End If
End Sub

Private Sub MSFlexGrid1_GotFocus()
   pbDetailGotFocus = True
   If xrFrame1.Enabled = False Then Exit Sub
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If pbCtrlPress Then
      If KeyCode = vbKeyControl Then pbCtrlPress = False
   End If
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   If Index = 6 Then Label2.Caption = TransStat(CInt(Value))
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   pbMoveDown = True
   pbMoveUpxx = True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF3, vbKeyReturn
         If oTrans.SearchTransaction(txtSearch(Index).Text, True) Then
            ClearFields
            InitGrid
            LoadMaster
            LoadDetail
         End If
    End Select
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("EB")
   End With

   pnIndex = Index
End Sub
Private Sub txtSearch_GotFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
End Sub
Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtSearch
      loTxt.Text = ""
   Next
   For Each loTxt In txtField
      loTxt.Text = ""
   Next
   For Each loTxt In txtOthers
      loTxt.Text = ""
   Next
End Sub
Private Sub txtOthers_GotFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("HT1")
   End With
   
   Select Case Index
      Case 0
         pbMoveUpxx = False
         pbMoveDown = True
      Case 6
         pbMoveUpxx = True
         pbMoveDown = False
      Case Else
         pbMoveDown = True
         pbMoveUpxx = True
   End Select
   
   pnIndex = Index
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lnRow As Integer
   Dim lbCancel As Boolean
   
   pnRow = MSFlexGrid1.Row
   
   Select Case Index
      Case 0
         Select Case KeyCode
            Case vbKeyReturn
               If oTrans.Detail(IIf(pnActiveRow > 0, pnActiveRow - 1, pnActiveRow), Index) = "" Then
                     If oTrans.SearchDetail(pnRow - 2, Index, txtOthers(Index)) Then
                        txtOthers(Index).SetFocus
                        Call LoadDetail
                        Call MoveToLastRec
                     End If
               Else
                  MsgBox "Unable to overwrite row data." + vbCrLf + _
                     "Delete the row if incorrect entry.", vbInformation
                  txtOthers(Index).SetFocus
               End If
            Case vbKeyF3
               If oTrans.Detail(IIf(pnActiveRow > 0, pnActiveRow - 1, pnActiveRow), Index) = "" Then
                  If oTrans.SearchDetail(pnRow - 2, Index, txtOthers(Index)) Then
                     Call LoadDetail
                     Call MoveToLastRec
                  End If
               Else
                  MsgBox "Unable to overwrite row data." + vbCrLf + _
                     "Delete the row if incorrect entry.", vbInformation
                  txtOthers(Index).SetFocus
               End If
            Case vbKeyAdd
               If pbCtrlPress Then Call cmdDetail_Click(1)
            Case vbKeySubtract
               If pbCtrlPress Then Call cmdDetail_Click(0)
         End Select
'Mac PH 08.11.12
'      Case 4
'         If KeyCode = vbKeyReturn Then
'            Call txtOthers_Validate(4, lbCancel)
'            If Not lbCancel Then
'               If MSFlexGrid1.Row + 1 < MSFlexGrid1.Rows Then
'                  MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'                  MSFlexGrid1_Click
'               End If
'            End If
'         End If
      Case Else
         Select Case KeyCode
            Case vbKeyF12
               If pbCtrlPress Then Call CopyT2AllEmp
         End Select
   End Select
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Or (pnActiveRow < 1) Then Exit Sub
   
   Dim lsTime As String
   
   'dont allow to pass time parameter when Employee Name is null.
   If oTrans.Detail(pnActiveRow - 2, 0) = "" Then
      ClearOthers
      Exit Sub
   End If
   
'65       If oTrans.Detail(pnActiveRow - 2, Index) <> "" Then
'66          txtOthers(Index) = oTrans.Detail(pnActiveRow - 2, Index)
'67          Exit Sub
'68       End If
   
   Select Case Index
      Case 1, 2
         oTrans.Detail(pnActiveRow - 2, Index) = getCTime(txtOthers(Index))
      Case 3, 4, 5, 6
         If InStr(txtOthers(Index), "p") = True Or InStr(txtOthers(Index), "pm") Then
            lsTime = txtOthers(Index) & "p"
            oTrans.Detail(pnActiveRow - 2, Index) = getCTime(lsTime)
         Else
            oTrans.Detail(pnActiveRow - 2, Index) = getCTime(txtOthers(Index))
         End If
         
         'Mac PH (08.23.12)
         'Moves down when pm out is validated.
         If Index = 4 Then
            pbCtrlPress = True
            Call Form_KeyDown(vbKeyDown, False)
            pbCtrlPress = False
         End If
   End Select
End Sub

Private Sub CopyT2AllEmp()
   Dim lnCtr As Integer
   pbCopy2All = True
   For lnCtr = 0 To oTrans.ItemCount - 1
      With oTrans
         .Detail(lnCtr, pnIndex) = txtOthers(pnIndex)
         
         MSFlexGrid1.TextMatrix(lnCtr + 2, pnIndex + 1) = .Detail(lnCtr, pnIndex)
      End With
   Next
   pbCopy2All = False
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer, ByVal Value As Variant)
   With txtOthers(Index)
      .Text = IFNull(Value, "")
   End With

'   If pnActiveRow < 1 Or pbCopy2All Or (MSFlexGrid1.Row = MSFlexGrid1.Rows - 1) Then Exit Sub
   If pnActiveRow < 1 Or pbCopy2All Then Exit Sub
   
   MSFlexGrid1.TextMatrix(pnActiveRow, Index + 1) = IFNull(Value, "")
End Sub

Private Sub ClearOthers()
   Dim lnCtr As Integer
   
   For lnCtr = 0 To 6
      txtOthers(lnCtr) = ""
   Next
   
'   pnRow = oTrans.ItemCount - 1
'   MoveToLastRec
End Sub
Private Sub flexFocus()
   With MSFlexGrid1
      If .Row > 15 Then .TopRow = 1
      .Row = 2
      pnActiveRow = .Row
      pbDetailGotFocus = True
   End With
End Sub

Private Sub SetGridRowColor(ByVal lnRow As Integer)
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .FillStyle = flexFillRepeat
      
      .Row = IIf(pnLastSelc = .Rows, pnLastSelc - 1, pnLastSelc)
      .RowSel = pnLastSelc - 1
      .Col = 1
      .ColSel = .Cols - 1
      .CellBackColor = &HFFFFFF

      .Row = lnRow
      .RowSel = lnRow
      .Col = 1
      .ColSel = .Cols - 1
      .CellBackColor = &HFF8080
      
      pnLastSelc = .Row
      If Not .RowIsVisible(.Row + IIf(.Row = .Rows - 1, 0, 1)) Then .TopRow = .Row
   End With
End Sub

Private Sub MoveToLastRec()
   With MSFlexGrid1
      .Row = .Rows - 1
      pnActiveRow = .Row
      pnRow = pnActiveRow - 2
      
      If Not .RowIsVisible(.Row) Then
         .TopRow = .Row - 10
      End If
      
      detailFieldChange
   End With
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oApp
      .xLogError Err.Number, Err.Description, pxeModuleName, lsProcName, Erl
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


