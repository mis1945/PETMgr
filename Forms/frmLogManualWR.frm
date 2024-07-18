VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLogManualWR 
   BorderStyle     =   0  'None
   Caption         =   "Log Manual"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   8955
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   15796
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin xrControl.xrFrame xrFrame1 
         Height          =   1335
         Left            =   105
         Tag             =   "wt0;fb0"
         Top             =   1700
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   2355
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtOthers 
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
            MaxLength       =   50
            TabIndex        =   12
            Top             =   80
            Width           =   4395
         End
         Begin VB.TextBox txtOthers 
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
            Index           =   2
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   870
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
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
            Index           =   3
            Left            =   6135
            MaxLength       =   50
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   500
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
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
            Index           =   4
            Left            =   6135
            MaxLength       =   50
            TabIndex        =   16
            Top             =   870
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
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
            Index           =   5
            Left            =   10275
            MaxLength       =   50
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   500
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
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
            Index           =   6
            Left            =   10275
            MaxLength       =   50
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   870
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
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
            Index           =   1
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   13
            Top             =   500
            Width           =   2415
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1st Q. Time Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   105
            TabIndex        =   20
            Top             =   945
            Width           =   1260
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1st Q. Time In"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   105
            TabIndex        =   19
            Top             =   570
            Width           =   1125
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   11
            Top             =   195
            Width           =   825
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2nd Q. Time In"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   4590
            TabIndex        =   21
            Top             =   570
            Width           =   1170
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2nd Q. Time Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   4575
            TabIndex        =   22
            Top             =   945
            Width           =   1305
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime In"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   9030
            TabIndex        =   23
            Top             =   570
            Width           =   960
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   9030
            TabIndex        =   24
            Top             =   945
            Width           =   1095
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   1575
         Left            =   105
         Tag             =   "wt0;fb0"
         Top             =   105
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   2778
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
            Index           =   9
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   7
            Top             =   1015
            Width           =   5715
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
            Index           =   10
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   5
            ToolTipText     =   "F3 / Enter to Search "
            Top             =   645
            Width           =   3675
         End
         Begin VB.ComboBox cmbField 
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
            Index           =   7
            ItemData        =   "frmLogManualWR.frx":0000
            Left            =   7680
            List            =   "frmLogManualWR.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   635
            Width           =   2925
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
            Index           =   2
            Left            =   7695
            MaxLength       =   50
            TabIndex        =   3
            Top             =   120
            Width           =   2355
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
            Left            =   7695
            MaxLength       =   50
            TabIndex        =   10
            Top             =   635
            Width           =   5000
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
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            TabIndex        =   6
            Top             =   1045
            Width           =   765
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reason"
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
            Left            =   120
            TabIndex        =   4
            Top             =   680
            Width           =   660
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   315
            Left            =   1725
            Tag             =   "et0;ht2"
            Top             =   210
            Width           =   2415
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
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
            Index           =   10
            Left            =   6165
            TabIndex        =   2
            Top             =   210
            Width           =   1410
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   105
            TabIndex        =   0
            Top             =   210
            Width           =   1335
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
            TabIndex        =   8
            Top             =   690
            Width           =   1185
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5760
         Left            =   105
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3060
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   10160
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3705
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
      Picture         =   "frmLogManualWR.frx":005B
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3075
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
      Picture         =   "frmLogManualWR.frx":07D5
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3705
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
      Picture         =   "frmLogManualWR.frx":0F4F
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1185
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
      Picture         =   "frmLogManualWR.frx":16C9
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3075
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmLogManualWR.frx":1E43
   End
   Begin xrControl.xrButton cmdDetail 
      Height          =   600
      Index           =   1
      Left            =   105
      TabIndex        =   30
      Top             =   1815
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
      Picture         =   "frmLogManualWR.frx":25BD
   End
   Begin xrControl.xrButton cmdDetail 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2445
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
      Picture         =   "frmLogManualWR.frx":364F
   End
   Begin xrControl.xrButton cmdDetail 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   33
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
      Picture         =   "frmLogManualWR.frx":3DC9
   End
End
Attribute VB_Name = "frmLogManualWR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeModuleName = "frmLogManualWR"
Private Const pxBRANCHCODES = "M001»A001»N001»H001»PHO1»PHO2"

Private WithEvents oTrans As clsLogManualWR
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim pnRow As Integer
Dim pnActiveRow As Integer

Dim pbCtrlPress As Boolean
Dim pbFormLoad As Boolean
Dim pbDetailGotFocus As Boolean
Dim pbCopy2All As Boolean
Dim pbByBranch As Boolean
Dim pnLastSelc As Integer
Dim pbHasValue As Boolean

'controlling get/next focus on detail
Dim pbMoveUpxx As Boolean
Dim pbMoveDown As Boolean

Property Let ByBranch(vValue As Boolean)
   pbByBranch = vValue
End Property

Private Sub cmbField_Validate(Index As Integer, Cancel As Boolean)
   Debug.Print cmbField(Index).ListIndex
   oTrans.Master("sBranchCd") = cmbField(Index).ListIndex
   oTrans.Master("") = cmbField(Index).ListIndex
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lnRep As Integer
   Dim pbState As Boolean
   
   Select Case Index
      Case 0 'cancel
         lnRep = MsgBox("Are you certain in aborting the current transaction?", vbQuestion + vbYesNo, "Confirm")
         
         If lnRep = vbYes Then
            If oTrans.InitTransaction Then
               Call ClearFields
               Call ClearOthers
               Call InitGrid
               Call InitForm(1)
               Call InitOthers(0)
            End If
         End If
      Case 1 'search
         If pnIndex = 0 Then
            Call oTrans.SearchDetail(pnRow, 0, txtOthers(pnIndex))
         ElseIf pnIndex = 1 Then
            Call oTrans.SearchMaster(Index, txtField(pnIndex))
         Else
            GoTo endProc
         End If
      Case 2 'save
         Call oTrans.DeleteDetail(oTrans.ItemCount)
         If oTrans.SaveTransaction Then Call cmdButton_Click(3)
      Case 3 'new
         Call ClearFields
         Call ClearOthers
         Call InitGrid
         Call InitForm(0)
         Call InitOthers(1)
      
         oTrans.NewTransaction
         
         If pbByBranch = False Then
            oTrans.Master("sBranchCd") = 1
         End If
         
         LoadMaster
         Call MSFlexGrid1_Click
         'MAC(07.02.12)
'         If pbByBranch Then
'            txtField(1).SetFocus
'         Else
'            cmbField(7).SetFocus
'         End If
         txtField(2).SetFocus
      Case 4 'close
         Unload Me
   End Select
   
endProc:
   Exit Sub
End Sub

Private Sub cmdDetail_Click(Index As Integer)
   Dim lnCtr As Integer
   Dim lnCtr1 As Integer
   Dim lnRow As Integer
   
   Select Case Index
      Case 0   'delete
         
         If oTrans.ItemCount = 0 And oTrans.Detail(0, 0) = "" Then GoTo endProc
         
         If pnActiveRow < 2 Then Exit Sub 'MAC(07.02.12)
         If oTrans.DeleteDetail(pnActiveRow - 2) Then 'MAC(07.02.12)
            LoadDetail
            pnActiveRow = pnActiveRow - 1
            If pnActiveRow < 2 Then pnActiveRow = 2
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
         
         If oTrans.Detail(pnRow, "xfullname") = "" Then GoTo endProc
         
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
            
            If MSFlexGrid1.Rows > 16 Then
               .ColWidth(1) = 3050
               .TopRow = pnRow
            Else
               .ColWidth(1) = 3300
            End If
               
            .TextMatrix(pnRow + 2, 0) = pnRow + 1
            .Row = pnRow + 2
            pnActiveRow = .Row
            
            Call detailFieldChange
         End With
         GoTo endWithFocus
      Case 2
         If oTrans.loadEmployee Then
            lnRow = oTrans.ItemCount
            
            With MSFlexGrid1
               MSFlexGrid1.Rows = lnRow + 2
               If MSFlexGrid1.Rows > 16 Then
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
               Call flexFocus
               Call detailFieldChange
            End With
            cmdDetail(2).Visible = False
         Else
            MSFlexGrid1.Rows = 3
         End If
         GoTo endWithFocus
   End Select
   
endProc:
   Exit Sub
endWithFocus:
   If xrFrame1.Enabled = True Then txtOthers(0).SetFocus
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
      oTrans.NewTransaction
      Call InitOthers(1)
      
      
      If pbByBranch Then
         txtField(1).Text = oApp.BranchName
      Else
         'Set the default value to 1
         oTrans.HasParent = True
         oTrans.Master(7) = 1
         cmbField(7).ListIndex = Val(oTrans.Master(7))
      End If
      
      bLoaded = True
      
      'If Not (oApp.BranchCode = "M001" Or oApp.BranchCode = "N001" Or oApp.BranchCode = "H001") Then
      If InStr(1, pxBRANCHCODES, oApp.BranchCode) = 0 Then
         cmbField(7).Visible = False
         txtField(1).Visible = True
         txtField(1).TabStop = False
         txtField(1).Locked = True
      Else
         If pbByBranch Then
            lblField(1) = "Branch Name"
            cmbField(7).Visible = False
            txtField(1).Visible = True
         Else
            lblField(1) = "Employee Level"
            cmbField(7).Visible = True
            txtField(1).Visible = False
         End If
      End If
   End If
   
   LoadMaster

   If Not pbFormLoad Then pbFormLoad = True
   pnActiveRow = 2
   pnRow = 2
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   With MSFlexGrid1
      Select Case KeyCode
      Case vbKeyReturn
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
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   Set oTrans = New clsLogManualWR
   Set oTrans.AppDriver = oApp
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Call InitGrid
   Call InitForm(0)
   Call InitOthers(0)
   Call ClearFields
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
   bLoaded = False

End Sub

Private Sub detailFieldChange()
   Dim lnCtr As Integer
   Dim lnRow As Integer

   lnRow = pnActiveRow
   If lnRow < 2 Then Exit Sub 'MAC(07.02.12)
   
   With oTrans
      For lnCtr = 0 To 6
         txtOthers(lnCtr) = IFNull(.Detail(lnRow - 2, lnCtr), "")
         Select Case lnCtr
               Case 1
            txtOthers(lnCtr).Enabled = oTrans.IsActiveAMInxxxx
         Case 2
            txtOthers(lnCtr).Enabled = oTrans.IsActiveAMOutxxx
         Case 3
            txtOthers(lnCtr).Enabled = oTrans.IsActivePMInxxxx
         Case 4
            txtOthers(lnCtr).Enabled = oTrans.IsActivePMOutxxx
         Case 5
            txtOthers(lnCtr).Enabled = oTrans.IsActiveOTimeInx
         Case 6
            txtOthers(lnCtr).Enabled = oTrans.IsActiveOTimeOut
         End Select
      Next
   End With
   
   If txtOthers(0).Enabled Then txtOthers(0).SetFocus
   SetGridRowColor (lnRow)
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

Private Sub LoadMaster()
   txtField(0) = oTrans.Master(0)
   
   If Not pbByBranch Then
      cmbField(7).ListIndex = oTrans.Master(7)
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
      
      If MSFlexGrid1.Rows > 16 Then
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
      
      'MAC(06.26.12)
      'pnActiveRow = .Rows - 1
   End With
   
End Sub

Private Sub MSFlexGrid1_Click()
   Dim lnCtr As Integer
   'Try this out...
'   If oTrans.ItemCount = 1 Then Exit Sub
      
   With oTrans
      pnActiveRow = MSFlexGrid1.Row
      pnRow = pnActiveRow - 2
      
      For lnCtr = 0 To 6
         txtOthers(lnCtr) = Format(IFNull(.Detail(pnRow, lnCtr), ""), "HH:MM AM/PM")
      Next
      
         Call detailFieldChange
   End With
   
   If xrFrame1.Enabled Then txtOthers(0).SetFocus
   
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer, ByVal Value As Variant)
   With txtOthers(Index)
      .Text = IFNull(Value, "")
   End With

'   If pnActiveRow < 1 Or pbCopy2All Or (MSFlexGrid1.Row = MSFlexGrid1.Rows - 1) Then Exit Sub
   If pnActiveRow < 1 Or pbCopy2All Then Exit Sub
   
   With MSFlexGrid1
      .TextMatrix(pnActiveRow, Index + 1) = IFNull(Value, "")
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   With txtField(Index)
      Select Case Index
         Case 2
            .Text = strLongDate(IFNull(Value, ""))
         Case Else
            .Text = IFNull(Value, "")
      End Select
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   
   Select Case Index
      Case 2
         txtField(Index) = strShortDate(txtField(Index))
   End Select
   
   
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   pbMoveUpxx = True
   pbMoveDown = True
   pnIndex = Index
End Sub
Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
         Case vbKeyF3
      If Index = 1 Then
         If oTrans.SearchMaster(Index, txtField(Index).Text) Then
            SetNextFocus
         End If
      End If
      If Index = 10 Then
         If oTrans.SearchMaster(Index, txtField(Index).Text) Then
            SetNextFocus
         End If
      End If
   Case vbKeyReturn
      If Index = 1 Then
         If txtField(Index) <> "" Then
            Call oTrans.SearchMaster(Index, txtField(Index).Text)
         Else
            oTrans.Master(Index) = txtField(Index).Text
         End If
      End If
   End Select
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

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      oTrans.Master(Index) = .Text
   End With
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
                     Call oTrans.SearchDetail(pnRow - 2, Index, txtOthers(Index))
                     Call MoveToLastRec
                     txtOthers(Index).SetFocus
               Else
                  MsgBox "Unable to overwrite row data." + vbCrLf + _
                     "Delete the row if incorrect entry.", vbInformation
                  txtOthers(Index).SetFocus
               End If
            Case vbKeyF3
               If oTrans.Detail(IIf(pnActiveRow > 0, pnActiveRow - 1, pnActiveRow), Index) = "" Then
                  Call oTrans.SearchDetail(pnRow - 2, Index, txtOthers(Index))
                  Call MoveToLastRec
               Else
                  MsgBox "Unable to overwrite row data." + vbCrLf + _
                     "Delete the row if incorrect entry.", vbInformation
                  txtOthers(Index).SetFocus
               End If
         End Select
      Case Else
         Select Case KeyCode
         Case vbKeyF12
               If pbCtrlPress Then Call CopyT2AllEmp
         End Select
   End Select
      
   Select Case KeyCode
      Case vbKeyAdd
         If pbCtrlPress Then Call cmdDetail_Click(1)
      Case vbKeySubtract
         If pbCtrlPress Then Call cmdDetail_Click(0)
   End Select
   
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Or (pnActiveRow < 2) Then Exit Sub
   
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
            lsTime = txtOthers(Index) '& "p" 'she 2015-06-10 para saan bakit nag add pa ulit ng "p"
            oTrans.Detail(pnActiveRow - 2, Index) = getCTime(lsTime)
         Else
            oTrans.Detail(pnActiveRow - 2, Index) = getCTime(txtOthers(Index))
         End If
         If Index = 4 Then
            pbCtrlPress = True
            Call Form_KeyDown(vbKeyDown, False)
            pbCtrlPress = False
         End If
   End Select
      
End Sub

Private Sub InitForm(ByVal fnEdit As Integer)
   Dim lnCtr As Integer
   
   xrFrame1.Enabled = False
   xrFrame2.Enabled = (fnEdit = 0)
   cmdButton(3).Visible = Not (fnEdit = 0)
   cmdButton(4).Visible = Not (fnEdit = 0)

   cmdButton(0).Visible = (fnEdit = 0)
   cmdButton(1).Visible = (fnEdit = 0)
   cmdButton(2).Visible = (fnEdit = 0)
   
   cmdDetail(0).Visible = (fnEdit = 0)
   cmdDetail(1).Visible = (fnEdit = 0)
   cmdDetail(2).Visible = (fnEdit = 0)
   
   If fnEdit = 1 Then
      pnActiveRow = 1
      pnRow = 0
   End If
End Sub

Private Sub InitOthers(ByVal fnEdit As Integer)
   Dim lnCtr As Integer
   
   Select Case fnEdit
   Case 0
      For lnCtr = 1 To 6
         txtOthers(lnCtr).Enabled = False
      Next
   Case 1
      xrFrame1.Enabled = True
      For lnCtr = 1 To 6
         txtOthers(lnCtr).Enabled = True
      Next
   End Select
End Sub

Private Sub ClearFields()
   Dim lnCtr As Integer
   Dim loTxt As TextBox
      
'   For lnCtr = 0 To 4
'      txtField(lnCtr) = ""
'   Next
   txtField(0) = ""
   txtField(1) = ""
   txtField(2) = ""
   txtField(10) = ""
   txtField(9) = ""
   For Each loTxt In txtField
      loTxt.Text = ""
      loTxt.BackColor = oApp.getColor("EB")
   Next
End Sub

Private Sub ClearOthers()
   Dim lnCtr As Integer
   Dim loTxt As TextBox
   
   For lnCtr = 0 To 6
      txtOthers(lnCtr) = ""
   Next
   
   For Each loTxt In txtOthers
      loTxt.BackColor = oApp.getColor("EB")
   Next
   
   pnRow = oTrans.ItemCount - 1
End Sub

Private Sub CopyT2AllEmp()
   Dim lnCtr As Integer
   pbCopy2All = True
   For lnCtr = 0 To oTrans.ItemCount - 1
      With oTrans
         .Detail(lnCtr, pnIndex) = txtOthers(pnIndex)
         
         'Mac PH 08.11.12
         'lnCtr + 1 >> lnCtr + 2
         MSFlexGrid1.TextMatrix(lnCtr + 2, pnIndex + 1) = .Detail(lnCtr, pnIndex)
      End With
   Next
   pbCopy2All = False
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



