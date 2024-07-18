VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLogManualApprvlWR 
   BorderStyle     =   0  'None
   Caption         =   "Log Manual Approval"
   ClientHeight    =   10290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   2445
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
      Picture         =   "frmLogManualApprvlWR.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   555
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
      Picture         =   "frmLogManualApprvlWR.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Approve"
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
      Picture         =   "frmLogManualApprvlWR.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   5
      Top             =   1815
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&DisApprv"
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
      Picture         =   "frmLogManualApprvlWR.frx":166E
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   630
      Index           =   1
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   1111
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
         Index           =   0
         Left            =   1620
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   2415
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
         Top             =   165
         Width           =   1365
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
         TabIndex        =   2
         Tag             =   "eb0;et0"
         Top             =   180
         Width           =   2355
      End
      Begin VB.Shape Shape3 
         Height          =   345
         Index           =   0
         Left            =   10485
         Top             =   150
         Width           =   2415
      End
      Begin VB.Shape Shape4 
         Height          =   405
         Index           =   0
         Left            =   10455
         Top             =   120
         Width           =   2475
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   8955
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   1200
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
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   13
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
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
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
            Index           =   5
            Left            =   10275
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   11
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
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   10
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
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   9
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
            Index           =   2
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   870
            Width           =   2415
         End
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
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   7
            Top             =   80
            Width           =   4395
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
            TabIndex        =   20
            Top             =   945
            Width           =   1095
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
            TabIndex        =   19
            Top             =   570
            Width           =   960
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
            TabIndex        =   18
            Top             =   945
            Width           =   1305
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
            TabIndex        =   17
            Top             =   570
            Width           =   1170
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
            TabIndex        =   16
            Top             =   195
            Width           =   825
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
            TabIndex        =   15
            Top             =   570
            Width           =   1125
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
            TabIndex        =   14
            Top             =   945
            Width           =   1260
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   1575
         Index           =   0
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
            Index           =   0
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   120
            Width           =   2415
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
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   25
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
            Index           =   2
            Left            =   7695
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   24
            Top             =   120
            Width           =   2355
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
            ItemData        =   "frmLogManualApprvlWR.frx":1DE8
            Left            =   7680
            List            =   "frmLogManualApprvlWR.frx":1DFB
            Style           =   2  'Dropdown List
            TabIndex        =   23
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
            Index           =   10
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "F3 / Enter to Search "
            Top             =   635
            Width           =   3675
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
            Index           =   9
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1015
            Width           =   5715
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
            TabIndex        =   31
            Top             =   690
            Width           =   1185
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
            TabIndex        =   30
            Top             =   210
            Width           =   1335
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
            TabIndex        =   29
            Top             =   210
            Width           =   1410
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
            TabIndex        =   28
            Top             =   680
            Width           =   660
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
            TabIndex        =   27
            Top             =   1045
            Width           =   765
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
End
Attribute VB_Name = "frmLogManualApprvlWR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeModuleName = "frmLogManualApprvl"
Private Const pxBRANCHCODES = "M001»H001»N001»PHO1»PHO2"

Private WithEvents oTrans As clsLogManualWR
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim pnRow As Integer
Dim pnActiveRow As Integer
Dim pbByBranch As Boolean
Dim pbCtrlPress As Boolean
Dim pbFormLoad As Boolean
Dim pbHaveRefNo As Boolean
Dim pbDtlLoaded As Boolean
Dim pnLastSelc As Integer
Private psTransNox As String
Public Property Let TransNox(Value As String)
   psTransNox = Value
End Property
Property Let ByBranch(vValue As Boolean)
   pbByBranch = vValue
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   Select Case Index
   Case 0   'close
      Unload Me
   Case 1   'search
      If pnIndex = 0 Or pnIndex = 1 Then
         pbDtlLoaded = False
         If oTrans.SearchTransaction(txtSearch(pnIndex).Text, pnIndex = 0) Then
            ClearFields
            LoadMaster
            LoadDetail
            txtOthers(0).SetFocus
         End If
      End If
   Case 2   'approve
      If oTrans.Master(0) = "" Then GoTo endProc
      If oTrans.CloseTransaction() Then
         Label2.Caption = TransStat(CInt(oTrans.Master(6)))
         MsgBox "Transaction was closed successfuly!!!", vbInformation, "Notice"
         ClearFields
      Else
         MsgBox "Transaction approval failed!!!", vbInformation, "Notice"
      End If
      GoTo endWithFocus
   Case 3   'disapprove
      If oTrans.Master(0) = "" Then GoTo endProc
      If oTrans.CancelTransaction Then
         Label2.Caption = TransStat(CInt(oTrans.Master(6)))
         MsgBox "Transaction was cancelled!!!", vbInformation, "Notice"
         ClearFields
      Else
         MsgBox "Transaction cancellation failed!!!", vbInformation, "Notice"
      End If
      GoTo endWithFocus
   End Select

endProc:
   Exit Sub
endWithFocus:
   txtSearch(0) = ""
   txtSearch(0).SetFocus
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
'
'   If bLoaded = False Then
'
'
'      If pbByBranch Then
'         txtField(1).Text = oApp.BranchName
'      Else
'         'Set the default value to 1
'         oTrans.HasParent = True
'         oTrans.Master(7) = 1
'         cmbField(7).ListIndex = Val(oTrans.Master(7))
'      End If
      
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
         If pbDtlLoaded Then Exit Sub
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
                        
            SetNextFocus
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
            SetPreviousFocus
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
   oTrans.TransStatus = IIf(InStr(1, pxBRANCHCODES, oApp.BranchCode) > 0, 10, 0)
'   oTrans.TransStatus = IIf(oApp.BranchCode = "M001" Or oApp.BranchCode = "H001" Or oApp.BranchCode = "N001", 10, 0)
   oTrans.InitTransaction
   
   If psTransNox <> "" Then
         '@@@ soft-monitor
      Call oTrans.OpenTransaction(psTransNox)
   End If

   Call ClearFields
   Call InitGrid
   
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
      Dim lnCtr As Integer
   
   For lnCtr = 0 To 2
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
      ElseIf lnCtr = 2 Then
         txtField(lnCtr) = strLongDate(oTrans.Master(lnCtr))
      Else
         txtField(lnCtr) = oTrans.Master(lnCtr)
      End If
   Next
   
   txtSearch(0) = txtField(0)
   txtField(2) = Format(txtField(2), "MMMM DD, YYYY")
   txtField(9) = IFNull(oTrans.Master(9), "")
   txtField(10) = IFNull(oTrans.Master(10), "")

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

   If oTrans.ItemCount = 1 Then Exit Sub

   With oTrans
      pnActiveRow = MSFlexGrid1.Row

      For lnCtr = 1 To 7
         txtOthers(lnCtr - 1) = Format(IFNull(.Detail(pnActiveRow - 1, lnCtr - 1), ""), "HH:MM:SS AM/PM")
      Next
      
   Call detailFieldChange
   End With
End Sub

Private Sub MSFlexGrid1_GotFocus()
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


Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         txtSearch(0).SetFocus
   End Select
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF3, vbKeyReturn
         pbDtlLoaded = False
         If oTrans.SearchTransaction(txtSearch(Index).Text, True) Then
            ClearFields
            LoadMaster
            LoadDetail
            txtOthers(0).SetFocus
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
      
      If .Text <> "" Then
         .SelStart = 0
         .SelLength = Len(.Text)
      End If
   End With
   
   pnIndex = Index
End Sub
Private Sub txtOthers_LostFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("EB")
   End With

   pnIndex = Index
End Sub
Private Sub txtOthers_GotFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
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

