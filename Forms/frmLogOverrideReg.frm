VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLogOverrideReg 
   BorderStyle     =   0  'None
   Caption         =   "Log Override Register"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   6930
      TabIndex        =   21
      Top             =   2490
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Delete"
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
      Picture         =   "frmLogOverrideReg.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   6930
      TabIndex        =   24
      Top             =   3750
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
      Picture         =   "frmLogOverrideReg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   6930
      TabIndex        =   19
      Top             =   1230
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
      Picture         =   "frmLogOverrideReg.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   6930
      TabIndex        =   20
      Top             =   1860
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
      Picture         =   "frmLogOverrideReg.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   6930
      TabIndex        =   23
      Top             =   3120
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
      Picture         =   "frmLogOverrideReg.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   4
      Left            =   6930
      TabIndex        =   25
      Top             =   3750
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
      Picture         =   "frmLogOverrideReg.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   6930
      TabIndex        =   22
      Top             =   3120
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
      Picture         =   "frmLogOverrideReg.frx":2CDC
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   6510
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   11483
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.CheckBox Check1 
         Caption         =   "Time Out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   4260
         TabIndex        =   31
         Tag             =   "wt0;fb0"
         Top             =   3510
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Time In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   4260
         TabIndex        =   30
         Tag             =   "wt0;fb0"
         Top             =   3165
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Time Out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   2535
         TabIndex        =   29
         Tag             =   "wt0;fb0"
         Top             =   3510
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Time In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   2535
         TabIndex        =   28
         Tag             =   "wt0;fb0"
         Top             =   3165
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Time Out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   810
         TabIndex        =   27
         Tag             =   "wt0;fb0"
         Top             =   3510
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Time In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   810
         TabIndex        =   26
         Tag             =   "wt0;fb0"
         Top             =   3165
         Width           =   1305
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1260
         Width           =   4815
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
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2415
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   4
         Top             =   795
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
         Index           =   5
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   2190
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
         Index           =   4
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1725
         Width           =   4815
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
         Index           =   7
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         Top             =   4470
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
         Index           =   14
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   16
         Top             =   4935
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
         Index           =   6
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   12
         Top             =   4005
         Width           =   4815
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
         Height          =   990
         Index           =   3
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   18
         Top             =   5400
         Width           =   4815
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "1ST QUARTER"
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
         Left            =   795
         TabIndex        =   34
         Tag             =   "et0;fb0"
         Top             =   2790
         Width           =   1335
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "2ND QUARTER"
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
         Left            =   2520
         TabIndex        =   33
         Tag             =   "et0;fb0"
         Top             =   2790
         Width           =   1365
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "OVERTIME"
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
         Left            =   4425
         TabIndex        =   32
         Tag             =   "et0;fb0"
         Top             =   2790
         Width           =   1005
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   180
         X2              =   6300
         Y1              =   3945
         Y2              =   3945
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   180
         X2              =   6300
         Y1              =   2670
         Y2              =   2670
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
         Height          =   300
         Left            =   4290
         TabIndex        =   2
         Tag             =   "eb0;et0"
         Top             =   180
         Width           =   2070
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   4245
         Top             =   150
         Width           =   2160
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   4215
         Top             =   120
         Width           =   2220
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
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   1350
         Width           =   615
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
         Index           =   21
         Left            =   105
         TabIndex        =   0
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         TabIndex        =   3
         Top             =   885
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requested"
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
         Left            =   105
         TabIndex        =   7
         Top             =   1815
         Width           =   930
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Reqstd."
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
         Left            =   105
         TabIndex        =   9
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Approved"
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
         TabIndex        =   13
         Top             =   4560
         Width           =   1260
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved"
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
         TabIndex        =   11
         Top             =   4095
         Width           =   795
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved Time"
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
         Left            =   105
         TabIndex        =   15
         Top             =   5025
         Width           =   1290
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   105
         TabIndex        =   17
         Top             =   5400
         Width           =   780
      End
      Begin VB.Shape Shape2 
         Height          =   975
         Index           =   2
         Left            =   4125
         Top             =   2925
         Width           =   1590
      End
      Begin VB.Shape Shape2 
         Height          =   975
         Index           =   1
         Left            =   2400
         Top             =   2925
         Width           =   1590
      End
      Begin VB.Shape Shape2 
         Height          =   975
         Index           =   0
         Left            =   660
         Top             =   2925
         Width           =   1590
      End
   End
End
Attribute VB_Name = "frmLogOverrideReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmLogOverrideReg"

Private WithEvents oTrans As clsLogOverride
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim psSelected() As String
Dim pnIndex As Integer

Private Sub Check1_Validate(Index As Integer, Cancel As Boolean)
   oTrans.Master(Index) = Check1(Index).Value
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   On Error Goto errProc

   Select Case Index
   Case 0   'save
      If oTrans.SaveTransaction(True) Then
      MsgBox "Transaction was successfuly updated.", vbInformation, "Notice"
         InitForm 0
         ClearFields
      End If
   Case 1   'cancel trans
      MsgBox oTrans.Master(0)
      If oTrans.CancelTransaction(oTrans.Master(0)) Then
         MsgBox "Transaction was cancelled successfuly!!!", vbInformation, "Notice"
      Else
         MsgBox "Cancelling failed!!!", vbInformation, "Notice"
      End If
   Case 2   'update
      If txtField(2) = "" Then Exit Sub
      
      If oTrans.Master(15) = xeStateCancelled Or _
         oTrans.Master(15) = xeStatePosted Then
         MsgBox "Updating Closed/Cancelled/Posted Transaction is strictly prohibited!!!", vbCritical, "Warning"
      Else
         InitForm 1
         txtField(1).SetFocus
      End If
   Case 3   'browse
      If oTrans.SearchTransaction("", False) Then LoadMaster
   Case 4   'cancel
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
      
      If lnRep = vbYes Then
         oTrans.InitTransaction
         Call ClearFields
         InitForm 0
      End If
   Case 5   'close
      Unload Me
   Case 6   'delete trans
      If oTrans.Master(0) <> "" Then
         If oTrans.DeleteTransaction Then
            MsgBox "Transaction was successfuly deleted.", vbInformation, "Notice"
            InitForm 0
            ClearFields
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
   On Error Goto errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      bLoaded = True
      txtField(0).SetFocus
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
   On Error Goto errProc

   CenterChildForm mdiMain, Me
   
   Set oTrans = New clsLogOverride
   Set oTrans.AppDriver = oApp
   
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
   Call InitForm(0)
   ClearFields
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub InitForm(ByVal fnEdit As Integer)
   Dim lnCtr As Integer
   Dim chkbox As CheckBox
   
'   xrFrame1.Enabled = Not (fnEdit = 0)
   cmdButton(0).Visible = Not (fnEdit = 0)
   cmdButton(4).Visible = Not (fnEdit = 0)
   txtField(14).Enabled = Not (fnEdit = 0)
   
   cmdButton(2).Visible = (fnEdit = 0)
   cmdButton(3).Visible = (fnEdit = 0)
   cmdButton(5).Visible = (fnEdit = 0)
   txtField(0).Enabled = (fnEdit = 0)
   txtField(1).Enabled = Not (fnEdit = 0)
   
   For Each chkbox In Check1
      chkbox.Enabled = Not (fnEdit = 0)
   Next
   
   For lnCtr = 3 To 7
      txtField(lnCtr).Enabled = Not (fnEdit = 0)
   Next
   
   If fnEdit = 0 Then Label2.Caption = ""
End Sub
Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
End Sub
Private Sub LoadMaster()
   Dim lnCtr As Integer
   With oTrans
      For lnCtr = 0 To 7
         Select Case lnCtr
            Case 2, 5, 7
               txtField(lnCtr) = strLongDate(.Master(lnCtr))
            Case Else
               txtField(lnCtr) = .Master(lnCtr)
         End Select
      Next
      For lnCtr = 8 To 13
         Check1(lnCtr).Value = CInt(.Master(lnCtr))
      Next
      txtField(14) = Format(.Master(14), "HH:MM AM/PM")
      Label2.Caption = TransStat(.Master(15))
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 2, 7
      txtField(Index) = strLongDate(IFNull(Value, ""))
   Case 15
      Label2.Caption = TransStat(CInt(Value))
   Case Else
      txtField(Index) = IFNull(Value, "")
   End Select
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3
      If Index = 1 Or Index = 4 Or Index = 6 Then
         If oTrans.SearchMaster(Index, txtField(Index).Text) Then
            SetNextFocus
         End If
      End If
   Case vbKeyReturn
      If Index = 1 Or Index = 4 Or Index = 6 Then
         If txtField(Index) <> "" Then
            Call oTrans.SearchMaster(Index, txtField(Index).Text)
         Else
            oTrans.Master(Index) = txtField(Index).Text
         End If
      End If
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   oTrans.Master(Index) = txtField(Index).Text
End Sub
Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   pnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
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

