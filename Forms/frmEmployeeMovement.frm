VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmEmployeeMovement 
   BorderStyle     =   0  'None
   Caption         =   "Employee Movement"
   ClientHeight    =   9060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrFrame xrFrame2 
      Height          =   8355
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   14737
      BackColor       =   12632256
      ClipControls    =   0   'False
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
         Index           =   24
         ItemData        =   "frmEmployeeMovement.frx":0000
         Left            =   1605
         List            =   "frmEmployeeMovement.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   5190
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
         Height          =   390
         Index           =   91
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   21
         Top             =   4680
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
         Height          =   390
         Index           =   90
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   19
         Top             =   4260
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
         Height          =   390
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
         Index           =   81
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2160
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
         Height          =   390
         Index           =   1
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2580
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
         Height          =   390
         Index           =   82
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   13
         Top             =   3000
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
         Height          =   390
         Index           =   83
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   15
         Top             =   3420
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
         Height          =   390
         Index           =   84
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   17
         Top             =   3840
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
         Height          =   390
         Index           =   85
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   25
         Top             =   5595
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
         Height          =   390
         Index           =   9
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   27
         Top             =   6015
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
         Height          =   390
         Index           =   87
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   29
         Top             =   6435
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
         Height          =   390
         Index           =   11
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   31
         Top             =   6870
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
         Height          =   390
         Index           =   13
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   33
         Top             =   7395
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
         Height          =   390
         Index           =   14
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   35
         Top             =   7815
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
         Height          =   390
         Index           =   80
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   3
         Top             =   795
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
         Height          =   390
         Index           =   89
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1635
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
         Height          =   390
         Index           =   88
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1215
         Width           =   4815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Level"
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
         Index           =   41
         Left            =   105
         TabIndex        =   22
         Top             =   5280
         Width           =   1365
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. Branch"
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
         Left            =   105
         TabIndex        =   20
         Top             =   4770
         Width           =   1080
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Branch"
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
         TabIndex        =   18
         Top             =   4350
         Width           =   1230
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   300
         X2              =   6420
         Y1              =   5130
         Y2              =   5130
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   180
         X2              =   6300
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Left            =   90
         TabIndex        =   6
         Top             =   1695
         Width           =   1005
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
         Index           =   5
         Left            =   105
         TabIndex        =   4
         Top             =   1305
         Width           =   615
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
         Caption         =   "Employee "
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
         TabIndex        =   2
         Top             =   885
         Width           =   930
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   180
         X2              =   6300
         Y1              =   2085
         Y2              =   2085
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Top             =   2250
         Width           =   450
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
         Index           =   2
         Left            =   105
         TabIndex        =   10
         Top             =   2670
         Width           =   405
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
         Index           =   3
         Left            =   105
         TabIndex        =   12
         Top             =   3090
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         TabIndex        =   14
         Top             =   3510
         Width           =   1005
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
         Index           =   7
         Left            =   105
         TabIndex        =   16
         Top             =   3930
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Level"
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
         Left            =   105
         TabIndex        =   24
         Top             =   5685
         Width           =   1050
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Pay"
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
         TabIndex        =   26
         Top             =   6105
         Width           =   900
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Benefits"
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
         Left            =   105
         TabIndex        =   28
         Top             =   6525
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Index           =   11
         Left            =   105
         TabIndex        =   30
         Top             =   6960
         Width           =   675
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reqstd By"
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
         Index           =   12
         Left            =   105
         TabIndex        =   32
         Top             =   7485
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Reqstd"
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
         Index           =   13
         Left            =   105
         TabIndex        =   34
         Top             =   7905
         Width           =   1080
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   39
      Top             =   6120
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
      Picture         =   "frmEmployeeMovement.frx":006F
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   37
      Top             =   5490
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
      Picture         =   "frmEmployeeMovement.frx":07E9
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   38
      Top             =   6120
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
      Picture         =   "frmEmployeeMovement.frx":0F63
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   36
      Top             =   5490
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
      Picture         =   "frmEmployeeMovement.frx":16DD
   End
End
Attribute VB_Name = "frmEmployeeMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmEmployeeMovement"

Private oDriver As clsFormDriver
Private WithEvents oTrans As clsEmployeeMovement
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Private pnIndex As Integer

Private Sub cmbField_Validate(Index As Integer, Cancel As Boolean)
   oTrans.Master(Index) = cmbField(Index).ListIndex
End Sub

Private Sub cmdButton_Click(Index As Integer)
Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   '''On Error GoTo errProc

   Select Case Index
   Case 0   'cancel
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
      If lnRep = vbYes Then
         InitForm 1
      End If
   Case 1   'search
   Case 2   'save
      If oTrans.SaveTransaction(True) Then
         MsgBox "Transaction Successfully Saved!!!", vbInformation, "Confirm"
         oTrans.NewTransaction
         InitForm 0
         LoadMaster
      Else
         MsgBox "Unable to Update Transaction!!!", vbCritical, "Warning"
      End If
   Case 3   'new
      oTrans.NewTransaction
      InitForm 0
      LoadMaster
      txtField(80).SetFocus
   Case 4   'close
      Unload Me
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   '''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
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
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me

   Set oTrans = New clsEmployeeMovement
   Set oTrans.AppDriver = oApp

   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

   InitForm 0
   oTrans.NewTransaction
   LoadMaster
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub InitForm(ByVal fnEdit As Integer)
   Dim lnCtr As Integer
   Dim loTxt As TextBox

   xrFrame2.Enabled = (fnEdit = 0)
   cmdButton(3).Visible = Not (fnEdit = 0)
   cmdButton(4).Visible = Not (fnEdit = 0)

   cmdButton(0).Visible = (fnEdit = 0)
   cmdButton(2).Visible = (fnEdit = 0)

   For Each loTxt In txtField
      loTxt = ""
   Next
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0) = .Master(0)
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 1, 14
      txtField(Index) = strLongDate(Value)
   Case 9, 11
      If IsNumeric(Value) Then
         txtField(Index) = Format(Value, "#,##0.00")
      Else
         txtField(Index) = ""
      End If
   Case Else
      txtField(Index) = Value
      If Index = 80 Then
         txtField(88) = oTrans.Master(88)
         txtField(89) = oTrans.Master(89)
         txtField(90) = oTrans.Master(90)
         txtField(91) = oTrans.Master(91)
         If IsNumeric(oTrans.Master(24)) Then
            cmbField(24).ListIndex = oTrans.Master(24)
         End If
      End If
   End Select
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3
      If Index >= 80 Then
         If oTrans.SearchMaster(Index, txtField(Index).Text) Then
            SetNextFocus
         End If
      End If
   Case vbKeyReturn
      If Index >= 80 Then
         If txtField(Index) <> "" Then
            Call oTrans.SearchMaster(Index, txtField(Index).Text)
         Else
            oTrans.Master(Index) = txtField(Index).Text
         End If
      End If
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 9, 11
      If IsNumeric(txtField(Index).Text) Then
         oTrans.Master(Index) = txtField(Index).Text
      End If
   Case Else
      oTrans.Master(Index) = txtField(Index).Text
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   Select Case Index
      Case 1, 14
         txtField(Index) = strShortDate(IFNull(oTrans.Master(Index), ""))
   End Select

   oDriver.ColumnIndex = Index
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

