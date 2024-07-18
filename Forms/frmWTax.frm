VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmWTax 
   BorderStyle     =   0  'None
   Caption         =   "WTax Table Maintenance"
   ClientHeight    =   8250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   KeyPreview      =   -1  'True
   LinkTopic       =   "WTAX Chart"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1185
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   2090
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
         Height          =   400
         Index           =   3
         Left            =   6315
         TabIndex        =   7
         Text            =   ".05"
         Top             =   600
         Width           =   1410
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
         Height          =   400
         Index           =   2
         Left            =   6315
         TabIndex        =   5
         Text            =   "25,500.00"
         Top             =   120
         Width           =   1410
      End
      Begin VB.ComboBox cmbField 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         ItemData        =   "frmWTax.frx":0000
         Left            =   1845
         List            =   "frmWTax.frx":0010
         TabIndex        =   3
         Text            =   "Semi-Monthly"
         Top             =   555
         Width           =   2190
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
         Height          =   400
         Index           =   0
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "August 16, 2011"
         Top             =   105
         Width           =   2190
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "In Excess"
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
         Left            =   4725
         TabIndex        =   6
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Amount"
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
         Left            =   4725
         TabIndex        =   4
         Top             =   195
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table Type"
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
         Left            =   255
         TabIndex        =   2
         Top             =   645
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effectivity Date"
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
         Left            =   240
         TabIndex        =   0
         Top             =   185
         Width           =   1290
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   5340
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1770
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   9419
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtAmntx 
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
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   31
         Text            =   "10,000.00"
         Top             =   4455
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
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
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   29
         Text            =   "10,000.00"
         Top             =   4035
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
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
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   27
         Text            =   "10,000.00"
         Top             =   3615
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
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
         Index           =   8
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "10,000.00"
         Top             =   3195
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
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
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   23
         Text            =   "10,000.00"
         Top             =   2775
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
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
         Index           =   1
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   11
         Text            =   "0"
         Top             =   255
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
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
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   21
         Text            =   "10,000.00"
         Top             =   2355
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
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
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   19
         Text            =   "10,000.00"
         Top             =   1935
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
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
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "10,000.00"
         Top             =   1515
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
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
         Index           =   3
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "7, 200.00"
         Top             =   1095
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
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
         Index           =   2
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   675
         Width           =   1755
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Married - W/ 4 Qualified Dependent"
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
         Left            =   855
         TabIndex        =   30
         Top             =   4515
         Width           =   3045
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Married - W/ 3 Qualified Dependent"
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
         Left            =   855
         TabIndex        =   28
         Top             =   4095
         Width           =   3045
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Married - W/ 2 Qualified Dependent"
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
         Left            =   855
         TabIndex        =   26
         Top             =   3675
         Width           =   3045
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Married - W/ 1 Qualified Dependent"
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
         Left            =   855
         TabIndex        =   24
         Top             =   3255
         Width           =   3045
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single - W/ 4 Qualified Dependent"
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
         Left            =   855
         TabIndex        =   22
         Top             =   2835
         Width           =   2940
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Exemption"
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
         Left            =   855
         TabIndex        =   10
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single - W/ 3 Qualified Dependent"
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
         Left            =   855
         TabIndex        =   20
         Top             =   2422
         Width           =   2940
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single - W/ 2 Qualified Dependent"
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
         Left            =   855
         TabIndex        =   18
         Top             =   2002
         Width           =   2940
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single - W/ 1 Qualified Dependent"
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
         Left            =   855
         TabIndex        =   16
         Top             =   1582
         Width           =   2940
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maried"
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
         Left            =   855
         TabIndex        =   14
         Top             =   1162
         Width           =   585
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single"
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
         Left            =   855
         TabIndex        =   12
         Top             =   742
         Width           =   540
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   6840
      TabIndex        =   32
      Top             =   7470
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
      Picture         =   "frmWTax.frx":003A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   7635
      TabIndex        =   33
      Top             =   7470
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
      Picture         =   "frmWTax.frx":07B4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   6840
      TabIndex        =   8
      Top             =   7470
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmWTax.frx":0F2E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   7635
      TabIndex        =   9
      Top             =   7470
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
      Picture         =   "frmWTax.frx":16A8
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmWTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmWTax"

Private WithEvents oTrans As clsWTaxChart
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnCtr As Integer
Dim pnIndex As Integer

Dim pcSalTypex As String

Property Let Period(ByVal Value As String)
   If InStr("D«W«S«M", Value) > 0 Then
      pcSalTypex = Value
   End If
End Property

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 3 'Cancel
      InitForm 0
      txtField(2).SetFocus
   Case 2 'Save
      If oTrans.SaveTransaction Then
         InitForm 0
         txtField(2).SetFocus
      Else
         txtAmntx(1).SetFocus
      End If
   Case 0 'Retrieve
      If oTrans.OpenTransaction(pcSalTypex, CCur(txtField(2)), CSng(txtField(3))) Then
         LoadDetail
         If oTrans.UpdateTransaction Then
            InitForm 1
            txtAmntx(1).SetFocus
         End If
      End If
   Case 1 'Close
      Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
'   On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      InitForm 0
      LoadDetail
   
      txtField(0).Text = Format(oApp.getConfiguration("WTaxDte"), "Mmmm dd, yyyy")
      Select Case pcSalTypex
      Case "M"
         cmbField(1).Text = "Monthly"
      Case "S"
         cmbField(1).Text = "Semi-Monthly"
      Case "W"
         cmbField(1).Text = "Weekly"
      Case "D"
         cmbField(1).Text = "Daily"
            
      End Select
      txtField(2).Text = "0.00"
      txtField(3).Text = "0.00"
      
      txtField(2).SetFocus
      
      bLoaded = True
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
'   On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   Set oTrans = New clsWTaxChart
   Set oTrans.AppDriver = oApp
   
   oTrans.InitTransaction

   pcSalTypex = "S"

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
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

Private Sub Form_Unload(Cancel As Integer)
   bLoaded = False
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub InitForm(ByVal fnEdit As Integer)
   xrFrame2.Enabled = (fnEdit = 0)
   cmdButton(0).Visible = (fnEdit = 0)
   cmdButton(1).Visible = (fnEdit = 0)
   
   xrFrame3.Enabled = Not (fnEdit = 0)
   cmdButton(2).Visible = Not (fnEdit = 0)
   cmdButton(3).Visible = Not (fnEdit = 0)
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Row As Integer, ByVal Index As Integer, ByVal Value As Variant)
   txtAmntx(Row + 1) = Format(Value, "#,##0.00")
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtAmntx_GotFocus(Index As Integer)
   With txtAmntx(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
      
      pnIndex = Index
   End With
End Sub

Private Sub txtAmntx_LostFocus(Index As Integer)
   With txtAmntx(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtAmntx_Validate(Index As Integer, Cancel As Boolean)
   oTrans.Detail(Index - 1, "nRangeFrm") = txtAmntx(Index).Text
End Sub

Private Sub LoadDetail()
   Dim loTxt As TextBox
   
   For Each loTxt In txtAmntx
      If oTrans.ItemCount > 0 Then
         loTxt.Text = Format(oTrans.Detail(loTxt.Index - 1, "nRangeFrm"), "#,##0.00")
      Else
         loTxt.Text = "0.00"
      End If
   Next
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
