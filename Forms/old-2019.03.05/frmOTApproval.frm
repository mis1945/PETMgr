VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmOTApproval 
   BorderStyle     =   0  'None
   Caption         =   "Overtime Approval"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   26
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
      Picture         =   "frmOTApproval.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   2250
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
      Picture         =   "frmOTApproval.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   24
      Top             =   2880
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
      Picture         =   "frmOTApproval.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   25
      Top             =   3510
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
      Picture         =   "frmOTApproval.frx":166E
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3150
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   3090
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   5556
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
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
         Index           =   1
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   15
         Top             =   795
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
         Height          =   800
         Index           =   5
         Left            =   1620
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   2205
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
         Index           =   3
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1755
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
         Height          =   420
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2190
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   2
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1275
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
         Index           =   4
         Left            =   4430
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1740
         Width           =   2190
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
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
         TabIndex        =   18
         Top             =   1830
         Width           =   915
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
         TabIndex        =   22
         Top             =   2205
         Width           =   780
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
         TabIndex        =   14
         Top             =   885
         Width           =   405
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
         TabIndex        =   11
         Top             =   210
         Width           =   1485
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1725
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   2175
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
         TabIndex        =   16
         Top             =   1365
         Width           =   930
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thru"
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
         Left            =   3990
         TabIndex        =   20
         Top             =   1830
         Width           =   375
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   4400
         Top             =   120
         Width           =   2220
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   4425
         Top             =   150
         Width           =   2160
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
         Left            =   4470
         TabIndex        =   13
         Tag             =   "eb0;et0"
         Top             =   180
         Width           =   2070
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1140
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   2011
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
         Height          =   420
         Index           =   0
         Left            =   1620
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   120
         Width           =   2190
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
         Height          =   420
         Index           =   1
         Left            =   1620
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   585
         Width           =   5000
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
         Top             =   210
         Width           =   1365
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
         Index           =   15
         Left            =   105
         TabIndex        =   2
         Top             =   675
         Width           =   870
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1335
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   1725
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   2355
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
         Height          =   420
         Index           =   2
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   9
         Top             =   765
         Width           =   2190
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
         Index           =   8
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "December 30, 2012"
         Top             =   105
         Width           =   2190
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
         Index           =   9
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   7
         Top             =   105
         Width           =   2190
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applied From"
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
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   4005
         TabIndex        =   6
         Top             =   210
         Width           =   210
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
         Index           =   16
         Left            =   105
         TabIndex        =   8
         Top             =   885
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmOTApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmOTApproval"

Private WithEvents oTrans As clsOvertime
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim pbSearched As Boolean

Private psTransNox As String
Public Property Let TransNox(Value As String)
   psTransNox = Value
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
         If pnIndex = 0 Then
            If oTrans.SearchTransaction(txtSearch(pnIndex).Text, False) Then
               Clearfields
               LoadMaster
               Call InitFields
            End If
         Else
            If oTrans.SearchTransaction(txtSearch(pnIndex).Text) Then
               Clearfields
               LoadMaster
               Call InitFields
            End If
         End If
         pnIndex = 3
      Else
         If oTrans.SearchTransaction("") Then
            Clearfields
            LoadMaster
            Call InitFields
         End If
      End If
   Case 2   'approve
      If oTrans.Master(0) <> "" And txtField(8) <> "" And txtField(9) <> "" Then
         If oTrans.PostTransaction(oTrans.Master(0), _
                                    txtField(8).Text, _
                                    txtField(9).Text) Then
            MsgBox "Transaction was closed successfuly!!!", vbInformation, "Notice"
         Else
            MsgBox "Closing transaction failed!!!", vbInformation, "Notice"
         End If
         Call Clearfields
      End If
      GoTo endWithFocus
   Case 3   'disapprove
      If oTrans.Master(0) <> "" And txtField(0) <> "" Then
         If oTrans.CancelTransaction Then
            MsgBox "Transaction was cancelled!!!", vbInformation, "Notice"
            txtField(8).Enabled = False
            txtField(9).Enabled = False
         Else
            MsgBox "Transaction cancellation failed!!!", vbInformation, "Notice"
         End If
         Call Clearfields
      End If
      GoTo endWithFocus
   End Select

endProc:
   Exit Sub
endWithFocus:
   txtSearch(0) = ""
   txtSearch(1) = ""
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
   
   If bLoaded = False Then
      bLoaded = True
   End If
   
   pbSearched = False
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

   Set oTrans = New clsOvertime
   Set oTrans.AppDriver = oApp
   
   If LCase(oApp.ProductID) = "petmgr" Then
      oTrans.TransStatus = 10
   Else
      oTrans.TransStatus = 0
   End If
   
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
      
   If psTransNox <> "" Then
      '@@@ soft-monitor
      Call oTrans.OpenTransaction(psTransNox)
   End If

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   Clearfields
   InitFields
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Public Sub LoadMaster()
   Dim lnCtr As Integer
   
   For lnCtr = 0 To 5
      Select Case lnCtr
         Case 1, 3, 4
            txtField(lnCtr).Text = strLongDate(oTrans.Master(lnCtr))
         Case Else
            txtField(lnCtr).Text = oTrans.Master(lnCtr)
      End Select
   Next
   
   txtField(8).Text = strLongDate(oTrans.Master("dAppldFrx"))
   txtField(9).Text = strLongDate(oTrans.Master("dAppldTox"))
   txtSearch(0) = txtField(0)
   txtSearch(1) = txtField(2)
   txtSearch(2) = strLongDate(oTrans.DateApproved)
   
   If oTrans.Master("cTranStat") = "4" Then
      Label2.Caption = "P-A-Y-E-D"
   Else
      Label2.Caption = TransStat(oTrans.Master("cTranstat"))
   End If
   
   pbSearched = True
End Sub

Private Sub Clearfields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
   txtSearch(0) = ""
   txtSearch(1) = ""
   Label2.Caption = ""
   
   pbSearched = False
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 10
      Label2.Caption = TransStat(CInt(Value))
   Case 1, 3, 8, 9
      txtField(Index) = strLongDate(Value)
   End Select
End Sub
Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   Select Case Index
      Case 8, 9
         txtField(Index) = strShortDate(txtField(Index))
   End Select
   
   pnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
   
   Select Case Index
      Case 8, 9
         txtField(Index) = strLongDate(txtField(Index))
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 8, 9
         oTrans.Master(8) = txtField(Index)
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
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   pnIndex = Index
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 0
         If oTrans.SearchTransaction(txtSearch(Index).Text, False) Then
            Clearfields
            LoadMaster
            Call InitFields
         End If
      Case 1
         If oTrans.SearchTransaction(txtSearch(Index).Text) Then
            Clearfields
            LoadMaster
            Call InitFields
         End If
      End Select
   End If
End Sub

Private Sub InitFields()
   xrFrame2.Enabled = (pbSearched = True)
   If pbSearched Then txtField(8).SetFocus
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

Private Sub txtSearch_Validate(Index As Integer, Cancel As Boolean)
   If Index = 2 Then
      oTrans.DateApproved = txtSearch(2)
      txtSearch(2) = strLongDate(oTrans.DateApproved)
   End If
End Sub

