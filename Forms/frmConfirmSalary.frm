VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmConfirmSalary 
   BorderStyle     =   0  'None
   Caption         =   "Employee Salary Confirmation"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   4890
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
      Picture         =   "frmConfirmSalary.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   3630
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
      Picture         =   "frmConfirmSalary.frx":077A
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   645
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   12315
      _ExtentX        =   21722
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
         Height          =   390
         Index           =   1
         Left            =   1290
         TabIndex        =   1
         Text            =   "Dela Cruz, Juan"
         Top             =   105
         Width           =   4815
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
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Width           =   870
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   5760
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1230
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   10160
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   5160
         Left            =   105
         Picture         =   "frmConfirmSalary.frx":0EF4
         Top             =   495
         Width           =   12090
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   105
         Top             =   60
         Width           =   2220
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   135
         Top             =   90
         Width           =   2145
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
         Left            =   180
         TabIndex        =   3
         Tag             =   "eb0;et0"
         Top             =   120
         Width           =   2055
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   5
      Top             =   4260
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Post"
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
      Picture         =   "frmConfirmSalary.frx":C0F4
   End
End
Attribute VB_Name = "frmConfirmSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmTimesheetAdjustmentsApprvl"

Private WithEvents oTrans As clsTimesheetAdj
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Private pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   On Error Goto errProc

   Select Case Index
   Case 0   'close
      Unload Me
   Case 1   'search
      If pnIndex = 0 Or pnIndex = 1 Then
         If pnIndex = 0 Then
            If oTrans.SearchTransaction(txtSearch(pnIndex).Text, False) Then
               ClearFields
               LoadMaster
            End If
         Else
            If oTrans.SearchTransaction(txtSearch(pnIndex).Text) Then
               ClearFields
               LoadMaster
            End If
         End If
         pnIndex = 2
      Else
         If oTrans.SearchMaster("") Then
            ClearFields
            LoadMaster
         End If
      End If
   Case 2   'approve
      If txtField(0) = "" And txtField(80) = "" Then GoTo endProc
      If oTrans.CloseTransaction(oTrans.Master(0)) Then
         MsgBox "Transaction was posted successfuly!!!", vbInformation, "Notice"
      Else
         MsgBox "Closing transaction failed!!!", vbInformation, "Notice"
      End If
      GoTo endWithFocus
   Case 3   'disapprove
      If txtField(0) = "" And txtField(80) = "" Then GoTo endProc
      If oTrans.CancelTransaction(oTrans.Master(0)) Then
         MsgBox "Transaction was cancelled!!!", vbInformation, "Notice"
      Else
         MsgBox "Transaction cancellation failed!!!", vbInformation, "Notice"
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
   'On Error Goto errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
   End If
'   txtSearch(0).SetFocus
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
   'On Error Goto errProc

   CenterChildForm mdiMain, Me

'   Set oTrans = New clsTimesheetAdj
'   Set oTrans.AppDriver = oApp
'
'   oTrans.Branch = oApp.BranchCode
'   oTrans.TransStatus = 0
'   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

'   ClearFields

   Label2.Caption = "OPEN"
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox

   For Each loTxt In txtField
      loTxt = ""
   Next

   Label2.Caption = ""
   txtSearch(0) = ""
   txtSearch(1) = ""
End Sub


Private Sub LoadMaster()
   Dim loTxt As TextBox

   With oTrans
      For Each loTxt In txtField
         Select Case loTxt.Index
         Case 1
            loTxt = strLongDate(.Master(loTxt.Index))
         Case Else
            loTxt = .Master(loTxt.Index)
         End Select
      Next
   
      txtSearch(0) = txtField(0)
      txtSearch(1) = txtField(80)
      txtSearch(2) = strLongDate(oTrans.DateApproved)
      
      Label2.Caption = TransStat(CInt(.Master(17)))
   End With

End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
      Case 1
         txtField(Index) = strLongDate(oTrans.Master(Index))
      Case 17
         Label2.Caption = TransStat(CInt(oTrans.Master(Index)))
      Case Else
         txtField(Index) = oTrans.Master(Index)
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   oTrans.Master(Index) = txtField(Index).Text
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
      Case 1
         txtField(Index) = strShortDate(oTrans.Master(Index))
   End Select
   
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

   pnIndex = Index
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 0
         If oTrans.SearchTransaction(txtSearch(Index).Text, False) Then
            ClearFields
            LoadMaster
         End If
      Case 1
         If oTrans.SearchTransaction(txtSearch(Index).Text) Then
            ClearFields
            LoadMaster
         End If
      End Select
   End If
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
      oTrans.DateApproved = txtSearch(Index)
      txtSearch(Index) = strLongDate(oTrans.DateApproved)
   End If
End Sub

