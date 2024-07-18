VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLeaveAdjustment 
   BorderStyle     =   0  'None
   Caption         =   "Leave Application"
   ClientHeight    =   5490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   4740
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   8361
      BackColor       =   12632256
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
         Index           =   83
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2655
         Width           =   975
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
         Index           =   81
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   7
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
         Height          =   800
         Index           =   4
         Left            =   1620
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   3690
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
         Index           =   3
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   13
         Top             =   3225
         Width           =   975
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
         Index           =   82
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2190
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
         Index           =   1
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   3
         Top             =   795
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
         Height          =   420
         Index           =   80
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1260
         Width           =   4815
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   180
         X2              =   6300
         Y1              =   3135
         Y2              =   3135
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Credit"
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
         TabIndex        =   10
         Top             =   2745
         Width           =   1065
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
         Left            =   105
         TabIndex        =   8
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No of Days"
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
         TabIndex        =   12
         Top             =   3315
         Width           =   960
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purpose"
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
         TabIndex        =   14
         Top             =   3690
         Width           =   720
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
         TabIndex        =   6
         Top             =   1815
         Width           =   615
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
         TabIndex        =   2
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
         TabIndex        =   0
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
         Width           =   2415
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
         TabIndex        =   4
         Top             =   1350
         Width           =   930
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   20
      Top             =   3330
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
      Picture         =   "frmLeaveAdjustment.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   18
      Top             =   2700
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
      Picture         =   "frmLeaveAdjustment.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   19
      Top             =   3330
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
      Picture         =   "frmLeaveAdjustment.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2070
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
      Picture         =   "frmLeaveAdjustment.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   16
      Top             =   2700
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
      Picture         =   "frmLeaveAdjustment.frx":1DE8
   End
End
Attribute VB_Name = "frmLeaveAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmLeaveAdjustment"

Private WithEvents oTrans As clsIndividualLeave
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim psSelected() As String
Dim pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc

   Select Case Index
   Case 0   'cancel
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
      If lnRep = vbYes Then
         InitForm 1
      End If
   Case 1   'search
      If pnIndex = 80 Then oTrans.SearchMaster pnIndex, IIf(txtField(pnIndex).Text <> "", txtField(pnIndex).Text, "%")
   Case 2   'save
      If oTrans.SaveTransaction Then
         MsgBox "Transaction Updated Successfully!!!", vbInformation, "Confirm"
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
   On Error GoTo errProc

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
   On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsIndividualLeave
   Set oTrans.AppDriver = oApp

   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

   Call InitForm(0)
   oTrans.NewTransaction
   Call LoadMaster
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
   Dim loTxt As TextBox

   For Each loTxt In txtField
      loTxt = ""
   Next

   xrFrame2.Enabled = (fnEdit = 0)
   cmdButton(3).Visible = Not (fnEdit = 0)
   cmdButton(4).Visible = Not (fnEdit = 0)

   cmdButton(0).Visible = (fnEdit = 0)
   cmdButton(1).Visible = (fnEdit = 0)
   cmdButton(2).Visible = (fnEdit = 0)

   
End Sub
Private Sub LoadMaster()
   Dim loTxt As TextBox

   For Each loTxt In txtField
      Select Case loTxt.Index
         Case 1
            loTxt = strLongDate(oTrans.Master(loTxt.Index))
         Case Else
            loTxt = oTrans.Master(loTxt.Index)
      End Select
   Next
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant)
   Dim lnCtr As Integer

   Select Case Index
      Case 80
         For lnCtr = 80 To 83
            txtField(lnCtr).Text = oTrans.Master(lnCtr)
         Next
      Case 1
         txtField(Index) = strLongDate(oTrans.Master(Index))
      Case Else
         txtField(Index).Text = oTrans.Master(Index)
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
   oTrans.Master(Index) = txtField(Index).Text
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   Select Case Index
      Case 1
         txtField(Index) = strShortDate(oTrans.Master(Index))
   End Select
   
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
