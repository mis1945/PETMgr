VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMealVoucher 
   BorderStyle     =   0  'None
   Caption         =   "Employee Meal Voucher"
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1755
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   600
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   3096
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
         Index           =   2
         Left            =   1875
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "M00111-000021"
         Top             =   720
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
         Left            =   6915
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1860
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   350
         Left            =   1920
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OPEN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   3
         Left            =   6960
         TabIndex        =   20
         Top             =   840
         Width           =   2235
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meal Type:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   780
         Width           =   1050
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
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
         Left            =   6240
         TabIndex        =   15
         Top             =   180
         Width           =   465
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No:"
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
         Left            =   225
         TabIndex        =   14
         Top             =   180
         Width           =   1485
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmMealVoucher.frx":0000
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1860
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmMealVoucher.frx":077A
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1230
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmMealVoucher.frx":0EF4
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmMealVoucher.frx":166E
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmMealVoucher.frx":1DE8
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   120
      TabIndex        =   12
      Top             =   1860
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmMealVoucher.frx":2562
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1230
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmMealVoucher.frx":2C5C
      CaptionAlign    =   0
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2235
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   2400
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   3942
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   6
         Left            =   5625
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "19EX929430"
         Top             =   1245
         Width           =   3015
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   6915
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2295
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
         Index           =   4
         Left            =   1860
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   3255
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
         Height          =   1320
         Index           =   5
         Left            =   1875
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmMealVoucher.frx":33D6
         Top             =   600
         Width           =   3255
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Left            =   6960
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   5670
         TabIndex        =   21
         Top             =   795
         Width           =   2130
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Control No:"
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
         Index           =   2
         Left            =   5640
         TabIndex        =   19
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name:"
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
         Index           =   6
         Left            =   105
         TabIndex        =   18
         Top             =   180
         Width           =   1620
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks:"
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
         Left            =   720
         TabIndex        =   17
         Top             =   540
         Width           =   840
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   120
      TabIndex        =   22
      Top             =   2490
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmMealVoucher.frx":33E4
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmMealVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmMealVoucher"

Private WithEvents oTrans As ggcMealVoucher.clsMealVoucher
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Private pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim lnMsg As String

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   Select Case Index
   Case 0   'Save
      If isEntryOk Then
         If oTrans.SaveTransaction Then
            MsgBox "Transaction save successfully!", vbInformation, pxeMODULENAME
            initButton (xeModeReady)
         Else
            MsgBox "Unable to save transaction!", vbInformation, pxeMODULENAME
         End If
      End If
   Case 1   'search
      If pnIndex = 1 Or pnIndex = 7 Then
            oTrans.SearchMaster pnIndex, txtField(pnIndex).Text
            If txtField(pnIndex).Text <> "" Then SetNextFocus
            Else
               If txtField(pnIndex).Text <> "" Then oTrans.SearchMaster pnIndex, txtField(pnIndex).Text
      End If
   Case 2   'Browse
         If oTrans.SearchTransaction = True Then
            Call LoadRecord
         Else
            lnMsg = MsgBox("No record Found!", vbCritical, pxeMODULENAME)
         End If
   Case 3   'Cancel
         lnMsg = MsgBox("Do you want to discard changes?", vbYesNo + vbQuestion, "Confirm")
         If lnMsg = vbYes Then
            ClearFields
            Call initButton(xeModeReady)
         End If
   Case 4 ' New
         If oTrans.NewTransaction = True Then
            Call ClearFields
            Call LoadRecord
            Call initButton(xeModeAddNew)
            txtField(1).SetFocus
         End If
   Case 5 ' Close
            Unload Me
   Case 6 ' Cancel Transaction
         If txtField(1) = "" And txtField(2) = "" Then
         MsgBox "Please select a record to cancel!!!", vbCritical, pxeMODULENAME
         Exit Sub
         End If
         If oTrans.Master("cTranSTat") = xeStateOpen Then
            lnMsg = MsgBox("Do you want to cancel this transaction???", vbYesNo + vbInformation, "Confirm")
               If lnMsg = vbYes Then
                  If oTrans.CancelTransaction Then
                  MsgBox "Transaction successfully cancelled!!!", vbInformation, pxeMODULENAME
                  ClearFields
                  initButton (xeModeReady)
                  End If
               End If
         Else
            MsgBox "Cannot update transaction already " + lblField(3).Caption + "!!!", vbCritical, pxeMODULENAME
         End If
   Case 7 ' Update
         If txtField(1) = "" And txtField(2) = "" Then
         MsgBox "Please select a record to update!!!", vbCritical, pxeMODULENAME
         Exit Sub
         End If
         If oTrans.Master("cTranSTat") = xeStateOpen Then
            If oTrans.UpdateTransaction Then
               initButton (xeModeUpdate)
               txtField(2).SetFocus
            End If
         Else
            MsgBox "Cannot update transaction already " + lblField(3).Caption + "!!!", vbCritical, pxeMODULENAME
         End If
   End Select

endProc:
   Exit Sub
endWithFocus:
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   'On Error GoTo errProc
   
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
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcMealVoucher.clsMealVoucher
   Set oTrans.AppDriver = oApp
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   
   ClearFields

   oTrans.InitTransaction
   oTrans.NewTransaction
   Call LoadRecord
   
   initButton (xeModeAddNew)

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub LoadRecord()

   With oTrans
      txtField(0) = IFNull(Format(.Master("sTransNox"), "@@@@-@@-@@@@@@"), "")
      txtField(1) = IFNull(Format(.Master("dTransact"), "MMM DD, YYYY"), "")
      txtField(2) = IFNull(.Master("sMealDesc"), "")
      txtField(3) = IFNull(Format(.Master("sContrlNo"), "@@@@@-@@@"), "")
      txtField(4) = IFNull(.Master("sClientNm"), "")
      txtField(5) = IFNull(.Master("sRemarksx"), "")
      txtField(6) = IFNull(Format(.Master("sVouchrNo"), "@@@@-@@-@@@@"), "")
      setTransTat (oTrans.Master("cTransTat"))
   End With
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   setTransTat -1
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(2).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(7).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow

   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(3).Visible = lbShow

   txtField(1).Enabled = lbShow
   txtField(2).Enabled = lbShow
   txtField(4).Enabled = lbShow
   txtField(5).Enabled = lbShow
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   With oTrans
      Select Case Index
         Case 0
              txtField(Index) = IFNull(.Master("sTransNox"), "")
         Case 1
              txtField(Index) = IFNull(Format(.Master("dTransact"), "MMM DD, YYYY"), "")
         Case 2
              txtField(Index) = IFNull(.Master("sMealDesc"), "")
         Case 3
              txtField(Index) = IFNull(.Master("sContrlNo"), "")
         Case 4
              txtField(Index) = IFNull(.Master("sClientNm"), "")
         Case 5
              txtField(Index) = IFNull(.Master("sRemarksx"), "")
         Case 6
              txtField(Index) = IFNull(.Master("sVouchrNo"), "")
      End Select
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Then
            oTrans.SearchMaster Index, .Text
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then oTrans.SearchMaster Index, .Text
         End If
         
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
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

Public Function setTransTat(lnStat As Integer)
   Select Case lnStat
      Case 0
         lblField(3) = "OPEN"
      Case 1
         lblField(3) = "PRINTED"
      Case 2
         lblField(3) = "POSTED"
      Case 3
         lblField(3) = "CANCELLED"
      Case 4
         lblField(3) = "VOIDED"
      Case Else
         lblField(3) = "UNKNOWN"
   End Select
End Function

Private Function isEntryOk() As Boolean
   If txtField(2).Text = "" Then
      MsgBox "Empty Meal Type Detected!" & vbCrLf & _
               "Pls Verify Entry Then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If

   If txtField(4).Text = "" Then
      MsgBox "Empty Employee Name Detected!" & vbCrLf & _
               "Pls Verify Entry Then Try Again!!!", vbCritical, "Warning"
      txtField(4).SetFocus
      GoTo EntryNotOK
   End If

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
With txtField(Index)
      Select Case Index
      Case 1
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
               .Text = Format(.Text, "MMM DD, YYYY")
         oTrans.Master("dTransact") = txtField(Index)
      Case 3
         oTrans.Master("sContrlNo") = txtField(Index)
      Case 5
         .Text = TitleCase(.Text)
         oTrans.Master("sRemarksx") = txtField(Index)
      End Select
     
   End With
End Sub
