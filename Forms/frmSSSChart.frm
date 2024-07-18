VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmSSSChart 
   BorderStyle     =   0  'None
   Caption         =   "SSS Chart"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5880
      Left            =   5925
      TabIndex        =   14
      Top             =   1050
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   10372
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   450
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   794
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Line Line3 
         X1              =   3120
         X2              =   7560
         Y1              =   210
         Y2              =   210
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   5880
      Left            =   1620
      Tag             =   "wt0;fb0"
      Top             =   1005
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   10372
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
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
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   675
         Width           =   1755
      End
      Begin VB.TextBox txtField 
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
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "7, 200.00"
         Top             =   1155
         Width           =   1755
      End
      Begin VB.TextBox txtField 
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
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "10,000.00"
         Top             =   1635
         Width           =   1755
      End
      Begin VB.TextBox txtField 
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
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "10,000.00"
         Top             =   2025
         Width           =   1755
      End
      Begin VB.TextBox txtField 
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
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   11
         Text            =   "10,000.00"
         Top             =   2415
         Width           =   1755
      End
      Begin VB.TextBox txtField 
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
         Index           =   0
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "0"
         Top             =   285
         Width           =   1755
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL CONTRIBUTION"
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
         Index           =   10
         Left            =   135
         TabIndex        =   12
         Top             =   3060
         Width           =   2175
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "30,000.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2370
         TabIndex        =   13
         Top             =   2955
         Width           =   1755
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Range Thru"
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
         Left            =   135
         TabIndex        =   2
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Credit"
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
         Left            =   135
         TabIndex        =   4
         Top             =   1215
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employer"
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
         Left            =   135
         TabIndex        =   6
         Top             =   1695
         Width           =   825
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
         Index           =   8
         Left            =   135
         TabIndex        =   8
         Top             =   2085
         Width           =   870
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Comp"
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
         Left            =   135
         TabIndex        =   10
         Top             =   2475
         Width           =   1440
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Range From"
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
         Left            =   135
         TabIndex        =   0
         Top             =   345
         Width           =   1065
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   3120
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
      Picture         =   "frmSSSChart.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   2490
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
      Picture         =   "frmSSSChart.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   2490
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
      Picture         =   "frmSSSChart.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   3120
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
      Picture         =   "frmSSSChart.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   1230
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
      Picture         =   "frmSSSChart.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1860
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
      Picture         =   "frmSSSChart.frx":2E7A
   End
End
Attribute VB_Name = "frmSSSChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmSSSChart"

Private oSkin As clsFormSkin
Private WithEvents oTrans As clsSSSChart
Attribute oTrans.VB_VarHelpID = -1
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim pnRow As Integer
Dim poCtrl As Variant

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRow As Integer
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
'   On Error GoTo errProc

   Select Case Index
   Case 1 'Add Detail
      oTrans.addDetail
      With MSHFlexGrid1
         .Rows = MSHFlexGrid1.Rows + 1
         .TopRow = MSHFlexGrid1.Rows - 10
         .Row = MSHFlexGrid1.Rows - 1
         .TextMatrix(.Row, 0) = "0.00"
         .TextMatrix(.Row, 1) = "0.00"
         .TextMatrix(.Row, 2) = "0.00"
         .TextMatrix(.Row, 3) = "0.00"
         .TextMatrix(.Row, 4) = "0.00"
         .TextMatrix(.Row, 5) = "0.00"
      End With
   Case 5 'Delete Detail
      Call oTrans.DeleteDetail(MSHFlexGrid1.Row - 1)
      Call MSHFlexGrid1.RemoveItem(MSHFlexGrid1.Row)
   Case 2 'Close
      Unload Me
   Case 3 'Update
      If oTrans.UpdateTransaction Then
         Call InitForm(xeModeUpdate)
      End If
   Case 4 'Cancel Update
      If MsgBox("Do you really want to undo changes?", vbCritical + vbYesNo, "Confirmation") = vbYes Then
         If oTrans.OpenTransaction Then
            Call LoadDetail
            Call InitForm(xeModeReady)
         End If
      End If
   Case 0 'Save
      If oTrans.SaveTransaction Then
         MsgBox "Updates save successfully!", vbOKOnly, "Confirmation"
         oTrans.OpenTransaction
         Call LoadDetail
         Call InitForm(xeModeReady)
      Else
         MsgBox "Unable to save changes!" & vbCrLf & _
                "Please check your entry and try again...", vbOKOnly, "Confirmation"
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
'   On Error GoTo errProc

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
'   On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsSSSChart
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.OpenTransaction
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

   Call InitGrid
   Call ClearFields(-1)
   Call InitForm(xeModeReady)
   
   Call LoadDetail

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub ClearFields(ByVal fnRow As Integer)
   Dim loTxt As TextBox
   Dim lnTotal As Currency
   
   If fnRow < 0 Then
      For Each loTxt In txtField
         loTxt = "0.00"
      Next
      
      lblTotal = "0.00"
   Else
      For Each loTxt In txtField
         loTxt = Format(oTrans.Detail(fnRow, loTxt.Index), "#,##0.00")
      Next
      
      lblTotal = Format(CCur(txtField(3)) + CCur(txtField(4)) + CCur(txtField(5)), "#,##0.00")
   End If
   
End Sub

Private Sub MSHFlexGrid1_Click()
   If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> "" Then
      Call ClearFields(MSHFlexGrid1.Row - 1)
      If xrFrame3.Enabled Then txtField(0).SetFocus
   Else
      Call ClearFields(-1)
   End If
End Sub

Private Sub MSHFlexGrid1_SelChange()
   With MSHFlexGrid1
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer, ByVal Value As Variant)
   txtField(Index) = Format(Value, "#,##0.00")
   lblTotal = Format(CCur(txtField(3)) + CCur(txtField(4)) + CCur(txtField(5)), "#,##0.00")
   
   MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, Index) = txtField(Index)
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
      Set poCtrl = txtField(Index)
   End With

   pnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub InitForm(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(4).Visible = lbShow
   cmdButton(5).Visible = lbShow
   
   cmdButton(2).Visible = Not lbShow
   cmdButton(3).Visible = Not lbShow
   
   xrFrame3.Enabled = lbShow
   
   If lbShow Then
      txtField(0).SetFocus
   End If
End Sub


Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSHFlexGrid1
      .Cols = 6
      .Rows = 2
      
      .Clear
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Row = 1
      
      'column title
      .TextMatrix(0, 0) = "From"
      .TextMatrix(0, 1) = "Thru"
      .TextMatrix(0, 2) = "Sal Crdt"
      .TextMatrix(0, 3) = "Employer"
      .TextMatrix(0, 4) = "Employee"
      .TextMatrix(0, 5) = "EC"
      
      'column width
      .ColWidth(0) = 1000
      .ColWidth(1) = 1000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000

      'column allinment
      .ColAlignment(0) = flexAlignRightCenter
      .ColAlignment(1) = flexAlignRightCenter
      .ColAlignment(2) = flexAlignRightCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      
      'set location
      .Row = 1
      .Col = 0
      .ColSel = .Cols - 1
      
      pnRow = 0
   End With
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   Dim ln13thMnth As Currency
   Dim lnBonusxx1 As Currency
   Dim lnPartyxx1 As Currency
   
   With MSHFlexGrid1
      .Rows = 2
      If oTrans.ItemCount = 0 Then
         .TextMatrix(1, 1) = "0.00"
         .TextMatrix(1, 2) = "0.00"
         .TextMatrix(1, 3) = "0.00"
         .TextMatrix(1, 4) = "0.00"
         .TextMatrix(1, 5) = "0.00"
      Else
         .Rows = oTrans.ItemCount + 1
         For lnCtr = 0 To oTrans.ItemCount - 1
            DoEvents
            .TextMatrix(lnCtr + 1, 0) = Format(oTrans.Detail(lnCtr, "sRangeFrm"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 1) = Format(oTrans.Detail(lnCtr, "sRangeTru"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 2) = Format(oTrans.Detail(lnCtr, "nSalCredt"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 3) = Format(oTrans.Detail(lnCtr, "nEmployer"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 4) = Format(oTrans.Detail(lnCtr, "nEmployee"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 5) = Format(oTrans.Detail(lnCtr, "nECProgrm"), "#,##0.00")
         Next
      End If
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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   oTrans.Detail(MSHFlexGrid1.Row - 1, Index) = CCur(txtField(Index))
   txtField(Index) = Format(oTrans.Detail(MSHFlexGrid1.Row - 1, Index), "#,##0.00")
   
   Select Case Index
   Case 3, 4, 5
      lblTotal = Format(CCur(txtField(3)) + CCur(txtField(4)) + CCur(txtField(5)), "#,##0.00")
   End Select
End Sub
