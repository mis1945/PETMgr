VERSION 5.00
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmServiceCharge 
   BorderStyle     =   0  'None
   Caption         =   "Service Charge EntryIncentive"
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3045
      Index           =   0
      Left            =   1665
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   5371
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.OptionButton optTaxable 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Non Taxable"
         Height          =   330
         Index           =   1
         Left            =   1410
         TabIndex        =   17
         Tag             =   "wt0;fb0"
         Top             =   1095
         Width           =   1230
      End
      Begin VB.OptionButton optTaxable 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Taxable"
         Height          =   330
         Index           =   0
         Left            =   1410
         TabIndex        =   16
         Tag             =   "wt0;fb0"
         Top             =   780
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   7965
         TabIndex        =   13
         Top             =   1740
         Width           =   1965
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   705
         Index           =   5
         Left            =   1290
         TabIndex        =   2
         Top             =   2250
         Width           =   8685
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1290
         TabIndex        =   1
         Top             =   1875
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1290
         TabIndex        =   0
         Top             =   1515
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   165
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   19
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incentive"
         Height          =   195
         Index           =   5
         Left            =   450
         TabIndex        =   18
         Top             =   825
         Width           =   660
      End
      Begin VB.Shape Shape2 
         Height          =   720
         Left            =   1290
         Top             =   750
         Width           =   1920
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6900
         TabIndex        =   12
         Top             =   1815
         Width           =   885
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   3
         Left            =   345
         TabIndex        =   7
         Top             =   2295
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refer. Date"
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   6
         Top             =   1935
         Width           =   825
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. Date"
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   5
         Top             =   1575
         Width           =   840
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   300
         TabIndex        =   4
         Top             =   210
         Width           =   675
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   330
         Left            =   1110
         Tag             =   "et0;ht2"
         Top             =   285
         Width           =   2265
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   555
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
      Picture         =   "frmServiceCharge.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   9
      Top             =   1200
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
      Picture         =   "frmServiceCharge.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   555
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
      Picture         =   "frmServiceCharge.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   11
      Top             =   1830
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
      Picture         =   "frmServiceCharge.frx":166E
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4140
      Left            =   1665
      TabIndex        =   14
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   3645
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   7303
      AllowBigSelection=   -1  'True
      AutoAdd         =   -1  'True
      AutoNumber      =   -1  'True
      BACKCOLOR       =   -2147483643
      BACKCOLORBKG    =   8421504
      BACKCOLORFIXED  =   -2147483633
      BACKCOLORSEL    =   -2147483635
      BORDERSTYLE     =   1
      COLS            =   2
      FILLSTYLE       =   0
      FIXEDCOLS       =   1
      FIXEDROWS       =   1
      FOCUSRECT       =   1
      EDITORBACKCOLOR =   -2147483643
      EDITORFORECOLOR =   -2147483640
      FORECOLOR       =   -2147483640
      FORECOLORFIXED  =   -2147483630
      FORECOLORSEL    =   -2147483634
      FORMATSTRING    =   ""
      Object.HEIGHT          =   4140
      GRIDCOLOR       =   12632256
      GRIDCOLORFIXED  =   0
      BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GRIDLINES       =   1
      GRIDLINESFIXED  =   2
      GRIDLINEWIDTH   =   1
      MOUSEICON       =   "frmServiceCharge.frx":1DE8
      MOUSEPOINTER    =   0
      REDRAW          =   -1  'True
      RIGHTTOLEFT     =   0   'False
      ROWS            =   2
      SCROLLBARS      =   3
      SCROLLTRACK     =   0   'False
      SELECTIONMODE   =   0
      Object.TOOLTIPTEXT     =   ""
      WORDWRAP        =   0   'False
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   75
      TabIndex        =   15
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Register"
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
      Picture         =   "frmServiceCharge.frx":1E04
   End
End
Attribute VB_Name = "frmServiceCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmServiceCharge"

Private oTrans As clsServiceChargeDistribution
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnCtr As Integer, pnIndex As Integer
Dim pbGridFocus As Boolean, pbSave As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   
   txtField_LostFocus pnIndex
   With GridEditor1
      Select Case Index
      Case 0 'Save
         If txtField(0).Text <> "" And txtField(3).Text > 0 Then
            If oTrans.SaveTransaction(True) = True Then
               MsgBox "Transaction Save Successfully!!", vbInformation, "Notice"
               InitForm
               initButton 0
               ClearFields
            Else
               MsgBox "Unable to save Transaction!!!"
               GoTo errProc
            End If
         End If
      Case 1 'New
         oTrans.InitTransaction
         oTrans.NewTransaction
         ClearFields
         Call initButton(1)
      Case 2 'Cancel
         ClearFields
         initButton (0)
      Case 3 'Close
         Unload Me
      Case 4 'Register
         frmServiceChargeReg.Show
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   With GridEditor1
      .Refresh
   End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hWnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsServiceChargeDistribution
   Set oTrans.AppDriver = oApp
   
   oTrans.TransStatus = 210
   oTrans.InitTransaction
   oTrans.NewTransaction
      
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   InitForm
   ClearFields
   initButton xeModeAddNew

   oTrans.Master("sBeneftID") = "11003"
   Call LoadDetail
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      If .Col = 5 Then
         oTrans.Detail(.Row - 1, 6) = CDbl(.TextMatrix(.Row, .Col))
         .TextMatrix(.Row, .Col) = Format(oTrans.Detail(.Row - 1, 6), "#,##0.00")
      End If
      Call ComputeTotal
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      
   End With
   pbGridFocus = True
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_KeyDown"
   ''On Error GoTo errProc
   
   If KeyCode = vbKeyReturn Then
      With GridEditor1
         If .Col = 6 Then
            .TextMatrix(.Row, .Col) = Format(CDbl(.TextMatrix(.Row, .Col)), "#,##0.00")
            .Refresh
            .SetFocus
         End If
      End With
      KeyCode = 0
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
   End With
End Sub

Private Sub optTaxable_Click(Index As Integer)
    If optTaxable(0).Value = Checked Then
        oTrans.Master("sBeneftID") = "11003" 'taxable
    Else
        oTrans.Master("sBeneftID") = "11009"
    End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 1 Then .Text = Format(.Text, "MM/DD/YY")
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   pbGridFocus = False
   pnIndex = Index
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(1).Visible = Not lbShow
   cmdButton(3).Visible = Not lbShow
   
   cmdButton(0).Visible = lbShow
   cmdButton(2).Visible = lbShow

End Sub

Private Sub InitForm()
   Dim lnCtr As Integer

   With GridEditor1
      .Cols = 6
      .Rows = 2
      .Font = "MS Sans Serif"

      'Column Title
      .TextMatrix(0, 1) = "Employee Name"
      .TextMatrix(0, 2) = "Dept. Name"
      .TextMatrix(0, 3) = "Position"
      .TextMatrix(0, 4) = "Emp. Type"
      .TextMatrix(0, 5) = "Amount"
      
      .Row = 0
      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'Column Width
      .ColWidth(0) = 500
      .ColWidth(1) = 3000
      .ColWidth(2) = 2500
      .ColWidth(3) = 1800
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000

      .ColEnabled(1) = False
      .ColEnabled(2) = False
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      
      .Row = 1
      
      oTrans.Master("sBeneftID") = "11003"

   End With
   
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   For Each loTxt In txtField
      pnCtr = loTxt.Index
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@"))
      Case 1, 2
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      End Select
   Next
   
   With GridEditor1
      .Rows = 2
      .Col = 1
      
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = 0#
   End With
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 1, 2 ' date
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
         oTrans.Master(Index) = .Text
      Case 3
         If Not IsNumeric(.Text) Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
         oTrans.Master(Index) = CDbl(.Text)
      Case 5
         oTrans.Master(Index) = .Text
      End Select
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

Private Sub LoadDetail()
   Dim lnCtr As Integer
   
   With GridEditor1
      .Rows = oTrans.ItemCount + 1
      For lnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, 2)
         .TextMatrix(lnCtr + 1, 2) = oTrans.Detail(lnCtr, 3)
         .TextMatrix(lnCtr + 1, 3) = oTrans.Detail(lnCtr, 4)
         .TextMatrix(lnCtr + 1, 4) = oTrans.Detail(lnCtr, 5)
         .TextMatrix(lnCtr + 1, 5) = "0.00"
      Next
   End With
End Sub

Private Sub ComputeTotal()
   Dim lnCtr As Integer
   Dim lnTotal As Double
   
   lnTotal = 0#
   
   With GridEditor1
      For lnCtr = 0 To oTrans.ItemCount - 1
         lnTotal = lnTotal + oTrans.Detail(lnCtr, 6)
      Next
   End With
   
   txtField(3).Text = Format(lnTotal, "#,#0.00")
   oTrans.Master("nTotalAmt") = lnTotal
End Sub
