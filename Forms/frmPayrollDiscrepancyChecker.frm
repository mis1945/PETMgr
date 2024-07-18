VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmPayrollDiscrepancyChecker 
   BorderStyle     =   0  'None
   Caption         =   "Payroll Discrepancy Checker"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   615
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   1085
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
         Index           =   1
         Left            =   5190
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         Top             =   90
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
         Height          =   400
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   1
         Top             =   85
         Width           =   2190
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIOD TO:"
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
         Index           =   4
         Left            =   3990
         TabIndex        =   2
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIOD FROM:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   0
         Top             =   165
         Width           =   1425
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5250
      Index           =   1
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1155
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   9260
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Frame chkFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   2310
         Left            =   30
         TabIndex        =   20
         Tag             =   "wt0;fb0"
         Top             =   2865
         Width           =   9495
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   2085
            Left            =   105
            TabIndex        =   21
            Tag             =   "wt0;fb0"
            Top             =   120
            Width           =   6885
            Begin VB.ComboBox lstBox 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "frmPayrollDiscrepancyChecker.frx":0000
               Left            =   1815
               List            =   "frmPayrollDiscrepancyChecker.frx":000D
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   1185
               Width           =   2250
            End
            Begin VB.TextBox txtField 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   0
               Left            =   1830
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   210
               Width           =   4470
            End
            Begin VB.TextBox txtField 
               Appearance      =   0  'Flat
               Height          =   630
               Index           =   1
               Left            =   1830
               MultiLine       =   -1  'True
               TabIndex        =   8
               Top             =   540
               Width           =   4470
            End
            Begin VB.Label lblField 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "REMARKS:"
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
               Index           =   2
               Left            =   135
               TabIndex        =   7
               Top             =   585
               Width           =   900
            End
            Begin VB.Label lblField 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CHECK STATUS:"
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
               Index           =   1
               Left            =   135
               TabIndex        =   9
               Top             =   1230
               Width           =   1365
            End
            Begin VB.Label lblField 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "EMPLOYEE NAME:"
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
               Left            =   135
               TabIndex        =   5
               Top             =   255
               Width           =   1515
            End
         End
         Begin VB.CommandButton chkButton 
            Caption         =   "Deduction"
            Height          =   375
            Index           =   4
            Left            =   7095
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1830
            Width           =   2295
         End
         Begin VB.CommandButton chkButton 
            Caption         =   "Gov. Deduction"
            Height          =   375
            Index           =   3
            Left            =   7095
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1425
            Width           =   2295
         End
         Begin VB.CommandButton chkButton 
            Caption         =   "Benefits"
            Height          =   375
            Index           =   2
            Left            =   7095
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1020
            Width           =   2295
         End
         Begin VB.CommandButton chkButton 
            Caption         =   "Loans"
            Height          =   375
            Index           =   1
            Left            =   7095
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   615
            Width           =   2295
         End
         Begin VB.CommandButton chkButton 
            Appearance      =   0  'Flat
            Caption         =   "Timesheet"
            Height          =   375
            Index           =   0
            Left            =   7095
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   210
            Width           =   2295
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2850
         Left            =   30
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   30
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5027
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1140
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
      Picture         =   "frmPayrollDiscrepancyChecker.frx":003C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   17
      Top             =   510
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmPayrollDiscrepancyChecker.frx":07B6
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   18
      Top             =   510
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
      Picture         =   "frmPayrollDiscrepancyChecker.frx":0F30
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1140
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
      Picture         =   "frmPayrollDiscrepancyChecker.frx":16AA
   End
End
Attribute VB_Name = "frmPayrollDiscrepancyChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeModuleName = "frmPayrollDiscrepancyChecker"

Private oSkin As clsFormSkin
Private bLoaded As Boolean

Private WithEvents p_oTrans As ggcPayroll.clsPayrollChecker
Attribute p_oTrans.VB_VarHelpID = -1

Private pnIndex As Integer
Private pnRow As Integer
Private pnActiveRow As Integer

Private pbCtrlPress As Boolean
Private pbFormLoad As Boolean
Private pbDetailGotFocus As Boolean

Private pdDate As Date

Private psSelectedTime As String
Private psTimeTempStrg As String
Private pnLastSelc As Integer

Private loRSTimesheet As Recordset
Private loRSLoans As Recordset
Private loRSBenefits As Recordset
Private loRSGovtDedc As Recordset
Private loRSDeductns As Recordset

Private Function retreiveDetail()
   Dim lnCtr As Integer
   Dim loRSTimesheet As Recordset
   Dim loRSLoans As Recordset
   Dim loRSBenefits As Recordset
   Dim loRSGovtDedc As Recordset
   Dim loRSDeductns As Recordset
   
   Dim ldValue As Date
   
   retreiveDetail = False
   
   If IsDate(txtSearch(0)) Then
      ldValue = CDate(txtSearch(0).Text)
   Else
      ldValue = CDate("1900-01-01")
   End If
   
   With p_oTrans
      If .SearchTransaction(ldValue) Then
         LoadDetail
         InitForm
      End If
      
      retreiveDetail = True
   End With
End Function

Private Sub LoadDetail()
   Dim lnRow As Integer
   
   With MSFlexGrid1
      .Rows = p_oTrans.ItemCount + 1
      
      For lnRow = 0 To p_oTrans.ItemCount - 1
         .TextMatrix(lnRow + 1, 0) = lnRow + 1
         .TextMatrix(lnRow + 1, 1) = p_oTrans.Detail(lnRow, "xbranchnm")
         .TextMatrix(lnRow + 1, 2) = p_oTrans.Detail(lnRow, "xemploynm")
         .TextMatrix(lnRow + 1, 3) = p_oTrans.Detail(lnRow, "sremarksx")
      Next
      
      If .Rows > 11 Then
         .ColWidth(3) = 3640
      Else
         .ColWidth(3) = 3890
      End If
      
      MSFlexGrid1_Click
   End With
End Sub

Private Sub InitForm()
   Dim lnCtr As Integer
   Dim loTxt As TextBox
   Dim lbShow As Boolean

   lbShow = p_oTrans.EditMode = xeModeUpdate
   For Each loTxt In txtField
      loTxt.BackColor = oApp.getColor("EB")
   Next

   Frame1.Enabled = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   
   cmdButton(0).Visible = Not lbShow
   cmdButton(1).Visible = Not lbShow
   txtSearch(0).Enabled = Not lbShow
   txtSearch(1).Enabled = Not lbShow
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Cols = 4
      .Rows = 2
      
      .Clear
      
      .Row = 0
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "Name"
      .TextMatrix(0, 3) = "Remarks"
      .RowHeightMin = 250
      
      'column width
      .ColWidth(0) = 370
      .ColWidth(1) = 2600
      .ColWidth(2) = 2600
      .ColWidth(3) = 3890
      
      'column alignment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      
      .Rows = 2
      
      .Row = 1
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub chkButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lors As Recordset
   Dim lnRow As Integer
   lnRow = pnActiveRow - 1
   
   Select Case Index
      Case 0 'Timesheet
         With p_oTrans
            If .EditMode = xeModeUpdate Or _
               .EditMode = xeModeReady Then
               Set lors = p_oTrans.getTimesheet(.Detail(lnRow, "sPayPerID"), .Detail(lnRow, "sEmployID"))
      
               If TypeName(lors) = "Nothing" Then
                  MsgBox "No timesheet found for this employee.", vbInformation, "Notice"
               Else
                  Dim loFormtimesheet As frmEmployeeTimesheet
                  Set loFormtimesheet = New frmEmployeeTimesheet
                  
                  Set loFormtimesheet.Data = lors
                  loFormtimesheet.Show
               End If
            End If
         End With
      Case 1 'Loans
         With p_oTrans
            If .EditMode = xeModeUpdate Or _
               .EditMode = xeModeReady Then
               Set lors = p_oTrans.getLoans(.Detail(lnRow, "sPayPerID"), .Detail(lnRow, "sEmployID"))
         
               If TypeName(lors) = "Nothing" Then
                  MsgBox "No loan found for this employee.", vbInformation, "Notice"
               Else
                  Dim loFormpayloans As frmPayrollLoans
                  Set loFormpayloans = New frmPayrollLoans
                  
                  Set loFormpayloans.Data = lors
                  loFormpayloans.Show
               End If
            End If
         End With
      Case 2 'Benifits
         With p_oTrans
            If .EditMode = xeModeUpdate Or _
               .EditMode = xeModeReady Then
               Set lors = p_oTrans.getBenefit(.Detail(lnRow, "sPayPerID"), .Detail(lnRow, "sEmployID"))
         
               If TypeName(lors) = "Nothing" Then
                  MsgBox "No Benefits found for this employee.", vbInformation, "Notice"
               Else
               
                  Dim loFormpaybenefit As frmPayrollBenefits
                  Set loFormpaybenefit = New frmPayrollBenefits
                  
                  Set loFormpaybenefit.Data = lors
                  loFormpaybenefit.Show
               End If
            End If
         End With
      Case 3 'Gov. Deduction
         With p_oTrans
            If .EditMode = xeModeUpdate Or _
               .EditMode = xeModeReady Then
               Set lors = p_oTrans.getGovtDeduction(.Detail(lnRow, "sPayPerID"), .Detail(lnRow, "sEmployID"))
         
               If TypeName(lors) = "Nothing" Then
                  MsgBox "No Gov. Deduction found for this employee.", vbInformation, "Notice"
               Else
               
                  Dim loFormpaygovdeduc As frmPayrollGovDeduction
                  Set loFormpaygovdeduc = New frmPayrollGovDeduction
                  
                  Set loFormpaygovdeduc.Data = lors
                  loFormpaygovdeduc.Show
               
               End If
            End If
         End With
      Case 4 'Deduction
         With p_oTrans
            If .EditMode = xeModeUpdate Or _
               .EditMode = xeModeReady Then
               Set lors = p_oTrans.getDeductions(.Detail(lnRow, "sPayPerID"), .Detail(lnRow, "sEmployID"))
         
               If TypeName(lors) = "Nothing" Then
                  MsgBox "No deduction found for this employee.", vbInformation, "Notice"
               Else
                  Dim loFormpaydeduc As frmPayrollDeduction
                  Set loFormpaydeduc = New frmPayrollDeduction
                     
                  Set loFormpaydeduc.Data = lors
                  loFormpaydeduc.Show
               End If
            End If
         End With
   End Select
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
   Exit Sub
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lnRep As Integer
   Dim lsOldProc As String
   
   Select Case Index
   Case 0   'Retrieve
      Call retreiveDetail
   Case 1   'Close
      Unload Me
   Case 2   'Save
      If Not (CDate(txtSearch(0).Text) = p_oTrans.PeriodFrom) Then
         MsgBox "Please load the correct transaction date before processing!"
         Exit Sub
      End If
      
      If p_oTrans.SaveTransaction Then
         MsgBox "Transaction save successfully.", vbInformation, "Notice"
         
         pnRow = 0
         p_oTrans.InitTransaction
         
         InitGrid
         InitForm
         ClearFields
      End If
   Case 3   'Cancel
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
      If lnRep = vbYes Then
         p_oTrans.InitTransaction
         pnRow = 0
         InitGrid
         InitForm
         ClearFields
      End If
   End Select
End Sub

Private Sub ClearFields()
   txtField(0) = ""
   txtField(1) = ""
   txtSearch(0) = ""
   txtSearch(1) = ""
   lstBox.ListIndex = 0
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   'On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
      
      pnActiveRow = 1
      pnRow = 2
   End If
   
   If Not pbFormLoad Then pbFormLoad = True
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
   Exit Sub
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

Private Sub lstBox_LostFocus()
   p_oTrans.Detail(pnRow, "cChckStat") = lstBox.ListIndex
End Sub

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      If p_oTrans.EditMode = xeModeReady Or _
         p_oTrans.EditMode = xeModeUpdate Then
         pnActiveRow = .Row
         
         Call detailFieldChange
      End If
   
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub detailFieldChange()
   Dim lnRow As Integer
   pnRow = pnActiveRow - 1
   
   With MSFlexGrid1
      With p_oTrans
         txtField(0).Text = p_oTrans.Detail(pnRow, "xemploynm")
         txtField(1).Text = p_oTrans.Detail(pnRow, "sremarksx")
         lstBox.ListIndex = CInt(p_oTrans.Detail(pnRow, "cchckstat"))
         
         Frame1.Enabled = CInt(p_oTrans.Detail(pnRow, "cchckstat")) = 0
         
         If Frame1.Enabled Then
            txtField(1).SetFocus
         Else
            chkButton(0).SetFocus
         End If
      End With
   End With
End Sub

Private Sub p_oTrans_MasterRetreived(ByVal Index As Integer, ByVal loValue As Variant)
   Select Case Index
      Case 100
         txtSearch(0) = Format(loValue, "Mmm dd, yyyy")
      Case 101
         txtSearch(1) = Format(loValue, "Mmm dd, yyyy")
   End Select
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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   If Index = 1 Then
      p_oTrans.Detail(pnRow, "sRemarksx") = txtField(Index)
      LoadDetail
   End If
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
   txtSearch(Index).BackColor = oApp.getColor("HT1")
   
   If Index = 0 Then
      If IsDate(txtSearch(Index)) Then
         txtSearch(Index) = Format(txtSearch(Index), "yyyy-mm-dd")
      End If
   End If
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 0 Then If KeyCode = vbKeyF3 Then Call retreiveDetail
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("EB")
      
      If Index = 0 Then
         If IsDate(txtSearch(Index)) Then
            txtSearch(Index) = Format(txtSearch(Index), "Mmm dd, yyyy")
         End If
      End If
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   pnActiveRow = 0
   pbFormLoad = False
   pbCtrlPress = False
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   Set p_oTrans = New ggcPayroll.clsPayrollChecker
   
   With p_oTrans
       Set .AppDriver = oApp
      .Branch = oApp.BranchCode
         
      If Not .InitTransaction Then Unload Me
   End With
     
   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   
   Call InitGrid
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
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

Private Sub txtSearch_Validate(Index As Integer, Cancel As Boolean)
   With txtSearch(Index)
      If Index = 0 Then
         If Not IsDate(.Text) Then
            .Text = ""
         Else
            .Text = Format(.Text, "MM/DD/YYYY")
         End If
      End If
   End With
End Sub
