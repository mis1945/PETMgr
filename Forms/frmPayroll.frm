VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmPayroll 
   BorderStyle     =   0  'None
   Caption         =   "Payroll Processing"
   ClientHeight    =   10035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7515
      Left            =   1590
      TabIndex        =   17
      Top             =   2400
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   13256
      _Version        =   393216
      Cols            =   3
      FixedCols       =   2
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1800
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   3175
      BackColor       =   12632256
      BorderStyle     =   1
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
         Index           =   4
         Left            =   5460
         TabIndex        =   10
         Text            =   "August 28, 2011"
         Top             =   1245
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
         Index           =   3
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "August 13, 2011"
         Top             =   1245
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
         Height          =   400
         Index           =   2
         Left            =   5460
         TabIndex        =   6
         Text            =   "August 31, 2011"
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
         Height          =   400
         Index           =   1
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "August 16, 2011"
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
         Height          =   400
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2190
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
         Left            =   5490
         TabIndex        =   2
         Tag             =   "eb0;et0"
         Top             =   165
         Width           =   2070
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   5445
         Top             =   135
         Width           =   2160
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   5415
         Top             =   105
         Width           =   2220
      End
      Begin VB.Image Image1 
         Height          =   1515
         Left            =   7860
         Top             =   120
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coverage Thru"
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
         Left            =   4110
         TabIndex        =   9
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coverage From"
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
         Left            =   135
         TabIndex        =   7
         Top             =   1325
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period Thru"
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
         Left            =   4110
         TabIndex        =   5
         Top             =   870
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period From"
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
         TabIndex        =   3
         Top             =   875
         Width           =   1065
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
         Top             =   200
         Width           =   1485
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   400
         Left            =   1725
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   2190
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1800
      Left            =   12135
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3175
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.Image Image2 
         Height          =   1635
         Left            =   60
         Picture         =   "frmPayroll.frx":0000
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2835
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   11
      Top             =   1200
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
      Picture         =   "frmPayroll.frx":DC9C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Confirm"
      AccessKey       =   "Confirm"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPayroll.frx":E416
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   16
      Top             =   2460
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "C&lose"
      AccessKey       =   "l"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPayroll.frx":EB90
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   15
      Top             =   1830
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Process"
      AccessKey       =   "Process"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPayroll.frx":F30A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   13
      Top             =   1830
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
      Picture         =   "frmPayroll.frx":FA84
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   12
      Top             =   585
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Print"
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
      Picture         =   "frmPayroll.frx":101FE
   End
End
Attribute VB_Name = "frmPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeModuleName = "frmPayroll"
Private Const pxeREG_OTRTE As Single = 1.25

Private p_oProgress As clsSpeedometer

Private WithEvents oTrans As ggcPayroll.clsPayrollNew
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private psEmpTypID As String
Private pcMainOffc As String
Dim poReport As clsReport
Dim poRecord As Recordset
Dim pbWithNeg As Boolean
Dim pbLoad As Boolean
Dim pnIndex As Integer

Property Let EmployeeType(ByVal Value As String)
   psEmpTypID = Value
End Property

Property Let isMainOffice(ByVal Value As Boolean)
   pcMainOffc = IIf(Value, 1, 0)
End Property

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0 'Save
      If oTrans.SaveTransaction Then
         If MsgBox("Payroll period created successfully..." & vbCrLf & _
                   "Process payroll?", vbYesNo + vbInformation, "Payroll") = vbYes Then
            Call oTrans.OpenTransaction(oTrans.Master("spayperid"))
            oTrans.CloseTransaction
         End If
         Call loadGrid
         initButton oTrans.EditMode
      End If
   Case 1 'Confirm/Post
      If pbWithNeg Then
         If MsgBox("This payroll has an employee with negative NET PAY!" & vbCrLf & _
                "Do you want to continue?", vbCritical + vbOKCancel, "Payroll Info") = vbCancel Then
            Exit Sub
         End If
      End If
      
      If LCase(oApp.ProductID) = "petmgr" Then
         If oTrans.PostTransaction Then
            MsgBox "Payroll Posted Successfully"
         End If
         Call LoadMaster
      End If
   Case 2 'Process/Close
      If oTrans.CloseTransaction Then
         MsgBox "Payroll created successfully!"
      End If
      Call LoadMaster
      Call loadGrid
   Case 3 'CancelUpdate
      oTrans.NewTransaction
      Call loadGrid
      initButton IIf(oTrans.Master("cPeriodxx") = "0", xeModeAddNew, xeModeUpdate)
   Case 4 'Close/Exit Form
      Unload Me
   Case 5 'Print
      Call ReportTrans(ViewReport)
   End Select

End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   With MSFlexGrid1
      .Refresh
   End With
      
   If Not pbLoad Then
      pbLoad = True
      initButton IIf(oTrans.Master("cPostedxx") = "0", xeModeAddNew, xeModeUpdate)
   End If
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
'   'On Error GoTo errProc

   CenterChildForm mdiMain, Me
      
   Set oTrans = New clsPayrollNew
   Set oTrans.AppDriver = oApp
   oTrans.EmployeeType = psEmpTypID
   oTrans.isMainOffice = pcMainOffc
   oTrans.TransStatus = 10
   oTrans.InitTransaction
   oTrans.NewTransaction
         
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   
   Set poReport = New clsReport
   
   LoadMaster
   InitGrid
   If oTrans.EditMode = xeModeUpdate Then loadGrid
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oTrans = Nothing
   pbLoad = False
End Sub

'----------------------------
'InitGrid()
'  'Set the property rows, columns, etc...
'----------------------------
Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Rows = 2
      .Cols = 17
      
      .Row = 0
      .RowHeight(0) = 320
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Row = 1
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Employee"
      .TextMatrix(0, 2) = "Branch"
      .TextMatrix(0, 3) = "Basic Pay"
      .TextMatrix(0, 4) = "Attendance"
      .TextMatrix(0, 5) = "Absences"
      .TextMatrix(0, 6) = "Tardiness"
      .TextMatrix(0, 7) = "Undertime"
      .TextMatrix(0, 8) = "Holiday"
      .TextMatrix(0, 9) = "Overtime"
      .TextMatrix(0, 10) = "Benefits"
      .TextMatrix(0, 11) = "Adjustment"
      .TextMatrix(0, 12) = "Advances"
      .TextMatrix(0, 13) = "Deductions"
      .TextMatrix(0, 14) = "Loans"
      .TextMatrix(0, 15) = "Govt Dedct"
      .TextMatrix(0, 16) = "Net Pay"
      
      .RowHeightMin = 320
      
      .ColWidth(0) = 635
      .ColWidth(1) = 3280
      .ColWidth(2) = 2400
      .ColWidth(3) = 1200
      .ColWidth(4) = 1200
      .ColWidth(5) = 1200
      .ColWidth(6) = 1200
      .ColWidth(7) = 1200
      .ColWidth(8) = 1200
      .ColWidth(9) = 1200
      .ColWidth(10) = 1200
      .ColWidth(11) = 1200
      .ColWidth(12) = 1200
      .ColWidth(13) = 1200
      .ColWidth(14) = 1200
      .ColWidth(15) = 1200
      .ColWidth(16) = 1200
      
      'column allinment
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      
      'set location
      .Row = 1
      .Col = 2
      .ColSel = .Cols - 1
   End With

End Sub

'----------------------------
'loadGrid
'  'Load Payroll Processing From Grid
'----------------------------
Private Sub loadGrid()
   Dim lnCtr As Integer
   Dim lnDailyPay As Currency
   
   pbWithNeg = False
   With MSFlexGrid1
      If oTrans.ItemCount < 1 Then
         .Rows = 2
         .TextMatrix(lnCtr + 1, 0) = Format(lnCtr + 1, "0000")
         .TextMatrix(lnCtr + 1, 1) = ""
         .TextMatrix(lnCtr + 1, 2) = ""
         .TextMatrix(lnCtr + 1, 3) = ""
         .TextMatrix(lnCtr + 1, 4) = ""
         .TextMatrix(lnCtr + 1, 5) = ""
         .TextMatrix(lnCtr + 1, 6) = ""
         .TextMatrix(lnCtr + 1, 7) = ""
         .TextMatrix(lnCtr + 1, 8) = ""
         .TextMatrix(lnCtr + 1, 9) = ""
         .TextMatrix(lnCtr + 1, 10) = ""
         .TextMatrix(lnCtr + 1, 11) = ""
         .TextMatrix(lnCtr + 1, 12) = ""
         .TextMatrix(lnCtr + 1, 13) = ""
         .TextMatrix(lnCtr + 1, 14) = ""
         .TextMatrix(lnCtr + 1, 15) = ""
         .TextMatrix(lnCtr + 1, 16) = ""
      Else
         If oTrans.ItemCount > 18 Then
               .ColWidth(0) = 600
               .ColWidth(1) = 3050
            Else
               .ColWidth(0) = 635
               .ColWidth(1) = 3280
            End If
            
'         Call LoadDetail
         Call LoadDetailNew
            
'         For lnCtr = 0 To oTrans.ItemCount - 1
'         poRecord.Sort = "sBranchNm, nBasicSal DESC, sEmployNm"
         poRecord.Sort = "sBranchNm, sEmployNm"
         poRecord.MoveFirst
         
         Do Until poRecord.EOF
            DoEvents
            
'            lnDailyPay = Round(IIf(oTrans.Master("sEmpTypID") = "T", oTrans.Detail(lnCtr, "nBasicSal") / (oTrans.Detail(lnCtr, "nAttendnc") + oTrans.Detail(lnCtr, "nAbsencex")), oTrans.Detail(lnCtr, "nBasicSal") / 13), 2)
            .Rows = 2 + lnCtr
            .RowHeight(lnCtr + 1) = 338
'            .TextMatrix(lnCtr + 1, 0) = Format(lnCtr + 1, "0000")
'            .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "sEmployNm")
'            .TextMatrix(lnCtr + 1, 2) = IFNull(oTrans.Detail(lnCtr, "sBranchNm"))
'            .TextMatrix(lnCtr + 1, 3) = Format(oTrans.Detail(lnCtr, "nBasicSal"), "#,##0.00")
'            .TextMatrix(lnCtr + 1, 4) = oTrans.Detail(lnCtr, "nAttendnc")
'            .TextMatrix(lnCtr + 1, 5) = Format(Round(oTrans.Detail(lnCtr, "nAbsencex") * lnDailyPay, 2), "#,##0.00")
'            .TextMatrix(lnCtr + 1, 6) = Format(Round(oTrans.Detail(lnCtr, "nTardines") * (lnDailyPay / 8 / 60), 2), "#,##0.00")
'            .TextMatrix(lnCtr + 1, 7) = Format(Round(oTrans.Detail(lnCtr, "nUndrTime") * (lnDailyPay / 8 / 60), 2), "#,##0.00")
'            .TextMatrix(lnCtr + 1, 8) = Format(IIf(IsNumeric(oTrans.Detail(lnCtr, "nHolidayx")) = True, oTrans.Detail(lnCtr, "nHolidayx"), 0), "#,##0.00")
'            .TextMatrix(lnCtr + 1, 9) = Format(Round(oTrans.Detail(lnCtr, "nOverTime") * (lnDailyPay / 8 / 60), 2), "#,##0.00")
'            .TextMatrix(lnCtr + 1, 10) = Format(oTrans.Detail(lnCtr, "nBenefits"), "#,##0.00")
'            .TextMatrix(lnCtr + 1, 11) = Format(oTrans.Detail(lnCtr, "nAdjustxx"), "#,##0.00")
'            .TextMatrix(lnCtr + 1, 12) = Format(oTrans.Detail(lnCtr, "nAdvances"), "#,##0.00")
'            .TextMatrix(lnCtr + 1, 13) = Format(oTrans.Detail(lnCtr, "nDeductnx"), "#,##0.00")
'            .TextMatrix(lnCtr + 1, 14) = Format(oTrans.Detail(lnCtr, "nLoanAmrt"), "#,##0.00")
'            .TextMatrix(lnCtr + 1, 15) = Format(oTrans.Detail(lnCtr, "nGovDedct"), "#,##0.00")
            
            .TextMatrix(lnCtr + 1, 0) = Format(lnCtr + 1, "0000")
            .TextMatrix(lnCtr + 1, 1) = poRecord("sEmployNm")
            .TextMatrix(lnCtr + 1, 2) = poRecord("sBranchNm")
            .TextMatrix(lnCtr + 1, 3) = Format(poRecord("nBasicSal"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 4) = poRecord("nAttendnc")
            .TextMatrix(lnCtr + 1, 5) = Format(poRecord("nAbsencex"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 6) = Format(poRecord("nTardines"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 7) = Format(poRecord("nUndrTime"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 8) = Format(poRecord("nHolidayx"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 9) = Format(poRecord("nOverTime"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 10) = Format(poRecord("nBenefits"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 11) = Format(poRecord("nAdjustxx"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 12) = Format(poRecord("nAdvances"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 13) = Format(poRecord("nDeductnx"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 14) = Format(poRecord("nLoanAmrt"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 15) = Format(poRecord("nGovDedct"), "#,##0.00")
            
            DoEvents
            
            'Compute for the net Pay
            .TextMatrix(lnCtr + 1, 16) = Format(CCur(.TextMatrix(lnCtr + 1, 3)) _
                                          + (CCur(.TextMatrix(lnCtr + 1, 8)) _
                                             + CCur(.TextMatrix(lnCtr + 1, 9)) _
                                             + CCur(.TextMatrix(lnCtr + 1, 10)) _
                                             + CCur(.TextMatrix(lnCtr + 1, 11))) _
                                          - (CCur(.TextMatrix(lnCtr + 1, 5)) _
                                             + CCur(.TextMatrix(lnCtr + 1, 6)) _
                                             + CCur(.TextMatrix(lnCtr + 1, 7)) _
                                             + CCur(.TextMatrix(lnCtr + 1, 12)) _
                                             + CCur(.TextMatrix(lnCtr + 1, 13)) _
                                             + CCur(.TextMatrix(lnCtr + 1, 14)) _
                                             + CCur(.TextMatrix(lnCtr + 1, 15))) _
                                         , "#,##0.00")
            If .TextMatrix(lnCtr + 1, 16) < 0 Then
               pbWithNeg = True
            End If
                                        
'         Next
            lnCtr = lnCtr + 1
            poRecord.MoveNext
         Loop
      End If
      
      If pbWithNeg Then MsgBox "This payroll has an employee with negative NET PAY!", vbInformation + vbOKOnly, "Payroll Info"
   
   End With
End Sub

Private Sub LoadMaster()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 1 To 4
         loTxt.Text = Format(oTrans.Master(loTxt.Index), "Mmmm DD, YYYY")
      Case Else
         loTxt.Text = oTrans.Master(loTxt.Index)
      End Select
   Next
   
   Label2 = TransStat(oTrans.Master("cPostedxx"))
   
End Sub

Private Sub initButton(mode As xeEditMode)
   xrFrame1.Enabled = mode = xeModeAddNew
   cmdButton(0).Visible = xrFrame1.Enabled
   cmdButton(3).Visible = xrFrame1.Enabled

   cmdButton(4).Visible = Not xrFrame1.Enabled
   cmdButton(2).Visible = Not xrFrame1.Enabled
   cmdButton(1).Visible = Not xrFrame1.Enabled
   
   If xrFrame1.Enabled Then
      txtField(2).SetFocus
   Else
      MSFlexGrid1.SetFocus
   End If
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

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
   Case 2, 4
      txtField(Index) = Format(oTrans.Master(Index), "MM/DD/YYYY")
   End Select
   
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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      oTrans.Master(Index) = .Text
   
      Select Case Index
      Case 2, 4
         .Text = Format(IIf(oTrans.Master(Index) = "", Now, oTrans.Master(Index)), "Mmmm DD, YYYY")
      End Select
   End With
End Sub

Private Function ReportTrans(ByVal ReportType As adReport) As Boolean
   Dim lrs As ADODB.Recordset
   Dim lbSwitch As Boolean
   Dim lnCtr As Integer
   Dim lnTmp As Integer
   Dim lsOldProc As String
   
   lsOldProc = "ReportTrans"
'   'On Error GoTo errProc
   
   ReportTrans = False
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "sField01", adVarChar, 40
   lrs.Fields.Append "sField02", adVarChar, 50
   lrs.Fields.Append "lField01", adCurrency, 30
   lrs.Fields.Append "nField01", adInteger
   lrs.Fields.Append "lField02", adCurrency, 30
   lrs.Fields.Append "lField03", adCurrency, 30
   lrs.Fields.Append "lField04", adCurrency, 30
   lrs.Fields.Append "lField05", adCurrency, 30
   lrs.Fields.Append "lField06", adCurrency, 30
   lrs.Fields.Append "lField07", adCurrency, 30
   lrs.Fields.Append "lField08", adCurrency, 30
   lrs.Fields.Append "lField09", adCurrency, 30
   lrs.Fields.Append "lField10", adCurrency, 30
   lrs.Fields.Append "lField11", adCurrency, 30
   lrs.Open
   
   lnTmp = 0
   With MSFlexGrid1
      lbSwitch = True
      For lnCtr = 1 To .Rows - 1
         lrs.AddNew
         lrs.Fields("sField01").Value = .TextMatrix(lnCtr, 2)         'Branch Name
         lrs.Fields("sField02").Value = .TextMatrix(lnCtr, 1)         'Employee Name
         lrs.Fields("lField01").Value = CCur(.TextMatrix(lnCtr, 3))   'Basic Salary
         lrs.Fields("nField01").Value = CInt(.TextMatrix(lnCtr, 4))   'Attendance
         lrs.Fields("lField02").Value = CCur(.TextMatrix(lnCtr, 10))  'Benefits
         lrs.Fields("lField03").Value = CCur(.TextMatrix(lnCtr, 9))   'Overtime
         lrs.Fields("lField04").Value = CCur(.TextMatrix(lnCtr, 8))   'Holiday
         lrs.Fields("lField05").Value = CCur(.TextMatrix(lnCtr, 11))  'Adjustment
         lrs.Fields("lField06").Value = CCur(.TextMatrix(lnCtr, 5))   'Absences
         lrs.Fields("lField07").Value = CCur(.TextMatrix(lnCtr, 6))   'Tardiness
         lrs.Fields("lField08").Value = CCur(.TextMatrix(lnCtr, 7))   'Undertime
         lrs.Fields("lField09").Value = CCur(.TextMatrix(lnCtr, 14))  'Loans
         lrs.Fields("lField10").Value = CCur(.TextMatrix(lnCtr, 13)) + _
                                        CCur(.TextMatrix(lnCtr, 12))  'Deductions/Advances
         lrs.Fields("lField11").Value = CCur(.TextMatrix(lnCtr, 15))  'Gov't Deductions
                                       
      Next
   End With
   
   poReport.InitReport
   Set poReport.ReportSource = lrs
   poReport.ReportID = "EmpPy1"
   poReport.ReportHeading1 = "P-A-Y-R-O-L-L"
   poReport.ReportHeading2 = "For the Period" & " " & Format(txtField(1).Text, "MMMM DD, YYYY") & " - " & Format(txtField(2).Text, "MMMM DD, YYYY")
   
   If ReportType = ViewReport Then
      poReport.ShowReport
   Else
      poReport.PrintReport
   End If
   ReportTrans = True
   
endProc:
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub LoadDetail()
   Dim lnCtr As Integer
   Dim lnDailyPay As Currency
   Set poRecord = New ADODB.Recordset

   With poRecord
      .Fields.Append "sEmployNm", adVarChar, 50
      .Fields.Append "sBranchNm", adVarChar, 40
      .Fields.Append "nBasicSal", adCurrency, 30
      .Fields.Append "nAttendnc", adInteger
      .Fields.Append "nAbsencex", adCurrency, 30
      .Fields.Append "nTardines", adCurrency, 30
      .Fields.Append "nUndrTime", adCurrency, 30
      .Fields.Append "nHolidayx", adCurrency, 30
      .Fields.Append "nOverTime", adCurrency, 30
      .Fields.Append "nBenefits", adCurrency, 30
      .Fields.Append "nAdjustxx", adCurrency, 30
      .Fields.Append "nAdvances", adCurrency, 30
      .Fields.Append "nDeductnx", adCurrency, 30
      .Fields.Append "nLoanAmrt", adCurrency, 30
      .Fields.Append "nGovDedct", adCurrency, 30
      
      .Open
   
      If oTrans.ItemCount = 0 Then Exit Sub
      
      For lnCtr = 0 To oTrans.ItemCount - 1
         lnDailyPay = Round(IIf(oTrans.Master("sEmpTypID") = "T", oTrans.Detail(lnCtr, "nBasicSal") / (oTrans.Detail(lnCtr, "nAttendnc") + oTrans.Detail(lnCtr, "nAbsencex")), oTrans.Detail(lnCtr, "nBasicSal") / 13), 2)
      
         .AddNew
         .Fields("sEmployNm") = oTrans.Detail(lnCtr, "sEmployNm")
         .Fields("sBranchNm") = IFNull(oTrans.Detail(lnCtr, "sBranchNm"))
         .Fields("nBasicSal") = oTrans.Detail(lnCtr, "nBasicSal")
         .Fields("nAttendnc") = oTrans.Detail(lnCtr, "nAttendnc")
         .Fields("nAbsencex") = Round(oTrans.Detail(lnCtr, "nAbsencex") * lnDailyPay, 2)
         .Fields("nTardines") = Round(oTrans.Detail(lnCtr, "nTardines") * (lnDailyPay / 8 / 60), 2)
         .Fields("nUndrTime") = Round(oTrans.Detail(lnCtr, "nUndrTime") * (lnDailyPay / 8 / 60), 2)
         .Fields("nHolidayx") = IIf(IsNumeric(oTrans.Detail(lnCtr, "nHolidayx")) = True, oTrans.Detail(lnCtr, "nHolidayx"), 0)
                     
         'compute overtime pay properly
         '2012.11.03 - kalyptus
         If oTrans.Master("dPeriodTo") > CDate("2012-10-31") Then
            .Fields("nOverTime") = Round(oTrans.Detail(lnCtr, "nOverTime") * (lnDailyPay / 8 / 60) * pxeREG_OTRTE, 2)
         Else
            .Fields("nOverTime") = Round(oTrans.Detail(lnCtr, "nOverTime") * (lnDailyPay / 8 / 60), 2)
         End If
         .Fields("nBenefits") = oTrans.Detail(lnCtr, "nBenefits")
         .Fields("nAdjustxx") = oTrans.Detail(lnCtr, "nAdjustxx")
         .Fields("nAdvances") = oTrans.Detail(lnCtr, "nAdvances")
         .Fields("nDeductnx") = oTrans.Detail(lnCtr, "nDeductnx")
         .Fields("nLoanAmrt") = oTrans.Detail(lnCtr, "nLoanAmrt")
         .Fields("nGovDedct") = oTrans.Detail(lnCtr, "nGovDedct")
         DoEvents
      Next
   End With
   
End Sub

Private Sub LoadDetailNew()
   Dim lnCtr As Integer
   Dim lnDailyPay As Currency
   Set poRecord = New ADODB.Recordset

   Set p_oProgress = New clsSpeedometer

   If oTrans.ItemCount > 0 Then
      With p_oProgress
         .InitProgress "Loading Payroll Sheet...", 1, 3
         .PrimaryRemarks = "Initializing Records..."
         .MoveProgress "Loading Employee Information..."
      End With
   End If
      
   With poRecord
      .Fields.Append "sEmployNm", adVarChar, 50
      .Fields.Append "sBranchNm", adVarChar, 40
      .Fields.Append "nBasicSal", adCurrency, 30
      .Fields.Append "nAttendnc", adInteger
      .Fields.Append "nAbsencex", adCurrency, 30
      .Fields.Append "nTardines", adCurrency, 30
      .Fields.Append "nUndrTime", adCurrency, 30
      .Fields.Append "nHolidayx", adCurrency, 30
      .Fields.Append "nOverTime", adCurrency, 30
      .Fields.Append "nBenefits", adCurrency, 30
      .Fields.Append "nAdjustxx", adCurrency, 30
      .Fields.Append "nAdvances", adCurrency, 30
      .Fields.Append "nDeductnx", adCurrency, 30
      .Fields.Append "nLoanAmrt", adCurrency, 30
      .Fields.Append "nGovDedct", adCurrency, 30
      
      .Open
   
      If oTrans.ItemCount = 0 Then
         Exit Sub
      End If
      
      setRepProgress oTrans.ItemCount, "Loading Employee Records..."
      DoEvents
      
      For lnCtr = 0 To oTrans.ItemCount - 1
         If moveRepProgress("Extracting " & oTrans.Detail(lnCtr, "sEmployNm") & "...") = False Then
            p_oProgress.CloseProgress
            MsgBox "Record Loading was Cancelled!!!", vbInformation, "Notice"
            Exit Sub
         End If
         
         If oTrans.Master("sEmpTypID") = "T" Then
            If oTrans.Detail(lnCtr, "cDivision") = 3 Then
               lnDailyPay = Round(oTrans.Detail(lnCtr, "nBasicSal") / 13, 2)
            Else
               lnDailyPay = Round(oTrans.Detail(lnCtr, "nBasicSal") / (oTrans.Detail(lnCtr, "nAttendnc") + oTrans.Detail(lnCtr, "nAbsencex")), 2)
            End If
         Else
            lnDailyPay = Round(oTrans.Detail(lnCtr, "nBasicSal") / 13, 2)
         End If
      
         .AddNew
         .Fields("sEmployNm") = oTrans.Detail(lnCtr, "sEmployNm")
         .Fields("sBranchNm") = IFNull(oTrans.Detail(lnCtr, "sBranchNm"))
         .Fields("nBasicSal") = oTrans.Detail(lnCtr, "nBasicSal")
         .Fields("nAttendnc") = oTrans.Detail(lnCtr, "nAttendnc")
         .Fields("nAbsencex") = Round(oTrans.Detail(lnCtr, "nAbsencex") * lnDailyPay, 2)
                     
         .Fields("nTardines") = oTrans.Detail(lnCtr, "nTardines")
         .Fields("nUndrTime") = oTrans.Detail(lnCtr, "nUndrTime")
         .Fields("nHolidayx") = IIf(IsNumeric(oTrans.Detail(lnCtr, "nHolidayx")) = True, oTrans.Detail(lnCtr, "nHolidayx"), 0)
         .Fields("nOverTime") = oTrans.Detail(lnCtr, "nOverTime")
                     
         .Fields("nBenefits") = oTrans.Detail(lnCtr, "nBenefits")
         '.Fields("nAdjustxx") = oTrans.Detail(lnCtr, "nAdjustxx") + oTrans.Detail(lnCtr, "nNightDif") + oTrans.Detail(lnCtr, "nOTNightD") + oTrans.Detail(lnCtr, "nAdjHldyx")
         .Fields("nAdjustxx") = oTrans.Detail(lnCtr, "nAdjustxx") + oTrans.Detail(lnCtr, "nNightDif") + oTrans.Detail(lnCtr, "nOTNightD")
         .Fields("nAdvances") = oTrans.Detail(lnCtr, "nAdvances")
         .Fields("nDeductnx") = oTrans.Detail(lnCtr, "nDeductnx")
         .Fields("nLoanAmrt") = oTrans.Detail(lnCtr, "nLoanAmrt")
         .Fields("nGovDedct") = oTrans.Detail(lnCtr, "nGovDedct")
         DoEvents
      Next
   End With
   
   If moveRepProgress("Processing Done") = False Then
      p_oProgress.CloseProgress
      MsgBox "Record Generation was Cancelled!!!", vbInformation, "Notice"
      Exit Sub
   End If
   
   p_oProgress.CloseProgress
   
End Sub

Private Sub setRepProgress(ByVal lnMaxValue As Long, Optional lvRemarks As Variant)
   With p_oProgress
      .SecMaxValue = lnMaxValue
      If Not IsMissing(lvRemarks) Then
         .PrimaryRemarks = lvRemarks
      End If
   End With
End Sub

Private Function moveRepProgress(ByVal lsSecRemarks As String, Optional lvPriRemarks As Variant) As Boolean
   With p_oProgress
      If Not IsMissing(lvPriRemarks) Then
         moveRepProgress = .MoveProgress(lsSecRemarks, lvPriRemarks)
      Else
         moveRepProgress = .MoveProgress(lsSecRemarks)
      End If
   End With
End Function

