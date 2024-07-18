VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmPayrollRegOld 
   BorderStyle     =   0  'None
   Caption         =   "Payroll Processing"
   ClientHeight    =   10035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15255
   Icon            =   "frmPayrollRegOld.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   660
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   1164
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   1
         Top             =   120
         Width           =   2190
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
         Index           =   24
         Left            =   105
         TabIndex        =   0
         Top             =   195
         Width           =   1365
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   13905
      TabIndex        =   16
      Top             =   5205
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
      Picture         =   "frmPayrollRegOld.frx":000C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   13905
      TabIndex        =   14
      Top             =   3945
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
      Picture         =   "frmPayrollRegOld.frx":0786
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   13905
      TabIndex        =   15
      Top             =   4575
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
      Picture         =   "frmPayrollRegOld.frx":0F00
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6810
      Left            =   105
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3105
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   12012
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
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1260
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   3175
      BackColor       =   12632256
      BorderStyle     =   1
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
         TabIndex        =   3
         Text            =   "M00111-000021"
         Top             =   120
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
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "August 16, 2011"
         Top             =   795
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
         TabIndex        =   10
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
         Index           =   4
         Left            =   5460
         TabIndex        =   12
         Text            =   "August 28, 2011"
         Top             =   1245
         Width           =   2190
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
         TabIndex        =   2
         Top             =   200
         Width           =   1485
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
         TabIndex        =   5
         Top             =   875
         Width           =   1065
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
         TabIndex        =   9
         Top             =   870
         Width           =   990
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
         TabIndex        =   11
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Image Image1 
         Height          =   1515
         Left            =   7860
         Top             =   120
         Width           =   2490
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   5415
         Top             =   105
         Width           =   2220
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   5445
         Top             =   135
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
         Left            =   5490
         TabIndex        =   4
         Tag             =   "eb0;et0"
         Top             =   165
         Width           =   2070
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2505
      Left            =   10650
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   4419
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.Image Image2 
         Height          =   2355
         Left            =   60
         Picture         =   "frmPayrollRegOld.frx":167A
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2835
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   13890
      TabIndex        =   17
      Top             =   2055
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "RB"
      AccessKey       =   "RB"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPayrollRegOld.frx":F316
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   13905
      TabIndex        =   18
      Top             =   2685
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "MB"
      AccessKey       =   "MB"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPayrollRegOld.frx":FBF0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   13905
      TabIndex        =   19
      Top             =   2685
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Others"
      AccessKey       =   "Others"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPayrollRegOld.frx":104CA
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   13890
      TabIndex        =   20
      Top             =   825
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "No Bank"
      AccessKey       =   "No Bank"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPayrollRegOld.frx":10C44
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   13890
      TabIndex        =   21
      Top             =   1440
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "BDO"
      AccessKey       =   "BDO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPayrollRegOld.frx":113BE
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   13905
      TabIndex        =   22
      Top             =   3315
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "SB"
      AccessKey       =   "SB"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPayrollRegOld.frx":11C98
   End
End
Attribute VB_Name = "frmPayrollRegOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmPayrollReg"
Private Const pxeREG_OTRTE As Single = 1.25

Private WithEvents oTrans As ggcPayroll.clsPayroll
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private psEmpTypID As String
Private pcMainOffc As String

Dim poReport As clsReport
Dim poRecord As Recordset

Dim pbCtrlPress As Boolean
Dim pnIndex As Integer
Dim pbLoad As Boolean
Dim pbFlexGrid As Boolean

Property Let EmployeeType(ByVal Value As String)
   psEmpTypID = Value
End Property

Property Let isMainOffice(ByVal Value As Boolean)
   pcMainOffc = IIf(Value, 1, 0)
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim loCls As clsRobinsonCMS
   Dim lnCtr As Integer
   Dim lnDta As Integer
   Dim loFrm As frmRobinsonsCMS
   Dim loFmx As frmChinaCMS
   Select Case Index
   Case 0 'Export to Robinsons CMS
      If oTrans.Master("cPostedxx") <> xeStatePosted Then Exit Sub
   
      Set loCls = New clsRobinsonCMS
      Set loCls.AppDriver = oApp
      loCls.InitTransaction
      loCls.NewTransaction
      loCls.Master("dTransact") = oApp.ServerDate
      loCls.Master("sPIRRefNo") = oTrans.Master("sPayPerID")
      
      For lnCtr = 1 To MSFlexGrid1.Rows - 1
         If MSFlexGrid1.TextMatrix(lnCtr, 18) = "00XX044" And CCur(MSFlexGrid1.TextMatrix(lnCtr, 16)) > 0 Then
'         If MSFlexGrid1.TextMatrix(lnCtr, 18) = "" And CCur(MSFlexGrid1.TextMatrix(lnCtr, 16)) > 0 Then
            loCls.addDetail
            loCls.Detail(loCls.ItemCount - 1, "sPayeAcct") = MSFlexGrid1.TextMatrix(lnCtr, 17)
            loCls.Detail(loCls.ItemCount - 1, "sPayeName") = MSFlexGrid1.TextMatrix(lnCtr, 1)
            loCls.Detail(loCls.ItemCount - 1, "nTranAmtx") = CCur(MSFlexGrid1.TextMatrix(lnCtr, 16))
         End If
      Next
      
      Set loFrm = New frmRobinsonsCMS
      Set loFrm.CMS = loCls
      
      loFrm.Show 1
      If loFrm.IsOkey Then
         Call loCls.SaveTransaction(True)
         loCls.exportCMS
         MsgBox "Export of Payroll successfully done!", vbInformation + vbOKOnly, "Success"
      End If
         
      Unload loFrm
         
   Case 1 'browse
      If oTrans.SearchTransaction(txtSearch, True) Then
         Call ClearFields
         Call InitGrid
         Call LoadMaster
         Call loadGrid
         Call flexFocus
         txtField(2).SetFocus
      Else
         txtSearch.SetFocus
      End If
   Case 2 'close
      Unload Me
   Case 3 'Export to MetroBank
      If oTrans.Master("cPostedxx") <> xeStatePosted Then Exit Sub
      
      Call ExportMB(MSFlexGrid1)
   Case 4
      If oTrans.Master("cPostedxx") <> xeStatePosted Then Exit Sub
      
      Set loFmx = New frmChinaCMS
      loFmx.txtField(0) = oTrans.Master("sPayPerID")
      loFmx.txtField(1) = Format(oTrans.Master("dCovergTo"), "Mmmm DD, YYYY")
      Set loFmx.Data = poRecord
      loFmx.Show 1
   Case 5 'Print
      Call ReportTrans(ViewReport)
   Case 6
      If oTrans.Master("cPostedxx") <> xeStatePosted Then Exit Sub
      Call ExportNoAccount(MSFlexGrid1)
   Case 7
      If oTrans.Master("cPostedxx") <> xeStatePosted Then Exit Sub
      Call ExportBDO(MSFlexGrid1)
   Case 8 ' Export to Security Bank Template
      If oTrans.Master("cPostedxx") <> xeStatePosted Then Exit Sub
      
      Call ExportSB(MSFlexGrid1)
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
   End If
   
   If Left(oApp.BranchCode, 1) = "M" Then
      cmdButton(4).Visible = False
   Else
      cmdButton(0).Visible = False
      cmdButton(3).Visible = False
      cmdButton(6).Visible = False
   End If
   
   pbCtrlPress = False
   pbFlexGrid = False
   txtSearch.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lnRow As Integer
   With MSFlexGrid1
      Select Case KeyCode
         Case vbKeyReturn, vbKeyDown, vbKeyUp
            Select Case KeyCode
               Case vbKeyReturn
                  If GetFocus = .hWnd Then Exit Sub
                  SetNextFocus
               Case vbKeyDown
                  If pbCtrlPress Then
                     If Not (.Row = .Rows - 1) Then
                        .Row = .Row + 1
                     End If
               
                     If Not .RowIsVisible(.Row + IIf(.Row = .Rows - 1, 0, 1)) Then .TopRow = .TopRow + 1
                        
                        lnRow = MSFlexGrid1.Row - 1
                        
                        KeyCode = 0
                        
                        .Col = 1
                        .ColSel = .Cols - 1
                     Else
                        SetNextFocus
                     End If
               Case vbKeyUp
                  If pbCtrlPress Then
                     If Not (.Row = 1) Then
                        .Row = .Row - 1
                     End If
               
                        If Not .RowIsVisible(.Row) Then .TopRow = .TopRow - 1
                           lnRow = MSFlexGrid1.Row - 1
                                   
                           KeyCode = 0
                           
                           .Col = 1
                           .ColSel = .Cols - 1
                  Else
                     SetPreviousFocus
                  End If
               End Select
         Case vbKeyControl
            pbCtrlPress = True
      End Select
   End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyControl
         If pbCtrlPress Then pbCtrlPress = False
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
'   On Error GoTo errProc

   CenterChildForm mdiMain, Me
      
   Set oTrans = New clsPayroll
   Set oTrans.AppDriver = oApp
   oTrans.EmployeeType = psEmpTypID
   oTrans.isMainOffice = pcMainOffc
   oTrans.TransStatus = 12
   oTrans.InitTransaction
         
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   
   Set poReport = New clsReport
   
   ClearFields
   InitGrid
   
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
      .Cols = 19
      
      .Row = 0
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      .RowHeight(0) = 338
      
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
      .TextMatrix(0, 17) = "Bank Account"
      .TextMatrix(0, 18) = "Bank"
      
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
      .ColWidth(17) = 0
      .ColWidth(18) = 0
      
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
         .TextMatrix(lnCtr + 1, 17) = ""
         .TextMatrix(lnCtr + 1, 18) = ""
      Else
         If oTrans.ItemCount > 18 Then
            .ColWidth(0) = 600
            .ColWidth(1) = 3050
         Else
            .ColWidth(0) = 635
            .ColWidth(1) = 3280
         End If
         
         Call LoadDetail
         
         poRecord.Sort = "sBranchNm, sEmployNm"
         poRecord.MoveFirst
         
         Do Until poRecord.EOF
            DoEvents
            
'         For lnCtr = 0 To oTrans.ItemCount - 1
'            lnDailyPay = Round(IIf(oTrans.Master("sEmpTypID") = "T", oTrans.Detail(lnCtr, "nBasicSal") / (oTrans.Detail(lnCtr, "nAttendnc") + oTrans.Detail(lnCtr, "nAbsencex")), oTrans.Detail(lnCtr, "nBasicSal") / 13), 2)
            .Rows = 2 + lnCtr
            .RowHeight(lnCtr + 1) = 338
            
'            .TextMatrix(lnCtr + 1, 0) = Format(lnCtr + 1, "0000")
'            .TextMatrix(lnCtr + 1, 1) = IFNull(oTrans.Detail(lnCtr, "sEmployNm"))
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
                                        
            .TextMatrix(lnCtr + 1, 17) = poRecord("sBnkActNo")
            .TextMatrix(lnCtr + 1, 18) = poRecord("sBankIDxx")
'         Next
            lnCtr = lnCtr + 1
            poRecord.MoveNext
         Loop
      End If
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
   txtSearch = oTrans.Master(0)
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

'Private Sub MSFlexGrid1_Click()
'   With MSFlexGrid1
'      pnActiveRow = .Row
'   End With
'   SetPreviousFocus
'End Sub

Private Sub MSFlexGrid1_GotFocus()
   pbFlexGrid = True
End Sub
Private Sub MSFlexGrid1_LostFocus()
   If pbFlexGrid Then pbFlexGrid = False
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
         Case 2, 4
            txtField(Index) = Format(oTrans.Master(Index), "MM/DD/YYYY")
            pbFlexGrid = True
         Case Else
            pbFlexGrid = False
      End Select

   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
      Case vbKeyDown
         If pbCtrlPress Then
            KeyCode = 0
         Else
            SetNextFocus
         End If
      Case vbKeyUp
         If pbCtrlPress Then
               KeyCode = 0
         Else
            SetPreviousFocus
         End If
   End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
   Case 2, 4
      txtField(Index) = Format(txtField(Index), "MMMM DD, YYYY")
   End Select
   
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub


Private Sub txtSearch_GotFocus()
   With txtSearch
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   pbFlexGrid = False
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF3, vbKeyReturn
         If oTrans.SearchTransaction(txtSearch, True) Then
            Call ClearFields
            Call LoadMaster
            Call loadGrid
            Call flexFocus
            txtField(2).SetFocus
         Else
            txtSearch.SetFocus
         End If
      Case vbKeyDown
         SetNextFocus
   End Select
End Sub

Private Sub txtSearch_LostFocus()
   With txtSearch
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next

   Label2.Caption = ""
End Sub

'Private Function ReportTrans(ByVal ReportType As adReport) As Boolean
'   Dim lrs As ADODB.Recordset
'   Dim lbSwitch As Boolean
'   Dim lnCtr As Integer
'   Dim lnTmp As Integer
'   Dim lsOldProc As String
'
'   lsOldProc = "ReportTrans"
''   On Error GoTo errProc
'
'   ReportTrans = False
'
'   Set lrs = New ADODB.Recordset
'
'   lrs.Fields.Append "sField01", adVarChar, 40
'   lrs.Fields.Append "sField02", adVarChar, 50
'   lrs.Fields.Append "lField01", adCurrency, 30
'   lrs.Fields.Append "nField01", adInteger
'   lrs.Fields.Append "lField02", adCurrency, 30
'   lrs.Fields.Append "lField03", adCurrency, 30
'   lrs.Fields.Append "lField04", adCurrency, 30
'   lrs.Fields.Append "lField05", adCurrency, 30
'   lrs.Fields.Append "lField06", adCurrency, 30
'   lrs.Fields.Append "lField07", adCurrency, 30
'   lrs.Fields.Append "lField08", adCurrency, 30
'   lrs.Fields.Append "lField09", adCurrency, 30
'   lrs.Fields.Append "lField10", adCurrency, 30
'   lrs.Fields.Append "lField11", adCurrency, 30
'   lrs.Open
'
'   lnTmp = 0
'   With MSFlexGrid1
'      lbSwitch = True
'      For lnCtr = 1 To .Rows - 1
'         lrs.AddNew
'         lrs.Fields("sField01").Value = .TextMatrix(lnCtr, 2)         'Branch Name
'         lrs.Fields("sField02").Value = .TextMatrix(lnCtr, 1)         'Employee Name
'         lrs.Fields("lField01").Value = CCur(.TextMatrix(lnCtr, 3))   'Basic Salary
'         lrs.Fields("nField01").Value = CInt(.TextMatrix(lnCtr, 4))   'Attendance
'         lrs.Fields("lField02").Value = CCur(.TextMatrix(lnCtr, 10))  'Benefits
'         lrs.Fields("lField03").Value = CCur(.TextMatrix(lnCtr, 9))   'Overtime
'         lrs.Fields("lField04").Value = CCur(.TextMatrix(lnCtr, 8))   'Holiday
'         lrs.Fields("lField05").Value = CCur(.TextMatrix(lnCtr, 11))  'Adjustment
'         lrs.Fields("lField06").Value = CCur(.TextMatrix(lnCtr, 5))   'Absences
'         lrs.Fields("lField07").Value = CCur(.TextMatrix(lnCtr, 6))   'Tardiness
'         lrs.Fields("lField08").Value = CCur(.TextMatrix(lnCtr, 7))   'Undertime
'         lrs.Fields("lField09").Value = CCur(.TextMatrix(lnCtr, 14))  'Loans
'         lrs.Fields("lField10").Value = CCur(.TextMatrix(lnCtr, 13)) + _
'                                        CCur(.TextMatrix(lnCtr, 12))  'Deductions/Advances
'         lrs.Fields("lField11").Value = CCur(.TextMatrix(lnCtr, 15))  'Gov't Deductions
'
'      Next
'   End With
'
'   poReport.InitReport
'   Set poReport.ReportSource = lrs
'   poReport.ReportID = "EmpPy1"
'   poReport.ReportHeading1 = "P-A-Y-R-O-L-L"
'   poReport.ReportHeading2 = "For the Period" & " " & Format(txtField(1).Text, "MMMM DD, YYYY") & " - " & Format(txtField(2).Text, "MMMM DD, YYYY")
'
'   If ReportType = ViewReport Then
'      poReport.ShowReport
'   Else
'      poReport.PrintReport
'   End If
'   ReportTrans = True
'
'endProc:
'   Set lrs = Nothing
'   Exit Function
'errProc:
'   ShowError lsOldProc & "( " & " )"
'End Function

Private Function ReportTrans(ByVal ReportType As adReport) As Boolean
   Dim lrs As ADODB.Recordset
   Dim lbSwitch As Boolean
   Dim lnCtr As Integer
   Dim lnTmp As Integer
   Dim lsOldProc As String
   Dim lsBranchCD As String
   
   lsOldProc = "ReportTrans"
'   On Error GoTo errProc
   
   ReportTrans = False
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "sField10", adVarChar, 50
   lrs.Fields.Append "nField01", adInteger
   lrs.Fields.Append "sField01", adVarChar, 15  'Bank Account
   lrs.Fields.Append "sField02", adVarChar, 70  'Employee Name
   lrs.Fields.Append "lField01", adCurrency     'Net Pay
   lrs.Fields.Append "lField02", adCurrency     'Net Pay 2
   
   lrs.Fields.Append "nField02", adInteger
   lrs.Fields.Append "sField03", adVarChar, 15  'Bank Account
   lrs.Fields.Append "sField04", adVarChar, 70  'Employee Name
   lrs.Fields.Append "lField03", adCurrency     'Net Pay
   lrs.Fields.Append "lField04", adCurrency     'Net Pay 2
   
   lrs.Open

   lnTmp = 0
   With poRecord
      .Sort = "sBranchNm, sBnkActNo"
      .MoveFirst
      lbSwitch = True
      Do Until .EOF
         If lsBranchCD <> .Fields("sBranchNm") Then
            lsBranchCD = .Fields("sBranchNm")
            lnCtr = 1
         Else
            lnCtr = lnCtr + 1
         End If
         
         If lnCtr Mod 2 = 1 Then
            lrs.AddNew
            lrs.Fields("sField10").Value = .Fields("sBranchNm")
            lrs.Fields("nField01").Value = lnCtr
            lrs.Fields("sField01").Value = .Fields("sBnkActNo")   'Bank Account
            lrs.Fields("sField02").Value = .Fields("sEmployNm")   'Employee Name
            If Trim(lrs.Fields("sField01").Value) <> "" Then
               lrs.Fields("lField01").Value = .Fields("nNetPayxx") 'Net Pay
               lrs.Fields("lField03").Value = 0
            Else
               lrs.Fields("lField01").Value = 0
               lrs.Fields("lField03").Value = .Fields("nNetPayxx") 'Net Pay
            End If
         Else
            lrs.Fields("nField02").Value = lnCtr
            lrs.Fields("sField03").Value = .Fields("sBnkActNo")   'Bank Account
            lrs.Fields("sField04").Value = .Fields("sEmployNm")   'Employee Name
            If Trim(lrs.Fields("sField03").Value) <> "" Then
               lrs.Fields("lField02").Value = .Fields("nNetPayxx") 'Net Pay
               lrs.Fields("lField04").Value = 0
            Else
               lrs.Fields("lField02").Value = 0
               lrs.Fields("lField04").Value = .Fields("nNetPayxx") 'Net Pay
            End If
         End If
         
         .MoveNext
      Loop
   End With
   
   poReport.InitReport
   lrs.Sort = ""
   Set poReport.ReportSource = lrs
   poReport.ReportID = "RbCMS2"
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

Private Sub flexFocus()
   With MSFlexGrid1
      If .Row > 15 Then .TopRow = 1
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
      .SetFocus
      pbFlexGrid = True
   End With
End Sub

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
      .Fields.Append "nNetPayxx", adCurrency, 30
      .Fields.Append "sBnkActNo", adVarChar, 15
      .Fields.Append "sBankIDxx", adVarChar, 9
      
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
         .Fields("sBnkActNo") = IFNull(oTrans.Detail(lnCtr, "sBnkActNo"))
         .Fields("sBankIDxx") = IFNull(oTrans.Detail(lnCtr, "sBankIDxx"))
      
         .Fields("nNetPayxx") = Format((.Fields("nBasicSal") _
                              + .Fields("nHolidayx") _
                              + .Fields("nOverTime") _
                              + .Fields("nBenefits") _
                              + .Fields("nAdjustxx")) _
                             - (.Fields("nAbsencex") _
                              + .Fields("nTardines") _
                              + .Fields("nUndrTime") _
                              + .Fields("nAdvances") _
                              + .Fields("nDeductnx") _
                              + .Fields("nLoanAmrt") _
                              + .Fields("nGovDedct")), "#,##0.00")
               
      Next
   End With
   
End Sub

Private Function ExportMB(flxGrd As MSFlexGrid) As Boolean
   Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook
   
   Dim lsSQL As String
   Dim lors As Recordset
   
   Dim XEmployID As String
   Dim XEmployNm As String
   Dim XBranchCD As String
   Dim XAcctNmbr As String
   Dim XNetPayxx As String
   
   Dim sFileName As String
   
   Dim lnLineCtr As Integer
   Dim lnItemCtr As Integer
   
   XEmployID = "A"
   XEmployNm = "B"
   XBranchCD = "C"
   XAcctNmbr = "D"
   XNetPayxx = "E"
   
   lnLineCtr = 6
   
   sFileName = "MetroBank Payroll"
    
   Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & sFileName & ".XLS")
   
   Set xlsheet = xlwbook.Sheets.Item(1)
   xlsheet.Range(XEmployID & (lnLineCtr + 1) & ":" & XNetPayxx & 65536).ClearContents
    
   With flxGrd
   
     For lnItemCtr = 1 To flxGrd.Rows - 1
         If .TextMatrix(lnItemCtr, 18) = "00XX006" And CCur(.TextMatrix(lnItemCtr, 16)) > 0 And Trim(.TextMatrix(lnItemCtr, 17)) <> "" Then
            lnLineCtr = lnLineCtr + 1
            xlsheet.Range(XEmployID & (lnLineCtr)).Value = lnLineCtr - 6
            xlsheet.Range(XEmployNm & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
            xlsheet.Range(XBranchCD & (lnLineCtr)).Value = Left(.TextMatrix(lnItemCtr, 17), 3)
            xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = Mid(.TextMatrix(lnItemCtr, 17), 4)
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = CCur(.TextMatrix(lnItemCtr, 16))
         End If
     Next
'     MsgBox xlsheet.Range("E" & Rows.Count).End(xlUp).Row
         
     'kalyptus-2013.08.19 02:03pm
     'delete remaining items in our template if ever template was updated...
     lnItemCtr = xlsheet.Range("E" & xlsheet.Rows.Count).End(xlUp).Row
     If lnItemCtr > lnLineCtr Then
         xlsheet.Range("A" & lnLineCtr + 1 & ":E" & lnItemCtr).Clear
     End If
   
   End With
    
   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & txtField(0) & ").XLS"
   xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
   xl.Quit
    
   Set xlwbook = Nothing
   Set xl = Nothing
End Function

Private Function ExportBDO(flxGrd As MSFlexGrid) As Boolean
   Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook
   
   Dim lsSQL As String
   Dim lors As Recordset
   
   Dim XEmployNm As String
   Dim XRemarksx As String
   Dim XAcctNmbr As String
   Dim XNetPayxx As String
   
   Dim sFileName As String
   
   Dim lnLineCtr As Integer
   Dim lnItemCtr As Integer
   
   XAcctNmbr = "A"
   XNetPayxx = "B"
   XEmployNm = "C"
   XRemarksx = "D"
   
   lnLineCtr = 6
   
   sFileName = "BDO ATM Payroll Converter for BOB"
    
   Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & sFileName & ".XLS")
   
   Set xlsheet = xlwbook.Sheets.Item(1)
   xlsheet.Range(XAcctNmbr & (lnLineCtr + 1) & ":" & XRemarksx & 65536).ClearContents
    
   With flxGrd
     For lnItemCtr = 1 To flxGrd.Rows - 1
         If .TextMatrix(lnItemCtr, 18) = "00XX024" And CCur(.TextMatrix(lnItemCtr, 16)) > 0 And Trim(.TextMatrix(lnItemCtr, 17)) <> "" Then
            lnLineCtr = lnLineCtr + 1
            xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = "'" & .TextMatrix(lnItemCtr, 17)
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = CCur(.TextMatrix(lnItemCtr, 16))
            xlsheet.Range(XEmployNm & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
            xlsheet.Range(XRemarksx & (lnLineCtr)).Value = ""
         End If
     Next
   End With
    
   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & txtField(0) & ").XLS"
   xl.ActiveWorkbook.Close
   xl.Quit
    
   Set xlwbook = Nothing
   Set xl = Nothing
End Function

Private Function ExportNoAccount(flxGrd As MSFlexGrid) As Boolean
   Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook
   
   Dim lsSQL As String
   Dim lors As Recordset
   
   Dim XEmployNo As String
   Dim XBranchNm As String
   Dim XEmployNm As String
   Dim XNetPayxx As String
   
   Dim sFileName As String
   
   Dim lnLineCtr As Integer
   Dim lnItemCtr As Integer
   
   XEmployNo = "A"
   XBranchNm = "B"
   XEmployNm = "C"
   XNetPayxx = "D"
   
   lnLineCtr = 6
   
   sFileName = "No Bank Account"
    
   Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & sFileName & ".XLS")
   Set xlsheet = xlwbook.Sheets.Item(1)
   xlsheet.Range(XEmployNo & (lnLineCtr + 1) & ":" & XNetPayxx & 65536).ClearContents
    
   xlsheet.Range("B4").Value = txtField(1) + " - " + txtField(2)
    
   With flxGrd
     For lnItemCtr = 1 To flxGrd.Rows - 1
         If .TextMatrix(lnItemCtr, 18) = "" And CCur(.TextMatrix(lnItemCtr, 16)) > 0 Then
            lnLineCtr = lnLineCtr + 1
            xlsheet.Range(XEmployNo & (lnLineCtr)).Value = lnLineCtr - 6
            xlsheet.Range(XBranchNm & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 2)
            xlsheet.Range(XEmployNm & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = CCur(.TextMatrix(lnItemCtr, 16))
         End If
     Next
         
     'kalyptus-2013.08.19 02:03pm
     'delete remaining items in our template if ever template was updated...
     lnItemCtr = xlsheet.Range("E" & xlsheet.Rows.Count).End(xlUp).Row
     If lnItemCtr > lnLineCtr Then
         xlsheet.Range("A" & lnLineCtr + 1 & ":E" & lnItemCtr).Clear
     End If
   
   End With
    
   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & txtField(0) & ").XLS"
   xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
   xl.Quit
    
   Set xlwbook = Nothing
   Set xl = Nothing
End Function

Private Function ExportSB(flxGrd As MSFlexGrid) As Boolean
   Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook
   
   Dim lsSQL As String
   Dim lors As Recordset
   
   Dim XEmployID As String
   Dim XEmployNm As String
   Dim XBranchCD As String
   Dim XAcctNmbr As String
   Dim XNetPayxx As String
   
   Dim sFileName As String
   
   Dim lnLineCtr As Integer
   Dim lnItemCtr As Integer
   
   XEmployNm = "A"
   XAcctNmbr = "B"
   XNetPayxx = "C"
   
   lnLineCtr = 8
   
   sFileName = "SBC_payroll_ver_1.7"
    
   Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & sFileName & ".XLS")
   
   Set xlsheet = xlwbook.Sheets.Item(1)
   xlsheet.Range(XEmployNm & (lnLineCtr + 1) & ":" & XNetPayxx & 65536).ClearContents
    
   With flxGrd
   
     For lnItemCtr = 1 To flxGrd.Rows - 1
         If .TextMatrix(lnItemCtr, 18) = "00XX022" And CCur(.TextMatrix(lnItemCtr, 16)) > 0 And Trim(.TextMatrix(lnItemCtr, 17)) <> "" Then
            lnLineCtr = lnLineCtr + 1
            xlsheet.Range(XEmployNm & (lnLineCtr)).Value = .TextMatrix(lnItemCtr, 1)
            xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = Mid(.TextMatrix(lnItemCtr, 17), 4)
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = CCur(.TextMatrix(lnItemCtr, 16))
         End If
     Next
     
     'kalyptus-2013.08.19 02:03pm
     'delete remaining items in our template if ever template was updated...
     lnItemCtr = xlsheet.Range("C" & xlsheet.Rows.Count).End(xlUp).Row
     If lnItemCtr > lnLineCtr Then
         xlsheet.Range("A" & lnLineCtr + 1 & ":C" & lnItemCtr).Clear
     End If
   
   End With
    
   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & txtField(0) & ").XLS"
   xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
   xl.Quit
    
   Set xlwbook = Nothing
   Set xl = Nothing
End Function
