VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frm13thMonth 
   BorderStyle     =   0  'None
   Caption         =   "13th Month Pay Computation"
   ClientHeight    =   9855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9855
   ScaleWidth      =   14880
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1845
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   3254
      BackColor       =   12632256
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
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "August 16, 2011"
         Top             =   825
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
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   1275
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
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "M00111-000021"
         Top             =   150
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
         Height          =   825
         Index           =   4
         Left            =   5820
         TabIndex        =   1
         Top             =   825
         Width           =   3870
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   7470
         Top             =   150
         Width           =   2220
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   7500
         Top             =   180
         Width           =   2160
      End
      Begin VB.Label Label3 
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
         Left            =   7545
         TabIndex        =   12
         Tag             =   "eb0;et0"
         Top             =   210
         Width           =   2070
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
         Left            =   540
         TabIndex        =   11
         Top             =   900
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
         Left            =   615
         TabIndex        =   10
         Top             =   1350
         Width           =   990
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   405
         Left            =   1815
         Tag             =   "et0;ht2"
         Top             =   240
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
         Left            =   120
         TabIndex        =   7
         Top             =   225
         Width           =   1485
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   4650
         TabIndex        =   0
         Tag             =   "et0;fb0"
         Top             =   900
         Width           =   1095
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   75
      TabIndex        =   4
      Top             =   1860
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
      Picture         =   "frm13thMonth.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   2
      Top             =   3105
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
      Picture         =   "frm13thMonth.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   2475
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
      Picture         =   "frm13thMonth.frx":0EF4
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7110
      Left            =   1605
      TabIndex        =   5
      Top             =   2490
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   12541
      _Version        =   393216
      Cols            =   3
      FocusRect       =   0
      SelectionMode   =   1
      MergeCells      =   1
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   75
      TabIndex        =   13
      Top             =   1230
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
      Picture         =   "frm13thMonth.frx":166E
   End
   Begin VB.Image Image2 
      Height          =   1845
      Left            =   11685
      Picture         =   "frm13thMonth.frx":1DE8
      Stretch         =   -1  'True
      Top             =   555
      Width           =   3090
   End
End
Attribute VB_Name = "frm13thMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frm13thMonth"

Private oSkin As clsFormSkin
Private p_oPayMiscx As clsPayMisc
Private WithEvents oTrans As cls13MonthPay4PayDivNew
Attribute oTrans.VB_VarHelpID = -1
Dim poReport As clsReport

Private bLoaded As Boolean
Dim pnCtr As Integer
Dim pnIndex As Integer
Dim pbTxtField As Boolean
Dim pasPeriod(24) As String

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 1
      If oTrans.PostTransaction Then
         MsgBox "13th Month Pay was confirmed successfully!", vbOKOnly, "Confirmation"
      Else
         MsgBox "Unabled to confirm 13th Month Pay!", vbOKOnly, "Confirmation"
      End If
   Case 2
      'Process 13th Month Pay here
      oTrans.CloseTransaction
      oTrans.NewTransaction
      Call LoadLedger
   Case 3
      Call ReportTrans(ViewReport)
   Case 4
      Unload Me
   End Select

End Sub

Private Sub Form_Activate()

10       Dim lsOldProc As String

20       lsOldProc = "Form_Activate"
30       'On Error GoTo errProc

40       oApp.MenuName = Me.Tag
50       Me.ZOrder 0

60       If bLoaded = False Then
70          bLoaded = True
80       End If

endProc:
100      Exit Sub
errProc:
110      ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF12 And oApp.UserLevel = xeEngineer Then
      oTrans.export2Excel
   End If
End Sub

Private Sub Form_Load()
10       Dim lsOldProc As String

20       lsOldProc = "Form_Load"
30       'On Error GoTo errProc

40       CenterChildForm mdiMain, Me

80       Set oTrans = New cls13MonthPay4PayDivNew
90       Set oTrans.AppDriver = oApp
91       oTrans.TransStatus = 10

110      oTrans.InitTransaction
111      oTrans.NewTransaction

112      If oTrans.EditMode = xeModeAddNew Then
113         oTrans.SaveTransaction
114      End If

120      Set oSkin = New clsFormSkin
130      Set oSkin.AppDriver = oApp
140      Set oSkin.Form = Me
150      oSkin.ApplySkin xeFormTransEqualLeft

151      Set p_oPayMiscx = New clsPayMisc
152      Set p_oPayMiscx.AppDriver = oApp

160      InitGrid
165      LoadMaster
170      LoadLedger
175      Set poReport = New clsReport


endProc:
180      Exit Sub
errProc:
190      ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub LoadMaster()
10       Dim loTxt As TextBox

20       For Each loTxt In txtField
30          loTxt = oTrans.Master(loTxt.Index)
40       Next

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



Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

   pnIndex = Index
   pbTxtField = True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With

   pnIndex = Index
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   Dim lsSQL As String
   Dim lors As Recordset

   With MSFlexGrid1
      .Rows = 2
      .Cols = 28

      .Row = 0
      .RowHeight(0) = 320

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next

      .MergeRow(0) = True

      'Set Column Header
      .Row = 1
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Employee"
      .TextMatrix(0, 2) = "Branch"

      'kalyptus - 2017.12.06 01:55pm
      'Set the header based on the entries in Payroll_Calendar_Period_Semi_Monthly
      lsSQL = "SELECT sPeriodID, sPeriodTo FROM Payroll_Calendar_Period_Semi_Monthly WHERE sPeriodID LIKE 'M00112%'"
      Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)

      lnCtr = 3
      Do Until lors.EOF
         lsSQL = lors("sPeriodTo") & "-" & Year(Date)
         .TextMatrix(0, lnCtr) = Format(lsSQL, "Mmm DD")

         lnCtr = lnCtr + 1
         lors.MoveNext
      Loop
      .TextMatrix(0, 27) = "Net Pay"


'      .TextMatrix(0, 3) = "Nov 30"
'      .TextMatrix(0, 4) = "Dec 15"
'      .TextMatrix(0, 5) = "Dec 31"
'      .TextMatrix(0, 6) = "Jan 15"
'      .TextMatrix(0, 7) = "Jan 31"
'      .TextMatrix(0, 8) = "Feb 15"
'      .TextMatrix(0, 9) = "Feb 28"
'      .TextMatrix(0, 10) = "Mar 15"
'      .TextMatrix(0, 11) = "Mar 31"
'      .TextMatrix(0, 12) = "Apr 15"
'      .TextMatrix(0, 13) = "Apr 30"
'      .TextMatrix(0, 14) = "May 15"
'      .TextMatrix(0, 15) = "May 31"
'      .TextMatrix(0, 16) = "Jun 15"
'      .TextMatrix(0, 17) = "Jun 30"
'      .TextMatrix(0, 18) = "Jul 15"
'      .TextMatrix(0, 19) = "Jul 31"
'      .TextMatrix(0, 20) = "Aug 15"
'      .TextMatrix(0, 21) = "Aug 31"
'      .TextMatrix(0, 22) = "Sep 15"
'      .TextMatrix(0, 23) = "Sep 30"
'      .TextMatrix(0, 24) = "Oct 15"
'      .TextMatrix(0, 25) = "Oct 31"
'      .TextMatrix(0, 26) = "Nov 15"

      .RowHeightMin = 320


      'Set Column Width
      .ColWidth(0) = 635
      .ColWidth(1) = 3280
      .ColWidth(2) = 2400
      For lnCtr = 3 To .Cols - 1
         .ColWidth(lnCtr) = 1200
      Next

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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   If Index = 4 Then
      oTrans.Master(Index) = txtField(Index)
   End If
End Sub

Private Sub LoadLedger()
   Dim lnRow As Integer
   Dim lnCol As Integer
   Dim lnTotal As Currency
   Dim lnGrand As Currency
   Dim ldDate1 As Date
   Dim ldDate2 As Date
   Dim lnMonth As Single

   ldDate1 = Now()
   With MSFlexGrid1
      If oTrans.ItemCount = 0 Then
         .Rows = 2
         For lnCol = 0 To .Cols - 1
            .TextMatrix(1, lnCol) = ""
         Next
      Else
         .Rows = oTrans.ItemCount + 1
         DoEvents
         For lnRow = 0 To oTrans.ItemCount - 1
            .TextMatrix(lnRow + 1, 0) = lnRow + 1
            .TextMatrix(lnRow + 1, 1) = oTrans.Detail(lnRow, 1)
            .TextMatrix(lnRow + 1, 2) = oTrans.Detail(lnRow, 2)

            lnGrand = 0
            lnMonth = 0
            For lnCol = 1 To 24
               lnTotal = Val(oTrans.Detail(lnRow, "nbasics" & Format(lnCol, "00"))) _
                       + oTrans.Detail(lnRow, "nattend" & Format(lnCol, "00")) _
                       + oTrans.Detail(lnRow, "nadjust" & Format(lnCol, "00"))
               lnTotal = Round(lnTotal - oTrans.Detail(lnRow, "ncolaxx" & Format(lnCol, "00")), 2)

               If lnTotal <> 0 Then
                  lnMonth = lnMonth + 1
               End If
               lnGrand = lnGrand + lnTotal
               .TextMatrix(lnRow + 1, 2 + lnCol) = Format(lnTotal, "#,##0.00")
               DoEvents
            Next

            .TextMatrix(lnRow + 1, 27) = Format(lnGrand / 12, "#,##0.00")

            'Temporarily use this code to save the 13th month pay to Payroll_Annual_Total

'            If oTrans.Detail(lnRow, "n13thMnth") <> p_oPayMiscx.Yamot(CCur(Round(lnGrand / 12, 2))) Then
'               Call save13thMP( _
'                  oTrans.Master("sPeriodID") _
'                , oTrans.Detail(lnRow, "sEmployID") _
'                , oTrans.Detail(lnRow, "sBranchCD") _
'                , lnMonth / 2 _
'                , CCur(Round(lnGrand / 12, 2)))
'            End If

               'kalyptus - 2021.11.03 05:18pm
               'Save if branch is the branch of the primary payroll manager of the division
               If oApp.BranchCode = oTrans.Master("sMainBrCD") Then
                  Call save13thMP( _
                     oTrans.Master("sPeriodID") _
                   , oTrans.Detail(lnRow, "sEmployID") _
                   , oTrans.Detail(lnRow, "sBranchCD") _
                   , lnMonth / 2 _
                   , CCur(Round(lnGrand / 12, 2)))
               End If

         Next

         ldDate2 = Now

      End If
   End With

'   MsgBox DateDiff("s", ldDate1, ldDate2)

End Sub

Private Function ReportTrans(ByVal ReportType As adReport) As Boolean
   Dim lrs As ADODB.Recordset
   Dim lbSwitch As Boolean
   Dim lnRow As Integer
   Dim lnCol As Integer
   Dim lsOldProc As String
   Dim lsBranchCD As String

   lsOldProc = "ReportTrans"
'   On Error GoTo errProc

   ReportTrans = False

   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "sField01", adVarChar, 50
   lrs.Fields.Append "sField02", adVarChar, 50
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "lField02", adCurrency
   lrs.Fields.Append "lField03", adCurrency
   lrs.Fields.Append "lField04", adCurrency
   lrs.Fields.Append "lField05", adCurrency
   lrs.Fields.Append "lField06", adCurrency
   lrs.Fields.Append "lField07", adCurrency
   lrs.Fields.Append "lField08", adCurrency
   lrs.Fields.Append "lField09", adCurrency
   lrs.Fields.Append "lField10", adCurrency
   lrs.Fields.Append "lField11", adCurrency
   lrs.Fields.Append "lField12", adCurrency
   lrs.Fields.Append "lField13", adCurrency
   lrs.Fields.Append "lField14", adCurrency
   lrs.Fields.Append "lField15", adCurrency
   lrs.Fields.Append "lField16", adCurrency
   lrs.Fields.Append "lField17", adCurrency
   lrs.Fields.Append "lField18", adCurrency
   lrs.Fields.Append "lField19", adCurrency
   lrs.Fields.Append "lField20", adCurrency
   lrs.Fields.Append "lField21", adCurrency
   lrs.Fields.Append "lField22", adCurrency
   lrs.Fields.Append "lField23", adCurrency
   lrs.Fields.Append "lField24", adCurrency

   lrs.Open

   With MSFlexGrid1
      For lnRow = 1 To .Rows - 1
         lrs.AddNew
         lrs.Fields("sField01").Value = .TextMatrix(lnRow, 2)
         lrs.Fields("sField02").Value = .TextMatrix(lnRow, 1)
         For lnCol = 1 To 24
            lrs.Fields("lField" & Format(lnCol, "00")).Value = CCur(.TextMatrix(lnRow, lnCol + 2))
         Next
      Next
   End With

   poReport.InitReport
   lrs.Sort = ""
   Set poReport.ReportSource = lrs
   poReport.ReportID = "EmpPy4"
   poReport.ReportHeading1 = "13TH MONTH PAY"
   poReport.ReportHeading2 = "For the Period" & " " & Format(CDate(oTrans.Master("sPeriodFr")), "MMMM DD, YYYY") & " - " & Format(CDate(oTrans.Master("sPeriodTo")), "MMMM DD, YYYY")

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

'save13thMP(otrans.Detail(lnRow,"sEmployID") , lnMonth, CCUR(ROUND(LNGRAND,2)))
Private Sub save13thMP(ByVal fsPeriodID As String, ByVal fsEmployID As String, ByVal fsBranchCD As String, ByVal fnMonth As Single, ByVal fn13thPay As Currency)
   Dim lsSQL As String
   Dim lors As Recordset

   lsSQL = "SELECT" & _
                  "  a.sPeriodID" & _
                  ", b.sEmployID" & _
                  ", a.sBranchCD" & _
                  ", b.sPositnID" & _
                  ", a.nMVPRnk01" & _
                  ", a.nAttend01" & _
                  ", a.nBasicPay" & _
                  ", a.n13ThMnth" & _
                  ", a.nBonusxxx" & _
                  ", a.nBonusxx1" & _
                  ", a.nPartyxx1" & _
                  ", a.cReleased" & _
                  ", a.sReleased" & _
                  ", a.dReleased" & _
                  ", a.sDeptIDxx" & _
                  ", a.sEmpLevID" & _
                  ", b.sEmpLevID xEmpLevID" & _
                  ", a.sEmployID xEmployID" & _
                  ", b.nBasicPay xBasicPay" & _
                  ", b.sDeptIDxx xDeptIDxx" & _
                  ", b.sPyBranch"

   lsSQL = lsSQL & _
                  ", b.sBranchCD xBranchCD" & _
          " FROM Employee_Master001 b" & _
               " LEFT JOIN Payroll_Annual_Total a ON b.sEmployID = a.sEmployID" & _
                     " AND a.sPeriodID = " & strParm(fsPeriodID) & _
          " WHERE b.sEmployID = " & strParm(fsEmployID)

   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set lors.ActiveConnection = Nothing
   Debug.Print lsSQL

   'Kalyptus - 2015.11.18 01:31pm
   'Comment out the second part of if since it seems to be the reason why confirmation of ranking was reset....
'   If IFNull(loRS("xEmployID"), "") = "" Or (IFNull(loRS("sEmployID"), "") = IFNull(loRS("xEmployID"), "")) Then
   If IFNull(lors("xEmployID"), "") = "" Then
      lors("sPeriodID") = fsPeriodID
      lors("sEmployID") = fsEmployID
      lors("cReleased") = "0"
   End If

   lors("sBranchCD") = fsBranchCD
   lors("nAttend01") = fnMonth
   lors("n13thMnth") = p_oPayMiscx.Yamot(fn13thPay)
   lors("sEmpLevID") = lors("xEmpLevID")
   lors("sDeptIDxx") = lors("xDeptIDxx")
   lors("nBasicPay") = lors("xBasicPay")

   If IFNull(lors("sEmployID"), "") <> IFNull(lors("xEmployID"), "") Then
      lsSQL = ADO2SQL(lors, "Payroll_Annual_Total", , , , "xEmployID»xEmpLevID»xBasicPay»xDeptIDxx»sPyBranch»xBranchCD")
   Else
      If lors("sBranchCD") <> IFNull(lors("xBranchCD"), "") Then
         'If employee changes branch then
         lsSQL = "DELETE FROM Payroll_Annual_Total WHERE sPeriodID = " & strParm(fsPeriodID) & " AND sEmployID = " & strParm(fsEmployID)
         'oApp.CommitTrans
         Call oApp.Execute(lsSQL, "Payroll_Annual_Total", lors("sBranchCD"))
         lsSQL = ADO2SQL(lors, "Payroll_Annual_Total", , , , "xEmployID»xEmpLevID»xBasicPay»xDeptIDxx»sPyBranch»xBranchCD")
      Else
         lsSQL = ADO2SQL(lors, "Payroll_Annual_Total", "sPeriodID = " & strParm(fsPeriodID) & " AND sEmployID = " & strParm(fsEmployID), , , "xEmployID»xEmpLevID»xBasicPay»xDeptIDxx»sPyBranch")
      End If
   End If

   If lsSQL <> "" Then
      Call oApp.Execute(lsSQL, "Payroll_Annual_Total", lors("sBranchCD"))
   End If
End Sub

