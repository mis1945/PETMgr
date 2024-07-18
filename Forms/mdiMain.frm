VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Payroll, Employee, and Timekeeping Manager"
   ClientHeight    =   6060
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11280
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "mdiMain.frx":424A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmeLog 
      Interval        =   1000
      Left            =   2040
      Top             =   1200
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Object.ToolTipText     =   "Edit Mode"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   526
            MinWidth        =   526
            Object.ToolTipText     =   "Accounts Receivable Monitoring"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   526
            MinWidth        =   526
            Object.ToolTipText     =   "Motorcycle Monitoring"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   526
            MinWidth        =   526
            Object.ToolTipText     =   "Spareparts Monitoring"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   526
            MinWidth        =   526
            Object.ToolTipText     =   "SMS Hotline Monitoring"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   526
            MinWidth        =   526
            Object.ToolTipText     =   "Payroll Monitoring"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   526
            MinWidth        =   526
            Object.ToolTipText     =   "Order Redeem"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7938
            MinWidth        =   7938
            Object.ToolTipText     =   "Branch"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
            Object.ToolTipText     =   "Current User"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1850
            MinWidth        =   1850
            Object.ToolTipText     =   "System Date"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "&Files"
      Begin VB.Menu mnuHoliday 
         Caption         =   "Holiday"
      End
      Begin VB.Menu mnuShift 
         Caption         =   "Shift"
      End
      Begin VB.Menu mnuSSSChart 
         Caption         =   "SSS Chart"
      End
      Begin VB.Menu mnuWTaxChart 
         Caption         =   "WTax Chart"
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParameters 
         Caption         =   "Parameter"
         Begin VB.Menu mnuDepartment 
            Caption         =   "Department"
         End
         Begin VB.Menu mnuPosition 
            Caption         =   "Position/Designation"
         End
         Begin VB.Menu mnuFPSep5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEvalCategory 
            Caption         =   "Evaluation Category"
         End
         Begin VB.Menu mnuFPSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTitle 
            Caption         =   "Title"
         End
         Begin VB.Menu mnuFPSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSalaryLevel 
            Caption         =   "Salary Level"
         End
         Begin VB.Menu mnuEmpLevel 
            Caption         =   "Employee Level"
         End
         Begin VB.Menu mnuSalaryRegion 
            Caption         =   "Salary Region"
         End
         Begin VB.Menu mnuFPSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBenefits 
            Caption         =   "Benefits"
         End
         Begin VB.Menu mnuDeductions 
            Caption         =   "Deductions"
         End
         Begin VB.Menu mnuLoans 
            Caption         =   "Loans"
         End
         Begin VB.Menu mnuFPSep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIdentification 
            Caption         =   "Identification"
         End
         Begin VB.Menu mnuHobbies 
            Caption         =   "Hobbies"
         End
         Begin VB.Menu mnuInterest 
            Caption         =   "Interest"
         End
         Begin VB.Menu mnuSports 
            Caption         =   "Sports"
         End
         Begin VB.Menu mnuNationality 
            Caption         =   "Nationality"
         End
         Begin VB.Menu mnuBHSep9 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOJTApplication 
            Caption         =   "OJT Application"
         End
         Begin VB.Menu mnuSchool 
            Caption         =   "School"
         End
         Begin VB.Menu mnuCourse 
            Caption         =   "Course"
         End
         Begin VB.Menu mnuBHSep12 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMealType 
            Caption         =   "Meal Type"
         End
         Begin VB.Menu mnuShiftMeal 
            Caption         =   "Shift Meal"
         End
      End
   End
   Begin VB.Menu mnuHumanResources 
      Caption         =   "&Human Resources"
      Begin VB.Menu mnuEmpRecord 
         Caption         =   "Employee Record"
      End
      Begin VB.Menu mnuFHSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHSeminars 
         Caption         =   "Seminars/Trainings"
      End
      Begin VB.Menu mnuFHSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEvaluation 
         Caption         =   "Evaluation"
      End
      Begin VB.Menu mnuLeaveReset 
         Caption         =   "Leave Credits Reset"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuApplications 
         Caption         =   "Applications"
         Begin VB.Menu mnuPayLoans 
            Caption         =   "Loans"
         End
         Begin VB.Menu mnuPayAdvances 
            Caption         =   "Advances"
         End
         Begin VB.Menu mnuTASep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLeaveApplication 
            Caption         =   "Leave"
         End
         Begin VB.Menu mnuTardinessApplication 
            Caption         =   "Tardiness"
         End
         Begin VB.Menu mnuUndertimeApplication 
            Caption         =   "Undertime"
         End
         Begin VB.Menu mnuOBTApplication 
            Caption         =   "Business Trip"
         End
         Begin VB.Menu mnuOBTWLogApplication 
            Caption         =   "Business Trip W/ Log"
         End
         Begin VB.Menu mnuOTApplication 
            Caption         =   "Overtime"
         End
      End
      Begin VB.Menu mnuTRequests 
         Caption         =   "Requests"
         Begin VB.Menu mnuEmpMovement 
            Caption         =   "Employee Movement"
         End
         Begin VB.Menu mnuEmployeeHiring 
            Caption         =   "Employee Hiring"
         End
         Begin VB.Menu mnuEmployeeTermination 
            Caption         =   "Employee Termination"
         End
         Begin VB.Menu mnuTRSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSuspensionApplication 
            Caption         =   "Suspension"
         End
         Begin VB.Menu mnuTRSep02 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTAShiftMovement 
            Caption         =   "Shift Movement"
         End
         Begin VB.Menu mnuTADayoffShifting 
            Caption         =   "Day-off Shifting"
         End
         Begin VB.Menu mnuTAScheduleShifting 
            Caption         =   "Schedule Shifting"
         End
         Begin VB.Menu mnuTAScheduleShiftingBatch 
            Caption         =   "Schedule Shifting(Batch)"
         End
         Begin VB.Menu mnuCompOff 
            Caption         =   "Compensation Off"
         End
      End
      Begin VB.Menu mnuTSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayAdjustments 
         Caption         =   "Employee Adjustments"
      End
      Begin VB.Menu mnuPayTimesheetAdj 
         Caption         =   "Timesheet Adjustment"
      End
      Begin VB.Menu mnuTLeaveAdj 
         Caption         =   "Leave Adjustments"
      End
      Begin VB.Menu mnuPayDeductions 
         Caption         =   "Employee Deductions"
      End
      Begin VB.Menu mnu13thMPAdjustment 
         Caption         =   "13th Month Pay Adjustment"
      End
      Begin VB.Menu mnuBPEntry 
         Caption         =   "Bayanihan Program"
      End
      Begin VB.Menu mnuTSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApprovals 
         Caption         =   "Approvals"
         Begin VB.Menu mnuEmpLoansApprvl 
            Caption         =   "Loans"
         End
         Begin VB.Menu mnuPayAdvancesApprvl 
            Caption         =   "Advances"
         End
         Begin VB.Menu mnuTVSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLeaveApproval 
            Caption         =   "Leave"
         End
         Begin VB.Menu mnuOBTApproval 
            Caption         =   "Business Trip"
         End
         Begin VB.Menu mnuOBTWLogApproval 
            Caption         =   "Business Trip W/ Log"
         End
         Begin VB.Menu mnuOTApproval 
            Caption         =   "Overtime"
         End
         Begin VB.Menu mnuTVShiftMovement 
            Caption         =   "Shift Movement"
         End
         Begin VB.Menu mnuTVDayoffShifting 
            Caption         =   "Day-off Shifting"
         End
         Begin VB.Menu mnuTVScheduleShifting 
            Caption         =   "Schedule Shifting"
         End
         Begin VB.Menu mnuTVScheduleShiftingBatch 
            Caption         =   "Schedule Shifting(Batch)"
         End
         Begin VB.Menu mnuCompOffApp 
            Caption         =   "Compensation Off"
         End
         Begin VB.Menu mnuTVSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEmpMovementApprvl 
            Caption         =   "Employee Movement"
         End
         Begin VB.Menu mnuTVEmployeeHiring 
            Caption         =   "Employee Hiring"
         End
         Begin VB.Menu mnuTVEmployeeTermination 
            Caption         =   "Employee Termination"
         End
         Begin VB.Menu mnuSuspensionApproval 
            Caption         =   "Suspension"
         End
         Begin VB.Menu mnuTardinessApproval 
            Caption         =   "Tardiness"
         End
         Begin VB.Menu mnuTVUndertime 
            Caption         =   "Undertime"
         End
         Begin VB.Menu mnuTVSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPayAdjustmentsApprvl 
            Caption         =   "Employee Adjustments Approval"
         End
         Begin VB.Menu mnuPayDeductionsApprvl 
            Caption         =   "Employee Deductions Approval"
         End
         Begin VB.Menu mnu13thMPAdjustmentApproval 
            Caption         =   "13th Month Pay Adjustment Approval"
         End
      End
      Begin VB.Menu mnuMealVoucher 
         Caption         =   "Meal Voucher"
      End
      Begin VB.Menu mnuMealPosting 
         Caption         =   "Meal Posting"
      End
   End
   Begin VB.Menu mnuAttendance 
      Caption         =   "&Attendance"
      Begin VB.Menu mnuLogProcess 
         Caption         =   "Process Log"
      End
      Begin VB.Menu mnuTimesheetSummary 
         Caption         =   "Timesheet Summary"
      End
      Begin VB.Menu mnuTimesheetConfirm 
         Caption         =   "Timesheet Confirmation"
      End
      Begin VB.Menu mnuASep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogManual 
         Caption         =   "Manual Log - By Branch"
      End
      Begin VB.Menu mnuLogManualLevl 
         Caption         =   "Manual Log - By Level"
      End
      Begin VB.Menu mnuASep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAForgot 
         Caption         =   "Forgot to Log In/Out"
      End
   End
   Begin VB.Menu mnuPayroll 
      Caption         =   "&Payroll"
      Begin VB.Menu mnuPayrollComputation 
         Caption         =   "Payroll Computation"
      End
      Begin VB.Menu mnuPayrollChecker 
         Caption         =   "Payroll Checker"
      End
      Begin VB.Menu mnuPayrollSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpLoans 
         Caption         =   "Review Loans"
      End
      Begin VB.Menu mnuEmpAdjustments 
         Caption         =   "Review Adjustments"
      End
      Begin VB.Menu mnuEmpAdvances 
         Caption         =   "Review Advances"
      End
      Begin VB.Menu mnuEmployeeBenefits 
         Caption         =   "Review Benefits"
      End
      Begin VB.Menu mnuEmpDeductions 
         Caption         =   "Review Deductions"
      End
      Begin VB.Menu mnuEmp13thMonth 
         Caption         =   "Release 13th Month && Bonus"
      End
   End
   Begin VB.Menu mnuAdministrator 
      Caption         =   "Ad&ministrator"
      Begin VB.Menu mnuAdmApproval 
         Caption         =   "Approval"
         Begin VB.Menu mnuLogManualApprvl 
            Caption         =   "Log Manual"
         End
         Begin VB.Menu mnuTimesheetAdjlApprvl 
            Caption         =   "Timesheet Adjustment"
         End
         Begin VB.Menu mnuForgotLogApprvl 
            Caption         =   "Forgot to Log In/Out"
         End
         Begin VB.Menu mnuLeaveAdjApprvl 
            Caption         =   "Leave Adjustment"
         End
         Begin VB.Menu mnuBayanihanApproval 
            Caption         =   "Bayanihan Program"
         End
         Begin VB.Menu mnuServiceChargeApp 
            Caption         =   "Service Charge"
         End
      End
      Begin VB.Menu mnuLogOverride 
         Caption         =   "Approval Code Generator"
      End
      Begin VB.Menu mnuMSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu13thProcess 
         Caption         =   "Process 13th Month Pay"
      End
      Begin VB.Menu mnuBonusEntry 
         Caption         =   "Bonus && 13th Month Pay Entry"
      End
      Begin VB.Menu mnuPRTM 
         Caption         =   "Payroll Related Transactions Monitor"
      End
      Begin VB.Menu mnuTMIssuance 
         Caption         =   "Tardiness Memo Issuance"
      End
      Begin VB.Menu mnuServiceCharge 
         Caption         =   "Service Charge"
      End
      Begin VB.Menu mnuMSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMSettings 
         Caption         =   "Settings"
         Begin VB.Menu mnuMSPayrollPeriod 
            Caption         =   "Payroll Period"
         End
         Begin VB.Menu mnuMSAttendance 
            Caption         =   "Attendance"
         End
         Begin VB.Menu mnuMSReminders 
            Caption         =   "Reminders/Monitors"
         End
         Begin VB.Menu mnuMSRequest 
            Caption         =   "Request/Application"
         End
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout"
      End
   End
   Begin VB.Menu mnuHistory 
      Caption         =   "H&istory"
      Begin VB.Menu mnuIHumanResources 
         Caption         =   "Human Resources"
         Begin VB.Menu mnuIHIOCs 
            Caption         =   "Inter-Office Communication"
         End
      End
      Begin VB.Menu mnuIApplication 
         Caption         =   "Application"
         Begin VB.Menu mnuPayLoansReg 
            Caption         =   "Loans"
         End
         Begin VB.Menu mnuPayAdvancesReg 
            Caption         =   "Advances"
         End
         Begin VB.Menu mnuIASep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLeave 
            Caption         =   "Leave"
         End
         Begin VB.Menu mnuOvertime 
            Caption         =   "Overtime"
         End
         Begin VB.Menu mnuIATardiness 
            Caption         =   "Tardiness"
         End
         Begin VB.Menu mnuIAUndertime 
            Caption         =   "Undertime"
         End
         Begin VB.Menu mnuOfficialBusiness 
            Caption         =   "Business Trip"
         End
         Begin VB.Menu mnuOfficialBusinessWLog 
            Caption         =   "Business Trip W/ Log"
         End
      End
      Begin VB.Menu mnuIRequest 
         Caption         =   "Request"
         Begin VB.Menu mnuEmpMovementReg 
            Caption         =   "Employee Movement"
         End
         Begin VB.Menu mnuIRHiring 
            Caption         =   "Employee Hiring"
         End
         Begin VB.Menu mnuIRTermination 
            Caption         =   "Employee Termination"
         End
         Begin VB.Menu mnuIRSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSuspension 
            Caption         =   "Suspension"
         End
         Begin VB.Menu mnuIRSep02 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIRShiftMovement 
            Caption         =   "Shift Movement"
         End
         Begin VB.Menu mnuIRDayoffShifting 
            Caption         =   "Day-off Shifting"
         End
         Begin VB.Menu mnuIRScheduleShifting 
            Caption         =   "Schedule Shifting"
         End
         Begin VB.Menu mnuIRScheduleShiftingBatch 
            Caption         =   "Schedule Shifting(Batch)"
         End
         Begin VB.Menu mnuCompOffReg 
            Caption         =   "Compensation Off"
         End
      End
      Begin VB.Menu mnuITttendance 
         Caption         =   "Attendance"
         Begin VB.Menu mnuLogOverrideReg 
            Caption         =   "Log Override"
         End
         Begin VB.Menu mnuITSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLogManualReg 
            Caption         =   "Manual Log"
         End
         Begin VB.Menu mnuITForgot 
            Caption         =   "Forgot To Log In/Out"
         End
      End
      Begin VB.Menu mnuIPayroll 
         Caption         =   "Payroll"
         Begin VB.Menu mnuPayComputationReg 
            Caption         =   "Payroll Computation"
         End
         Begin VB.Menu mnuLoanAdjustmentReg 
            Caption         =   "Loan Adjustment"
         End
      End
      Begin VB.Menu mnuPayAdjustmentsReg 
         Caption         =   "Payroll Adjustments"
      End
      Begin VB.Menu mnuPayDeductionsReg 
         Caption         =   "Payroll Deductions"
      End
      Begin VB.Menu mnuPayTimesheetAdjReg 
         Caption         =   "Timesheet Adjustment"
      End
      Begin VB.Menu mnuTMIssuanceReg 
         Caption         =   "Tardiness Memo Issuance"
      End
      Begin VB.Menu mnuTMLeaveAdjReg 
         Caption         =   "Leave Adjustment"
      End
      Begin VB.Menu mnu13thMPAdjustmentReg 
         Caption         =   "13th Month Pay Adjustment"
      End
      Begin VB.Menu mnuBayanihanReg 
         Caption         =   "Bayanihan Program"
      End
      Begin VB.Menu mnuMealPostingRegister 
         Caption         =   "Meal Posting"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Re&ports"
      Begin VB.Menu mnuGeneralReport 
         Caption         =   "General Reports"
      End
      Begin VB.Menu mnuRSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelfieLogReport 
         Caption         =   "Selfie Log Reports"
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuContactTracker 
         Caption         =   "Contact Tracker"
      End
      Begin VB.Menu mnuGenTransfer 
         Caption         =   "General Transfer"
         Begin VB.Menu mnuDocTransfer 
            Caption         =   "Document Transfer"
         End
         Begin VB.Menu mnuDocTransPosting 
            Caption         =   "Document Transfer Posting"
         End
      End
      Begin VB.Menu mnuImportAttendance 
         Caption         =   "Import Attendance"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnMsg As String

Private Const pxeJavaPath As String = "D:\GGC_Java_Systems\"
Private Const pxBRANCHCODES = "M001»M0W1»A001»N001»H002"

'Needed this simple utility to reset the loan
'since we are also capturing the loan from the old payroll
'If the new payroll is really up and running then feel free to delete the
'sub-procedure...
'kalyptus - 2012.05.02

Private Sub MDIForm_DblClick()
   Dim loObj As clsBatchLeave
   Set loObj = New clsBatchLeave
   Set loObj.AppDriver = oApp
   loObj.InitTransaction
   loObj.NewTransaction
   loObj.SaveTransaction
   
'   Dim lsSQL As String
'   Dim loRS As Recordset
'   Dim lsDate As String
'
'   If oApp.UserLevel = xeEngineer Then
'      lsDate = InputBox("Please enter the date", "Loan Reset")
'
'      If Not IsDate(lsDate) Then Exit Sub
'
'      lsSQL = "SELECT" & _
'                     "  a.sTransNox" & _
'                     ", IFNULL(b.nBalancex + b.nAmountxx, a.nBalancex) nBalancex" & _
'             " FROM Employee_Loan_Master a" & _
'                  " , Employee_Loan_Ledger b" & _
'             " WHERE a.sTransNox = b.sTransNox" & _
'               " AND b.dTransact = " & dateParm(lsDate)
'      Set loRS = oApp.Connection.Execute(lsSQL, , adCmdText)
'
'      Do Until loRS.EOF
'         lsSQL = "UPDATE Employee_Loan_Master" & _
'                    " SET nBalancex = " & loRS("nBalancex") & _
'                       ", cHoldDedx = " & strParm(0) & _
'                " WHERE sTransNox = " & strParm(loRS("sTransNox"))
'         Call oApp.Execute(lsSQL, "Employee_Loan_Master")
'
'         lsSQL = "DELETE" & _
'                " FROM Employee_Loan_Ledger " & _
'                " WHERE sTransNox = " & strParm(loRS("sTransNox")) & _
'                  " AND dTransact = " & dateParm(lsDate)
'         Call oApp.Execute(lsSQL, "Employee_Loan_Master")
'
'         loRS.MoveNext
'      Loop
'   End If
End Sub

Private Sub mnu13thMonthComputation_Click()
   
End Sub

Private Sub mnu13thMPAdjustment_Click()
   frm13thMonthAdjustments.Tag = "mnu13thMPAdjustment"
   frm13thMonthAdjustments.Show
End Sub

Private Sub mnu13thMPAdjustmentApproval_Click()
   frm13thMonthAdjustmentsApprvl.Tag = "mnu13thMPAdjustmentApproval"
   frm13thMonthAdjustmentsApprvl.Show
End Sub

Private Sub mnu13thMPAdjustmentReg_Click()
   frm13thMonthAdjustmentsReg.Tag = "mnu13thMPAdjustmentReg"
   frm13thMonthAdjustmentsReg.Show
End Sub

Private Sub mnu13thProcess_Click()
   frm13thMonth.Tag = "mnu13thProcess"
   frm13thMonth.Show
End Sub

Private Sub mnuAForgot_Click()
   frmForgot2Swipe.Tag = "mnuAForgot"
   frmForgot2Swipe.Show
End Sub

Private Sub mnuBayanihanApproval_Click()
   frmBayanihanApproval.Tag = "mnuBayanihanApproval"
   frmBayanihanApproval.Show
End Sub

Private Sub mnuBayanihanReg_Click()
   frmBayanihanReg.Tag = "mnuBayanihanReg"
   frmBayanihanReg.Show
End Sub

Private Sub mnuBenefits_Click()
   frmBenefits.Tag = "mnuBenefits"
   frmBenefits.Show
End Sub

Private Sub mnuBonusEntry_Click()
   frmEmp13thMonth.Tag = "mnuBonusEntry"
   frmEmp13thMonth.Show
End Sub

Private Sub mnuBPEntry_Click()
   frmBayanihan.Tag = "mnuBPEntry"
   frmBayanihan.Show
End Sub

Private Sub mnuCompOff_Click()
   frmCompensationOffApplication.Tag = "mnuCompOff"
   frmCompensationOffApplication.Show
End Sub

Private Sub mnuCompOffApp_Click()
   frmCompensationOffApproval.Tag = "mnuCompOffApp"
   frmCompensationOffApproval.Show
End Sub

Private Sub mnuCompOffReg_Click()
   frmCompensationOffReg.Tag = "mnuCompOffReg"
   frmCompensationOffReg.Show
End Sub

Private Sub mnuContactTracker_Click()
   frmContactTrace.Tag = "mnuContactTracker"
   frmContactTrace.Show
End Sub

Private Sub mnuCourse_Click()
'   frmCourse.Tag = "mnuCourse"
'   frmCourse.Show

   frmMealOjtCourse.Tag = "mnuCourse"
   frmMealOjtCourse.Show
End Sub

Private Sub mnuDeductions_Click()
   frmDeductions.Tag = "mnuDeductions"
   frmDeductions.Show
End Sub

Private Sub mnuDepartment_Click()
   frmDepartmentneo.Tag = "mnuDepartment"
   frmDepartmentneo.Show
End Sub

Private Sub mnuDocTransfer_Click()
   frmGeneralTransfer.Tag = "mnuDocTransfer"
   frmGeneralTransfer.Show
End Sub

Private Sub mnuDocTransPosting_Click()
   frmGeneralTransferPosting.Tag = "mnuDocTransPosting"
   frmGeneralTransferPosting.Show
End Sub

Private Sub mnuEmp13thMonth_Click()
   frmEmp13thMonthRelease.Tag = "mnuEmp13thMonth"
   frmEmp13thMonthRelease.Show
End Sub

Private Sub mnuEmpAdjustments_Click()
   frmEmpAdjustments.Tag = "mnuEmpAdjustments"
   frmEmpAdjustments.Show
End Sub

Private Sub mnuEmpAdvances_Click()
   frmEmpAdvances.Tag = "mnuEmpAdvances"
   frmEmpAdvances.Show
End Sub

Private Sub mnuEmpDeductions_Click()
   frmEmpDeductions.Tag = "mnuEmpDeductions"
   frmEmpDeductions.Show
End Sub

Private Sub mnuEmpLoansApprvl_Click()
   frmEmployeeLoansApprvl.Tag = "mnuEmpLoansApprvl"
   frmEmployeeLoansApprvl.Show
End Sub

Private Sub mnuEvalCategory_Click()
'   frmEvalCategory.Tag = "mnuEvalCategory"
'   frmEvalCategory.Show
End Sub

Private Sub mnuEvaluation_Click()
   frmEmpAppraisal.Tag = "mnuEvaluation"
   frmEmpAppraisal.Show
End Sub

Private Sub mnuForgotLogApprvl_Click()
   frmForgot2SwipeApprvl.Tag = "mnuForgotLogApprvl"
   frmForgot2SwipeApprvl.Show
End Sub

Private Sub mnuGeneralReport_Click()
   Dim loReports As clsPETRep
   Dim loRepViewer As frmRepViewer

      Set loReports = New clsPETRep
         With loReports
            Set .AppDriver = oApp
               If .ShowReport Then
               Set loRepViewer = New frmRepViewer
               Set loRepViewer.ReportSource = .Source

               loRepViewer.Show
               .closeReport
         End If
   End With
End Sub

Private Sub mnuHSeminars_Click()
   frmTrainings.Tag = "mnuHSeminars"
   frmTrainings.Show
End Sub

Private Sub mnuIATardiness_Click()
   frmTardinessReg.Tag = "mnuIATardiness"
   frmTardinessReg.Show
End Sub

Private Sub mnuIAUndertime_Click()
   frmUndertimeReg.Tag = "mnuIAUndertime"
   frmUndertimeReg.Show
End Sub

Private Sub mnuImportAttendance_Click()


   Dim loCls As clsLogCapture
   Set loCls = New clsLogCapture

   Set loCls.AppDriver = oApp
   loCls.Import
   
End Sub

Private Sub mnuIRDayoffShifting_Click()
   frmDayOffShftReg.Tag = "mnuIRDayoffShifting"
   frmDayOffShftReg.Show
End Sub

Private Sub mnuIRScheduleShifting_Click()
   frmShiftSchedReg.Tag = "mnuIRScheduleShifting"
   frmShiftSchedReg.Show
End Sub

Private Sub mnuIRScheduleShiftingBatch_Click()
    frmBatchShftReg.Tag = "mnuIRScheduleShiftingBatch"
    frmBatchShftReg.Show
End Sub

Private Sub mnuITForgot_Click()
   frmForgot2SwipeReg.Tag = "mnuITForgot"
   frmForgot2SwipeReg.Show
End Sub

Private Sub mnuLeaveAdjApprvl_Click()
   frmLeaveAdjApproval.Tag = "mnuLeaveAdjApprvl"
   frmLeaveAdjApproval.Show
End Sub

Private Sub mnuLeaveReset_Click()
   If oApp.getConfiguration("BatchLve") <> "1" Then
      MsgBox "Reset of leave is not yet initiated. !!!" & vbCrLf & vbCrLf & _
            "For inquiry, please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      Exit Sub
   End If
   
   frmLeaveReset.Tag = "mnuLeaveReset"
   frmLeaveReset.Show
End Sub

Private Sub mnuLoanAdjustmentReg_Click()
   frmLoanAdjustmentReg.Tag = "mnuLoanAdjustmentReg"
   frmLoanAdjustmentReg.Show
End Sub

Private Sub mnuLogManualApprvl_Click()
'   frmLogManualApprvl2.Tag = "mnuLogManualApprvl"
'   frmLogManualApprvl2.Show
   frmLogManualApprvlWR.Tag = "mnuLogManualApprvl"
   frmLogManualApprvlWR.Show
End Sub

Private Sub mnuLogManualLevl_Click()
'   frmLogManual2.Tag = "mnuLogManualLevl"
'   frmLogManual2.ByBranch = False
'   frmLogManual2.Show
   
   frmLogManualWR.Tag = "mnuLogManualLevl"
   frmLogManualWR.ByBranch = False
   frmLogManualWR.Show
End Sub

Private Sub mnuLogManualReg_Click()
'   frmLogManualReg2.Tag = "mnuLogManualReg"
'   frmLogManualReg2.Show
   
   frmLogManualRegWR.Tag = "mnuLogManualReg"
   frmLogManualRegWR.Show
End Sub

Private Sub mnuLogout_Click()
   lnMsg = MsgBox("Do you want to terminate current session???", vbYesNo + vbQuestion, "Confirm")
   If lnMsg = vbYes Then
   End
   End If
End Sub

Private Sub mnuMealPosting_Click()
   frmMealPosting.Tag = "mnuMealPosting"
   frmMealPosting.Show
End Sub

Private Sub mnuMealPostingRegister_Click()
'   frmMealPostingRegister.Tag = "mnuMealPostingRegister"
'   frmMealPostingRegister.Show
End Sub

Private Sub mnuMealType_Click()
   frmMealType.Tag = "mnuMealType"
   frmMealType.Show
End Sub

Private Sub mnuMealVoucher_Click()
   frmMealVoucher.Tag = "mnuMealVoucher"
   frmMealVoucher.Show
End Sub

Private Sub mnuNationality_Click()
   frmCountry.Tag = "mnuNationality"
   frmCountry.Show
End Sub

Private Sub mnuOBTWLogApplication_Click()
   frmOBWithLogApp.Tag = "mnuOBTWLogApplication"
   frmOBWithLogApp.Show
End Sub

Private Sub mnuOBTWLogApproval_Click()
   frmOBWithLogApprvl.Tag = "mnuOBTWLogApproval"
   frmOBWithLogApprvl.Show
End Sub

Private Sub mnuOfficialBusinessWLog_Click()
   frmOBWithLogReg.Tag = "mnuOfficialBusinessWLog"
   frmOBWithLogReg.Show
End Sub

Private Sub mnuOJTApplication_Click()
   frmMealOjtApplication.Tag = "mnuOJTApplication"
   frmMealOjtApplication.Show
End Sub

Private Sub mnuPayAdjustments_Click()
   frmEmployeeAdjustments.Tag = "mnuPayAdjustments"
   frmEmployeeAdjustments.Show
End Sub

Private Sub mnuPayAdjustmentsApprvl_Click()
   frmEmployeeAdjustmentsApprvl.Tag = "mnuPayAdjustmentsApprvl"
   frmEmployeeAdjustmentsApprvl.Show
End Sub

Private Sub mnuPayAdjustmentsReg_Click()
   frmEmployeeAdjustmentsReg.Tag = "mnuPayAdjustmentsReg"
   frmEmployeeAdjustmentsReg.Show
End Sub

Private Sub mnuPayAdvances_Click()
   frmEmployeeAdvances.Tag = "mnuPayAdvances"
   frmEmployeeAdvances.Show
End Sub

Private Sub mnuPayAdvancesApprvl_Click()
   frmEmployeeAdvancesApprvl.Tag = "mnuPayAdvancesApprvl"
   frmEmployeeAdvancesApprvl.Show
End Sub

Private Sub mnuPayAdvancesReg_Click()
   frmEmployeeAdvancesReg.Tag = "mnuPayAdvancesReg"
   frmEmployeeAdvancesReg.Show
End Sub

Private Sub mnuPayComputationReg_Click()
   Dim lcEmpTypex As String
   
   lcEmpTypex = getEmpType

   If lcEmpTypex <> "" Then
      frmPayrollReg.EmployeeType = Left(lcEmpTypex, 1)
      frmPayrollReg.isMainOffice = Right(lcEmpTypex, 1)
      frmPayrollReg.Caption = "Payroll Computation - " & IIf(lcEmpTypex = "T", "Trainees", IIf(lcEmpTypex = "R", "Regular Employees", "Level A Employees"))
      
      If Right(lcEmpTypex, 1) = "0" Then
         frmPayrollReg.Caption = frmPayrollReg.Caption & "(Branch)"
      End If
      
      frmPayrollReg.Tag = "mnuPayComputationReg"
      frmPayrollReg.Show
   End If
End Sub

Private Sub mnuPayDeductions_Click()
   frmEmployeeDeductions.Tag = "mnuPayDeductions"
   frmEmployeeDeductions.Show
End Sub

Private Sub mnuPayDeductionsApprvl_Click()
   frmEmployeeDeductionsApprvl.Tag = "mnuPayDeductionsApprvl"
   frmEmployeeDeductionsApprvl.Show
End Sub

Private Sub mnuPayDeductionsReg_Click()
   frmEmployeeDeductionsReg.Tag = "mnuPayDeductionsReg"
   frmEmployeeDeductionsReg.Show
End Sub

Private Sub mnuEmpLevel_Click()
   frmEmpLevel.Tag = "mnuEmpLevel"
   frmEmpLevel.Show
End Sub

Private Sub mnuEmpLoans_Click()
   frmEmpLoans.Tag = "mnuEmpLoans"
   frmEmpLoans.Show
End Sub

Private Sub mnuEmployeeBenefits_Click()
   frmEmpBenefits.Tag = "mnuEmployeeBenefits"
   frmEmpBenefits.Show
End Sub

Private Sub mnuEmpMovement_Click()
   frmEmployeeMovement.Tag = "mnuEmpMovement"
   frmEmployeeMovement.Show
End Sub

Private Sub mnuEmpMovementApprvl_Click()
   frmEmployeeMovementApprvl.Tag = "mnuEmpMovementApprvl"
   frmEmployeeMovementApprvl.Show
End Sub

Private Sub mnuEmpMovementReg_Click()
   frmEmployeeMovementReg.Tag = "mnuEmpMovementReg"
   frmEmployeeMovementReg.Show
End Sub

Private Sub mnuEmpRecord_Click()
   frmEmployee.Tag = "mnuEmpRecord"
   frmEmployee.Show
End Sub

Private Sub mnuHobbies_Click()
   frmHobbies.Tag = "mnuHobbies"
   frmHobbies.Show
End Sub

Private Sub mnuHoliday_Click()
   frmHoliday.Tag = "mnuHoliday"
   frmHoliday.Show
End Sub

Private Sub mnuIdentification_Click()
   frmIdentification.Tag = "mnuIdentification"
   frmIdentification.Show
End Sub

Private Sub mnuInterest_Click()
   frmInterests.Tag = "mnuInterest"
   frmInterests.Show
End Sub

Private Sub mnuLeave_Click()
   frmLeaveReg.Tag = "mnuLeave"
   frmLeaveReg.Show
End Sub

Private Sub mnuLeaveApplication_Click()
   frmLeaveApplication.Tag = "mnuLeaveApplication"
   frmLeaveApplication.Show
End Sub

Private Sub mnuLeaveApproval_Click()
   frmLeaveApproval.Tag = "mnuLeaveApproval"
   frmLeaveApproval.Show
End Sub

Private Sub mnuLoans_Click()
   frmLoans.Tag = "mnuLoans"
   frmLoans.Show
End Sub

Private Sub mnuLogManual_Click()
'   frmLogManual2.Tag = "mnuLogManual"
'   frmLogManual2.ByBranch = True
'   frmLogManual2.Show
   
   frmLogManualWR.Tag = "mnuLogManual"
   frmLogManualWR.ByBranch = True
   frmLogManualWR.Show
End Sub

Private Sub mnuLogOverride_Click()
   frmLogOverride.Tag = "mnuLogOverride"
   frmLogOverride.Show
End Sub

Private Sub mnuLogOverrideReg_Click()
   frmLogOverrideReg.Tag = "mnuLogOverrideReg"
   frmLogOverrideReg.Show
End Sub

Private Sub mnuLogProcess_Click()
   frmLogProcess.Tag = "mnuLogProcess"
   frmLogProcess.Show
End Sub

Private Sub mnuOBTApplication_Click()
   frmOBApplication.Tag = "mnuOBTApplication"
   frmOBApplication.Show
End Sub

Private Sub mnuOBTApproval_Click()
   frmOBApproval.Tag = "mnuOBTApproval"
   frmOBApproval.Show
End Sub

Private Sub mnuOfficialBusiness_Click()
   frmOBReg.Tag = "mnuOfficialBusiness"
   frmOBReg.Show
End Sub

Private Sub mnuOTApplication_Click()
   frmOTApplication.Tag = "mnuOTApplication"
   frmOTApplication.Show
End Sub

Private Sub mnuOTApproval_Click()
   frmOTApproval.Tag = "mnuOTApproval"
   frmOTApproval.Show
End Sub

Private Sub mnuOvertime_Click()
   frmOTReg.Tag = "mnuOvertime"
   frmOTReg.Show
End Sub

Private Sub mnuPayLoans_Click()
   frmEmployeeLoans.Tag = "mnuPayLoans"
   frmEmployeeLoans.Show
End Sub

Private Sub mnuPayLoansReg_Click()
   frmEmployeeLoansReg.Tag = "mnuPayLoansReg"
   frmEmployeeLoansReg.Show
End Sub

Private Sub mnuPayrollChecker_Click()
   frmPayrollDiscrepancyChecker.Tag = "mnuPayrollChecker"
   frmPayrollDiscrepancyChecker.Show
End Sub

Private Sub mnuPayrollComputation_Click()
   Dim lcEmpTypex As String
   
   lcEmpTypex = getEmpType

   If lcEmpTypex <> "" Then
      frmPayroll.EmployeeType = Left(lcEmpTypex, 1)
'      frmPayroll.isMainOffice = Right(lcEmpTypex, 1)
      frmPayroll.isMainOffice = IIf(Left(lcEmpTypex, 1) = "T", "0", "1")
      frmPayroll.Caption = "Payroll Computation - " & IIf(Left(lcEmpTypex, 1) = "T", "Trainees", IIf(Left(lcEmpTypex, 1) = "R", "Regular Employees", "Level A Employees"))
      
'      If Right(lcEmpTypex, 1) = "0" Then
'         frmPayroll.Caption = frmPayroll.Caption & "(Branch)"
'      End If
      
      frmPayroll.Tag = "mnuPayrollComputation"
      frmPayroll.Show
   End If
End Sub

Private Sub mnuPayTimesheetAdj_Click()
   frmTimesheetAdjustments.Tag = "mnuPayTimesheetAdj"
   frmTimesheetAdjustments.Show
End Sub

Private Sub mnuPayTimesheetAdjReg_Click()
   frmTimesheetAdjustmentsReg.Tag = "mnuPayTimesheetAdjReg"
   frmTimesheetAdjustmentsReg.Show
End Sub

Private Sub mnuPosition_Click()
   frmPosition.Tag = "mnuPosition"
   frmPosition.Show
End Sub

Private Sub mnuPostLeaveReset_Click()
   frmLeaveResetPosting.Tag = "mnuPostLeaveReset"
   frmLeaveResetPosting.Show
End Sub

Private Sub mnuPRTM_Click()
'   Dim loCls As clsPRTM
'   Set loCls = New clsPRTM
'   Set loCls.AppDriver = oApp
'   loCls.isShowMsg = True
'   loCls.ProcessReport
End Sub

Private Sub mnuSalaryLevel_Click()
   frmSalLevel.Tag = "mnuSalaryLevel"
   frmSalLevel.Show
End Sub

Private Sub mnuSalaryRegion_Click()
   frmSalRegion.Tag = "mnuSalaryRegion"
   frmSalRegion.Show
End Sub

Private Sub mnuSchool_Click()
   frmMealOjtSchool.Tag = "mnuSchool"
   frmMealOjtSchool.Show
End Sub

Private Sub mnuSelfieLogReport_Click()
   Dim loSelfLogViewer As frmSelfieLogCriteria
   Set loSelfLogViewer = New frmSelfieLogCriteria
   Set loSelfLogViewer.AppDriver = oApp
   loSelfLogViewer.Show 1
   
   If (loSelfLogViewer.Employee <> "") Then
      Dim lsArguments As String
      Dim lnResult As Long
      
      If (Dir(pxeJavaPath & "selfie_route.bat") <> "") Then
         lsArguments = oApp.ProductID & " " & oApp.UserID _
         & " " & loSelfLogViewer.Employee _
         & " " & Format(loSelfLogViewer.DateFrom, "YYYY-mm-dd") _
         & " " & Format(loSelfLogViewer.DateThru, "YYYY-mm-dd")
            
         lnResult = (RMJExecute(pxeJavaPath & "selfie_route.bat " & lsArguments))
         
         If (lnResult = 1) Then
            MsgBox "Unable to load DCP Route for this date. Unable to load map.!", vbInformation, "Notice"
         ElseIf (lnResult = 2) Then
            MsgBox "System error. Please inform MIS Support to fix the issue.", vbInformation, "Notice"
         End If
      Else 'path check
         MsgBox "File Path Does'nt Exist  " & pxeJavaPath & "selfie_route.bat" & "   Please Inform MIS Dept !!", vbInformation, "Notice"
      End If
   End If
End Sub

Private Sub mnuServiceCharge_Click()
   frmServiceCharge.Tag = "mnuServiceCharge"
   frmServiceCharge.Show
End Sub

Private Sub mnuServiceChargeApp_Click()
   frmServiceChargeApproval.Tag = "mnuServiceChargeApp"
   frmServiceChargeApproval.Show
End Sub

'Private Sub mnuSalLvlRegion_Click()
'   frmSalLvlRegion.Tag = "mnuSalLvlRegion"
'   frmSalLvlRegion.Show
'End Sub

Private Sub mnuShift_Click()
   frmShift.Tag = "mnuShift"
   frmShift.Show
End Sub

Private Sub mnuShiftMeal_Click()
   frmShiftMealSchedue.Tag = "mnuShiftMeal"
   frmShiftMealSchedue.Show
End Sub

Private Sub mnuSports_Click()
   frmSports.Tag = "mnuSports"
   frmSports.Show
End Sub

Private Sub mnuSSSChart_Click()
   frmSSSChart.Tag = "mnuSSSChart"
   frmSSSChart.Show
End Sub

Private Sub mnuSuspension_Click()
   frmSuspensionReg.Tag = "mnuSuspension"
   frmSuspensionReg.Show
End Sub

Private Sub mnuSuspensionApplication_Click()
   frmSuspensionApplication.Tag = "mnuSuspensionApplication"
   frmSuspensionApplication.Show
End Sub

Private Sub mnuSuspensionApproval_Click()
   frmSuspensionApproval.Tag = "mnuSuspensionApproval"
   frmSuspensionApproval.Show
End Sub

Private Sub mnuTADayoffShifting_Click()
   frmDayOffShftApplication.Tag = "mnuTADayoffShifting"
   frmDayOffShftApplication.Show
End Sub

Private Sub mnuTardinessApplication_Click()
   frmTardiness.Tag = "mnuTardinessApplication"
   frmTardiness.Show
End Sub

Private Sub mnuTardinessApproval_Click()
   frmTardinessApproval.Tag = "mnuTardinessApproval"
   frmTardinessApproval.Show
End Sub

Private Sub mnuTAScheduleShifting_Click()
   frmShiftSchedApplication.Tag = "mnuTAScheduleShifting"
   frmShiftSchedApplication.Show
End Sub

Private Sub mnuTAScheduleShiftingBatch_Click()
   frmBatchShiftChange.Tag = "mnuTAScheduleShiftingBatch"
   frmBatchShiftChange.Show
End Sub

Private Sub mnuTimesheetAdjlApprvl_Click()
   frmTimesheetAdjustmentsApprvl.Tag = "mnuTimesheetAdjlApprvl"
   frmTimesheetAdjustmentsApprvl.Show
End Sub

Private Sub mnuTimesheetConfirm_Click()
   frmTimesheetConfirmation.Tag = "mnuTimesheetConfirm"
   frmTimesheetConfirmation.Show
End Sub

Private Sub mnuTimesheetSummary_Click()
   frmTimesheetSummary.Tag = "mnuTimesheetSummary"
   frmTimesheetSummary.Show
End Sub

Private Sub mnuTitle_Click()
   frmTitle.Tag = "mnuTitle"
   frmTitle.Show
End Sub

Private Sub mnuTLeaveAdj_Click()
   frmLeaveAdjustment.Tag = "mnuTLeaveAdj"
   frmLeaveAdjustment.Show
End Sub

Private Sub mnuTMIssuance_Click()
   frmTardinessMemoIssuance.Tag = "mnuTMIssuance"
   frmTardinessMemoIssuance.Show
End Sub

Private Sub mnuTMIssuanceReg_Click()
   frmTardinessMemoIssuanceReg.Tag = "mnuTMIssuanceReg"
   frmTardinessMemoIssuanceReg.Show
End Sub

Private Sub mnuTMLeaveAdjReg_Click()
   frmLeaveAdjReg.Tag = "mnuTMLeaveAdjReg"
   frmLeaveAdjReg.Show
End Sub

Private Sub mnuTVDayoffShifting_Click()
   frmDayOffShftApproval.Tag = "mnuTVDayoffShifting"
   frmDayOffShftApproval.Show
End Sub

Private Sub mnuTVScheduleShifting_Click()
   frmShiftSchedApproval.Tag = "mnuTVScheduleShifting"
   frmShiftSchedApproval.Show
End Sub

Private Function getLastPeriod(ByVal fsEmployID As String) As Date
   Dim lsSQL As String
   Dim lors As Recordset
   
   lsSQL = "SELECT" & _
                  " a.dCovergTo" & _
          " FROM Payroll_Period a" & _
              " LEFT JOIN Payroll_Summary b ON a.sPayPerID = b.sPayPerID" & _
          " WHERE b.sEmployID = " & strParm(fsEmployID) & _
          " ORDER BY a.dCovergTo DESC LIMIT 1"
   Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
   
   If lors.EOF Then
      getLastPeriod = Format(oApp.ServerDate, "yyyy-mm-dd")
   Else
      getLastPeriod = lors("dCovergTo") + 1
   End If
   
End Function

Private Sub quickTransfer(ByVal fsTable As String, ByVal fsFilter As String, ByVal fsBranchCD As String)
'   Dim lors As Recordset
'   Dim lsSQL As String
'
'   Set lors = GetRecordSet(oApp.Connection, fsTable, fsFilter)
'   Do Until lors.EOF
'      lsSQL = ADO2SQL(lors, fsTable)
'      lsSQL = Replace(lsSQL, "INSERT INTO", "REPLACE INTO")
'
'      Call send2Log( _
'         oApp.Connection, _
'         oApp.BranchCode, _
'         oApp.BranchCode, _
'         lsSQL, _
'         fsTable, _
'         fsBranchCD, _
'         oApp.UserID, _
'         oApp.ServerDate, _
'         True)
'
'      lors.MoveNext
'   Loop
   
End Sub

Private Sub mnuTVScheduleShiftingBatch_Click()
    frmBatchShftApproval.Tag = "mnuTVScheduleShiftingBatch"
    frmBatchShftApproval.Show
End Sub

Private Sub mnuTVUndertime_Click()
   frmUndertimeApproval.Tag = "mnuTVUndertime"
   frmUndertimeApproval.Show
End Sub

Private Sub mnuUndertimeApplication_Click()
   frmUndertime.Tag = "mnuUndertimeApplication"
   frmUndertime.Show
End Sub

Private Sub mnuWTaxChart_Click()
'   frmWTax.Tag = "mnuWTaxChart"
'   frmWTax.Period = "S"
'   frmWTax.Show
   frmWTaxChart.Show
End Sub

Private Sub tmeLog_Timer()
   Dim lsSQL As String
   Dim lors As Recordset
   Dim loCls As clsEmployeeMovement
   Dim ldDateFrom As Date
   Dim lcDivision As String
   Dim loDiviosn As Recordset
   
   'she 2021-11-30
   Set loDiviosn = New Recordset
   lsSQL = "SELECT * FROM Branch_Others WHERE sBranchCd = " & strParm(oApp.getConfiguration("PetMstr"))
   Debug.Print lsSQL
   loDiviosn.Open lsSQL, oApp.Connection, , adCmdText
   
   If Not loDiviosn.EOF Then
      'she 2022-04-08
      If oApp.BranchCode = "H002" Then
         lcDivision = 3
      Else
         lcDivision = loDiviosn("cDivision")
      End If
   Else
      lcDivision = ""
   End If
   
   DoEvents
   lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", a.sEmployID" & _
                  ", a.sBranchCD" & _
                  ", a.xBranchCD" & _
                  ", a.dEffectve" & _
          " FROM Employee_Movement a" & _
               " LEFT JOIN Branch_Others b on a.sBranchCd = b.sBranchCd" & _
               " LEFT JOIN Division c ON b.cDivision = c.sDivsnCde" & _
          " WHERE a.cTranStat IN ('1', '2')" & _
            " AND a.dEffectve < " & dateParm(oApp.ServerDate) & _
            " AND a.dEffectve > " & dateParm("2013-04-26") & _
            " AND c.sDivsnCde = " & strParm(lcDivision) & _
          " ORDER BY a.dEffectve, a.sTransNox"
   Debug.Print lsSQL
   If InStr(1, pxBRANCHCODES, oApp.BranchCode, vbTextCompare) = 0 Then
'   If oApp.BranchCode <> "M001" Then
      Exit Sub
   End If

   Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)

   DoEvents

   If lors.EOF Then Exit Sub

   Set loCls = New clsEmployeeMovement
   Set loCls.AppDriver = oApp
   loCls.HasParent = True
   loCls.InitTransaction

   DoEvents
   Do Until lors.EOF
      DoEvents
      If LCase(oApp.ProductID) = "petmgr" Then
         'Its from the main office so send updates to all branches...
         If loCls.OpenTransaction(lors("sTransNox")) Then
            Call loCls.PostTransaction(lors("sTransNox"))
         End If
      Else
         'if monitor is not from main office then just post the movement
         lsSQL = "UPDATE Employee_Movement" & _
                " SET cTranStat = " & strParm(xeStatePosted) & _
                " WHERE sTransNox = " & strParm(lors("sTransNox"))
         oApp.Connection.Execute lsSQL, , adCmdText
      End If

       'From this Branch employee is assigned to other branch
       If IFNull(loCls.Master("sBranchCD")) <> "" _
      And loCls.Master("sBranchCD") <> oApp.BranchCode _
      And IFNull(loCls.Master("xBranchCD"), "") = oApp.BranchCode _
      And InStr(1, pxBRANCHCODES, lors("sBranchCD")) = 0 Then

          DoEvents
          ldDateFrom = getLastPeriod(loCls.Master("sEmployID"))
          Call quickTransfer("Employee_Log", _
                             "sEmployID = " & strParm(loCls.Master("sEmployID")) & _
                        " AND dTransact BETWEEN " & dateParm(ldDateFrom) & " AND " & dateParm(loCls.Master("dEffectve")), _
                             loCls.Master("sBranchCD"))
          DoEvents

          Call quickTransfer("Employee_Timesheet", _
                             "sEmployID = " & strParm(loCls.Master("sEmployID")) & _
                        " AND dTransact BETWEEN " & dateParm(ldDateFrom) & " AND " & dateParm(loCls.Master("dEffectve")), _
                             loCls.Master("sBranchCD"))
          DoEvents

          Call quickTransfer("Employee_Leave", _
                             "sEmployID = " & strParm(loCls.Master("sEmployID")) & _
                        " AND dApproved BETWEEN " & dateParm(ldDateFrom) & " AND " & dateParm(loCls.Master("dEffectve")), _
                             loCls.Master("sBranchCD"))
          DoEvents

          Call quickTransfer("Employee_Business_Trip", _
                             "sEmployID = " & strParm(loCls.Master("sEmployID")) & _
                        " AND dApproved BETWEEN " & dateParm(ldDateFrom) & " AND " & dateParm(loCls.Master("dEffectve")), _
                             loCls.Master("sBranchCD"))
      End If

      lors.MoveNext
      DoEvents
   Loop
End Sub


