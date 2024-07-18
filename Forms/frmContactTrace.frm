VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmContactTrace 
   BorderStyle     =   0  'None
   Caption         =   "Employee Contact Tracker"
   ClientHeight    =   10050
   ClientLeft      =   105
   ClientTop       =   0
   ClientWidth     =   20100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   20100
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   105
      TabIndex        =   9
      Top             =   9600
      Width           =   18435
      _ExtentX        =   32517
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   840
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   18435
      _ExtentX        =   32517
      _ExtentY        =   1482
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   9600
         MaxLength       =   50
         TabIndex        =   5
         Top             =   75
         Width           =   4815
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   3
         Top             =   420
         Width           =   4815
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   9600
         MaxLength       =   50
         TabIndex        =   7
         Top             =   420
         Width           =   4815
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   1
         Top             =   75
         Width           =   4815
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay"
         Height          =   195
         Index           =   1
         Left            =   8580
         TabIndex        =   4
         Top             =   135
         Width           =   675
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Town/City"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         Height          =   195
         Index           =   25
         Left            =   8100
         TabIndex        =   6
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
         Height          =   195
         Index           =   24
         Left            =   90
         TabIndex        =   0
         Top             =   135
         Width           =   630
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4050
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1425
      Width           =   18435
      _ExtentX        =   32517
      _ExtentY        =   7144
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3525
         Left            =   90
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   405
         Width           =   18225
         _ExtentX        =   32147
         _ExtentY        =   6218
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Residence Probable Contact"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   8
         Top             =   135
         Width           =   2445
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   4050
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   5505
      Width           =   18435
      _ExtentX        =   32517
      _ExtentY        =   7144
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3525
         Left            =   90
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   405
         Width           =   18225
         _ExtentX        =   32147
         _ExtentY        =   6218
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch / Office Probable Contact"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   12
         Top             =   135
         Width           =   2865
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   18750
      TabIndex        =   13
      Top             =   555
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
      Picture         =   "frmContactTrace.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   18750
      TabIndex        =   14
      Top             =   1815
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
      Picture         =   "frmContactTrace.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   18750
      TabIndex        =   15
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Export"
      AccessKey       =   "Export"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmContactTrace.frx":0EF4
   End
End
Attribute VB_Name = "frmContactTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeModuleName = "frmContactTrace"

Private oSkin As clsFormSkin

Private bLoaded As Boolean

Private sProvIDxx As String
Private sTownIDxx As String
Private sBrgyIDxx As String
Private sEmployID As String

Private oRS1 As Recordset
Private oRS2 As Recordset

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0
      Call ExportExcel
   Case 1
      Call loadRecord1
      Call loadRecord2
   Case 2
      Unload Me
   End Select
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

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance

   ClearFields
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub ClearFields()
   sProvIDxx = ""
   sTownIDxx = ""
   sBrgyIDxx = ""
   sEmployID = ""
   
   txtSearch(0) = ""
   txtSearch(1) = ""
   txtSearch(2) = ""
   txtSearch(3) = ""
   
   Set oRS1 = New Recordset
   Set oRS2 = New Recordset
   
   InitGrid
   InitGrid2
End Sub

Private Sub InitGrid2()
   Dim lnCtr As Integer
   With MSFlexGrid2
      .Cols = 11
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
      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "Employee"
      .TextMatrix(0, 3) = "Dept"
      .TextMatrix(0, 4) = "Address"
      .TextMatrix(0, 5) = "Barangay"
      .TextMatrix(0, 6) = "Town/City"
      .TextMatrix(0, 7) = "Province"
      .TextMatrix(0, 8) = "Contact No"
      .TextMatrix(0, 9) = "Last Log Branch"
      .TextMatrix(0, 10) = "Last Log"
      
      'column width
      .ColWidth(0) = 570
      .ColWidth(1) = 1900
      .ColWidth(2) = 2300
      .ColWidth(3) = 800
      .ColWidth(4) = 3000
      .ColWidth(5) = 1500
      .ColWidth(6) = 1500
      .ColWidth(7) = 1500
      .ColWidth(8) = 1300
      .ColWidth(9) = 1900
      .ColWidth(10) = 1900

      'column allinment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignLeftCenter
      .ColAlignment(5) = flexAlignLeftCenter
      .ColAlignment(6) = flexAlignLeftCenter
      .ColAlignment(7) = flexAlignLeftCenter
      .ColAlignment(8) = flexAlignLeftCenter
      .ColAlignment(9) = flexAlignLeftCenter
      .ColAlignment(10) = flexAlignLeftCenter
      
      'set location
      .Row = 1
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Cols = 9
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
      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "Employee"
      .TextMatrix(0, 3) = "Dept"
      .TextMatrix(0, 4) = "Address"
      .TextMatrix(0, 5) = "Barangay"
      .TextMatrix(0, 6) = "Town/City"
      .TextMatrix(0, 7) = "Province"
      .TextMatrix(0, 8) = "Contact No"
      
      'column width
      .ColWidth(0) = 590
      .ColWidth(1) = 3000
      .ColWidth(2) = 3000
      .ColWidth(3) = 900
      .ColWidth(4) = 3200
      .ColWidth(5) = 2000
      .ColWidth(6) = 2000
      .ColWidth(7) = 2000
      .ColWidth(8) = 1500

      'column allinment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignLeftCenter
      .ColAlignment(5) = flexAlignLeftCenter
      .ColAlignment(6) = flexAlignLeftCenter
      .ColAlignment(7) = flexAlignLeftCenter
      .ColAlignment(8) = flexAlignLeftCenter
      
      'set location
      .Row = 1
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Function getSQ_Master() As String
   getSQ_Master = "SELECT" & _
                     "  b.sCompnyNm" & _
                     ", IFNULL(TRIM(CONCAT(b.sHouseNox, ' ', b.sAddressx)), '') xAddressx" & _
                     ", IFNULL(e.sBrgyName, '') sBrgyName" & _
                     ", IFNULL(c.sTownName, '') sTownName" & _
                     ", IFNULL(d.sProvName, '') sProvName" & _
                     ", b.sMobileNo" & _
                     ", f.sBranchNm" & _
                     ", IFNULL(g.sDeptCode, '') sDeptCode" & _
                     ", a.sBranchCd" & _
                     ", a.sEmployID" & _
                  " FROM Employee_Master001 a" & _
                        " LEFT JOIN Branch f" & _
                           " ON a.sBranchCd = f.sBranchCd" & _
                        " LEFT JOIN Department g" & _
                           " ON a.sDeptIDxx = g.sDeptIDxx" & _
                     ", Client_Master b" & _
                     ", TownCity c" & _
                     ", Province d" & _
                     ", Barangay e" & _
                  " WHERE a.sEmployID = b.sClientID" & _
                     " AND b.sTownIDxx = c.sTownIDxx" & _
                     " AND c.sProvIDxx = d.sProvIDxx" & _
                     " AND b.sBrgyIDxx = e.sBrgyIDxx" & _
                     " AND a.cRecdStat = '1'"
End Function

Private Sub loadRecord1()
   Dim lsSQL As String
   
   lsSQL = getSQ_Master
   
   If sProvIDxx <> "" Then lsSQL = AddCondition(lsSQL, "d.sProvIDxx = " & strParm(sProvIDxx))
   If sTownIDxx <> "" Then lsSQL = AddCondition(lsSQL, "c.sTownIDxx = " & strParm(sTownIDxx))
   If sBrgyIDxx <> "" Then lsSQL = AddCondition(lsSQL, "e.sBrgyIDxx = " & strParm(sBrgyIDxx))
   If sEmployID <> "" Then lsSQL = AddCondition(lsSQL, "a.sEmployID = " & strParm(sEmployID))
   
   lsSQL = lsSQL & " ORDER BY d.sProvName, e.sBrgyName, c.sTownName, xAddressx, f.sBranchNm, b.sCompnyNm"
               
   With oRS1
      If .State = adStateOpen Then .Close
      
      .Open lsSQL, oApp.Connection, , , adCmdText
      Set .ActiveConnection = Nothing
      
      InitGrid
      
      Debug.Print lsSQL
      If .EOF Then
         MsgBox "No record found for the given criteria.", vbInformation, "Notice"
         Exit Sub
      End If
   End With
   
   With MSFlexGrid1
      Dim lnCtr As Integer
      
      .Rows = oRS1.RecordCount + 1
      lnCtr = 0
      
      ProgressBar1.Min = 0
      ProgressBar1.Max = oRS1.RecordCount
      ProgressBar1.Value = 0
      
      oRS1.MoveFirst
      Do Until oRS1.EOF
         DoEvents
      
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = oRS1("sBranchNm")
         .TextMatrix(lnCtr + 1, 2) = oRS1("sCompnyNm")
         .TextMatrix(lnCtr + 1, 3) = oRS1("sDeptCode")
         .TextMatrix(lnCtr + 1, 4) = oRS1("xAddressx")
         .TextMatrix(lnCtr + 1, 5) = oRS1("sBrgyName")
         .TextMatrix(lnCtr + 1, 6) = oRS1("sTownName")
         .TextMatrix(lnCtr + 1, 7) = oRS1("sProvName")
         .TextMatrix(lnCtr + 1, 8) = oRS1("sMobileNo")
         
         lnCtr = lnCtr + 1
         oRS1.MoveNext
         
         ProgressBar1.Value = ProgressBar1.Value + 1
      Loop
      
      ProgressBar1.Value = 0
   End With
End Sub

Private Sub loadRecord2()
   InitGrid2
   
   If oRS1.RecordCount = 0 Then Exit Sub
   
   Dim lnCtr As Integer
   Dim lsBranchCD As String
   
   With oRS1
      .MoveFirst
      
      Do Until .EOF
         If InStr(1, lsBranchCD, .Fields("sBranchCd")) = 0 Then
            lsBranchCD = lsBranchCD + .Fields("sBranchCd") + "»"
         End If
      
         .MoveNext
      Loop
      
      lsBranchCD = Mid(lsBranchCD, 1, Len(lsBranchCD) - 1)
      
      Dim lsCondition As String
      
      If lsBranchCD <> "" Then
         Dim lasSplit() As String
      
         lsCondition = "("
         
         lasSplit = Split(lsBranchCD, "»")
         
         For lnCtr = 0 To UBound(lasSplit)
            lsCondition = lsCondition + strParm(lasSplit(lnCtr)) + ","
         Next
         
         lsCondition = "a.sBranchCd IN " + Mid(lsCondition, 1, Len(lsCondition) - 1) + ")"
      End If
      
      Dim lsSQL As String
      
      lsSQL = AddCondition(getSQ_Master, lsCondition)
      lsSQL = lsSQL & " ORDER BY f.sBranchNm, g.sDeptCode, b.sCompnyNm, e.sBrgyName, c.sTownName, d.sProvName"
      
      With oRS2
         If .State = adStateOpen Then .Close
         
         Debug.Print lsSQL
         .Open lsSQL, oApp.Connection, , , adCmdText
         Set .ActiveConnection = Nothing
         
         InitGrid2
         
         If .EOF Then Exit Sub
      End With
      
      With MSFlexGrid2
         .Rows = oRS2.RecordCount + 1
         lnCtr = 0
         
         .ColWidth(4) = IIf(.Rows > 13, 2750, 3000)
         
         Dim lsBranchNm As String
         Dim lsLogTimex As String
         
         ProgressBar1.Min = 0
         ProgressBar1.Max = oRS2.RecordCount
         ProgressBar1.Value = 0
         
         oRS2.MoveFirst
         Do Until oRS2.EOF
            DoEvents
         
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = oRS2("sBranchNm")
            .TextMatrix(lnCtr + 1, 2) = oRS2("sCompnyNm")
            .TextMatrix(lnCtr + 1, 3) = oRS2("sDeptCode")
            .TextMatrix(lnCtr + 1, 4) = oRS2("xAddressx")
            .TextMatrix(lnCtr + 1, 5) = oRS2("sBrgyName")
            .TextMatrix(lnCtr + 1, 6) = oRS2("sTownName")
            .TextMatrix(lnCtr + 1, 7) = oRS2("sProvName")
            .TextMatrix(lnCtr + 1, 8) = oRS2("sMobileNo")
            
            Call getEmpLastLog(oRS2("sEmployID"), lsBranchNm, lsLogTimex)
            
            .TextMatrix(lnCtr + 1, 9) = lsBranchNm
            .TextMatrix(lnCtr + 1, 10) = lsLogTimex
            
            lnCtr = lnCtr + 1
            oRS2.MoveNext
            
            ProgressBar1.Value = ProgressBar1.Value + 1
         Loop
         
         ProgressBar1.Value = 0
      End With
   End With
End Sub

Private Sub ExportExcel()
   On Error Resume Next
   
   If oRS1.RecordCount = 0 Then
      MsgBox "No record to export.", vbInformation, "Notice"
      Exit Sub
   End If

   Dim ExlObj As Object
   
   Set ExlObj = CreateObject("excel.application")      ' Initialize the excel object
   ExlObj.Workbooks.Add                                ' Add an excel workbook
   
   ExlObj.Visible = True
   
   With ExlObj.ActiveSheet
     ' Print the heading and columns
      '.Cells(1, 1).Value = "Employee Contact Tracker (Residence)"
      '.Cells(1, 1).Font.Name = "Verdana"
      '.Cells(1, 1).Font.Bold = True:
      
      .Cells(1, 1).Value = "Branch":      .Cells(1, 2).Value = "Employee"
      .Cells(1, 3).Value = "Dept.":       .Cells(1, 4).Value = "Address"
      .Cells(1, 5).Value = "Barangay":    .Cells(1, 6).Value = "Town/City"
      .Cells(1, 7).Value = "Province":    .Cells(1, 8).Value = "Contact No."
   End With
   
   Dim lnCtr As Integer
   
   For lnCtr = 1 To 8
      ExlObj.ActiveSheet.Cells(1, lnCtr).Font.Bold = True
   Next
   
   lnCtr = 2
   
   With oRS1
      ProgressBar1.Max = .RecordCount
      ProgressBar1.Value = 0
   
      .MoveFirst
      Do Until .EOF
         DoEvents
      
         ExlObj.ActiveSheet.Cells(lnCtr, 1).Value = .Fields("sBranchNm")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 1).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 2).Value = .Fields("sCompnyNm")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 2).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 3).Value = .Fields("sDeptCode")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 3).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 4).Value = .Fields("xAddressx")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 4).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 5).Value = .Fields("sBrgyName")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 5).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 6).Value = .Fields("sTownName")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 6).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 7).Value = .Fields("sProvName")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 7).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
      
         ExlObj.ActiveSheet.Cells(lnCtr, 8).Value = .Fields("sMobileNo")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 8).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         .MoveNext
         lnCtr = lnCtr + 1
         
         ProgressBar1.Value = ProgressBar1.Value + 1
      Loop
      
      ProgressBar1.Value = 0
   End With
   
   If MsgBox("Do you want to export BRANCH/OFFICE PROBABLE CONTACT?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   
   If oRS2.RecordCount = 0 Then
      MsgBox "No record to export.", vbInformation, "Notice"
      Exit Sub
   End If
      
   ExlObj.Worksheets.Add
   ExlObj.Sheets("Sheet2").Activate

   With ExlObj.ActiveSheet
      ' Print the heading and columns
       '.Cells(1, 1).Value = "Employee Contact Tracker (Branch/Office)"
       '.Cells(1, 1).Font.Name = "Verdana"
       '.Cells(1, 1).Font.Bold = True:
       
       .Cells(1, 1).Value = "Branch":              .Cells(1, 2).Value = "Employee"
       .Cells(1, 3).Value = "Dept.":               .Cells(1, 4).Value = "Address"
       .Cells(1, 5).Value = "Barangay":            .Cells(1, 6).Value = "Town/City"
       .Cells(1, 7).Value = "Province":            .Cells(1, 8).Value = "Contact No."
       .Cells(1, 9).Value = "Last Log Branch":     .Cells(1, 10).Value = "Last Log"
   End With
       
   For lnCtr = 1 To 10
      ExlObj.ActiveSheet.Cells(1, lnCtr).Font.Bold = True
   Next
   
   lnCtr = 2
   
   Dim lsBranchNm As String
   Dim lsLogTimex  As String
   
   With oRS2
      ProgressBar1.Max = .RecordCount
      ProgressBar1.Value = 0
   
      .MoveFirst
      Do Until .EOF
      
         ExlObj.ActiveSheet.Cells(lnCtr, 1).Value = .Fields("sBranchNm")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 1).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 2).Value = .Fields("sCompnyNm")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 2).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 3).Value = .Fields("sDeptCode")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 3).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 4).Value = .Fields("xAddressx")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 4).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 5).Value = .Fields("sBrgyName")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 5).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 6).Value = .Fields("sTownName")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 6).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         ExlObj.ActiveSheet.Cells(lnCtr, 7).Value = .Fields("sProvName")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 7).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
      
         ExlObj.ActiveSheet.Cells(lnCtr, 8).Value = .Fields("sMobileNo")
         'ExlObj.ActiveCell.Worksheet.Cells(lnCtr, 8).AutoFormat xlRangeAutoFormatList1, 0, regular, 3, 1, 1
         
         Call getEmpLastLog(.Fields("sEmployID"), lsBranchNm, lsLogTimex)
            
         ExlObj.ActiveSheet.Cells(lnCtr, 9).Value = lsBranchNm
         ExlObj.ActiveSheet.Cells(lnCtr, 10).Value = lsLogTimex
         
         .MoveNext
         lnCtr = lnCtr + 1
         
         ProgressBar1.Value = ProgressBar1.Value + 1
      Loop
      
      ProgressBar1.Value = 0
   End With
End Sub

Private Sub getEmpLastLog(ByVal fsEmployID As String, ByRef fsBranchNm As String, ByRef fsLogTime As String)
   Dim lors As Recordset
   Dim lsSQL As String
   
   If fsEmployID = "" Then Exit Sub
   
   lsSQL = "SELECT b.sBranchNm, CONCAT(a.dTransact, ' ', IFNULL(a.dOTimeOut, IFNULL(a.dOTimeOut, IFNULL(a.dPMOutxxx, IFNULL(a.dPMInxxxx, IFNULL(a.dAMOutxxx, IFNULL(a.dAMInxxxx, '00:00:00'))))))) xTime" & _
            " FROM (SELECT * FROM Employee_Log WHERE sEmployID = 'xEmployID'" & _
                     " UNION SELECT * FROM Employee_Monitoring WHERE sEmployID = 'xEmployID'" & _
                     " ORDER BY dTransact DESC LIMIT 1) a" & _
               ", Branch b" & _
            " WHERE a.sBranchCd = b.sBranchCd"
            
  lsSQL = Replace(lsSQL, "xEmployID", fsEmployID)
  
  Set lors = New Recordset
  
  lors.Open lsSQL, oApp.Connection, , adCmdText
  Set lors.ActiveConnection = Nothing
  
  If lors.RecordCount = 0 Then GoTo endProc
  
  fsBranchNm = lors("sBranchNm")
  fsLogTime = lors("xTime")
  
endProc:
  Set lors = Nothing
End Sub

Private Function getProvince(ByVal fsValue As String) As String
   Dim lsSQL As String
   Dim lors As Recordset
   
   getProvince = ""
   
   fsValue = Trim(fsValue)
   
   If fsValue = "" Then
      sProvIDxx = ""
      Exit Function
   End If
   
   lsSQL = "SELECT sProvIDxx, sProvName FROM Province WHERE cRecdStat = '1' AND sProvName LIKE " & strParm(fsValue & "%")
   
   Set lors = New Recordset
   
   With lors
      lors.Open lsSQL, oApp.Connection, , , adCmdText
      Set lors.ActiveConnection = Nothing
      
      If lors.RecordCount = 0 Then
         sProvIDxx = ""
      ElseIf lors.RecordCount = 1 Then
         sProvIDxx = .Fields("sProvIDxx")
         getProvince = .Fields("sProvName")
      Else
         lsSQL = KwikBrowse(oApp, lors, "sProvIDxx»sProvName", "ID»Description", "@»@")
         
         If lsSQL = "" Then
            sProvIDxx = ""
         Else
            Dim lasSplit() As String
            
            lasSplit = Split(lsSQL, "»")
            
            sProvIDxx = lasSplit(0)
            getProvince = lasSplit(1)
         End If
      End If
   End With

   Set lors = Nothing
End Function

Private Function getEmployee(ByVal fsValue As String) As String
   Dim lsSQL As String
   Dim lors As Recordset
   
   getEmployee = ""
   
   fsValue = Trim(fsValue)
   
   If fsValue = "" Then
      sEmployID = ""
      Exit Function
   End If
   
   lsSQL = "SELECT" & _
               "  a.sEmployID" & _
               ", b.sCompnyNm" & _
               ", c.sBranchNm" & _
            " FROM Employee_Master001 a" & _
                  " LEFT JOIN Branch c ON a.sBranchCd = c.sBranchCd" & _
               ", Client_Master b" & _
            " WHERE a.sEmployID = b.sClientID" & _
               " AND a.cRecdStat = '1'" & _
               " AND b.sCompnyNm LIKE " & strParm(fsValue & "%")
   
   Set lors = New Recordset
   
   With lors
      lors.Open lsSQL, oApp.Connection, , , adCmdText
      Set lors.ActiveConnection = Nothing
      
      If lors.RecordCount = 0 Then
         sEmployID = ""
      ElseIf lors.RecordCount = 1 Then
         sEmployID = .Fields("sEmployID")
         getEmployee = .Fields("sCompnyNm")
      Else
         lsSQL = KwikBrowse(oApp, lors, "sEmployID»sCompnyNm»sBranchNm", "ID»Name»Branch", "@»@»@")
         
         If lsSQL = "" Then
            sEmployID = ""
         Else
            Dim lasSplit() As String
            
            lasSplit = Split(lsSQL, "»")
            
            sEmployID = lasSplit(0)
            getEmployee = lasSplit(1)
         End If
      End If
   End With

   Set lors = Nothing
End Function

Private Function getTown(ByVal fsValue As String) As String
   Dim lsSQL As String
   Dim lors As Recordset
   
   getTown = ""
   
   fsValue = Trim(fsValue)
   
   If fsValue = "" Then
      sTownIDxx = ""
      Exit Function
   End If
   
   lsSQL = "SELECT" & _
               "  a.sTownIDxx" & _
               ", a.sTownName" & _
               ", b.sProvName" & _
            " FROM TownCity a" & _
               ", Province b" & _
            " WHERE a.sProvIDxx = b.sProvIDxx" & _
               " AND a.cRecdStat = '1'" & _
               " AND a.sTownName LIKE " & strParm(fsValue & "%")
               
   If sProvIDxx <> "" Then lsSQL = AddCondition(lsSQL, "a.sProvIDxx = " & strParm(sProvIDxx))
   
   Set lors = New Recordset
   
   With lors
      lors.Open lsSQL, oApp.Connection, , , adCmdText
      Set lors.ActiveConnection = Nothing
      
      If lors.RecordCount = 0 Then
         sTownIDxx = ""
      ElseIf lors.RecordCount = 1 Then
         sTownIDxx = .Fields("sTownIDxx")
         getTown = .Fields("sTownName")
      Else
         lsSQL = KwikBrowse(oApp, lors, "sTownIDxx»sTownName»sProvName", "ID»Description»Province", "@»@»@")
         
         If lsSQL = "" Then
            sTownIDxx = ""
         Else
            Dim lasSplit() As String
            
            lasSplit = Split(lsSQL, "»")
            
            sTownIDxx = lasSplit(0)
            getTown = lasSplit(1)
         End If
      End If
   End With

   Set lors = Nothing
End Function

Private Function getBarangay(ByVal fsValue As String) As String
   Dim lsSQL As String
   Dim lors As Recordset
   
   getBarangay = ""
   
   fsValue = Trim(fsValue)
   
   If fsValue = "" Then
      sBrgyIDxx = ""
      Exit Function
   End If
   
   lsSQL = "SELECT" & _
               "  c.sBrgyIDxx" & _
               ", c.sBrgyName" & _
               ", a.sTownName" & _
               ", b.sProvName" & _
            " FROM TownCity a" & _
               ", Province b" & _
               ", Barangay c" & _
            " WHERE a.sProvIDxx = b.sProvIDxx" & _
               " AND c.sTownIDxx = a.sTownIDxx" & _
               " AND c.sBrgyName LIKE " & strParm(fsValue & "%")
               
   If sProvIDxx <> "" Then lsSQL = AddCondition(lsSQL, "a.sProvIDxx = " & strParm(sProvIDxx))
   If sTownIDxx <> "" Then lsSQL = AddCondition(lsSQL, "c.sTownIDxx = " & strParm(sTownIDxx))
   
   Set lors = New Recordset
   
   With lors
      lors.Open lsSQL, oApp.Connection, , , adCmdText
      Set lors.ActiveConnection = Nothing
      
      If lors.RecordCount = 0 Then
         sBrgyIDxx = ""
      ElseIf lors.RecordCount = 1 Then
         sBrgyIDxx = .Fields("sBrgyIDxx")
         getBarangay = .Fields("sBrgyName")
      Else
         lsSQL = KwikBrowse(oApp, lors, "sBrgyIDxx»sBrgyName»sTownName»sProvName", "ID»Description»Town/City»Province", "@»@»@»@")
         
         If lsSQL = "" Then
            sBrgyIDxx = ""
         Else
            Dim lasSplit() As String
            
            lasSplit = Split(lsSQL, "»")
            
            sBrgyIDxx = lasSplit(0)
            getBarangay = lasSplit(1)
         End If
      End If
   End With

   Set lors = Nothing
End Function

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

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub MSFlexGrid2_Click()
   With MSFlexGrid2
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 0
         txtSearch(Index) = getProvince(txtSearch(Index))
      Case 1
         txtSearch(Index) = getEmployee(txtSearch(Index))
      Case 2
         txtSearch(Index) = getTown(txtSearch(Index))
      Case 3
         txtSearch(Index) = getBarangay(txtSearch(Index))
      End Select
   End If
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub
