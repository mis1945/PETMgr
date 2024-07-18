VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmEmployeeTimesheet 
   BorderStyle     =   0  'None
   Caption         =   "Employee Timesheet"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   16950
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5235
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   9234
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5145
         Left            =   30
         TabIndex        =   0
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   9075
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   15615
      TabIndex        =   1
      Top             =   510
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
      Picture         =   "frmEmployeeTimesheet.frx":0000
   End
End
Attribute VB_Name = "frmEmployeeTimesheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim pnRow As Integer
Dim pnActiveRow As Integer

Dim pbFormLoad As Boolean
Dim pbDetailGotFocus As Boolean
Dim pbInUpdateMode As Boolean

Dim pdDate As Date

Dim psSelectedTime As String
Dim psTimeTempStrg As String
Dim pnLastSelc As Integer

Private p_oRS As Recordset

Property Set Data(ByVal foValue As Recordset)
   Set p_oRS = foValue
End Property

Private Sub cmdButton_Click(Index As Integer)
 Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   pnActiveRow = 1
   pbFormLoad = False
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      lnRow = p_oRS.RecordCount
   
      .Rows = lnRow + 2
      
      lnCtr = 0
      p_oRS.MoveFirst
      Do Until p_oRS.EOF
         .TextMatrix(lnCtr + 2, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 2, 1) = Format(p_oRS("dTransact"), "Mmm dd, yyyy")
         .TextMatrix(lnCtr + 2, 2) = IIf(IFNull(p_oRS("dAMInxxxx"), "") = "", "", Format(p_oRS("dAMInxxxx"), "hh:mm AM/PM"))
         .TextMatrix(lnCtr + 2, 3) = IIf(IFNull(p_oRS("dAMOutxxx"), "") = "", "", Format(p_oRS("dAMOutxxx"), "hh:mm AM/PM"))
         .TextMatrix(lnCtr + 2, 4) = IIf(IFNull(p_oRS("dPMInxxxx"), "") = "", "", Format(p_oRS("dPMInxxxx"), "hh:mm AM/PM"))
         .TextMatrix(lnCtr + 2, 5) = IIf(IFNull(p_oRS("dPMOutxxx"), "") = "", "", Format(p_oRS("dPMOutxxx"), "hh:mm AM/PM"))
         .TextMatrix(lnCtr + 2, 6) = IIf(IFNull(p_oRS("dOTimeInx"), "") = "", "", Format(p_oRS("dOTimeInx"), "hh:mm AM/PM"))
         .TextMatrix(lnCtr + 2, 7) = IIf(IFNull(p_oRS("dOTimeOut"), "") = "", "", Format(p_oRS("dOTimeOut"), "hh:mm AM/PM"))
         .TextMatrix(lnCtr + 2, 8) = IIf(p_oRS("cAbsentxx") = "1", "YES", "NO")
         .TextMatrix(lnCtr + 2, 9) = IIf(p_oRS("cLeavexxx") = "1", "YES", "NO")
         .TextMatrix(lnCtr + 2, 10) = IIf(p_oRS("cHolidayx") = "1", "YES", "NO")
         .TextMatrix(lnCtr + 2, 11) = IIf(p_oRS("cDeductxx") = "1", "YES", "NO")
         .TextMatrix(lnCtr + 2, 12) = IIf(p_oRS("cRestDayx") = "1", "YES", "NO")
         .TextMatrix(lnCtr + 2, 13) = p_oRS("nTardyxxx")
         .TextMatrix(lnCtr + 2, 14) = p_oRS("nOverTime")
         .TextMatrix(lnCtr + 2, 15) = p_oRS("nUndrTime")
         
         lnCtr = lnCtr + 1
         p_oRS.MoveNext
      Loop
      
      pnRow = 0
      .Row = 2
      
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc
   
   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   
   Call InitGrid
   Call LoadDetail
endProc:
   Exit Sub
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Cols = 16
      .Rows = 2
      .RowHeightMin = 300

      .MergeCells = flexMergeFree
      
      .Clear
      
      .Row = 0
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = ""
      .TextMatrix(0, 2) = "1ST QUARTER"
      .TextMatrix(0, 3) = "1ST QUARTER"
      .TextMatrix(0, 4) = "2ND QUARTER"
      .TextMatrix(0, 5) = "2ND QUARTER"
      .TextMatrix(0, 6) = "OVERTIME"
      .TextMatrix(0, 7) = "OVERTIME"
      .TextMatrix(0, 8) = ""
      .TextMatrix(0, 9) = ""
      .TextMatrix(0, 10) = ""
      .TextMatrix(0, 11) = ""
      .TextMatrix(0, 12) = ""
      .TextMatrix(0, 13) = ""
      .TextMatrix(0, 14) = ""
      .TextMatrix(0, 15) = ""
      
      .Row = 1
            
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      'column title
      .TextMatrix(1, 0) = "NO."
      .TextMatrix(1, 1) = "DATE"
      .TextMatrix(1, 2) = "IN"
      .TextMatrix(1, 3) = "OUT"
      .TextMatrix(1, 4) = "IN"
      .TextMatrix(1, 5) = "OUT"
      .TextMatrix(1, 6) = "IN"
      .TextMatrix(1, 7) = "OUT"
      .TextMatrix(1, 8) = "ABSENT"
      .TextMatrix(1, 9) = "LEAVE"
      .TextMatrix(1, 10) = "HOLIDAY"
      .TextMatrix(1, 11) = "DEDUCT"
      .TextMatrix(1, 12) = "RESTDAY"
      .TextMatrix(1, 13) = "TARDI"
      .TextMatrix(1, 14) = "OTIME"
      .TextMatrix(1, 15) = "UTIME"
      
      .MergeRow(0) = True
      .MergeRow(1) = True
      
      .ColWidth(0) = 400
      .ColWidth(1) = 1300
      
      'column alignment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignCenterCenter
      .ColAlignment(2) = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignCenterCenter
      .ColAlignment(4) = flexAlignCenterCenter
      .ColAlignment(5) = flexAlignCenterCenter
      .ColAlignment(6) = flexAlignCenterCenter
      .ColAlignment(7) = flexAlignCenterCenter
      .ColAlignment(8) = flexAlignCenterCenter
      .ColAlignment(9) = flexAlignCenterCenter
      .ColAlignment(10) = flexAlignCenterCenter
      .ColAlignment(11) = flexAlignCenterCenter
      .ColAlignment(12) = flexAlignCenterCenter
      .ColAlignment(13) = flexAlignCenterCenter
      .ColAlignment(14) = flexAlignCenterCenter
      .ColAlignment(15) = flexAlignCenterCenter
      
      .Rows = 3
      .TextMatrix(2, 0) = "1"
      .Row = 2
      
      pnLastSelc = .Row
      pnRow = 2
   End With
End Sub

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub
