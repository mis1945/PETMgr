VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmPayrollLoans 
   BorderStyle     =   0  'None
   Caption         =   "Employee Loans"
   ClientHeight    =   3105
   ClientLeft      =   105
   ClientTop       =   0
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3105
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2490
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   4392
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2400
         Left            =   30
         TabIndex        =   0
         Top             =   30
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4233
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   7380
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
      Picture         =   "frmPayrollLoans.frx":0000
   End
End
Attribute VB_Name = "frmPayrollLoans"
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

Dim pbCtrlPress As Boolean
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
Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      lnRow = p_oRS.RecordCount
   
      .Rows = lnRow + 1
      
      lnCtr = 0
      p_oRS.MoveFirst
      Do Until p_oRS.EOF
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = p_oRS("sReferNox")
         .TextMatrix(lnCtr + 1, 2) = Format(p_oRS("dTransact"), "Mmm dd, yyyy")
         .TextMatrix(lnCtr + 1, 3) = Format(p_oRS("nPNValuex"), "###,###.00")
         .TextMatrix(lnCtr + 1, 4) = Format(p_oRS("nABalance"), "###,###.00")
         .TextMatrix(lnCtr + 1, 5) = Format(p_oRS("nMonAmort"), "###,###.00")
                 
         lnCtr = lnCtr + 1
         p_oRS.MoveNext
      Loop
      
      .Row = 1
      pnRow = pnActiveRow - 1
      
   End With
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Cols = 6
      .Rows = 1

      .MergeCells = flexMergeFree
      
      .Clear
      
      .Row = 0
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .TextMatrix(0, 0) = "NO."
      .TextMatrix(0, 1) = "REFER NO."
      .TextMatrix(0, 2) = "DATE "
      .TextMatrix(0, 3) = "PN VALUE"
      .TextMatrix(0, 4) = "BALANCE"
      .TextMatrix(0, 5) = "AMORTZTN"
      .RowHeightMin = 300
      
      'column width
      .ColWidth(0) = 400
      .ColWidth(1) = 1300
      .ColWidth(2) = 1300
      .ColWidth(3) = 1300
      .ColWidth(4) = 1300
      .ColWidth(5) = 1300
      
      'column alignment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   pnActiveRow = 1
   pbFormLoad = False
   pbCtrlPress = False
End Sub

