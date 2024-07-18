VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmChinaCMS 
   BorderStyle     =   0  'None
   Caption         =   "China Bank CMS"
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5205
      Left            =   105
      TabIndex        =   10
      Top             =   3015
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9181
      _Version        =   393216
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2415
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4260
      BackColor       =   12632256
      ClipControls    =   0   'False
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
         Height          =   420
         Index           =   8
         Left            =   6990
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "Dec. 10, 2012"
         Top             =   1830
         Visible         =   0   'False
         Width           =   2070
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
         Height          =   420
         Index           =   7
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "NISSAN PANGASINAN"
         Top             =   1365
         Width           =   7380
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
         Height          =   420
         Index           =   6
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "1790873119"
         Top             =   915
         Width           =   2070
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
         Height          =   420
         Index           =   1
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1830
         Width           =   2070
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
         Height          =   420
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2070
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activation Date"
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
         Left            =   5430
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit Description"
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
         Index           =   6
         Left            =   105
         TabIndex        =   4
         Top             =   1455
         Width           =   1485
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit Account#"
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
         Left            =   105
         TabIndex        =   2
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   105
         TabIndex        =   6
         Top             =   1920
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll ID"
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
         Top             =   210
         Width           =   930
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1725
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   2070
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9615
      TabIndex        =   17
      Top             =   1860
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
      Picture         =   "frmChinaCMS.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9615
      TabIndex        =   15
      Top             =   570
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Export"
      AccessKey       =   "E"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmChinaCMS.frx":077A
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   645
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   8250
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1138
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   11
         Left            =   7110
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   105
         Width           =   2025
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   10
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   105
         Width           =   1335
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Left            =   5850
         TabIndex        =   13
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Items"
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
         Index           =   14
         Left            =   3165
         TabIndex        =   11
         Top             =   210
         Width           =   960
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   9615
      TabIndex        =   16
      Top             =   1215
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
      Picture         =   "frmChinaCMS.frx":0EF4
   End
End
Attribute VB_Name = "frmChinaCMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmChinaCMS"

Private oSkin As clsFormSkin
Private bOk As Boolean
Private poReport As clsReport
Private poRecord As Recordset

Property Set Data(Value As Recordset)
   Set poRecord = Value
End Property

Property Get IsOkey() As Boolean
   IsOkey = bOk
End Property

Private Sub cmdButton_Click(Index As Integer)
   
   Me.Hide
   Select Case Index
   Case 0
      bOk = True
   Case 1
      bOk = False
   Case 5
      bOk = False
      Call ReportTrans
   End Select
End Sub

Private Sub Form_Activate()
   LoadMaster
   LoadDetail
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
'   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me
      
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   
   Set poReport = New clsReport
   
   InitGrid

   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set poReport = Nothing
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
   Case 1
      If IsDate(txtField(Index)) Then
         txtField(Index) = Format(txtField(Index), "MM/DD/YYYY")
      End If
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
   
   Select Case Index
   Case 1
      If IsDate(txtField(Index)) Then
         txtField(Index) = Format(txtField(Index), "Mmm. DD, YYYY")
      End If
   End Select
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Rows = 2
      .Cols = 4
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
      .TextMatrix(0, 1) = "Bank Account "
      .TextMatrix(0, 2) = "Employee"
      .TextMatrix(0, 3) = "Net Pay"
      
      .ColWidth(0) = 800
      .ColWidth(1) = 2200
      .ColWidth(2) = 4200
      .ColWidth(3) = 1650
      
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

Private Sub LoadMaster()
   Dim loTxt As TextBox
   
'   For Each loTxt In txtField
'      loTxt = IFNull(oTrans.Master(loTxt.Index))
'   Next
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   Dim lnAmt As Currency
   
   With MSFlexGrid1
      
      poRecord.Filter = "sBankIDxx = " & strParm("00XX003") & " AND sBnkActNo <> ''"
      
      If poRecord.RecordCount = 0 Then
         .Rows = 1
         .Rows = 2
         Exit Sub
      End If
      
      poRecord.Sort = "sEmployNm, sBnkActNo"
      poRecord.MoveFirst
      
      .Rows = poRecord.RecordCount + 1
      
      For lnCtr = 0 To poRecord.RecordCount - 1
         .RowHeight(lnCtr + 1) = 338
         .TextMatrix(lnCtr + 1, 0) = lnCtr
         .TextMatrix(lnCtr + 1, 1) = poRecord("sEmployNm")
         .TextMatrix(lnCtr + 1, 2) = poRecord("sBnkActNo")
         .TextMatrix(lnCtr + 1, 3) = Format(poRecord("nNetPayxx"), "#,##0.00")
         
         lnAmt = lnAmt + poRecord("nNetPayxx")
         poRecord.MoveNext
      Next
      
      txtField(10) = poRecord.RecordCount
      txtField(11) = Format(lnAmt, "#,##0.00")
      
   End With
   
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   oTrans.Master(Index) = txtField(Index)
End Sub

Private Function ReportTrans() As Boolean
   Dim lrs As ADODB.Recordset
   Dim lnCtr As Integer
   Dim lnTmp As Integer
   Dim lsOldProc As String
   Dim p_oRepApp As CRAXDRT.Application
   Dim p_oReport As CRAXDRT.Report
   Dim lsRemarks As String
   
   lsOldProc = "ReportTrans"
   ''On Error GoTo errProc
   
   ReportTrans = False
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 4
   lrs.Fields.Append "sField01", adVarChar, 100
   lrs.Fields.Append "sField02", adVarChar, 100
   lrs.Fields.Append "lField01", adCurrency, 7
   lrs.Open
   With poRecord
      .Filter = "sBankIDxx = " & strParm("00XX003") & " AND sBnkActNo <> ''"
      
      If poRecord.RecordCount = 0 Then
         MsgBox "There are no data to be printed at this time!", vbInformation, "Warning"
         Exit Function
      End If
      
      .Sort = "sEmployNm, sBnkActNo"
      .MoveFirst
      Do Until poRecord.EOF
         lrs.AddNew
         lnCtr = lnCtr + 1
         lrs("nField01") = lnCtr
         lrs("sField01") = UCase(.Fields("sEmployNm"))
         lrs("sField02") = .Fields("sBnkActNo")
         lrs("lField01") = .Fields("nNetPayxx")
         poRecord.MoveNext
      Loop
   End With
   
   'Initialize Report Class
   poReport.InitReport

   Set p_oRepApp = New Application
   Set p_oReport = p_oRepApp.OpenReport(oApp.AppPath & "\Reports\ChinaBankCMS.rpt")
   Set p_oRepApp = Nothing
   
   p_oReport.Sections("RH").ReportObjects("txtBank").SetText "CHINA BANKING CORPORATION"
   p_oReport.Sections("RH").ReportObjects("txtCompany").SetText txtField(7)
   p_oReport.Sections("RH").ReportObjects("txtDate").SetText Format(txtField(1), "MMMM DD, YYYY")
      
   Debug.Print NumToText(CDbl(txtField(11)))
      
   lsRemarks = "As per our teller card payroll agreement, please debit our S/A #" & txtField(6) & _
              " amounting to " & NumToText(CDbl(txtField(11))) & " Only (" & Format(CDbl(txtField(11)), "#,##0.00") & ")" & _
              " and credit to the account of the following:"
      
   p_oReport.Sections("RH").ReportObjects("txtRemarks").SetText lsRemarks
   
   p_oReport.DiscardSavedData
   p_oReport.FieldMappingType = crAutoFieldMapping
   p_oReport.Database.SetDataSource lrs
   
   Set poReport.ReportObject = p_oReport
   poReport.ReportID = "X"
   
   poReport.ShowReport
   
   ReportTrans = True
   
endProc:
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

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

