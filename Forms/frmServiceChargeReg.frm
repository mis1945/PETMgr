VERSION 5.00
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmServiceChargeReg 
   BorderStyle     =   0  'None
   Caption         =   "Service Charge Approval"
   ClientHeight    =   8250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2775
      Index           =   0
      Left            =   1665
      Tag             =   "wt0;fb0"
      Top             =   1155
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   4895
      BackColor       =   12632256
      Begin VB.OptionButton optTaxable 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Non Taxable"
         Height          =   330
         Index           =   1
         Left            =   1395
         TabIndex        =   25
         Tag             =   "wt0;fb0"
         Top             =   990
         Width           =   1230
      End
      Begin VB.OptionButton optTaxable 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Taxable"
         Height          =   330
         Index           =   0
         Left            =   1395
         TabIndex        =   24
         Tag             =   "wt0;fb0"
         Top             =   675
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
         TabIndex        =   9
         Top             =   1605
         Width           =   1965
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   660
         Index           =   5
         Left            =   1290
         TabIndex        =   2
         Top             =   2070
         Width           =   8685
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   1290
         TabIndex        =   1
         Top             =   1710
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1290
         TabIndex        =   0
         Top             =   1380
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
         Left            =   780
         TabIndex        =   23
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incentive"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   22
         Top             =   705
         Width           =   660
      End
      Begin VB.Shape Shape2 
         Height          =   720
         Left            =   1290
         Top             =   615
         Width           =   1920
      End
      Begin VB.Shape Shape3 
         Height          =   345
         Index           =   0
         Left            =   6660
         Top             =   375
         Width           =   2415
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
         Height          =   285
         Left            =   6690
         TabIndex        =   17
         Tag             =   "eb0;et0"
         Top             =   405
         Width           =   2355
      End
      Begin VB.Shape Shape4 
         Height          =   405
         Index           =   0
         Left            =   6630
         Top             =   345
         Width           =   2475
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
         TabIndex        =   8
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   3
         Left            =   510
         TabIndex        =   7
         Top             =   2085
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refer. Date"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   6
         Top             =   1770
         Width           =   825
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. Date"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   1440
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
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4140
      Left            =   1665
      TabIndex        =   10
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   3975
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
      MOUSEICON       =   "frmServiceChargeReg.frx":0000
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
      Index           =   0
      Left            =   105
      TabIndex        =   11
      Top             =   3750
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
      Picture         =   "frmServiceChargeReg.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   105
      TabIndex        =   12
      Top             =   555
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
      Picture         =   "frmServiceChargeReg.frx":0796
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   630
      Index           =   1
      Left            =   1665
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   1111
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
         Left            =   7170
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   90
         Width           =   2355
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
         Left            =   1980
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
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
         Index           =   15
         Left            =   5385
         TabIndex        =   16
         Top             =   165
         Width           =   1230
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
         Index           =   14
         Left            =   465
         TabIndex        =   15
         Top             =   195
         Width           =   1365
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   105
      TabIndex        =   18
      Top             =   1200
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
      Picture         =   "frmServiceChargeReg.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   105
      TabIndex        =   19
      Top             =   1845
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
      Picture         =   "frmServiceChargeReg.frx":17EA
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   105
      TabIndex        =   20
      Top             =   2490
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
      Picture         =   "frmServiceChargeReg.frx":20C4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   105
      TabIndex        =   21
      Top             =   3120
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
      Picture         =   "frmServiceChargeReg.frx":299E
   End
End
Attribute VB_Name = "frmServiceChargeReg"
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
Private poMiscx As clsPayMisc

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   
   txtField_LostFocus pnIndex
   With GridEditor1
      Select Case Index
      Case 0 'Close
         Unload Me
      Case 1 'Browse
         If oTrans.SearchTransaction(oApp.BranchCode, False) = True Then
            Call LoadData
         End If
      Case 2 'BDO
         If oTrans.Master("cTranStat") <> xeStatePosted Then Exit Sub
         Call ExportBDO
         MsgBox "Data Exported Successfully!!"
      Case 3 'MB
         If oTrans.Master("cTranStat") <> xeStatePosted Then Exit Sub
         Call ExportMB
         MsgBox "Data Exported Successfully!!"
      Case 4 'SB
         If oTrans.Master("cTranStat") <> xeStatePosted Then Exit Sub
         Call ExportSB
         MsgBox "Data Exported Successfully!!"
      Case 5 'no account
         If oTrans.Master("cTranStat") <> xeStatePosted Then Exit Sub
         Call ExportNoAccount
         MsgBox "Data Exported Successfully!!"
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
   
   Set poMiscx = New clsPayMisc
   Set poMiscx.AppDriver = oApp
      
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   InitForm
   ClearFields
   initButton xeModeAddNew

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
      .ColEnabled(5) = False
      
      .Row = 1

   End With
   
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   txtField(0).Text = ""
   txtField(1).Text = ""
   txtField(2).Text = ""
   txtField(3).Text = ""
   txtField(5).Text = ""
   
   txtSearch(0).Text = ""
   txtSearch(1).Text = ""
      
   
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

Private Sub LoadData()
   Dim lnCtr As Integer
   
   txtField(0).Text = oTrans.Master(0)
   txtField(1).Text = oTrans.Master(1)
   txtField(2).Text = oTrans.Master(2)
   txtField(3).Text = oTrans.Master(3)
   txtField(5).Text = oTrans.Master(5)
   
   txtSearch(0).Text = oTrans.Master(0)
   txtSearch(1).Text = oTrans.Master(2)
   
   If oTrans.Master("sBeneftID") = "11009" Then 'non taxable
        optTaxable(1).Value = Checked
        optTaxable(0).Value = Unchecked
   Else
        optTaxable(0).Value = Checked
        optTaxable(1).Value = Unchecked
   End If
      
   With GridEditor1
      .Rows = oTrans.ItemCount + 1
      
      For lnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, 2)
         .TextMatrix(lnCtr + 1, 2) = oTrans.Detail(lnCtr, 3)
         .TextMatrix(lnCtr + 1, 3) = oTrans.Detail(lnCtr, 4)
         .TextMatrix(lnCtr + 1, 4) = oTrans.Detail(lnCtr, 5)
         .TextMatrix(lnCtr + 1, 5) = Format(oTrans.Detail(lnCtr, 6), "#,##0.00")
      Next
   End With
   
   If oTrans.Master("cTranStat") = "4" Then
      Label2.Caption = "APPLIED"
   Else
      Label2.Caption = TransStat(CInt(oTrans.Master("cTranStat")))
   End If
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

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      If Index = 0 Then
         oTrans.SearchTransaction oApp.BranchCode, True
      Else
         oTrans.SearchTransaction txtSearch(Index), False
      End If
   End If
End Sub

Private Function ExportBDO() As Boolean
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
   
   Dim loRecord As Recordset
      
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
    
    lsSQL = "SELECT " & _
               " a.sTransNox" & _
               ", e.sCompnyNm" & _
               ", c.sBnkActNo" & _
               ", c.sBankIDxx" & _
               ", b.nAmountxx" & _
            " FROM Employee_SCD_Master a" & _
            ", Employee_SCD_Detail b" & _
            ", Employee_Master001 c" & _
                  " LEFT JOIN Banks d ON c.sBankIDxx = d.sBankIDxx" & _
                  " LEFT JOIN Client_Master e ON c.sEmployID = e.sClientID" & _
            " WHERE a.sTransNox = b.sTransNox" & _
            " AND b.sEmployID = c.sEmployID" & _
            " AND a.cTranStat = '2'" & _
            " AND a.sTransNox = " & strParm(txtField(0).Text) & _
            " AND c.sBankIDxx = '00XX024'" & _
            " AND CAST(REVERSE(DECODE(UNHEX(b.nAmountxx), 'PETMGR')) AS CHAR) > 0"
      Debug.Print lsSQL
      Set loRecord = New Recordset
      loRecord.Open lsSQL, oApp.Connection, , , adCmdText
      
      If Not loRecord.EOF Then
         For lnItemCtr = 0 To loRecord.RecordCount - 1
            lnLineCtr = lnLineCtr + 1
            xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = "'" & loRecord("sBnkActNo")
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = poMiscx.Paway(loRecord("nAmountxx"))
            xlsheet.Range(XEmployNm & (lnLineCtr)).Value = loRecord("sCompnyNm")
            xlsheet.Range(XRemarksx & (lnLineCtr)).Value = ""
            loRecord.MoveNext
         Next
      End If

   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & txtField(0) & ").XLS"
   xl.ActiveWorkbook.Close
   xl.Quit
    
   Set xlwbook = Nothing
   Set xl = Nothing
End Function

Private Function ExportMB() As Boolean
   Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook
   
   Dim lsSQL As String
   Dim loRecord As Recordset
   
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
    
   lsSQL = "SELECT " & _
               " a.sTransNox" & _
               ", e.sCompnyNm" & _
               ", c.sBnkActNo" & _
               ", c.sBankIDxx" & _
               ", b.nAmountxx" & _
            " FROM Employee_SCD_Master a" & _
            ", Employee_SCD_Detail b" & _
            ", Employee_Master001 c" & _
                  " LEFT JOIN Banks d ON c.sBankIDxx = d.sBankIDxx" & _
                  " LEFT JOIN Client_Master e ON c.sEmployID = e.sClientID" & _
            " WHERE a.sTransNox = b.sTransNox" & _
            " AND b.sEmployID = c.sEmployID" & _
            " AND a.cTranStat = '2'" & _
            " AND a.sTransNox = " & strParm(txtField(0).Text) & _
            " AND c.sBankIDxx = '00XX006'" & _
            " AND CAST(REVERSE(DECODE(UNHEX(b.nAmountxx), 'PETMGR')) AS CHAR) > 0"
      Debug.Print lsSQL
      Set loRecord = New Recordset
      loRecord.Open lsSQL, oApp.Connection, , , adCmdText
   
      If Not loRecord.EOF Then
         For lnItemCtr = 0 To loRecord.RecordCount - 1
            lnLineCtr = lnLineCtr + 1
            xlsheet.Range(XEmployID & (lnLineCtr)).Value = lnLineCtr - 6
            xlsheet.Range(XEmployNm & (lnLineCtr)).Value = loRecord("sCompnyNm")
            xlsheet.Range(XBranchCD & (lnLineCtr)).Value = Left(loRecord("sBnkActNo"), 3)
            xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = Mid(loRecord("sBnkActNo"), 4)
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = poMiscx.Paway(loRecord("nAmountxx"))
            loRecord.MoveNext
         Next
      End If
     
     'kalyptus-2013.08.19 02:03pm
     'delete remaining items in our template if ever template was updated...
     lnItemCtr = xlsheet.Range("E" & xlsheet.Rows.Count).End(xlUp).Row
     If lnItemCtr > lnLineCtr Then
         xlsheet.Range("A" & lnLineCtr + 1 & ":E" & lnItemCtr).Clear
     End If
   
    
   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & txtField(0) & ").XLS"
   xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
   xl.Quit
    
   Set xlwbook = Nothing
   Set xl = Nothing
End Function

Private Function ExportSB() As Boolean
   Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook
   
   Dim lsSQL As String
   Dim loRecord As Recordset
   
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
   
   lsSQL = "SELECT " & _
               " a.sTransNox" & _
               ", e.sCompnyNm" & _
               ", c.sBnkActNo" & _
               ", c.sBankIDxx" & _
               ", b.nAmountxx" & _
            " FROM Employee_SCD_Master a" & _
            ", Employee_SCD_Detail b" & _
            ", Employee_Master001 c" & _
                  " LEFT JOIN Banks d ON c.sBankIDxx = d.sBankIDxx" & _
                  " LEFT JOIN Client_Master e ON c.sEmployID = e.sClientID" & _
            " WHERE a.sTransNox = b.sTransNox" & _
            " AND b.sEmployID = c.sEmployID" & _
            " AND a.cTranStat = '2'" & _
            " AND a.sTransNox = " & strParm(txtField(0).Text) & _
            " AND c.sBankIDxx = '00XX022'" & _
            " AND CAST(REVERSE(DECODE(UNHEX(b.nAmountxx), 'PETMGR')) AS CHAR) > 0"
      Debug.Print lsSQL
      Set loRecord = New Recordset
      loRecord.Open lsSQL, oApp.Connection, , , adCmdText
   
      If Not loRecord.EOF Then
         For lnItemCtr = 0 To loRecord.RecordCount - 1
            lnLineCtr = lnLineCtr + 1
            xlsheet.Range(XEmployNm & (lnLineCtr)).Value = loRecord("sCompnyNm")
            xlsheet.Range(XAcctNmbr & (lnLineCtr)).Value = loRecord("sBnkActNo")
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = poMiscx.Paway(loRecord("nAmountxx"))
            loRecord.MoveNext
         Next
      End If
     
     
     'kalyptus-2013.08.19 02:03pm
     'delete remaining items in our template if ever template was updated...
     lnItemCtr = xlsheet.Range("C" & xlsheet.Rows.Count).End(xlUp).Row
     If lnItemCtr > lnLineCtr Then
         xlsheet.Range("A" & lnLineCtr + 1 & ":C" & lnItemCtr).Clear
     End If
   
   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & txtField(0) & ").XLS"
   xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
   xl.Quit
    
   Set xlwbook = Nothing
   Set xl = Nothing
End Function

Private Function ExportNoAccount() As Boolean
   Dim xl As New Excel.Application
   Dim xlsheet As Excel.Worksheet
   Dim xlwbook As Excel.Workbook
   
   Dim lsSQL As String
   Dim loRecord As Recordset
   
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
    
   lsSQL = "SELECT " & _
               " a.sTransNox" & _
               ", e.sCompnyNm" & _
               ", c.sBnkActNo" & _
               ", c.sBankIDxx" & _
               ", b.nAmountxx" & _
               ", f.sBranchNm" & _
            " FROM Employee_SCD_Master a" & _
            ", Employee_SCD_Detail b" & _
            ", Employee_Master001 c" & _
                  " LEFT JOIN Banks d ON c.sBankIDxx = d.sBankIDxx" & _
                  " LEFT JOIN Client_Master e ON c.sEmployID = e.sClientID" & _
                  " LEFT JOIN Branch f ON c.sBranchCd = f.sBranchCd" & _
            " WHERE a.sTransNox = b.sTransNox" & _
            " AND b.sEmployID = c.sEmployID" & _
            " AND a.cTranStat = '2'" & _
            " AND a.sTransNox = " & strParm(txtField(0).Text) & _
            " AND c.sBankIDxx = ''" & _
            " AND CAST(REVERSE(DECODE(UNHEX(b.nAmountxx), 'PETMGR')) AS CHAR) > 0"
      Debug.Print lsSQL
      Set loRecord = New Recordset
      loRecord.Open lsSQL, oApp.Connection, , , adCmdText
      
      If Not loRecord.EOF Then
         For lnItemCtr = 0 To loRecord.RecordCount - 1
            lnLineCtr = lnLineCtr + 1
            xlsheet.Range(XEmployNo & (lnLineCtr)).Value = lnLineCtr - 6
            xlsheet.Range(XBranchNm & (lnLineCtr)).Value = loRecord("sBranchNm")
            xlsheet.Range(XEmployNm & (lnLineCtr)).Value = loRecord("sCompnyNm")
            xlsheet.Range(XNetPayxx & (lnLineCtr)).Value = poMiscx.Paway(loRecord("nAmountxx"))
            loRecord.MoveNext
         Next
      End If
  
     'kalyptus-2013.08.19 02:03pm
     'delete remaining items in our template if ever template was updated...
     lnItemCtr = xlsheet.Range("E" & xlsheet.Rows.Count).End(xlUp).Row
     If lnItemCtr > lnLineCtr Then
         xlsheet.Range("A" & lnLineCtr + 1 & ":E" & lnItemCtr).Clear
     End If
   
   xlwbook.SaveAs oApp.AppPath & "\Temp\" & sFileName & "(" & txtField(0) & ").XLS"
   xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
   xl.Quit
    
   Set xlwbook = Nothing
   Set xl = Nothing
End Function
