VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmTardinessMemoIssuance 
   BorderStyle     =   0  'None
   Caption         =   "Tardiness Memo Issuance"
   ClientHeight    =   8085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   14100
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   2280
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   1215
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4022
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3675
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Apr 25, 2011"
         Top             =   1215
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Apr 25, 2011"
         Top             =   1215
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "M00111-000021"
         Top             =   135
         Width           =   1800
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Apr 25, 2011"
         Top             =   795
         Width           =   1905
      End
      Begin VB.Label lblRemarks 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4TH QUARTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3465
         TabIndex        =   23
         Top             =   1860
         Width           =   2130
      End
      Begin VB.Line Line2 
         X1              =   3570
         X2              =   3435
         Y1              =   1245
         Y2              =   1560
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From/Thru"
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
         Left            =   105
         TabIndex        =   16
         Top             =   1290
         Width           =   885
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans No."
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
         Left            =   75
         TabIndex        =   15
         Top             =   195
         Width           =   900
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Left            =   1530
         Tag             =   "et0;ht2"
         Top             =   225
         Width           =   1800
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
         Index           =   0
         Left            =   75
         TabIndex        =   14
         Top             =   870
         Width           =   405
      End
      Begin VB.Label lblStatus 
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
         Index           =   0
         Left            =   3555
         TabIndex        =   13
         Tag             =   "eb0;et0"
         Top             =   210
         Width           =   1950
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   3510
         Top             =   180
         Width           =   2040
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   3465
         Top             =   150
         Width           =   2115
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   615
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   1085
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   900
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Apr 25, 2011"
         Top             =   90
         Width           =   4260
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Apr 25, 2011"
         Top             =   90
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   10335
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Apr 25, 2011"
         Top             =   90
         Width           =   1905
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   9390
         TabIndex        =   4
         Top             =   150
         Width           =   450
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
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
         Index           =   4
         Left            =   75
         TabIndex        =   0
         Top             =   165
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans No"
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
         Left            =   5850
         TabIndex        =   2
         Top             =   150
         Width           =   780
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2280
      Left            =   7350
      Tag             =   "wt0;fb0"
      Top             =   1215
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   4022
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   1305
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "First Warning"
         Top             =   1065
         Width           =   5250
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   1305
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1470
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1305
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Sayson, Marlon A."
         Top             =   660
         Width           =   5250
      End
      Begin VB.Label lblStatus 
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
         Index           =   1
         Left            =   4500
         TabIndex        =   24
         Tag             =   "eb0;et0"
         Top             =   180
         Width           =   1950
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   1
         Left            =   4455
         Top             =   150
         Width           =   2040
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   1
         Left            =   4410
         Top             =   120
         Width           =   2115
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IOC Level"
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
         TabIndex        =   22
         Top             =   1125
         Width           =   825
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tardiness"
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
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   1530
         Width           =   840
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
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
         Left            =   135
         TabIndex        =   18
         Top             =   720
         Width           =   870
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2820
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
      Picture         =   "frmTardinessMemoIssuance.frx":0000
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4440
      Left            =   1605
      TabIndex        =   25
      Top             =   3525
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   7832
      _Version        =   393216
      Cols            =   3
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2205
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
      Picture         =   "frmTardinessMemoIssuance.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1575
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
      Picture         =   "frmTardinessMemoIssuance.frx":0EF4
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   4440
      Left            =   7350
      TabIndex        =   26
      Top             =   3525
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   7832
      _Version        =   393216
      Cols            =   3
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
End
Attribute VB_Name = "frmTardinessMemoIssuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmTardinessMemoIssuance"

Private oSkin As clsFormSkin
Private oTrans As clsTardyMemo
Private pnActiveRow As Integer

Dim poReport As clsReport
Dim pbLoad As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0
      Unload Me
   Case 1
      oTrans.SearchTransaction "", False
      Call LoadMaster
      Call LoadDetail
      LoadDetailInfo MSFlexGrid1.Row
   Case 2
      If MsgBox("This will print the IOC(re: Habitual Tardiness) of Employee(s)." & vbCrLf & _
                "Please load printer with coupon!", vbInformation + vbOKCancel, "Notification") = vbOK Then
         ReportTrans ViewReport
         
         If Not oTrans.CloseTransaction Then
            MsgBox "IOC(re: Habitual Tardiness) closed/printed successfully!", vbInformation + vbOKOnly, "Notification"
         End If
      Else
         MsgBox "Unable to close/print IOC(re: Habitual Tardiness)!" & vbCrLf & _
                "Closing/Printing was cancelled by user...", vbInformation + vbOKOnly, "Notification"
      End If
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   With MSFlexGrid1
      .Refresh
   End With
      
   With MSFlexGrid2
      .Refresh
   End With
      
   If Not pbLoad Then
      txtSearch(0) = oTrans.BranchName
      txtSearch(1) = ""
      txtSearch(2) = ""
      
      If UCase(oApp.ProductID) <> "PETMGR" Then
         txtSearch(0).Enabled = False
      End If
      
      LoadMaster
      InitGrid
      LoadDetail
      LoadDetailInfo MSFlexGrid1.Row
      pbLoad = True
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
         
   Set oTrans = New clsTardyMemo
   Set oTrans.AppDriver = oApp
   oTrans.TransStatus = 2
   oTrans.InitTransaction
            
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   pbLoad = False
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Rows = 2
      .Cols = 3
      
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
      .TextMatrix(0, 2) = "Tardiness"
      
      .RowHeightMin = 320
      
      .ColWidth(0) = 635
      .ColWidth(1) = 3520 + 300
      .ColWidth(2) = 1200
      
      'column allinment
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      
      'set location
      .Row = 1
      .Col = 2
      .ColSel = .Cols - 1
   End With


   With MSFlexGrid2
      .Rows = 2
      .Cols = 6
      
      .Row = 0
      .RowHeight(0) = 320
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Row = 1
      .TextMatrix(0, 0) = "Date"
      .TextMatrix(0, 1) = "In"
      .TextMatrix(0, 2) = "Out"
      .TextMatrix(0, 3) = "In"
      .TextMatrix(0, 4) = "Out"
      .TextMatrix(0, 5) = "Tardiness"
      
      .RowHeightMin = 320
      
      .ColWidth(0) = 1300 + 300
      .ColWidth(1) = 1000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      
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

Private Sub MSFlexGrid1_SelChange()
   If pnActiveRow <> MSFlexGrid1.Row Then
      pnActiveRow = MSFlexGrid1.Row
            
      Call LoadDetailInfo(pnActiveRow)
   End If
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
      Select Case Index
      Case 0
         Call oTrans.SearchBranch(txtSearch(0), False, True)
         txtSearch(0) = oTrans.BranchName
      Case 1, 2
         Call oTrans.SearchTransaction(txtSearch(Index), Index = 1)
         Call LoadMaster
         Call LoadDetail
         Call LoadDetailInfo(MSFlexGrid1.Row)
      End Select
   End If
End Sub

Private Sub LoadMaster()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 1 To 3
         loTxt.Text = Format(oTrans.Master(loTxt.Index), "Mmm. DD, YYYY")
      Case Else
         loTxt.Text = oTrans.Master(loTxt.Index)
      End Select
   Next
   
   lblRemarks = oTrans.Master("sDescript")
   Select Case oTrans.Master("cTranStat")
   Case "2"
      lblStatus(0) = "Open"
   Case "4"
      lblStatus(0) = "Posted"
   Case Else
      lblStatus(0) = "Unknown"
   End Select
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Rows = 2
      .TextMatrix(lnCtr + 1, 0) = Format(lnCtr + 1, "0000")
      .TextMatrix(lnCtr + 1, 1) = ""
      .TextMatrix(lnCtr + 1, 2) = ""
      
      If oTrans.ItemCount >= 1 Then
         
         If oTrans.ItemCount > 12 Then
            .ColWidth(1) = 3520
         Else
            .ColWidth(1) = 3520 + 300
         End If
         
         For lnCtr = 0 To oTrans.ItemCount - 1
            DoEvents
            
            .Rows = 2 + lnCtr
            .RowHeight(lnCtr + 1) = 338
            
            .TextMatrix(lnCtr + 1, 0) = Format(lnCtr + 1, "0000")
            .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "sRecpntxx")
            .TextMatrix(lnCtr + 1, 2) = Format(oTrans.Detail(lnCtr, "nTardyxxx"), "#,##0.00")
            
            DoEvents
         Next
      End If
      
   End With
End Sub

Private Sub LoadDetailInfo(ByVal fnRow As Integer)
   Dim lsSQL As String
   Dim lors As Recordset
   Dim loTxt As TextBox
   Dim lnTotal As Currency
   Dim lnRow As Integer
   
   Dim lsOldProc As String
   
   lsOldProc = "LoadDetailInfo"
'   'On Error GoTo errProc
   
   'Load info to the textboxes - txtDetail
   If fnRow < 0 Then
      For Each loTxt In txtDetail
         loTxt = ""
      Next
   Else
      txtDetail(1) = IFNull(oTrans.Detail(fnRow - 1, 1), "")
      txtDetail(10) = Format(oTrans.Detail(fnRow - 1, 10), "#,##0.00")
      Select Case oTrans.Detail(fnRow - 1, 12)
      Case "0"
         txtDetail(12) = "First Warning"
      Case "1"
         txtDetail(12) = "Last Warning"
      Case "2"
         txtDetail(12) = "For Suspension"
      Case Else
         txtDetail(12) = ""
      End Select
   End If
   
   Select Case oTrans.Detail(fnRow - 1, "cTranStat")
   Case "0"
      lblStatus(1) = "Open"
   Case "1"
      lblStatus(1) = "Printed"
   Case "2"
      lblStatus(1) = "Confirmed"
   Case "3"
      lblStatus(1) = "Cancelled"
   Case "4"
      lblStatus(1) = "Void"
   Case Else
      lblStatus(1) = "Unknown"
   End Select
   
   lsSQL = "SELECT" & _
                  "  a.dTransact" & _
                  ", b.dAMInxxxx" & _
                  ", b.dAMOutxxx" & _
                  ", b.dPMInxxxx" & _
                  ", b.dPMOutxxx" & _
                  ", (a.nTardyxxx + IFNULL(c.nAdjTardi, 0) - IFNULL(d.nOmittedx, 0)) nTardyxxx" & _
          " FROM Employee_Timesheet a" & _
               " LEFT JOIN Employee_Log b ON a.sEmployID = b.sEmployID AND a.dTransact = b.dTransact" & _
               " LEFT JOIN Employee_Timesheet_Adjustment c ON a.sEmployID = c.sEmployID AND a.dTransact = c.dTransact AND c.cTranStat NOT IN ('0', '3')" & _
               " LEFT JOIN Employee_Tardiness d ON a.sEmployID = d.sEmployID AND a.dTransact = d.dRequestd AND c.cTranStat NOT IN ('0', '3')" & _
          " WHERE a.sEmployID = " & strParm(oTrans.Detail(fnRow - 1, "sEmployTo")) & _
            " AND a.dTransact BETWEEN " & dateParm(oTrans.Master("sPeriodFr")) & " AND " & dateParm(oTrans.Master("sPeriodTo")) & _
            " AND (a.nTardyxxx + IFNULL(c.nAdjTardi, 0) - IFNULL(d.nOmittedx, 0)) > 0" & _
            " AND a.cRestDayx <> '2'"
            
   Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
      
   With MSFlexGrid2
      .Rows = 1
   
      lnRow = 1
      
      If lors.RecordCount > 12 Then
         .ColWidth(0) = 1300
      Else
         .ColWidth(0) = 1300 + 300
      End If
      
      Do Until lors.EOF
         .Rows = .Rows + 1
         .TextMatrix(lnRow, 0) = lors("dTransact")
         .TextMatrix(lnRow, 1) = Format(IFNull(lors("dAMInxxxx")), "HH:MM AMPM")
         .TextMatrix(lnRow, 2) = Format(IFNull(lors("dAMOutxxx")), "HH:MM AMPM")
         .TextMatrix(lnRow, 3) = Format(IFNull(lors("dPMInxxxx")), "HH:MM AMPM")
         .TextMatrix(lnRow, 4) = Format(IFNull(lors("dPMOutxxx")), "HH:MM AMPM")
         .TextMatrix(lnRow, 5) = IFNull(lors("nTardyxxx"))
         
         lnRow = lnRow + 1
         lors.MoveNext
      Loop
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Function ReportTrans(ByVal ReportType As adReport) As Boolean
   Dim lnCtr As Integer
   Dim lnDtl As Long
   Dim lsSQL As String
   Dim loRSSht As Recordset
   Dim loRSRep As Recordset
   Dim lsOldProc As String
         
   'lsOldProc = pxeMODULENAME & "." & "printTrans"
   'On Error GoTo errProc
   
   Set loRSRep = New Recordset
   
   With loRSRep
      .Fields.Append "sField01", adVarChar, 50
      .Fields.Append "sField02", adVarChar, 50
      .Fields.Append "sField03", adVarChar, 50
      
      .Fields.Append "sField04", adVarChar, 12
      .Fields.Append "sField05", adVarChar, 12
      .Fields.Append "sField06", adVarChar, 12
      .Fields.Append "sField07", adVarChar, 12
      .Fields.Append "sField08", adVarChar, 12
      
      .Fields.Append "sField09", adVarChar, 250
      .Fields.Append "sField10", adVarChar, 250
      
      .Fields.Append "sField11", adVarChar, 25
      .Fields.Append "sField12", adVarChar, 50
      
      .Fields.Append "nField01", adInteger
      .Fields.Append "nField02", adInteger
      .Open
   End With
      
   For lnCtr = 0 To oTrans.ItemCount - 1
      If oTrans.Detail(lnCtr, "cTranStat") = "0" Then
         lsSQL = "SELECT" & _
                        "  a.dTransact" & _
                        ", b.dAMInxxxx" & _
                        ", b.dAMOutxxx" & _
                        ", b.dPMInxxxx" & _
                        ", b.dPMOutxxx" & _
                        ", (a.nTardyxxx + IFNULL(c.nAdjTardi, 0) - IFNULL(d.nOmittedx, 0)) nTardyxxx" & _
                " FROM Employee_Timesheet a" & _
                     " LEFT JOIN Employee_Log b ON a.sEmployID = b.sEmployID AND a.dTransact = b.dTransact" & _
                     " LEFT JOIN Employee_Timesheet_Adjustment c ON a.sEmployID = c.sEmployID AND a.dTransact = c.dTransact AND c.cTranStat NOT IN ('0', '3')" & _
                     " LEFT JOIN Employee_Tardiness d ON a.sEmployID = d.sEmployID AND a.dTransact = d.dRequestd AND c.cTranStat NOT IN ('0', '3')" & _
                " WHERE a.sEmployID = " & strParm(oTrans.Detail(lnCtr, "sEmployTo")) & _
                  " AND a.dTransact BETWEEN " & dateParm(oTrans.Master("sPeriodFr")) & " AND " & dateParm(oTrans.Master("sPeriodTo")) & _
                  " AND (a.nTardyxxx + IFNULL(c.nAdjTardi, 0) - IFNULL(d.nOmittedx, 0)) > 0" & _
                  " AND a.cRestDayx <> '2'"
         
         Set loRSSht = oApp.Connection.Execute(lsSQL, , adCmdText)
         
'         lnDtl = 0
         Do Until loRSSht.EOF
            lnDtl = lnDtl + 1
'            If lnDtl <= 8 Then
               loRSRep.AddNew
            
               loRSRep("sField04") = Format(loRSSht("dTransact"), "MM/DD/YYYY")
               loRSRep("sField05") = Format(loRSSht("dAMInxxxx"), "HH:MM AMPM")
               loRSRep("sField06") = Format(loRSSht("dAMOutxxx"), "HH:MM AMPM")
               loRSRep("sField07") = Format(loRSSht("dPMInxxxx"), "HH:MM AMPM")
               loRSRep("sField08") = Format(loRSSht("dPMOutxxx"), "HH:MM AMPM")
               loRSRep("nField02") = Format(loRSSht("nTardyxxx"), "#,##0")
'            Else
'               loRSRep("sField04") = "Others"
'               loRSRep("sField05") = ""
'               loRSRep("sField06") = ""
'               loRSRep("sField07") = ""
'               loRSRep("sField08") = ""
'               loRSRep("nField02") = Format(loRSRep("nField02") + loRSSht("nTardyxxx"), "#,##0")
'            End If
            
            loRSRep("sField01") = oTrans.Detail(lnCtr, "sRecpntxx")
            loRSRep("sField02") = oTrans.Detail(lnCtr, "sSenderxx")
            loRSRep("sField11") = Format(oTrans.Master("dTransact"), "Mmmm DD, YYYY")
            loRSRep("sField12") = Format(oTrans.Master("sPeriodFr"), "YYYY/MM/DD") & "-" & Format(oTrans.Master("sPeriodTo"), "YYYY/MM/DD")
            
            Select Case oTrans.Detail(lnCtr, "cIOCTypex")
            Case "0"
               loRSRep("sField03") = "NOTICE OF HABITUAL TARDINESS"
               loRSRep("sField09") = "     May this notice serves as a WARNING!  Failure to comply with Company Rules & Regulation may cause the company to apply the necessary sanctions indicated in the Employees Manual."
               loRSRep("sField10") = "     Please do not hesitate to contact your Branch/Department Head if you have questions or need clarification."
            Case "1"
               loRSRep("sField03") = "NOTICE OF HABITUAL TARDINESS"
               loRSRep("sField09") = "     May this notice serves as a LAST WARNING!  Failure to comply with Company Rules & Regulation may cause the company to apply the necessary sanctions indicated in the Employees Manual."
               loRSRep("sField10") = "     Please do not hesitate to contact your Branch/Department Head if you have questions or need clarification."
            Case Else
               loRSRep("sField03") = "NOTICE OF SUSPENSION DUE TO HABITUAL TARDINESS"
               loRSRep("sField09") = "     This memo serves as a NOTICE OF SUSPENSION!  Failure to comply with Company Rules & Regulation may cause the company to apply the necessary sanctions/termination indicated in the Employees Manual."
               loRSRep("sField10") = "     Please submit your written explanation to the HRM Dept. within 48 hours upon the receipt of this memorandum."
            End Select
                     
            loRSRep("nField01") = Format(oTrans.Detail(lnCtr, "nTardyxxx"), "#,##0")
                     
            loRSSht.MoveNext
         Loop
      End If
   Next
      
   If lnDtl > 0 Then
      Set poReport = New clsReport
      poReport.InitReport
      Set poReport.ReportSource = loRSRep
      poReport.ReportID = "EmpTM1"
      poReport.ReportHeading1 = "NO HEADER"
      
      If ReportType = ViewReport Then
         poReport.ShowReport
      Else
         poReport.PrintReport
      End If
   Else
      MsgBox "All records has been printed!", vbOKOnly, "Notification"
   End If
   ReportTrans = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

