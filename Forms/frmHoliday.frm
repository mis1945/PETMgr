VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrcontrol.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHoliday 
   BorderStyle     =   0  'None
   Caption         =   "Holidays"
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   12645
      TabIndex        =   6
      Top             =   1935
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
      Picture         =   "frmHoliday.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   12645
      TabIndex        =   2
      Top             =   1305
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Retreive"
      AccessKey       =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmHoliday.frx":077A
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   4800
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1275
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   8467
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.CheckBox chkField 
         Caption         =   "Deduct Undertime"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   13
         Left            =   1065
         TabIndex        =   28
         Tag             =   "wt0;fb0"
         Top             =   4170
         Width           =   2085
      End
      Begin VB.CheckBox chkField 
         Caption         =   "Deduct Tardiness"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   12
         Left            =   1065
         TabIndex        =   27
         Tag             =   "wt0;fb0"
         Top             =   3795
         Width           =   2085
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
         Index           =   11
         Left            =   4095
         MaxLength       =   50
         TabIndex        =   25
         Text            =   "0"
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   10
         Left            =   4095
         MaxLength       =   50
         TabIndex        =   23
         Text            =   "1"
         Top             =   2790
         Width           =   975
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
         Index           =   9
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   21
         Text            =   "1"
         Top             =   3240
         Width           =   975
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
         Index           =   8
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   19
         Text            =   "1"
         Top             =   2790
         Width           =   975
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
         Index           =   80
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   17
         Text            =   "New Years Day"
         Top             =   2265
         Width           =   4000
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
         Index           =   81
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   15
         Text            =   "New Years Day"
         Top             =   1815
         Width           =   4000
      End
      Begin VB.ComboBox cmbField 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "frmHoliday.frx":0EF4
         Left            =   1050
         List            =   "frmHoliday.frx":0F07
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1005
         Width           =   2415
      End
      Begin VB.CheckBox chkField 
         Caption         =   "Recurring Holiday"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   1065
         TabIndex        =   13
         Tag             =   "wt0;fb0"
         Top             =   1425
         Width           =   2355
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
         Index           =   2
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "January 1"
         Top             =   555
         Width           =   2415
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
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "New Years Day"
         Top             =   105
         Width           =   4000
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   120
         X2              =   5070
         Y1              =   2730
         Y2              =   2730
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   105
         X2              =   5055
         Y1              =   1755
         Y2              =   1755
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime"
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
         Left            =   3000
         TabIndex        =   24
         Top             =   3330
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rest Day"
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
         Left            =   3000
         TabIndex        =   22
         Top             =   2880
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Present"
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
         TabIndex        =   20
         Top             =   3330
         Width           =   675
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Absent"
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
         TabIndex        =   18
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Town"
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
         Left            =   105
         TabIndex        =   16
         Top             =   2355
         Width           =   450
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
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
         Left            =   105
         TabIndex        =   14
         Top             =   1905
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   420
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
         TabIndex        =   9
         Top             =   645
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Holiday"
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
         Left            =   105
         TabIndex        =   7
         Top             =   195
         Width           =   645
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   690
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   1217
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   1
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
         Height          =   420
         Index           =   0
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "2025"
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "YEAR"
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
         Index           =   24
         Left            =   105
         TabIndex        =   0
         Top             =   210
         Width           =   510
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   12645
      TabIndex        =   4
      Top             =   1305
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
      Picture         =   "frmHoliday.frx":0F4A
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   12645
      TabIndex        =   5
      Top             =   1935
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
      Picture         =   "frmHoliday.frx":16C4
   End
   Begin xrControl.xrButton cmdDetail 
      Height          =   600
      Left            =   12645
      TabIndex        =   3
      Top             =   675
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Add Holiday"
      AccessKey       =   "Add Holiday"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4815
      Left            =   5370
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1275
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   8493
      _Version        =   393216
      FocusRect       =   0
      BorderStyle     =   0
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
Attribute VB_Name = "frmHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmHoliday"

Private WithEvents oTrans As clsHoliday
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim pbFlexGrid As Boolean
Dim pbSearch As Boolean

Private Sub chkField_Validate(Index As Integer, Cancel As Boolean)
   If oTrans.Detail(MSFlexGrid1.Row - 1, Index) = chkField(Index).Value Then Exit Sub
   
   oTrans.Detail(MSFlexGrid1.Row - 1, Index) = chkField(Index).Value
   chkField(Index).Value = oTrans.Detail(MSFlexGrid1.Row - 1, Index)
End Sub

'Private Sub chkField_Validate(Cancel As Boolean)
'   If oTrans.Detail(MSFlexGrid1.Row - 1, 4) <> chkField.Value Then
'      oTrans.Detail(MSFlexGrid1.Row - 1, 4) = chkField.Value
'   End If
'End Sub

Private Sub cmbField_Validate(Cancel As Boolean)
   oTrans.Detail(MSFlexGrid1.Row - 1, 5) = cmbField.ListIndex
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0   'Retrieve
      If IsNumeric(txtSearch(0)) And _
         Len(txtSearch(0)) = 4 Then
         
         oTrans.HolidayYear = txtSearch(0)
         
         If oTrans.OpenTransaction Then
            InitForm 1
            
            LoadDetail
         
            With MSFlexGrid1
               If .Row > 10 Then .TopRow = .TopRow + 1
               .Col = 1
               .ColSel = .Cols - 1
               .SetFocus
               pbFlexGrid = True
            End With
            
            Call detailFieldChange
            
            pbSearch = False
            txtfield(1).SetFocus
         End If
      Else
         txtSearch(0).SetFocus
      End If
   Case 1
      Unload Me
   Case 2
      If oTrans.SaveTransaction Then
         InitForm 0
         LoadDetail
         InitGrid
         GoTo endWithFocus
      End If
   Case 3
      If oTrans.InitTransaction Then
         InitForm 0
         LoadDetail
         InitGrid
         GoTo endWithFocus
      End If
   End Select
   
endProc:
   Exit Sub
endWithFocus:
   txtSearch(0) = ""
   txtSearch(0).SetFocus
   GoTo endProc
End Sub

Private Sub cmdDetail_Click()
   Dim lnRow As Integer
   Dim loTxt As TextBox

   lnRow = oTrans.ItemCount

   oTrans.Detail(lnRow, 1) = ""

   If oTrans.ItemCount > lnRow Then
      MSFlexGrid1.Rows = oTrans.ItemCount + 1
      MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = MSFlexGrid1.Row

      For Each loTxt In txtfield
         loTxt = ""
      Next

      chkField(4).Value = 0
      chkField(12).Value = 0
      chkField(13).Value = 0
      
      cmbField.ListIndex = -1
      
      txtfield(1).SetFocus
   End If

End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
   End If

   pbFlexGrid = False
   pbSearch = False
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim loTxt As TextBox
   Dim lnRow As Integer
   
   If pbFlexGrid Then
      With MSFlexGrid1
         Select Case KeyCode
         Case vbKeyReturn
            SetNextFocus
         Case vbKeyUp
            If Not (.Row = 1) Then
               .Row = .Row - 1
            End If
            
            If Not .RowIsVisible(.Row) Then .TopRow = .TopRow - 1
            
            lnRow = MSFlexGrid1.Row - 1
         
            Call detailFieldChange(lnRow)
            
            KeyCode = 0
         Case vbKeyDown
            If Not (.Row = .Rows - 1) Then
               .Row = .Row + 1
            End If
            
            If Not .RowIsVisible(.Row) Then .TopRow = .TopRow + 1
            
            lnRow = MSFlexGrid1.Row - 1
         
            Call detailFieldChange(lnRow)
            
            KeyCode = 0
         End Select
         
         .Col = 1
         .ColSel = .Cols - 1
      End With
   ElseIf pbSearch Then
      If KeyCode = vbKeyReturn Then Call cmdButton_Click(0)
   Else
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End If
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
'   On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsHoliday
   Set oTrans.AppDriver = oApp

   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance

   Call InitForm(0)
   Call InitGrid
   Call LoadDetail

   txtSearch(0).Text = ""
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub InitForm(ByVal fnEdit As Integer)
   Dim loTxt As TextBox

   xrFrame2.Enabled = Not (fnEdit = 0)
   cmdButton(2).Visible = Not (fnEdit = 0)
   cmdButton(3).Visible = Not (fnEdit = 0)
   cmdDetail.Visible = Not (fnEdit = 0)

   cmdButton(0).Visible = (fnEdit = 0)
   cmdButton(1).Visible = (fnEdit = 0)
   xrFrame1.Enabled = (fnEdit = 0)

   For Each loTxt In txtfield
      loTxt = ""
   Next

   chkField(4).Value = 0
   chkField(12).Value = 0
   chkField(13).Value = 0
   cmbField.ListIndex = -1

End Sub

Private Sub MSFlexGrid1_Click()
   Dim lnRow As Integer
   Dim loTxt As TextBox

   lnRow = MSFlexGrid1.Row - 1

   For Each loTxt In txtfield
      If loTxt.Index = 2 Then
         If IsDate(oTrans.Detail(lnRow, 3) & "/" & oTrans.Detail(lnRow, 2)) Then
            loTxt = Format(oTrans.Detail(lnRow, 3) & "/" & oTrans.Detail(lnRow, 2), "Mmmm DD")
         Else
            loTxt = ""
         End If
      Else
         loTxt = oTrans.Detail(lnRow, loTxt.Index)
      End If
   Next

   chkField(4).Value = oTrans.Detail(lnRow, 4)
   chkField(12).Value = oTrans.Detail(lnRow, 12)
   chkField(13).Value = oTrans.Detail(lnRow, 13)
   cmbField.ListIndex = oTrans.Detail(lnRow, 5)
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Cols = 3
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
      .TextMatrix(0, 0) = "NO"
      .TextMatrix(0, 1) = "HOLIDAY"
      .TextMatrix(0, 2) = "DATE"
'      .RowHeightMin = 338
      
      'column width
      .ColWidth(0) = 600
      .ColWidth(1) = 3905
      .ColWidth(2) = 2500

      'column allinment
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      
      'set location
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer

   With MSFlexGrid1
      If oTrans.ItemCount > 10 Then
         .ColWidth(1) = 3650
      Else
         .ColWidth(1) = 3905
      End If
      
      For lnRow = 0 To oTrans.ItemCount - 1
         .Rows = lnRow + 2
         .TextMatrix(lnRow + 1, 0) = lnRow + 1
         .TextMatrix(lnRow + 1, 1) = oTrans.Detail(lnRow, "sHolidyNm")
         If IsDate(oTrans.Detail(lnRow, 3) & "/" & oTrans.Detail(lnRow, 2)) Then
            .TextMatrix(lnRow + 1, 2) = Format(oTrans.Detail(lnRow, 3) & "/" & oTrans.Detail(lnRow, 2), "Mmmm DD")
         Else
            .TextMatrix(lnRow + 1, 2) = ""
         End If
      Next
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

Private Sub MSFlexGrid1_GotFocus()
   pbFlexGrid = True
End Sub

'Private Sub MSFlexGrid1_SelChange()
'   Dim lnRow As Integer
'   Dim loTxt As TextBox
'
'   With MSFlexGrid1
'
'      lnRow = MSFlexGrid1.Row - 1
'
'      For Each loTxt In txtField
'         If loTxt.Index = 2 Then
'            If IsDate(oTrans.Detail(lnRow, 3) & "/" & oTrans.Detail(lnRow, 2)) Then
'               loTxt = Format(oTrans.Detail(lnRow, 3) & "/" & oTrans.Detail(lnRow, 2), "Mmmm DD")
'            Else
'               loTxt = ""
'            End If
'         Else
'            loTxt = oTrans.Detail(lnRow, loTxt.Index)
'         End If
'      Next
'
'      chkField.Value = oTrans.Detail(lnRow, 4)
'      cmbField.ListIndex = oTrans.Detail(lnRow, 5)
'
'      .Col = 1
'      .ColSel = .Cols - 1
'
'   End With
'
'   If txtField(1).Enabled = True Then txtField(1).SetFocus
'End Sub

Private Sub oTrans_DetailRetrieved(ByVal Row As Integer, ByVal Index As Variant)
   Select Case Index
   Case 4, 12, 13
      chkField(Index).Value = oTrans.Detail(Row, Index)
   Case 5
      cmbField.ListIndex = oTrans.Detail(Row, Index)
   Case Else
      txtfield(Index) = oTrans.Detail(Row, Index)
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
      Case 1
         pbFlexGrid = True
      Case Else
         pbFlexGrid = False
   End Select
   
   With txtfield(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
      Select Case Index
      Case 80, 81
         Call oTrans.SearchDetail(MSFlexGrid1.Row - 1, Index, txtfield(Index))
      End Select
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtfield(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 1
      oTrans.Detail(MSFlexGrid1.Row - 1, Index) = txtfield(Index)
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = txtfield(Index)
   Case 2
      If Not IsDate(txtfield(Index)) Then
         Cancel = True
      Else
         txtfield(Index) = Format(txtfield(2), "Mmmm DD")
         MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txtfield(Index)
         oTrans.Detail(MSFlexGrid1.Row - 1, 2) = Day(txtfield(Index))
         oTrans.Detail(MSFlexGrid1.Row - 1, 3) = Month(txtfield(Index))
      End If
   Case Else
      oTrans.Detail(MSFlexGrid1.Row - 1, Index) = txtfield(Index)
   End Select
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
   pbSearch = True
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      If IsNumeric(txtSearch(0)) And _
         Len(txtSearch(0)) = 4 Then
         
         oTrans.HolidayYear = txtSearch(0)
         
         If oTrans.OpenTransaction Then
            InitForm 1
            
            LoadDetail
         
            With MSFlexGrid1
               If .Row > 15 Then .TopRow = 1
               .Col = 1
               .ColSel = .Cols - 1
               .SetFocus
               pbFlexGrid = True
            End With
            
            Call detailFieldChange
            
            pbSearch = False
            txtfield(1).SetFocus
         End If
      End If
   End If
End Sub

Private Sub detailFieldChange(Optional ByVal lnRow As Integer)
   Dim loTxt As TextBox
   
   For Each loTxt In txtfield
      If IsMissing(lnRow) Then lnRow = 0
      If loTxt.Index = 2 Then
         If IsDate(oTrans.Detail(lnRow, 3) & "/" & oTrans.Detail(lnRow, 2)) Then
            loTxt = Format(oTrans.Detail(lnRow, 3) & "/" & oTrans.Detail(lnRow, 2), "Mmmm DD")
         Else
            loTxt = ""
         End If
      Else
         loTxt = oTrans.Detail(lnRow, loTxt.Index)
      End If
   Next

   chkField(4).Value = oTrans.Detail(lnRow, 4)
   chkField(12).Value = oTrans.Detail(lnRow, 12)
   chkField(13).Value = oTrans.Detail(lnRow, 13)
   
   cmbField.ListIndex = oTrans.Detail(lnRow, 5)
End Sub
