VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSalRegion 
   BorderStyle     =   0  'None
   Caption         =   "Salary Regions"
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1802
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   3175
      BackColor       =   12632256
      BorderStyle     =   1
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
         Left            =   1620
         TabIndex        =   12
         Top             =   1200
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
         Left            =   1620
         TabIndex        =   10
         Top             =   690
         Width           =   4815
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
         Index           =   0
         Left            =   1620
         TabIndex        =   8
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Wage"
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
         Left            =   165
         TabIndex        =   11
         Top             =   1290
         Width           =   1365
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1725
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   2415
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   165
         TabIndex        =   9
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title ID"
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
         Index           =   0
         Left            =   165
         TabIndex        =   7
         Top             =   210
         Width           =   675
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   6000
      TabIndex        =   6
      Top             =   8235
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSalRegion.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   5220
      TabIndex        =   4
      Top             =   8235
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Browse"
      AccessKey       =   "B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSalRegion.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   4440
      TabIndex        =   3
      Top             =   8235
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Save"
      AccessKey       =   "S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSalRegion.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   3660
      TabIndex        =   1
      Top             =   8220
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Update"
      AccessKey       =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSalRegion.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   2880
      TabIndex        =   0
      Top             =   8220
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&New"
      AccessKey       =   "N"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSalRegion.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   6000
      TabIndex        =   5
      Top             =   8235
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Close"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSalRegion.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   4440
      TabIndex        =   2
      Top             =   8235
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Delete"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSalRegion.frx":2CDC
      PicturePos      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3810
      Left            =   105
      TabIndex        =   13
      Top             =   4065
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   6720
      _Version        =   393216
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
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1650
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   2370
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   2910
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   19
         Text            =   "4,000.00"
         Top             =   1065
         Width           =   2415
      End
      Begin VB.TextBox txtDetail 
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
         Index           =   1
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   15
         Text            =   "Junior Supervisor 1"
         Top             =   105
         Width           =   4815
      End
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   17
         Text            =   "4,000.00"
         Top             =   585
         Width           =   2415
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEI Guide"
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
         Left            =   165
         TabIndex        =   18
         Top             =   1155
         Width           =   885
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Level"
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
         Left            =   165
         TabIndex        =   14
         Top             =   195
         Width           =   1050
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   165
         TabIndex        =   16
         Top             =   675
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmSalRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmSalRegion"
Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private poDetail As Recordset
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Private pcSalGrpID As String
Private psSalGrpDs As String

Dim pnCtr As Integer
Dim pnIndex As Integer
Dim pbFlexGrid As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc
   Select Case Index
   Case 0
      oDriver.RecordCancelUpdate
'      initFrame (1)
      LoadDetail
   Case 1
      oDriver.BrowseRecord
      LoadDetail
      
      With MSFlexGrid1
         If .Row > 10 Then .TopRow = .TopRow + 1
         .Col = 1
         .ColSel = .Cols - 1
         .SetFocus
         pbFlexGrid = True
      End With
      
   Case 2
      oDriver.RecordSave
   Case 3
      oDriver.RecordUpdate
   Case 4
      oDriver.RecordNew
      LoadDetail
   Case 5
      Unload Me
   Case 6
      oDriver.RecordDelete
      ClearFields
      LoadDetail
   End Select
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      oDriver.RecordNew
      oDriver.DisableTextbox 0
      bLoaded = True
   End If
   
   pbFlexGrid = False
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim lors As Recordset
   
   lsOldProc = "Form_Load"
   On Error GoTo errProc
   
   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
    'Kalyptus - 2023.11.17 11:54am
    'Get the Salary Matrix Group of the current user based on loaded Branch
    lsSQL = "SELECT a.cSalGrpID, a.sSalGrpDs" & _
           " FROM Salary_Matrix_Group a" & _
                " LEFT JOIN Salary_Matrix_Group_Member b ON a.cSalGrpID = b.cSalGrpID" & _
                " LEFT JOIN Branch_Others c ON c.cDivision = b.sDivsnCde" & _
           " WHERE c.sBranchCD = " & strParm(oApp.BranchCode) & _
             " AND a.cRecdStat = '1'"
    Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
    
    If Not lors.EOF Then
        pcSalGrpID = lors("cSalGrpID")
        psSalGrpDs = lors("sSalGrpDs")
    Else
        pcSalGrpID = ""
        psSalGrpDs = ""
    End If
   
   With oDriver
      .RecQuery = "SELECT" _
                     & "  sSalRegID" _
                     & ", sSalRegNm" _
                     & ", nRegAmtxx" _
                     & ", sTaxRegID" _
                     & ", cRecdStat" _
                     & ", sModified" _
                     & ", dModified" _
                  & " FROM Salary_Region"
           
      .BrowseQuery = "SELECT" _
                        & "  sSalRegID" _
                        & ", sSalRegNm" _
                     & " FROM Salary_Region" _
                     & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY sSalRegID"
      .InitRecForm
   
      .BrowseFTitle(0) = "Code"
      .BrowseFTitle(1) = "Name"
      .BrowseFFormat(0) = "@@@"
   
      .FieldFormat(0) = "@@@"
      .FieldSize(0) = Len(.FieldFormat(0))
      .FieldStart = 1
      
   End With

   InitGrid
   LoadDetail

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oDriver = Nothing
End Sub


Private Sub MSFlexGrid1_DblClick()
   Dim lnRow As Integer
   Dim loTxt As TextBox
   
   With MSFlexGrid1
      lnRow = .Row
      
      For Each loTxt In txtDetail
         loTxt = .TextMatrix(lnRow, loTxt.Index)
      Next
      
      .Col = 1
      .ColSel = .Cols - 1
      
   End With
   
   If xrFrame3.Enabled Then txtDetail(2).SetFocus
End Sub

Private Sub MSFlexGrid1_GotFocus()
   pbFlexGrid = True
End Sub

'Private Sub MSFlexGrid1_SelChange()
'   Dim lnRow As Integer
'   Dim loTxt As TextBox
'
'   With MSFlexGrid1
'      lnRow = .Row
'
'      For Each loTxt In txtDetail
'         loTxt = .TextMatrix(lnRow, loTxt.Index)
'      Next
'
'      .Col = 1
'      .ColSel = .Cols - 1
'
'   End With
'
''   If xrFrame3.Enabled Then txtDetail(2).SetFocus
'End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   xrFrame3.Enabled = oDriver.EditMode = xeModeUpdate
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   On Error GoTo errProc
   
   If Not oDriver.SetValue(0, GetNextCode("Salary_Region", "sSalRegID", False, oApp.Connection)) Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(1) = ""
   oDriver.FieldValue(2) = 0#
   oDriver.FieldValue(3) = ""
   oDriver.FieldValue(4) = xeRecStateActive
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_WillSave"
   On Error GoTo errProc
   
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Description detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   End If
   
   If oDriver.FieldValue(2) <= 0# Then
      MsgBox "Invalid Minimum Regional Wage Salary detected!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      Cancel = True
   End If
   
   SaveDetail

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", False
End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
   With txtDetail(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtDetail_Validate(Index As Integer, Cancel As Boolean)
   With MSFlexGrid1
      If IsNumeric(txtDetail(Index)) Then
         .TextMatrix(.Row, Index) = Format(txtDetail(Index), "#,##0.00")
      End If
   
      txtDetail(Index) = Format(.TextMatrix(.Row, Index), "#,##0.00")
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   oDriver.ColumnIndex = Index
   pnIndex = Index
   pbFlexGrid = False
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim loTxt As TextBox
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
            
            For Each loTxt In txtDetail
               loTxt = .TextMatrix(.Row, loTxt.Index)
            Next
            
            KeyCode = 0
         Case vbKeyDown
            If Not (.Row = .Rows - 1) Then
               .Row = .Row + 1
            End If
            
            If Not .RowIsVisible(.Row) Then .TopRow = .TopRow + 1
            
            For Each loTxt In txtDetail
               loTxt = .TextMatrix(.Row, loTxt.Index)
            Next
            
            KeyCode = 0
         Case vbKeyF12
            If oDriver.EditMode = xeModeReady And oApp.UserLevel = xeEngineer Then
               Call createAdjustment
            End If
         Case vbKeyF11
            If oApp.UserLevel = xeEngineer Then
               Call createClearance
            End If
         End Select
         
         .Col = 1
         .ColSel = .Cols - 1
         
      End With
   Else
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub
Private Sub txtDetail_LostFocus(Index As Integer)
   With txtDetail(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub
Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_LostFocus"
   On Error GoTo errProc
         
   If Index = 2 Then
      txtField(Index).Text = TitleCase(txtField(Index).Text)
   End If
   Cancel = Not oDriver.ValidateField(Index)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Cols = 4
      .Rows = 2
      
      .Row = 0
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Row = 1
      'column title
      .TextMatrix(0, 0) = "ID"
      .TextMatrix(0, 1) = "Salary Level"
      .TextMatrix(0, 2) = "Salary Guide"
      .TextMatrix(0, 3) = "DEI Guide"
      .RowHeightMin = 338
      
      'column width
      .ColWidth(0) = 500
      .ColWidth(1) = 3100
      .ColWidth(2) = 1506
      .ColWidth(3) = 1506
      
      'column alignment
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignRightCenter
      .ColAlignment(3) = flexAlignRightCenter
      
      'set position
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub LoadDetail()
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim loTxt As TextBox
   Dim lsOldProc As String
   
   lsOldProc = "LoadDetail"
   On Error GoTo errProc
   
   lsSQL = "SELECT" & _
                  "  b.sSalLvlID" & _
                  ", b.sSalLvlNm" & _
                  ", a.nAmountxx" & _
                  ", a.nDEIGuide" & _
                  ", a.sSalRegID" & _
                  ", a.cSalGrpID" & _
          " FROM Salary_Level b" & _
               " LEFT JOIN Salary_Level_Region a ON a.sSalLvlID = b.sSalLvlID AND a.sSalRegID = " & strParm(oDriver.FieldValue(0)) & _
                    " AND b.cSalGrpID = a.cSalGrpID" & _
          " WHERE b.cSalGrpID = " & strParm(pcSalGrpID) & _
            " AND b.cRecdStat = '1'" & _
          " ORDER BY sSalLvlID"
      
   Set poDetail = New Recordset
   poDetail.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set poDetail.ActiveConnection = Nothing
   
   Debug.Print lsSQL
   
   With MSFlexGrid1
      If poDetail.EOF Then
         .TextMatrix(1, 0) = "001"
         .TextMatrix(1, 1) = ""
         .TextMatrix(1, 2) = "0.00"
         .TextMatrix(1, 3) = "0.00"
      Else
         lnCtr = 1
         .Rows = poDetail.RecordCount + 1
         
         If .Rows > 10 Then
            .ColWidth(1) = 2850
         ElseIf .Rows < 11 Then
            InitGrid
         End If
         
         Do Until poDetail.EOF
            .TextMatrix(lnCtr, 0) = poDetail(0)
            .TextMatrix(lnCtr, 1) = poDetail(1)
            .TextMatrix(lnCtr, 2) = Format(IFNull(poDetail(2), 0), "#,##0.00")
            .TextMatrix(lnCtr, 3) = Format(IFNull(poDetail(3), 0), "#,##0.00")
            lnCtr = lnCtr + 1
            poDetail.MoveNext
         Loop
      End If
      
      For Each loTxt In txtDetail
         loTxt = .TextMatrix(1, loTxt.Index)
      Next
      
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", False
End Sub

Private Sub SaveDetail()
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "SaveDetail"
   On Error GoTo errProc
   
   'kalyptus - 2023.11.17 12:03pm
   'Add Field cSalGrpID to the Salary Matrix of the current Salary Matrix Group of each Salary Region
   
   With MSFlexGrid1
      lnCtr = 1
      poDetail.MoveFirst
      Do Until poDetail.EOF
         If IFNull(poDetail(4), "") = "" Then
            poDetail(4) = oDriver.FieldValue(0)
            poDetail(2) = CCur(.TextMatrix(lnCtr, 2))
            poDetail(3) = CCur(.TextMatrix(lnCtr, 3))
            'kalyptus - 2023.11.2023 12:05PM
            poDetail(5) = pcSalGrpID
            lsSQL = ADO2SQL(poDetail _
                          , "Salary_Level_Region" _
                          , _
                          , oApp.UserID _
                          , oApp.ServerDate _
                          , "sSalLvlNm")
         Else
            'kalyptus - 2023.11.2023 12:05PM
            lsSQL = "SELECT " & _
                            "  sSalRegID" & _
                            ", sSalLvlID" & _
                            ", nAmountxx" & _
                            ", nDEIGuide" & _
                            ", cSalGrpID" & _
                   " FROM Salary_Level_Region" & _
                   " WHERE sSalRegID = " & strParm(poDetail("sSalRegID")) & _
                     " AND sSalLvlID = " & strParm(poDetail("sSalLvlID")) & _
                     " AND cSalGrpID = " & strParm(pcSalGrpID)
                     
            Set lors = New Recordset
            lors.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
            Set lors.ActiveConnection = Nothing
            
            If lors.EOF Then
               lors.AddNew
               lors("sSalRegID") = poDetail("sSalRegID")
               lors("sSalLvlID") = poDetail("sSalLvlID")
               lors("nAmountxx") = CCur(.TextMatrix(lnCtr, 2))
               lors("nDEIGuide") = CCur(.TextMatrix(lnCtr, 3))
               'kalyptus - 2023.11.2023 12:05PM
               lors("cSalGrpID") = pcSalGrpID
               
               lsSQL = ADO2SQL(lors _
                             , "Salary_Level_Region" _
                             , _
                             , oApp.UserID _
                             , oApp.ServerDate)
            Else
               lors("nAmountxx") = CCur(.TextMatrix(lnCtr, 2))
               lors("nDEIGuide") = CCur(.TextMatrix(lnCtr, 3))
            
               'kalyptus - 2023.11.2023 12:05PM
               lsSQL = ADO2SQL(lors _
                             , "Salary_Level_Region" _
                             , " sSalRegID = " & strParm(poDetail("sSalRegID")) & _
                           " AND sSalLvlID = " & strParm(poDetail("sSalLvlID")) & _
                           " AND cSalGrpID = " & strParm(pcSalGrpID) _
                             , oApp.UserID _
                             , oApp.ServerDate)
            
            End If
         End If
         
         If lsSQL <> "" Then
            oApp.Execute lsSQL, "Salary_Level_Region"
         End If
         
         lnCtr = lnCtr + 1
         poDetail.MoveNext
      Loop
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", False
End Sub

Private Sub initFrame(nValue As Integer)
'   xrFrame1.Enabled = (nValue = 0)
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

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
   For Each loTxt In txtDetail
      loTxt = ""
   Next
End Sub

Private Sub createClearance()
   Dim lsSQL As String
   Dim lors As Recordset
   
   Dim loSec As clsPayMisc

   Set loSec = New clsPayMisc
   Set loSec.AppDriver = oApp

   'Decrypt salary level before anything else...
   lsSQL = "TRUNCATE tmpTable01"
   oApp.Connection.Execute lsSQL, , adCmdText
   
   lsSQL = "SELECT * FROM Employee_Master001 WHERE cRecdStat = '1'"
   Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
   
   MsgBox lors.RecordCount
   
   Do Until lors.EOF
      DoEvents
      If IFNull(lors("sSalLvlID")) <> "" Then
         lsSQL = "INSERT INTO tmpTable01" & _
                " SET sField001 = " & strParm(lors("sEmployID")) & _
                   ", sField002 = " & strParm(loSec.hxDecrypt(lors("sSalLvlID")))
         oApp.Connection.Execute lsSQL, , adCmdText
      End If
      
      DoEvents
      
      lors.MoveNext
   Loop
   
   Debug.Print "Done!!!"

End Sub

'09 -> Adjustment Code
'Kalyptus - 2013.10.29
Private Sub createAdjustment()
   Dim lsSQL As String
   Dim lors As Recordset
   Dim loTrans As clsEmployeeMovement
   Dim loSec As clsPayMisc
   Dim lsBasicPay As String
   
   Set loSec = New clsPayMisc
   Set loSec.AppDriver = oApp
   
   Dim ldEffectve As String
   Dim ldRequestd As String
   Dim lxTagRegID As String
   Dim lnRegAmtxx As Currency
       
   'Decrypt salary level before anything else...
   lsSQL = "TRUNCATE tmpTable01"
   oApp.Connection.Execute lsSQL, , adCmdText
   
   lsSQL = "SELECT * FROM Employee_Master001" & _
          " WHERE sSalRegID = " & strParm(oDriver.FieldValue(0)) & _
            " AND cRecdStat = '1'" & _
            " AND cSalTypex = 'R'" & _
            " AND sSalLvlID IS NOT NULL"
   Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
   
   Do Until lors.EOF
      lsSQL = "INSERT INTO tmpTable01" & _
             " SET sField001 = " & strParm(lors("sEmployID")) & _
                ", sField002 = " & strParm(loSec.hxDecrypt(lors("sSalLvlID")))
      oApp.Connection.Execute lsSQL, , adCmdText
      DoEvents
      lors.MoveNext
   Loop
   
   'Whoooh... we're ready to make the testing of employees
   'kalyptus - 2023.11.17 01:20pm
   'Load employees with a division belonging to the Salary Matrix Group
'            " AND f.cPayDivCd = " & strParm(oApp.Config("cPayDivCd"))
   
   lsSQL = "SELECT sLastName, sFrstName, sCompnyNm, a.*, c.nAmountxx" & _
          " FROM Employee_Master001 a" & _
              " LEFT JOIN tmpTable01 e ON a.sEmployID = e.sField001" & _
              " LEFT JOIN Client_Master b ON a.sEmployID = b.sClientID" & _
              " LEFT JOIN Branch_Others f ON a.sSgBranch = f.sBranchCD" & _
              " LEFT JOIN Salary_Level_Region c ON a.sSalRegID = c.sSalRegID AND e.sField002 = c.sSalLvlID" & _
                " AND c.cSalGrpID = " & strParm(pcSalGrpID) & _
          " WHERE a.sSalRegID = " & strParm(oDriver.FieldValue(0)) & _
            " AND a.cRecdStat = '1'" & _
            " AND a.cSalTypex = 'R'" & _
            " AND a.sSalLvlID IS NOT NULL" & _
            " AND a.sEmpLevID = '0'" & _
            " AND f.cDivision IN (SELECT sDivsnCde FROM Salary_Matrix_Group_Member WHERE cSalGrpID = " & strParm(pcSalGrpID) & ")"

            
   Debug.Print lsSQL
   Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
   
   Set loTrans = New clsEmployeeMovement
   Set loTrans.AppDriver = oApp
   loTrans.InitTransaction
   loTrans.HasParent = True
      
   oApp.BeginTrans
      
   'Kalyptus - 2018.09.22 04:21pm
   'Set the effective date here...
   If Not lors.EOF Then
      ldRequestd = "2024/01/24"
      ldEffectve = "2023/12/05"
      lxTagRegID = oDriver.FieldValue(3)
      lnRegAmtxx = oDriver.FieldValue(2)
   Else
      ldRequestd = ""
      ldEffectve = ""
      lxTagRegID = ""
      lnRegAmtxx = 0
   End If
   
   Do Until lors.EOF
      
      Debug.Print lors.AbsolutePosition & "»" & lors.RecordCount & "»" & lors("sCompnyNm")
      
      DoEvents
      
'      'And InStr(1, "M05411000386»M03307001237»M00113001208", loRS("sEmployID"), vbTextCompare) = 0
'      If IFNull(lors("nAmountxx"), 0) <> 0 Then
'         lsBasicPay = lors("nAmountxx")
'         If lors("nAmountxx") > loSec.Paway(lors("nBasicPay")) Then
'            loTrans.NewTransaction
'            loTrans.Master("sMoveCode") = "09"
'            loTrans.Master("sEmployID") = lors("sEmployID")
'            loTrans.Master("dTransact") = oApp.ServerDate
'            loTrans.Master("xBranchCD") = lors("sBranchCD")
'            loTrans.Master("nBasicOld") = lors("nBasicPay")
'            loTrans.Master("nBasicPay") = lsBasicPay
'            loTrans.Master("dRequestd") = ldRequestd
'            loTrans.Master("dEffectve") = ldEffectve
'            Call loTrans.SaveTransaction(False)
'            Call loTrans.CloseTransaction(loTrans.Master("sTransNox"))
'
'         End If
'      End If
'
      'KALYPTUS - 2018.09.22 04:15pm
      'Test TaxReg of Employee versus Tax Region of Salary Region
      If IFNull(lors("sTaxRegID"), "") = lxTagRegID Then
        If lnRegAmtxx <> 0 Then
           If lnRegAmtxx > IFNull(lors("nSalaryxx"), 0) Then
              lsSQL = "UPDATE Employee_Master001" & _
                     " SET nSalaryxx = " & lnRegAmtxx & _
                     " WHERE sEmployID = " & strParm(lors("sEmployID"))
              Call oApp.Execute(lsSQL, "Employee_Master001")
           End If
        End If
      End If
      
      lors.MoveNext
      DoEvents
   Loop
       
   'kalyptus - 2018.09.22 04:24pm
   'Check from Region_Wage if this is an update of minimum wage...
   If lxTagRegID <> "" Then
       'UPDATE WAGE_Movement history
       lsSQL = "SELECT *" & _
              " FROM Region_Wage" & _
              " WHERE sRegionID = " & strParm(lxTagRegID) & _
                " AND dDateThru IS NULL" & _
              " ORDER BY dDateFrom DESC LIMIT 1"
       Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
       If lors.EOF Then
           lsSQL = "INSERT INTO Region_Wage" & _
                  " SET sRegionID = " & strParm(lxTagRegID) & _
                     ", dDateFrom = " & strParm(ldEffectve) & _
                     ", dDateThru = NULL" & _
                     ", nMinWages = " & lnRegAmtxx / 26 & _
                     ", nColaAmtx = 0"
       Else
           'kalyptus - 2023.11.17 01:27pm
           'Replace the condition <If loRS("nMinWages") Then> since nMinWages represents daily pay
           '  and lnRegAmtxx represents the monthly pay
           'If loRS("nMinWages") Then
           If lors("nMinWages") = (lnRegAmtxx / 26) Then
               lsSQL = ""
           Else
                lsSQL = "UPDATE Region_Wage" & _
                       " SET dDateThru = " & dateParm(CDate(ldEffectve) - 1) & _
                       " WHERE sRegionID = " & strParm(lxTagRegID) & _
                         " AND dDateFrom = " & dateParm(lors("dDateFrom"))
                Call oApp.Execute(lsSQL, "Region_Wage")
                
                lsSQL = "INSERT INTO Region_Wage" & _
                       " SET sRegionID = " & strParm(lxTagRegID) & _
                          ", dDateFrom = " & strParm(ldEffectve) & _
                          ", dDateThru = NULL" & _
                          ", nMinWages = " & lnRegAmtxx / 26 & _
                          ", nColaAmtx = 0"
           
           End If
       End If
      If lsSQL <> "" Then
           Call oApp.Execute(lsSQL, "Region_Wage")
      End If
      
      'UPDATE Tax Region
       lsSQL = "SELECT *" & _
              " FROM Region" & _
              " WHERE sRegionID = " & strParm(lxTagRegID)
       Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
       If lors.EOF Then
            lsSQL = ""
       Else
           If lors("nMinWages") = lnRegAmtxx Then
               lsSQL = ""
           Else
                lsSQL = "UPDATE Region" & _
                       " SET nMinWages = " & lnRegAmtxx / 26 & _
                          ", nMinWage2 = " & lnRegAmtxx / 26 & _
                       " WHERE sRegionID = " & strParm(lxTagRegID)
            End If
       End If
   
      If lsSQL <> "" Then
           Call oApp.Execute(lsSQL, "Region_Wage")
      End If
   
   End If
   
   oApp.CommitTrans

   MsgBox "Tapos na po"
End Sub

