VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmDepartmentneo 
   BorderStyle     =   0  'None
   Caption         =   "Departments"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   5295
      TabIndex        =   7
      Top             =   3930
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmDepartmentneo.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   4515
      TabIndex        =   6
      Top             =   3930
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmDepartmentneo.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   3735
      TabIndex        =   4
      Top             =   3930
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Update"
      AccessKey       =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDepartmentneo.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   2955
      TabIndex        =   0
      Top             =   3930
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&New"
      AccessKey       =   "N"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDepartmentneo.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   6075
      TabIndex        =   8
      Top             =   3930
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmDepartmentneo.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   4515
      TabIndex        =   5
      Top             =   3930
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Delete"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDepartmentneo.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3225
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   5689
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.CheckBox chkActive 
         Caption         =   "isActive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5295
         TabIndex        =   15
         Top             =   2820
         Width           =   1080
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
         Index           =   3
         Left            =   1590
         TabIndex        =   1
         Top             =   1950
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
         Index           =   4
         Left            =   1590
         TabIndex        =   3
         Top             =   1020
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
         Index           =   1
         Left            =   1590
         TabIndex        =   2
         Top             =   1485
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
         TabIndex        =   10
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Division:"
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
         TabIndex        =   14
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Dept:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1125
         Width           =   945
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
         Caption         =   "Dept Name:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dept. ID"
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
         TabIndex        =   11
         Top             =   210
         Width           =   750
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   6075
      TabIndex        =   9
      Top             =   3930
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmDepartmentneo.frx":2CDC
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmDepartmentneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmDepartmentneo"
Private Const pxeDeptID As String = "000"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnCtr As Integer
Dim pnIndex As Integer

Private Sub chkActive_Click()
   If chkActive.Value = True Then
      oDriver.FieldValue(5) = 1
   Else
      oDriver.FieldValue(5) = 0
   End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc
   
   Select Case Index
   Case 0
      oDriver.RecordCancelUpdate
   Case 1
      oDriver.BrowseRecord
   Case 2
      If chkActive.Value = True Then
         oDriver.FieldValue(5) = 1
      Else
         oDriver.FieldValue(5) = 0
      End If
      oDriver.RecordSave
   Case 3
      oDriver.RecordUpdate
   Case 4
      oDriver.RecordNew
   Case 5
      Unload Me
   Case 6
      oDriver.RecordDelete
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
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsSQL As String
   Dim lsOldProc As String
   
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
   
   With oDriver
      .RecQuery = "SELECT" _
                     & "  sDeptIDxx" _
                     & ", sDeptName" _
                     & ", sDeptCode" _
                     & ", sDivsnCde" _
                     & ", sMainIDxx" _
                     & ", cRecdStat" _
                     & ", sModified" _
                     & ", dModified" _
                  & " FROM Department "

      .BrowseQuery = "SELECT" _
                        & "  sDeptIDxx" _
                        & ", sDeptName" _
                        & ", sDeptCode" _
                     & " FROM Department" _
                     & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY sDeptIDxx"
      .InitRecForm
   
      .BrowseFTitle(0) = "Code"
      .BrowseFTitle(1) = "Name"
      .BrowseFFormat(0) = "@@@"
   

      
      .LookupQuery(3) = "SELECT sDivsnCde, sDivsnDsc" & _
                                 " FROM Division" & _
                                 " WHERE cRecdStat = '1'"
    
      .LookupReference(3) = "sDivsnCde»sDivsnDsc"
      .LookupColumn(3) = "sDivsnDsc"
      .LookupTitle(3) = "Desc"
      
      .LookupQuery(4) = "SELECT sDeptIDxx, sDeptName" & _
                                 " FROM Department" & _
                                 " WHERE cRecdStat = '1'" & _
                                 " AND sDeptIDxx IN(SELECT sMainIDxx FROM Department)"
    
      .LookupReference(4) = "sDeptIDxx»sDeptName"
      .LookupColumn(4) = "sDeptName"
      .LookupTitle(4) = "Main Dept"
   
'      .LookupQuery(1) = "SELECT sDeptIDxx, sDeptName" & _
'                                 " FROM Department" & _
'                                 " WHERE cRecdStat = '1'"
'
'      .LookupReference(1) = "sDeptIDxx»sDeptName"
'      .LookupColumn(1) = "sDeptName"
'      .LookupTitle(1) = "Dept"

      .FieldFormat(0) = "@@@@"
      .FieldSize(0) = Len(.FieldFormat(0))
      .FieldStart = 1
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oDriver = Nothing
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lsDeptID As String
   
   lsOldProc = "oDriver_InitValue"
   On Error GoTo errProc
   
   lsSQL = "SELECT * FROM Department Where sDeptIDxx NOT LIKE 'A%' Order BY sDeptIDxx DESC LIMIT 1"
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, , , adCmdText
   
   lsDeptID = lors("sDeptIDxx") + 1
   
   If Not oDriver.SetValue(0, Format(lsDeptID, pxeDeptID)) Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(1) = ""
   oDriver.FieldValue(2) = ""
   oDriver.FieldValue(3) = ""
   oDriver.FieldValue(4) = ""
   oDriver.FieldValue(5) = xeRecStateActive
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsSQL As String
   Dim lors As Recordset
   
   lsSQL = "SELECT a.sDeptName, b.sDivsnDsc" & _
            " FROM Department a" & _
               " LEFT JOIN Division b ON a.sDivsnCde = b.sDivsnCde" & _
            " WHERE  a.sMainIDxx = " & strParm(Ifnull(oDriver.FieldValue(3), ""))
   Debug.Print lsSQL
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, , , adCmdText
   
   If Not lors.EOF Then
      txtField(1).Text = Ifnull(lors("sDeptName"), "")
      txtField(3).Text = Ifnull(lors("sDivsnDsc"), "")
   Else
      txtField(1).Text = ""
      txtField(3).Text = ""
   End If
   chkActive.Value = oDriver.FieldValue(5)
   
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Description detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   End If

   oDriver.FieldValue(5) = chkActive.Value

End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
      
   oDriver.ColumnIndex = Index
   pnIndex = Index
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

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(pnIndex)
         If KeyCode = vbKeyF3 Then
            oDriver.RecordSearch .Text
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then oDriver.RecordSearch .Text
         End If
      End With
      KeyCode = 0
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_LostFocus"
   On Error GoTo errProc
    Select Case Index
      Case Else
         txtField(Index).Text = TitleCase(txtField(Index).Text)
         Cancel = Not oDriver.ValidateField(Index)
      End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
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

