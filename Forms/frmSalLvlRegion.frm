VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSalLvlRegion 
   BorderStyle     =   0  'None
   Caption         =   "Salary Level Regions"
   ClientHeight    =   4665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2835
      Left            =   255
      Tag             =   "wt0;fb0"
      Top             =   660
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   5001
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   1785
         TabIndex        =   7
         Top             =   2055
         Width           =   2415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   1785
         TabIndex        =   5
         Top             =   1605
         Width           =   2415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   1785
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   1785
         TabIndex        =   1
         Top             =   315
         Width           =   2415
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEI Guide"
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
         Index           =   3
         Left            =   330
         TabIndex        =   6
         Top             =   2145
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Index           =   2
         Left            =   330
         TabIndex        =   4
         Top             =   1695
         Width           =   675
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Index           =   1
         Left            =   1890
         Tag             =   "et0;ht2"
         Top             =   1065
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Index           =   0
         Left            =   1890
         Tag             =   "et0;ht2"
         Top             =   405
         Width           =   2415
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sal. Lvl. ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   330
         TabIndex        =   2
         Top             =   1050
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sal. Reg, ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   0
         Top             =   405
         Width           =   1260
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   6480
      TabIndex        =   8
      Top             =   3855
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
      Picture         =   "frmSalLvlRegion.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   5700
      TabIndex        =   9
      Top             =   3855
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
      Picture         =   "frmSalLvlRegion.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   4920
      TabIndex        =   10
      Top             =   3855
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
      Picture         =   "frmSalLvlRegion.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   4140
      TabIndex        =   11
      Top             =   3855
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
      Picture         =   "frmSalLvlRegion.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   3360
      TabIndex        =   12
      Top             =   3855
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
      Picture         =   "frmSalLvlRegion.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   6480
      TabIndex        =   13
      Top             =   3855
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
      Picture         =   "frmSalLvlRegion.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   4920
      TabIndex        =   14
      Top             =   3855
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
      Picture         =   "frmSalLvlRegion.frx":2CDC
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmSalLvlRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmSalLvlRegion"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim psSelected() As String
Dim pnCtr As Integer
Dim pnIndex As Integer

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
      If validateData Then oDriver.RecordSave
   Case 3
      oDriver.RecordUpdate
   Case 4
      oDriver.RecordNew
      txtField(0).Enabled = True
      txtField(0).SetFocus
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
      txtField(0).Enabled = True
      txtField(0).SetFocus
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
                     & "  sSalRegID" _
                     & ", sSalLvlID" _
                     & ", nAmountxx" _
                     & ", nDEIGuide" _
                     & ", sModified" _
                     & ", dModified" _
                  & " FROM Salary_Level_Region"

           
      .BrowseQuery = "SELECT" _
                        & "  sSalRegID" _
                        & ", sSalLvlID" _
                        & ", nAmountxx" _
                        & ", nDEIGuide" _
                     & " FROM Salary_Level_Region" _
                     & " ORDER BY sSalRegID"
                     
      .InitRecForm
   
      .BrowseFTitle(0) = "Sal. Reg. ID"
      .BrowseFTitle(1) = "Sal. Lvl. ID"
      .BrowseFTitle(2) = "Amount"
      .BrowseFTitle(3) = "DEI Guide"
      .BrowseFFormat(0) = "@@@"
   
      .FieldFormat(0) = "@@@"
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
   
   lsOldProc = "oDriver_InitValue"
   On Error GoTo errProc

   oDriver.FieldValue(0) = ""
   oDriver.FieldValue(1) = ""
   oDriver.FieldValue(2) = ""
   oDriver.FieldValue(3) = ""
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Description detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
   End With
   If Index = 2 Or Index = 3 Then txtField(Index).Text = ""
   
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
   Select Case KeyCode
      Case vbKeyF3
         If Index = 0 Then
            Call searchSalaryRegion
         ElseIf Index = 1 Then
            Call searchSalaryLevel
         Else: Exit Sub
         End If
   End Select
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKey0 To vbKey9
      Case vbKeyF3, 46 'period
      Case Else
         KeyAscii = 0
         Beep
   End Select
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
      
   txtField(Index).Text = TitleCase(txtField(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub searchSalaryLevel(Optional SalLvl As Variant)
   Dim lsBrowse As String
   Dim lsSQL As String
   Dim lrs As ADODB.Recordset
   Dim lsOldProc As String

   lsOldProc = "searchSalaryLevel"
   On Error GoTo errProc
   lsSQL = "SELECT" _
               & "  sSalLvlID" _
               & ", sSalLvlNm" _
            & " FROM Salary_Level" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive)

  If Not IsMissing(SalLvl) Then lsSQL = lsSQL & "AND sSalLvlNm LIKE " & strParm(SalLvl & "%")

  Set lrs = New ADODB.Recordset

  If lrs.State = adStateOpen Then lrs.Close
  lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If Not lrs.EOF Then
      lsBrowse = KwikBrowse(oApp, lrs _
                  , "sSalLvlID»sSalLvlNm" _
                  , "Sal. Lvl. ID»Description")

      If lsBrowse <> "" Then
         psSelected = Split(lsBrowse, "»")
         oDriver.LookupValue(0) = psSelected(0)
         txtField(1).Text = lrs.Fields("sSalLvlID")
      End If
   End If
   
   txtField(1).Tag = txtField(1).Text
   Set lrs = Nothing

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & IFNull(SalLvl) & " )"
End Sub
Private Sub searchSalaryRegion(Optional SalRegion As Variant)
   Dim lsBrowse As String
   Dim lsSQL As String
   Dim lrs As ADODB.Recordset
   Dim lsOldProc As String

   lsOldProc = "searchSalaryRegion"
   On Error GoTo errProc
   lsSQL = "SELECT" _
               & "  sSalRegID" _
               & ", sSalRegNm" _
            & " FROM Salary_Region" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive)

  If Not IsMissing(SalRegion) Then lsSQL = lsSQL & "AND sSalLvlNm LIKE " & strParm(SalRegion & "%")

  Set lrs = New ADODB.Recordset

  If lrs.State = adStateOpen Then lrs.Close
  lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If Not lrs.EOF Then
      lsBrowse = KwikBrowse(oApp, lrs _
                  , "sSalRegID»sSalRegNm" _
                  , "Sal. Region. ID»Description")

      If lsBrowse <> "" Then
         psSelected = Split(lsBrowse, "»")
         oDriver.LookupValue(0) = psSelected(0)
         txtField(0).Text = lrs.Fields("sSalRegID")
      End If
   End If
   
   txtField(0).Tag = txtField(0).Text
   Set lrs = Nothing

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & IFNull(SalRegion) & " )"
End Sub
Private Function validateData() As Boolean
   If txtField(0).Text = "" Or _
      txtField(1).Text = "" Or _
      txtField(2).Text = "" Or _
      txtField(3).Text = "" Then
      
      MsgBox "Blank fields detected!!!", vbCritical, "Warning"
      txtField(0).SetFocus
      validateData = False
   Else
      validateData = True
   End If
End Function
