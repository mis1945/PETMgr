VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMealOjtSchool 
   BorderStyle     =   0  'None
   Caption         =   "School"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3270
      Left            =   135
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   5768
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1815
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   300
         Width           =   1425
      End
      Begin VB.Frame Frame1 
         Caption         =   "School Details:"
         Height          =   2370
         Left            =   180
         TabIndex        =   8
         Tag             =   "wt0;fb0"
         Top             =   720
         Width           =   5295
         Begin VB.CheckBox chk 
            Caption         =   "College"
            Height          =   195
            Index           =   2
            Left            =   3735
            TabIndex        =   6
            Tag             =   "wt0;fb0"
            Top             =   1740
            Width           =   1155
         End
         Begin VB.CheckBox chk 
            Caption         =   "University"
            Height          =   195
            Index           =   3
            Left            =   3735
            TabIndex        =   7
            Tag             =   "wt0;fb0"
            Top             =   1980
            Width           =   1200
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1620
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1305
            Width           =   3300
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   690
            Index           =   2
            Left            =   1620
            MultiLine       =   -1  'True
            TabIndex        =   2
            Text            =   "frmMealOjtSchool.frx":0000
            Top             =   585
            Width           =   3300
         End
         Begin VB.CheckBox chk 
            Caption         =   "Senior High School"
            Height          =   195
            Index           =   0
            Left            =   1605
            TabIndex        =   4
            Tag             =   "wt0;fb0"
            Top             =   1725
            Width           =   2145
         End
         Begin VB.CheckBox chk 
            Caption         =   "Technical"
            Height          =   195
            Index           =   1
            Left            =   1605
            TabIndex        =   5
            Tag             =   "wt0;fb0"
            Top             =   1965
            Width           =   1200
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1620
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   255
            Width           =   3300
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Town/City:"
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   11
            Top             =   1395
            Width           =   1245
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   10
            Top             =   630
            Width           =   1485
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "School Name:"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   9
            Top             =   345
            Width           =   1245
         End
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1890
         Tag             =   "et0;ht2"
         Top             =   405
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "School ID:"
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
         Left            =   360
         TabIndex        =   12
         Top             =   345
         Width           =   1095
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   5055
      TabIndex        =   13
      Top             =   4050
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
      Picture         =   "frmMealOjtSchool.frx":0006
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   4275
      TabIndex        =   14
      Top             =   4050
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
      Picture         =   "frmMealOjtSchool.frx":0780
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   3495
      TabIndex        =   15
      Top             =   4050
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
      Picture         =   "frmMealOjtSchool.frx":0EFA
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   3495
      TabIndex        =   16
      Top             =   4050
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
      Picture         =   "frmMealOjtSchool.frx":1674
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   1935
      TabIndex        =   17
      Top             =   4050
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
      Picture         =   "frmMealOjtSchool.frx":1DEE
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   5055
      TabIndex        =   18
      Top             =   4050
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
      Picture         =   "frmMealOjtSchool.frx":2568
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   2715
      TabIndex        =   19
      Top             =   4050
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
      Picture         =   "frmMealOjtSchool.frx":2CE2
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmMealOjtSchool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmMealOjtSchool"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean
Dim psValue As String

Dim pnIndex As Integer

Private Sub chk_Click(Index As Integer)
   Dim lsOldProc As String
   lsOldProc = "chk_Click"

   Select Case Index
      Case 0
         oDriver.FieldValue(4) = chk(0).Value
      Case 1
         oDriver.FieldValue(5) = chk(1).Value
      Case 2
         oDriver.FieldValue(6) = chk(2).Value
      Case 3
         oDriver.FieldValue(7) = chk(3).Value
   End Select
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc
   
   Select Case Index
   Case 0
      oDriver.RecordCancelUpdate
      ClearFields
      initField
   Case 1
    oDriver.BrowseRecord
    initField 1
   Case 2
      If oDriver.RecordSave Then
         MsgBox "Successfully save!"
         initField
      Else
         MsgBox "Unable to save record/update!"
      End If
   Case 3
      If txtField(0).Text = "" Then
         MsgBox "Please select record to update!"
      Else
         oDriver.RecordUpdate
      If txtField(0).Text <> "" Then initField 1
      End If
   Case 4
      oDriver.RecordNew
      initField
   Case 5
      Unload Me
   Case 6
      If txtField(0).Text = "" Or txtField(1).Text = "" Then
         MsgBox "Please select record to delete!"
      Else
         If oDriver.RecordDelete Then
            MsgBox "Record successfully deleted!"
            initField
         Else
            MsgBox "Unable to delete record!"
         End If
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub ClearFields()
Dim loTxt As TextBox

For Each loTxt In txtField
   loTxt = ""
Next

End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   'On Error GoTo errProc
   
    oApp.MenuName = Me.Tag
    Me.ZOrder 0
    
   If bLoaded = False Then
      oDriver.RecordNew
      initField
      oDriver.DisableTextbox 0
      bLoaded = True
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   'On Error GoTo errProc
   
   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   oDriver.RecQuery = "SELECT" _
                           & "  sSchoolID" _
                           & ", sSchoolNm" _
                           & ", sAddressx" _
                           & ", sTownIDxx" _
                           & ", cSHSxxxxx" _
                           & ", cTechnicl" _
                           & ", cCollegex" _
                           & ", cUnivrsty" _
                           & ", cRecdStat" _
                           & ", sModified" _
                           & ", dModified" _
                        & " FROM School"
                        
   oDriver.BrowseQuery = "SELECT" _
                           & "  sSchoolID" _
                           & ", sSchoolNm" _
                        & " FROM School" _
                        & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY sSchoolNm"
                        
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Code"
   oDriver.BrowseFTitle(1) = "Description"
                        
   oDriver.LookupQuery(3) = "SELECT" _
                           & "  sTownIDxx" _
                           & ", sTownName" _
                           & ", sZippCode" _
                        & " FROM TownCity" _
                        & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY sTownName"
   
   oDriver.LookupReference(3) = "sTownIDxx»sTownName»sZippCode"
   oDriver.LookupColumn(3) = "sTownIDxx»sTownName»sZippCode"
   oDriver.LookupTitle(3) = "sTownIDxx»sTownName»sZippCode"

   oDriver.FieldStart = 1
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   chk(0).Enabled = False
   chk(1).Enabled = False
   chk(2).Enabled = False
   chk(3).Enabled = False
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   chk(0).Enabled = True
   chk(1).Enabled = True
   chk(2).Enabled = True
   chk(3).Enabled = True
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   Dim lnCtr As Integer
   
   lsOldProc = "oDriver_InitValue"
   ''On Error GoTo errProc
   
   With oDriver
      For lnCtr = 0 To 8
         Select Case lnCtr
         Case 0
            If .SetValue(lnCtr, GetNextCode("School", "sSchoolID", False, oApp.Connection)) = False Then Exit Sub
         Case 1 To 3
            oDriver.FieldValue(lnCtr) = ""
         Case 4 To 7
            .FieldValue(lnCtr) = 0
         Case 8
            oDriver.FieldValue(lnCtr) = xeRecStateActive
         End Select
      Next
   End With
   
   oDriver.FieldReference(0) = True
   psValue = ""

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_LoadOtherData()
   psValue = txtField(3).Text
End Sub

'Private Sub oDriver_InitValue()
'   Dim lsOldProc As String
'
'   lsOldProc = "oDriver_InitValue"
'   'On Error GoTo errProc
'
'   If oDriver.SetValue(0, GetNextCode("School", "sSchoolID", False, oApp.Connection)) = False Then Exit Sub
'   oDriver.FieldReference(0) = True
'   oDriver.FieldValue(1) = txtField(1)
'   oDriver.FieldValue(2) = txtField(2)
'   oDriver.FieldValue(3) = txtField(3)
'   chk(0).Value = IIf(IFNull(oDriver.FieldValue(4), "0") = "0", vbUnchecked, vbChecked)
'   chk(1).Value = IIf(IFNull(oDriver.FieldValue(5), "0") = "0", vbUnchecked, vbChecked)
'   chk(2).Value = IIf(IFNull(oDriver.FieldValue(6), "0") = "0", vbUnchecked, vbChecked)
'   chk(3).Value = IIf(IFNull(oDriver.FieldValue(7), "0") = "0", vbUnchecked, vbChecked)
'   oDriver.FieldValue(8) = xeRecStateActive
'   psValue = ""
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )"
'End Sub


Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid School Name detected!!", vbCritical, "Warning"
      Cancel = True
      txtField(1).SetFocus
   End If

   If oDriver.FieldValue(2) = "" Then
      MsgBox "Invalid School Address detected!!", vbCritical, "Warning"
      Cancel = True
      txtField(2).SetFocus
   End If
   
   If oDriver.FieldValue(3) = "" Then
      MsgBox "Invalid Town/City detected!!", vbCritical, "Warning"
      Cancel = True
      txtField(3).SetFocus
   End If
   
'      If oDriver.FieldValue(3) = "" Then
'      SearchTown
'      End If
'
'   End If
   
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   oDriver.ColumnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   Dim lsLookup As String

   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(3)
         If KeyCode = vbKeyF3 Then
               oDriver.RecordSearch .Text
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then
                  oDriver.RecordSearch .Text
            End If
         End If
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   Dim lsLookup As String
   
   lsOldProc = "txtField_Validate"
   ''On Error GoTo errProc
   
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 3
         Cancel = Not oDriver.ValidateField(Index)
         
         If Not Cancel Then
            If txtField(3).Text <> "" Then SearchTown
         End If
         psValue = .Text
      Case Else
         Cancel = Not oDriver.ValidateField(Index)
         oDriver.FieldValue(Index) = .Text
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True

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

'Private Sub initField(lnStat As Integer, Optional lnEnabled As Integer = 0)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   If lbShow Then
'      txtField(1).Text = oDriver.FieldValue(1)
'      txtField(2).Text = oDriver.FieldValue(2)
'      txtField(3).Text = SearchTown
'      chk(0).Value = IIf(IFNull(oDriver.FieldValue(4), "0") = "0", vbUnchecked, vbChecked)
'      chk(1).Value = IIf(IFNull(oDriver.FieldValue(5), "0") = "0", vbUnchecked, vbChecked)
'      chk(2).Value = IIf(IFNull(oDriver.FieldValue(6), "0") = "0", vbUnchecked, vbChecked)
'      chk(3).Value = IIf(IFNull(oDriver.FieldValue(7), "0") = "0", vbUnchecked, vbChecked)
'   Else
'      txtField(1).Text = ""
'      txtField(2).Text = ""
'      txtField(3).Text = ""
'      chk(0).Value = vbUnchecked
'      chk(1).Value = vbUnchecked
'      chk(2).Value = vbUnchecked
'      chk(3).Value = vbUnchecked
'   End If
'
'   txtField(1).Enabled = (lnEnabled = 0)
'   txtField(2).Enabled = (lnEnabled = 0)
'   txtField(3).Enabled = (lnEnabled = 0)
'   chk(0).Enabled = (lnEnabled = 0)
'   chk(1).Enabled = (lnEnabled = 0)
'   chk(2).Enabled = (lnEnabled = 0)
'   chk(3).Enabled = (lnEnabled = 0)
'End Sub
Private Sub initField(Optional lnStat As Integer = 0)
    Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, True, False)
   If lbShow Then
         txtField(1).Text = ""
         txtField(2).Text = ""
         txtField(3).Text = ""
         chk(0).Value = vbUnchecked
         chk(1).Value = vbUnchecked
         chk(2).Value = vbUnchecked
         chk(3).Value = vbUnchecked
      Else
         txtField(3).Text = SearchTown
         chk(0).Value = IIf(IFNull(oDriver.FieldValue(4), "0") = "0", vbUnchecked, vbChecked)
         chk(1).Value = IIf(IFNull(oDriver.FieldValue(5), "0") = "0", vbUnchecked, vbChecked)
         chk(2).Value = IIf(IFNull(oDriver.FieldValue(6), "0") = "0", vbUnchecked, vbChecked)
         chk(3).Value = IIf(IFNull(oDriver.FieldValue(7), "0") = "0", vbUnchecked, vbChecked)
      End If
End Sub

Function SearchTown()
   Dim lsOldProc As String
   Dim lors As ADODB.Recordset
   Dim lsSQL As String
   Dim lsTownName As String
   
   lsOldProc = "SearchTown"
   ''On Error GoTo errProc
   
   If txtField(3).Text = "" Then Exit Function
   
   lsSQL = "SELECT sTownIDxx" _
               & ", sTownName" _
               & ", sZippCode" _
               & ", cRecdStat" _
            & " FROM TownCity" _
            & " WHERE sTownIDxx = " & strParm(oDriver.FieldValue(3)) _
            & " ORDER BY sTownIDxx"
   
   Set lors = New ADODB.Recordset
   With lors
      .Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
         If Not .EOF() Then
            lsTownName = lors("sTownName")
         Else
            GoTo endProc
         End If
         
   End With
   
   SearchTown = lsTownName
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & oDriver.FieldValue(2) & "»" & oDriver.FieldValue(3) & " )", True
End Function
