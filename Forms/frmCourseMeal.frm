VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMealOjtCourse 
   BorderStyle     =   0  'None
   Caption         =   "Course"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2820
      Left            =   135
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   4974
      BorderStyle     =   1
      Begin VB.Frame Frame1 
         Caption         =   "Course Details:"
         Height          =   1920
         Left            =   120
         TabIndex        =   4
         Tag             =   "wt0;fb0"
         Top             =   750
         Width           =   5295
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1635
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   390
            Width           =   2760
         End
         Begin VB.CheckBox chk 
            Caption         =   "College"
            Height          =   195
            Index           =   1
            Left            =   1620
            TabIndex        =   3
            Tag             =   "wt0;fb0"
            Top             =   1605
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "Technical"
            Height          =   195
            Index           =   0
            Left            =   1620
            TabIndex        =   2
            Tag             =   "wt0;fb0"
            Top             =   1350
            Width           =   1155
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1620
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   825
            Width           =   2760
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Course Name:"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   15
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Course Code:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   5
            Top             =   885
            Width           =   1485
         End
      End
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
         Left            =   1740
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Course ID:"
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
         Left            =   285
         TabIndex        =   6
         Top             =   345
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1815
         Tag             =   "et0;ht2"
         Top             =   405
         Width           =   1425
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   5025
      TabIndex        =   7
      Top             =   3600
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
      Picture         =   "frmCourseMeal.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   4245
      TabIndex        =   8
      Top             =   3600
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
      Picture         =   "frmCourseMeal.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   3465
      TabIndex        =   9
      Top             =   3600
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
      Picture         =   "frmCourseMeal.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   3465
      TabIndex        =   10
      Top             =   3600
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
      Picture         =   "frmCourseMeal.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   1905
      TabIndex        =   11
      Top             =   3600
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
      Picture         =   "frmCourseMeal.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   5025
      TabIndex        =   12
      Top             =   3600
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
      Picture         =   "frmCourseMeal.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   2685
      TabIndex        =   13
      Top             =   3600
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
      Picture         =   "frmCourseMeal.frx":2CDC
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmMealOjtCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmMealOjtCourse"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer

Private Sub chk_Click(Index As Integer)
   Select Case Index
    Case 0
      oDriver.FieldValue(3) = chk(0).Value
    Case 1
      oDriver.FieldValue(4) = chk(1).Value
   End Select
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
         initField 0
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
            initField 0
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
      initField 0
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
                           & "  sCourseID" _
                           & ", sCourseNm" _
                           & ", sCourseCD" _
                           & ", cTechnicl" _
                           & ", cCollegex" _
                           & ", cRecdStat" _
                           & ", sModified" _
                           & ", dModified" _
                        & " FROM Course"
   oDriver.BrowseQuery = "SELECT" _
                           & "  sCourseID" _
                           & ", sCourseNm" _
                        & " FROM Course" _
                        & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY sCourseNm"
   
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Code"
   oDriver.BrowseFTitle(1) = "Description"

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
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   chk(0).Enabled = True
   chk(1).Enabled = True
   
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   'On Error GoTo errProc
   
   If oDriver.SetValue(0, GetNextCode("Course", "sCourseID", False, oApp.Connection)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(1) = txtField(1)
   oDriver.FieldValue(2) = txtField(2)
   chk(0).Value = IIf(IFNull(oDriver.FieldValue(3), "0") = "0", vbUnchecked, vbChecked)
   chk(1).Value = IIf(IFNull(oDriver.FieldValue(4), "0") = "0", vbUnchecked, vbChecked)
   oDriver.FieldValue(5) = xeRecStateActive

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Course Description detected!!", vbCritical, "Warning"
      Cancel = True
   End If

   If oDriver.FieldValue(2) = "" Then
      MsgBox "Invalid Course Code detected!!", vbCritical, "Warning"
      Cancel = True
   End If
   
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
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
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "txtField_Validate"
   'On Error GoTo errProc
   With txtField(Index)
      Select Case Index
         Case 1
            .Text = TitleCase(.Text)
            Cancel = Not oDriver.ValidateField(Index)
         Case 2
           .Text = TitleCase(.Text)
            oDriver.FieldValue(Index) = .Text
      End Select
   End With
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
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

Private Sub initField(Optional lnStat As Integer = 0)
   Dim lbShow As Boolean
   lbShow = IIf(lnStat = 0, True, False)
      If lbShow Then
         txtField(1).Text = ""
         txtField(2).Text = ""
         chk(0).Value = vbUnchecked
         chk(1).Value = vbUnchecked
      Else
         txtField(1).Text = oDriver.FieldValue(1)
         txtField(2).Text = oDriver.FieldValue(2)
         chk(0).Value = IIf(IFNull(oDriver.FieldValue(3), "0") = "0", vbUnchecked, vbChecked)
         chk(1).Value = IIf(IFNull(oDriver.FieldValue(4), "0") = "0", vbUnchecked, vbChecked)
      End If
End Sub

'Private Sub initField(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   If lbShow Then
'      txtField(1).Text = oDriver.FieldValue(1)
'      txtField(2).Text = oDriver.FieldValue(2)
'      chk(0).Value = IIf(IFNull(oDriver.FieldValue(3), "0") = "0", vbUnchecked, vbChecked)
'      chk(1).Value = IIf(IFNull(oDriver.FieldValue(4), "0") = "0", vbUnchecked, vbChecked)
'   Else
'      txtField(1).Text = ""
'      txtField(2).Text = ""
'      chk(0).Value = vbUnchecked
'      chk(1).Value = vbUnchecked
'   End If
'End Sub


