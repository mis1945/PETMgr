VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMealType 
   BorderStyle     =   0  'None
   Caption         =   "Meal Type"
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
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   600
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   4974
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1695
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1920
         Width           =   2760
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1695
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1575
         Width           =   2760
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1695
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   750
         Width           =   2760
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
         Left            =   1695
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   315
         Width           =   1425
      End
      Begin VB.Frame Frame1 
         Caption         =   "Time"
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Tag             =   "wt0;fb0"
         Top             =   1200
         Width           =   5295
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Coverage Thru:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Coverage From:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1245
         End
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1755
         Tag             =   "et0;ht2"
         Top             =   405
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Meal Description:"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   5
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Meal ID:"
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
         Left            =   210
         TabIndex        =   1
         Top             =   360
         Width           =   810
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   5040
      TabIndex        =   9
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
      Picture         =   "frmMealType.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   4260
      TabIndex        =   10
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
      Picture         =   "frmMealType.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   3480
      TabIndex        =   11
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
      Picture         =   "frmMealType.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   3480
      TabIndex        =   12
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
      Picture         =   "frmMealType.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   1920
      TabIndex        =   13
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
      Picture         =   "frmMealType.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   5040
      TabIndex        =   14
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
      Picture         =   "frmMealType.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   2700
      TabIndex        =   15
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
      Picture         =   "frmMealType.frx":2CDC
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmMealType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmMealType"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc
   
   Select Case Index
   Case 0
      oDriver.RecordCancelUpdate
      ClearFields
      initField 0, 1
   Case 1
    oDriver.BrowseRecord
    initField 1, 1
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
      If txtField(0).Text <> "" Then initField 1, 0
      End If
   Case 4
      oDriver.RecordNew
      initField 0
   Case 5
      Unload Me
   Case 6
      If txtField(0).Text = "" Then
      MsgBox "Please select record to delete!"
      Else
      oDriver.RecordDelete
      initField 0, 1
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
                           & "  sMealCode" _
                           & ", sMealDesc" _
                           & ", tCvrgFrom" _
                           & ", tCvrgThru" _
                           & ", cRecdStat" _
                           & ", sModified" _
                           & ", dModified" _
                        & " FROM Meal_Type"
   oDriver.BrowseQuery = "SELECT" _
                           & "  sMealCode" _
                           & ", sMealDesc" _
                        & " FROM Meal_Type" _
                        & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY sMealDesc"
   
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

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   'On Error GoTo errProc
   
'   If oDriver.SetValue(0, GetNextCode("Meal_Type", "sMealCode", False, oApp.Connection)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(0) = "UN"
   oDriver.FieldValue(1) = txtField(1)
   oDriver.FieldValue(2) = Format(txtField(2), "hh:mm:ss AM/PM")
   oDriver.FieldValue(3) = Format(txtField(3), "hh:mm:ss AM/PM")
   oDriver.FieldValue(4) = xeRecStateActive

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Empty Description detected!!!", vbCritical, "Warning"
      Cancel = True
   End If

   If oDriver.FieldValue(2) = "" Then
      MsgBox "Invalid Time Coverage From detected!!!", vbCritical, "Warning"
      Cancel = True
   End If

   If oDriver.FieldValue(3) = "" Then
      MsgBox "Invalid Time Coverage Thru detected!!!!!!", vbCritical, "Warning"
      Cancel = True
   End If
   
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
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
           If Not IsDate(.Text) Then .Text = oApp.ServerDate
               .Text = Format(.Text, "hh:mm:ss AM/PM")
            oDriver.FieldValue(Index) = .Text
         Case 3
            If Not IsDate(.Text) Then .Text = oApp.ServerDate
               .Text = Format(.Text, "hh:mm:ss AM/PM")
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

Private Sub initField(lnStat As Integer, Optional lnEnabled As Integer = 0)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   If lbShow Then
      txtField(1).Text = oDriver.FieldValue(1)
      txtField(2).Text = Format(oDriver.FieldValue(2), "hh:mm:ss AM/PM")
      txtField(3).Text = Format(oDriver.FieldValue(3), "hh:mm:ss AM/PM")
   Else
      txtField(1).Text = ""
      txtField(2).Text = Format(oApp.ServerDate, "hh:mm:ss AM/PM")
      txtField(3).Text = Format(oApp.ServerDate, "hh:mm:ss AM/PM")
   End If
   
   txtField(1).Enabled = (lnEnabled = 0)
   txtField(2).Enabled = (lnEnabled = 0)
   txtField(3).Enabled = (lnEnabled = 0)
End Sub
