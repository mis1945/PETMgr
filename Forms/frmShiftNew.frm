VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmShiftNew 
   BorderStyle     =   0  'None
   Caption         =   "Shifts"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4950
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   8731
      BackColor       =   12632256
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
         Left            =   1155
         TabIndex        =   5
         Top             =   1150
         Width           =   5100
      End
      Begin VB.CheckBox chkFlexi 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "FLEXI TIME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   23
         Tag             =   "et0;fb0"
         Top             =   4440
         Width           =   1530
      End
      Begin VB.CheckBox chkPaid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "PAID BREAK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   22
         Tag             =   "wt0;fb0"
         Top             =   4440
         Width           =   1560
      End
      Begin VB.TextBox txtField 
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
         Index           =   9
         Left            =   4845
         TabIndex        =   11
         Top             =   1635
         Width           =   1400
      End
      Begin VB.TextBox txtField 
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
         Index           =   8
         Left            =   1605
         TabIndex        =   7
         Top             =   1635
         Width           =   1400
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Left            =   2430
         TabIndex        =   14
         Top             =   3120
         Width           =   1400
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   2430
         TabIndex        =   16
         Top             =   3570
         Width           =   1400
      End
      Begin VB.TextBox txtField 
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
         Index           =   5
         Left            =   1605
         TabIndex        =   9
         Top             =   2115
         Width           =   1400
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   4755
         TabIndex        =   19
         Top             =   3165
         Width           =   1400
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   7
         Left            =   4755
         TabIndex        =   21
         Top             =   3615
         Width           =   1400
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
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   690
         Width           =   5100
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
         Left            =   1155
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meal Sched"
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
         Index           =   11
         Left            =   45
         TabIndex        =   4
         Top             =   1245
         Width           =   1035
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "2nd Quarter"
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
         Index           =   10
         Left            =   4185
         TabIndex        =   17
         Tag             =   "et0;fb0"
         Top             =   2835
         Width           =   1035
      End
      Begin VB.Shape Shape2 
         Height          =   1200
         Index           =   1
         Left            =   4050
         Top             =   2955
         Width           =   2220
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "1st Quarter"
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
         Left            =   1830
         TabIndex        =   12
         Tag             =   "et0;fb0"
         Top             =   2850
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         Height          =   1200
         Index           =   0
         Left            =   1740
         Top             =   2955
         Width           =   2220
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approve OT"
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
         Left            =   3615
         TabIndex        =   10
         Top             =   1725
         Width           =   1125
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hours Of Work"
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
         Left            =   90
         TabIndex        =   6
         Top             =   1725
         Width           =   1290
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IN"
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
         Left            =   1875
         TabIndex        =   13
         Top             =   3195
         Width           =   180
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OUT"
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
         Left            =   1830
         TabIndex        =   15
         Top             =   3630
         Width           =   390
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Break Time"
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
         Left            =   390
         TabIndex        =   8
         Top             =   2205
         Width           =   990
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IN"
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
         Left            =   4290
         TabIndex        =   18
         Top             =   3225
         Width           =   180
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OUT"
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
         Left            =   4290
         TabIndex        =   20
         Top             =   3675
         Width           =   405
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1260
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   2415
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
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
         Left            =   105
         TabIndex        =   2
         Top             =   750
         Width           =   975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift ID"
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
         Left            =   105
         TabIndex        =   0
         Top             =   210
         Width           =   690
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   5700
      TabIndex        =   29
      Top             =   5760
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
      Picture         =   "frmShiftNew.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   4920
      TabIndex        =   28
      Top             =   5760
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
      Picture         =   "frmShiftNew.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   4140
      TabIndex        =   25
      Top             =   5760
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
      Picture         =   "frmShiftNew.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   3360
      TabIndex        =   26
      Top             =   5760
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
      Picture         =   "frmShiftNew.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   4140
      TabIndex        =   24
      Top             =   5760
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
      Picture         =   "frmShiftNew.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   5700
      TabIndex        =   30
      Top             =   5760
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
      Picture         =   "frmShiftNew.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   2580
      TabIndex        =   27
      Top             =   5760
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
      Picture         =   "frmShiftNew.frx":2CDC
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmShiftNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeModuleName = "frmShift"

Private WithEvents oRecord As clsShiftNew
Attribute oRecord.VB_VarHelpID = -1
Private oSkin As clsFormSkin


Private p_nActiveRow As Integer

Dim pnCtr As Integer
Dim pnIndex As Integer
Dim pbLoaded As Boolean
Dim pbSearched As Boolean

Private Sub chkFlexi_Click()
   oRecord.Master("cFlexiTym") = chkFlexi.Value
End Sub

Private Sub chkPaid_Click()
   oRecord.Master("cPaidBrke") = chkPaid.Value
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsProcName As String
   Dim lnCtr As Integer
   Dim lbAdd As Boolean
   
   lsProcName = "cmdButton_Click"
   'On Error GoTo errProc
   
   Select Case Index
   Case 0 'Cancel Edit Mode
      If oRecord.EditMode = xeModeAddNew Then
         oRecord.InitRecord
         Call LoadRecord
         initButton xeModeReady
      ElseIf oRecord.OpenRecord(oRecord.Master("sShiftIDx")) Then
         Call LoadRecord
         initButton xeModeReady
      End If
   Case 1 'Browse
      If oRecord.SearchRecord Then
         Call LoadRecord
         initButton xeModeReady
      End If
   Case 2 'Save
      If oRecord.SaveRecord Then
         MsgBox "Record Save Successfully...", vbInformation, "INFO"
         oRecord.InitRecord
         Call LoadRecord
         initButton xeModeReady
      End If
   Case 3 'Update
      If oRecord.UpdateRecord Then
         initButton xeModeUpdate
         txtField(80).SetFocus
      End If
   Case 4 'New
      If oRecord.NewRecord Then
         initButton xeModeAddNew
         
         If txtField(80).Enabled Then txtField(80).SetFocus
         Call LoadRecord
      End If
   Case 5 'Close
      Unload Me
   Case 6 'Delete
      If oRecord.DeleteRecord Then
         MsgBox "Record Successfully deleted...", vbInformation, "INFO"
         
         Call LoadRecord
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   'On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If Not pbLoaded Then
      pbLoaded = True
   End If
   
   pbSearched = False
   
   If txtField(80).Enabled Then txtField(80).SetFocus
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   'On Error GoTo errProc
   
   Call CenterChildForm(mdiMain, Me)
   
   Set oRecord = New clsShiftNew
   Set oRecord.AppDriver = oApp
   oRecord.InitRecord
   
   oRecord.NewRecord
   initButton xeModeAddNew
   LoadRecord
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   pbLoaded = False
End Sub

Private Sub oRecord_MasterRetrieved(ByVal Index As Integer)
   txtField(Index) = oRecord.Master(Index)
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
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
      If Index = 80 Then
         If oRecord.SearchMaster(Index, txtField(Index).Text) Then
            SetNextFocus
         End If
      End If
   Case vbKeyReturn
      If Index = 80 Then
         If txtField(Index) <> "" Then
            Call oRecord.SearchMaster(Index, txtField(Index).Text)
         Else
            oRecord.Master(Index) = txtField(Index).Text
         End If
      End If
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
   'On Error GoTo errProc
   Select Case Index
   Case 3, 4, 6, 7
      txtField(Index).Text = getCTime(txtField(Index).Text)
   End Select
   
   oRecord.Master(Index) = txtField(Index).Text
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(3).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   
   'kalyptus - 2019.03.14 09:54am
   'Do not allow delete for users aside from engineer
   cmdButton(6).Enabled = oApp.UserLevel = xeEngineer
   
   cmdButton(0).Visible = lbShow
   cmdButton(2).Visible = lbShow
   
   xrFrame1.Enabled = lbShow
End Sub

Private Sub LoadRecord()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 3, 4, 6, 7
         loTxt = oRecord.Master(loTxt.Index)
      Case Else
         loTxt = oRecord.Master(loTxt.Index)
      End Select
   Next
   
   chkFlexi.Value = oRecord.Master("cFlexiTym")
   chkPaid.Value = oRecord.Master("cPaidBrke")
   
   'kalyptus - 2019.03.14 10:00am
   'Do not allow update of the following fields for users other than Engineer(s)
   txtField(80).Locked = Not (oApp.BranchCode = "H002" Or oApp.UserLevel = xeEngineer)
   
   If Not (oRecord.EditMode = xeModeAddNew) Then
      txtField(3).Locked = Not (oApp.UserLevel = xeEngineer)
      txtField(5).Locked = Not (oApp.UserLevel = xeEngineer)
      txtField(8).Locked = Not (oApp.UserLevel = xeEngineer)
      txtField(9).Locked = Not (oApp.UserLevel = xeEngineer)
      chkFlexi.Enabled = (oApp.UserLevel = xeEngineer)
      chkPaid.Enabled = (oApp.UserLevel = xeEngineer)
   End If
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oApp
      .xLogError Err.Number, Err.Description, pxeModuleName, lsProcName, Erl
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

