VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmLogProcess 
   BorderStyle     =   0  'None
   Caption         =   "Log Process"
   ClientHeight    =   7860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13245
   Enabled         =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   6090
      Left            =   255
      Tag             =   "wt0;fb0"
      Top             =   1470
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   10742
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   7410
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1605
         Width           =   2355
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   7
         Left            =   7410
         MaxLength       =   50
         TabIndex        =   14
         Top             =   4230
         Width           =   2355
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   6
         Left            =   7410
         MaxLength       =   50
         TabIndex        =   12
         Top             =   3705
         Width           =   2355
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   7410
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3180
         Width           =   2355
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   7410
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2655
         Width           =   2355
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   7410
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2130
         Width           =   2355
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5595
         Left            =   255
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   9869
         _Version        =   393216
         Enabled         =   -1  'True
         FocusRect       =   2
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   7410
         MaxLength       =   50
         TabIndex        =   3
         Top             =   690
         Width           =   3495
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
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
         Index           =   1
         Left            =   6120
         TabIndex        =   2
         Top             =   780
         Width           =   930
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OT Out"
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
         Index           =   7
         Left            =   6105
         TabIndex        =   13
         Top             =   4320
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OT In"
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
         Index           =   6
         Left            =   6105
         TabIndex        =   11
         Top             =   3795
         Width           =   480
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out PM"
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
         Index           =   5
         Left            =   6105
         TabIndex        =   9
         Top             =   3270
         Width           =   1155
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time In PM"
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
         Index           =   4
         Left            =   6105
         TabIndex        =   7
         Top             =   2745
         Width           =   1005
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out AM"
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
         Left            =   6105
         TabIndex        =   5
         Top             =   2220
         Width           =   1155
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time In AM"
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
         Left            =   6105
         TabIndex        =   4
         Top             =   1695
         Width           =   1005
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   735
      Left            =   255
      Tag             =   "wt0;fb0"
      Top             =   690
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   1296
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtSearch 
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
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   1
         Top             =   135
         Width           =   3030
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   24
         Left            =   225
         TabIndex        =   0
         Top             =   180
         Width           =   600
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   11940
      TabIndex        =   16
      Top             =   2910
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
      Picture         =   "frmLogProcess1.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   11940
      TabIndex        =   15
      Top             =   2250
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
      Picture         =   "frmLogProcess1.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   11940
      TabIndex        =   18
      Top             =   2250
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
      Picture         =   "frmLogProcess1.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   11940
      TabIndex        =   19
      Top             =   2910
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
      Picture         =   "frmLogProcess1.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   11940
      TabIndex        =   17
      Top             =   1590
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmLogProcess1.frx":1DE8
   End
End
Attribute VB_Name = "frmLogProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmHoliday"
'
'Private oTrans As clsLogProcess
'Private oSkin As clsFormSkin
'Private oDriver As New clsFormDriver
'Private bLoaded As Boolean
'
'Dim pnIndex As Integer
'Dim pnRow As Integer
'Dim pdDate As Date
'
'Private Sub Form_Activate()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Activate"
'   On Error Goto errProc
'
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'
'   If bLoaded = False Then
'      bLoaded = True
'   End If
'   txtSearch(0).Text = Format(oTrans.DateTransact, "MM/DD/YYYY")
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   End Select
'End Sub
'
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
''   On Error GoTo errProc
'
'   'CenterChildForm mdiMain, Me
'
'   Set oDriver = New clsFormDriver
'   Set oDriver.AppDriver = oApp
'   Set oDriver.MainForm = Me
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransMaintenance
'
'   Set oTrans = New clsLogProcess
'   Set oTrans.AppDriver = oApp
'   oTrans.Branch = oApp.BranchCode
'   oTrans.InitTransaction
'
'   Call InitGrid
'   Call InitForm(0)
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oTrans = Nothing
'   Set oSkin = Nothing
'End Sub
'
'Private Sub MSFlexGrid1_Click()
'   Dim lnRow As Integer
'   Dim loTxt As TextBox
'
'   lnRow = MSFlexGrid1.Row - 1
'   pnRow = lnRow
'
'   With oTrans
'      'txtField(0).Text = IFNull(.TranDate(lnRow), "")
'      txtField(1).Text = MSFlexGrid1.TextMatrix(lnRow + 1, 2)
'      txtField(2).Text = IFNull(Format(.AMInxxxx(lnRow), "HH:MM:SS AM/PM"), "")
'      txtField(3).Text = IFNull(Format(.AMOutxxx(lnRow), "HH:MM:SS AM/PM"), "")
'      txtField(4).Text = IFNull(Format(.PMInxxxx(lnRow), "HH:MM:SS AM/PM"), "")
'      txtField(5).Text = IFNull(Format(.PMOutxxx(lnRow), "HH:MM:SS AM/PM"), "")
'      txtField(6).Text = IFNull(Format(.OTimeInx(lnRow), "HH:MM:SS AM/PM"), "")
'      txtField(7).Text = IFNull(Format(.OTimeOut(lnRow), "HH:MM:SS AM/PM"), "")
'   End With
'End Sub
'
'Private Sub InitGrid()
'   Dim lnCtr As Integer
'   With MSFlexGrid1
'      .Cols = 3
'      .Rows = 2
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 0) = "No"
'      .TextMatrix(0, 1) = "Employee ID"
'      .TextMatrix(0, 2) = "Employee Name"
'
'      .Row = 0
'
'      'column alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 1
'      Next
'      'column width
'      .ColWidth(0) = 330
'      .ColWidth(1) = 2000
'      .ColWidth(2) = 3000
'
'      'column allinment
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .Row = 1
'   End With
'End Sub
'
'Private Sub LoadDetail()
'   Dim lnRow As Integer
'
'   With MSFlexGrid1
'      .Enabled = True
'
'      For lnRow = 0 To oTrans.ItemCount - 1
'         .Rows = lnRow + 2
'         .TextMatrix(lnRow + 1, 0) = lnRow + 1
'         .TextMatrix(lnRow + 1, 1) = oTrans.EmployID(lnRow)
'         .TextMatrix(lnRow + 1, 2) = oTrans.LastName(lnRow) & ", " & _
'                                     oTrans.FrstName(lnRow) & " " & _
'                                     oTrans.MiddName(lnRow)
'      Next
'      If lnRow + 1 > 18 Then
'         .ColWidth(2) = 1850
'      Else
'         .ColWidth(2) = 2000
'      End If
'   End With
'End Sub
'
'Private Sub cmdButton_Click(Index As Integer)
'   Select Case Index
'   Case 0   'Retrieve
'      If IsDate(txtSearch(0)) Then
'         If oTrans.loadEmpLog(txtSearch(0)) Then
'            pdDate = CDate(txtSearch(0))
'            LoadDetail
'            InitForm 0
'         End If
'      Else
'         MsgBox "Invalid date detected!!!", vbCritical, "Warning"
'         txtSearch(0).SetFocus
'      End If
'   Case 1   'Close
'      Unload Me
'   Case 2   'Save
'      If oTrans.SaveTransaction Then
'         MsgBox "ola"
'         InitForm 0
'      End If
'   Case 3   'Cancel
'      If oTrans.InitTransaction Then
'         InitForm 0
'         pnRow = 0
'         ClearFields
'      End If
'   Case 4   'Update
'      If oTrans.UpdateTransaction(pdDate) Then InitForm 1
'   End Select
'End Sub
'
'Private Sub InitForm(ByVal fnEdit As Integer)
'   Dim lnCtr As Integer
'
'   cmdButton(2).Visible = Not (fnEdit = 0)
'   cmdButton(3).Visible = Not (fnEdit = 0)
'
'   cmdButton(0).Visible = (fnEdit = 0)
'   cmdButton(1).Visible = (fnEdit = 0)
'   cmdButton(4).Visible = (fnEdit = 0)
'
'   For lnCtr = 2 To 7
'      txtField(lnCtr).Enabled = Not (fnEdit = 0)
'   Next
'End Sub
'
'Private Sub ClearFields()
'   Dim lnCtr As Integer
'
'   For lnCtr = 2 To 7
'      txtField(lnCtr) = ""
'   Next
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   Select Case Index
'      Case 2
'         oTrans.AMInxxxx(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
'      Case 3
'         oTrans.AMInxxxx(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
'      Case 4
'         oTrans.PMInxxxx(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
'      Case 5
'         oTrans.PMOutxxx(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
'      Case 6
'         oTrans.OTimeInx(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
'      Case 7
'         oTrans.OTimeOut(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
'   End Select
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   oDriver.ColumnIndex = Index
'   pnIndex = Index
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
