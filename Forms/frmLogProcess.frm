VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLogProcess 
   BorderStyle     =   0  'None
   Caption         =   "Log Process"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   615
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   1085
      BackColor       =   12632256
      ClipControls    =   0   'False
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
         Height          =   400
         Index           =   0
         Left            =   765
         MaxLength       =   50
         TabIndex        =   1
         Top             =   90
         Width           =   2190
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
         Index           =   24
         Left            =   105
         TabIndex        =   0
         Top             =   165
         Width           =   405
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   8310
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   1185
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   14658
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin xrControl.xrFrame xrFrame3 
         Height          =   8085
         Left            =   105
         Tag             =   "wt0;fb0"
         Top             =   105
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   14261
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   420
            Index           =   0
            Left            =   3975
            MaskColor       =   &H00FFFFC0&
            TabIndex        =   5
            Tag             =   "wt0;fb0"
            Top             =   1500
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   420
            Index           =   1
            Left            =   3975
            TabIndex        =   8
            Tag             =   "wt0;fb0"
            Top             =   2025
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   420
            Index           =   2
            Left            =   3975
            TabIndex        =   11
            Tag             =   "wt0;fb0"
            Top             =   2550
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   420
            Index           =   3
            Left            =   3975
            TabIndex        =   27
            Tag             =   "wt0;fb0"
            Top             =   3075
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   420
            Index           =   4
            Left            =   3975
            TabIndex        =   16
            Tag             =   "wt0;fb0"
            Top             =   3600
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   420
            Index           =   5
            Left            =   3975
            TabIndex        =   19
            Tag             =   "wt0;fb0"
            Top             =   4125
            Width           =   225
         End
         Begin xrControl.xrButton cmdMove 
            Height          =   420
            Index           =   0
            Left            =   4260
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   2535
            Visible         =   0   'False
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   741
            Caption         =   "<<"
            AccessKey       =   "<<"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin xrControl.xrButton cmdMove 
            Height          =   420
            Index           =   1
            Left            =   4260
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   3075
            Visible         =   0   'False
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   741
            Caption         =   ">>"
            AccessKey       =   ">>"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1500
            Width           =   2415
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
            Index           =   5
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   18
            Top             =   4125
            Width           =   2415
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
            Index           =   4
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   15
            Top             =   3600
            Width           =   2415
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
            Index           =   3
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   13
            Top             =   3075
            Width           =   2415
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
            Index           =   2
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   10
            Top             =   2550
            Width           =   2415
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
            Index           =   1
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   7
            Top             =   2025
            Width           =   2415
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time In AM"
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
            Index           =   13
            Left            =   120
            TabIndex        =   3
            Top             =   1590
            Width           =   990
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Out AM"
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
            Index           =   12
            Left            =   120
            TabIndex        =   6
            Top             =   2115
            Width           =   1155
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time In PM"
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
            Left            =   120
            TabIndex        =   9
            Top             =   2640
            Width           =   990
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Out PM"
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
            Left            =   120
            TabIndex        =   12
            Top             =   3165
            Width           =   1155
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OT In"
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
            Left            =   120
            TabIndex        =   14
            Top             =   3690
            Width           =   465
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OT Out"
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
            TabIndex        =   17
            Top             =   4215
            Width           =   630
         End
         Begin VB.Label lblField 
            Alignment       =   2  'Center
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Index           =   0
            Left            =   105
            TabIndex        =   2
            Top             =   300
            Width           =   4590
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   8085
         Left            =   4980
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   105
         Width           =   8430
         _ExtentX        =   14870
         _ExtentY        =   14261
         _Version        =   393216
         Rows            =   3
         Cols            =   3
         FixedRows       =   2
         FixedCols       =   2
         WordWrap        =   -1  'True
         Enabled         =   0   'False
         FocusRect       =   0
         FillStyle       =   1
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   25
      Top             =   1815
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
      Picture         =   "frmLogProcess.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   22
      Top             =   555
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
      Picture         =   "frmLogProcess.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   24
      Top             =   1185
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
      Picture         =   "frmLogProcess.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   26
      Top             =   1815
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
      Picture         =   "frmLogProcess.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   23
      Top             =   1185
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
      Picture         =   "frmLogProcess.frx":1DE8
   End
End
Attribute VB_Name = "frmLogProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmLogProcess"

Private oTrans As clsLogProcess
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim pnRow As Integer
Dim pnActiveRow As Integer

Dim pbCtrlPress As Boolean
Dim pbFormLoad As Boolean
Dim pbDetailGotFocus As Boolean
Dim pbInUpdateMode As Boolean

Dim pdDate As Date

Dim psSelectedTime As String
Dim psTimeTempStrg As String
Dim pnLastSelc As Integer

Private Sub cmdMove_Click(Index As Integer)
   Call moveTime(Index)
   SetPreviousFocus
End Sub

Private Sub cmdMove_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
      If pbCtrlPress Then KeyCode = 0
   End If
   SetPreviousFocus
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
   End If
   
   If Not pbFormLoad Then pbFormLoad = True
   
   pbInUpdateMode = False
   txtSearch(0).Text = Format(oTrans.DateTransact, "MM/DD/YYYY")
   
   pnActiveRow = 2
   pnRow = 3
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lnRow As Integer
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown, vbKeyRight, vbKeyLeft
      With MSFlexGrid1
         Select Case KeyCode
         Case vbKeyReturn
            If GetFocus = .hwnd Then Exit Sub
            SetNextFocus
         Case vbKeyDown
            If pbCtrlPress Then
               If pnActiveRow = .Rows - 1 Then Exit Sub
               
               If Not (pnActiveRow = 0) Then
                  pnActiveRow = pnActiveRow + 1
               End If
               
               .Row = pnActiveRow
               
               If Not .RowIsVisible(.Row + IIf(.Row = .Rows - 1, 0, 1)) Then .TopRow = .Row - 1
   
               Call detailFieldChange
            End If
         Case vbKeyUp
            If pbCtrlPress Then
               If .Row > 2 Then
                  If Not (pnActiveRow = 1) Then
                     pnActiveRow = pnActiveRow - 1
                  End If
                  
                  .Row = pnActiveRow
      
                  If Not .RowIsVisible(.Row) Then .TopRow = .TopRow - 1
   
                  Call detailFieldChange
               End If
   
            End If
         Case vbKeyRight '1
            If Not pbInUpdateMode Or Not pbCtrlPress Then Exit Sub
            Call moveTime(1)
         Case vbKeyLeft '0
            If Not pbInUpdateMode Or Not pbCtrlPress Then Exit Sub
            Call moveTime(0)
         End Select
      End With
   Case vbKeyControl
      pbCtrlPress = True
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   Set oTrans = New clsLogProcess
   Set oTrans.AppDriver = oApp
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Call InitGrid
   Call InitForm(0)
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
   
   pnActiveRow = 1
   pbFormLoad = False
   pbCtrlPress = False
End Sub

Private Sub detailFieldChange()
   Dim lnRow As Integer

   lnRow = pnActiveRow - 2
   pnRow = lnRow

   With oTrans
      lblField(0).Caption = oTrans.LastName(lnRow) & ", " & _
                                     oTrans.FrstName(lnRow) & " " & _
                                     oTrans.MiddName(lnRow)
      txtField(0).Text = IFNull(Format(.AMInxxxx(lnRow), "HH:MM:SS AM/PM"), "")
      txtField(1).Text = IFNull(Format(.AMOutxxx(lnRow), "HH:MM:SS AM/PM"), "")
      txtField(2).Text = IFNull(Format(.PMInxxxx(lnRow), "HH:MM:SS AM/PM"), "")
      txtField(3).Text = IFNull(Format(.PMOutxxx(lnRow), "HH:MM:SS AM/PM"), "")
      txtField(4).Text = IFNull(Format(.OTimeInx(lnRow), "HH:MM:SS AM/PM"), "")
      txtField(5).Text = IFNull(Format(.OTimeOut(lnRow), "HH:MM:SS AM/PM"), "")
   End With
   
   Call isDFieldModified
End Sub

Private Sub isDFieldModified()
   Dim lnCtr As Integer
   Dim lnRow As Integer
   Dim lbDetChanged As Boolean
   Dim loTxt As TextBox

   For Each loTxt In txtField
      loTxt.BackColor = oApp.getColor("EB")
   Next

   lnRow = pnActiveRow - 2
   pnRow = lnRow
   lbDetChanged = False
   
   With MSFlexGrid1
'      If Format(IFNull(.TextMatrix(lnRow + 2, 2), ""), "HH:MM:SS AM/PM") <> _
'         Format(IFNull(txtField(0).Text, ""), "HH:MM:SS AM/PM") Then MsgBox "ola" 'txtField(0).BackColor = oApp.getColor("HT1")
      
      If Format(IFNull(.TextMatrix(lnRow + 2, 2), ""), "HH:MM:SS AM/PM") <> _
         Format(IFNull(txtField(0).Text, ""), "HH:MM:SS AM/PM") Then lbDetChanged = True
      If Format(IFNull(.TextMatrix(lnRow + 2, 3), ""), "HH:MM:SS AM/PM") <> _
         Format(IFNull(txtField(1).Text, ""), "HH:MM:SS AM/PM") Then lbDetChanged = True
      If Format(IFNull(.TextMatrix(lnRow + 2, 4), ""), "HH:MM:SS AM/PM") <> _
         Format(IFNull(txtField(2).Text, ""), "HH:MM:SS AM/PM") Then lbDetChanged = True
      If Format(IFNull(.TextMatrix(lnRow + 2, 5), ""), "HH:MM:SS AM/PM") <> _
         Format(IFNull(txtField(3).Text, ""), "HH:MM:SS AM/PM") Then lbDetChanged = True
      If Format(IFNull(.TextMatrix(lnRow + 2, 6), ""), "HH:MM:SS AM/PM") <> _
         Format(IFNull(txtField(4).Text, ""), "HH:MM:SS AM/PM") Then lbDetChanged = True
      If Format(IFNull(.TextMatrix(lnRow + 2, 7), ""), "HH:MM:SS AM/PM") <> _
         Format(IFNull(txtField(5).Text, ""), "HH:MM:SS AM/PM") Then lbDetChanged = True
   
      .Row = lnRow + 2
      .Col = 1
      .CellBackColor = IIf(lbDetChanged, vbYellow, &H8000000F)
      SetGridRowColor (.Row)
   End With
   
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Cols = 8
      .Rows = 2

      .MergeCells = flexMergeFree
      
      .Clear
      
      .Row = 0
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .MergeRow(0) = True
      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = ""
      .TextMatrix(0, 2) = "1ST QUARTER"
      .TextMatrix(0, 3) = "1ST QUARTER"
      .TextMatrix(0, 4) = "2ND QUARTER"
      .TextMatrix(0, 5) = "2ND QUARTER"
      .TextMatrix(0, 6) = "OVERTIME"
      .TextMatrix(0, 7) = "OVERTIME"
      
      .Row = 1
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      'column title
      .TextMatrix(1, 0) = "NO."
      .TextMatrix(1, 1) = "EMPLOYEE NAME"
      .TextMatrix(1, 2) = "In"
      .TextMatrix(1, 3) = "Out"
      .TextMatrix(1, 4) = "In"
      .TextMatrix(1, 5) = "Out"
      .TextMatrix(1, 6) = "In"
      .TextMatrix(1, 7) = "Out"
      .RowHeightMin = 338
      
      'column width
      .ColWidth(0) = 592
      .ColWidth(1) = 3000
      .ColWidth(2) = 1200
      .ColWidth(3) = 1200
      .ColWidth(4) = 1200
      .ColWidth(5) = 1200
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      
      'column alignment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignCenterCenter
      .ColAlignment(4) = flexAlignCenterCenter
      .ColAlignment(5) = flexAlignCenterCenter
      .ColAlignment(6) = flexAlignCenterCenter
      .ColAlignment(7) = flexAlignCenterCenter
      
      .Rows = 3
      .TextMatrix(2, 0) = "1"
      
      .Row = 2
      pnLastSelc = .Row
      pnRow = 2
      SetGridRowColor (.Row)
   End With
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer

   With MSFlexGrid1
      .Rows = oTrans.ItemCount + 2
      
      For lnRow = 0 To oTrans.ItemCount - 1

         .TextMatrix(lnRow + 2, 0) = lnRow + 1
         .TextMatrix(lnRow + 2, 1) = oTrans.LastName(lnRow) & ", " & _
                                     oTrans.FrstName(lnRow) '& " " & _
                                     oTrans.MiddName(lnRow)
         .TextMatrix(lnRow + 2, 2) = oTrans.AMInxxxx(lnRow)
         .TextMatrix(lnRow + 2, 3) = oTrans.AMOutxxx(lnRow)
         .TextMatrix(lnRow + 2, 4) = oTrans.PMInxxxx(lnRow)
         .TextMatrix(lnRow + 2, 5) = oTrans.PMOutxxx(lnRow)
         .TextMatrix(lnRow + 2, 6) = oTrans.OTimeInx(lnRow)
         .TextMatrix(lnRow + 2, 7) = oTrans.OTimeOut(lnRow)
      Next

      .Enabled = True
      
      If lnRow + 1 > 18 Then
         .ColWidth(1) = 2750
      Else
         .ColWidth(1) = 2900
      End If
      
      .Row = 2
      pnRow = pnActiveRow - 2
      Call MSFlexGrid1_Click
   End With
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0   'Retrieve
      If retreiveDetail Then
         pnActiveRow = 1
         txtSearch(0).SetFocus
      End If
   Case 1   'Close
      Unload Me
   Case 2   'Save
      'kalyptus - 2017.05.15 04:36pm
      'Do not allow saving if transaction displayed in the form is different from the loaded transaction date
      If Not (CDate(txtSearch(0).Text) = oTrans.DateTransact And oTrans.DateTransact = pdDate) Then
         MsgBox "Please load the correct transaction date before processing!"
         Exit Sub
      End If
      
      If oTrans.SaveTransaction Then
         Call InitForm(0)
         ClearFields
         Call InitGrid
         
         If retreiveDetail Then
            oTrans.processLog (pdDate)
         End If
      End If
   Case 3   'Cancel
      If oTrans.InitTransaction Then
         InitForm 0
         pnRow = 0

         ClearFields
         InitGrid
         retreiveDetail
      End If
   Case 4   'Update
'      oTrans.processLog txtSearch(0)
      If oTrans.UpdateTransaction(pdDate) Then
         InitForm 1
'         pbDetailGotFocus = True
'         txtSearch(0).Enabled = True
'         txtSearch(0).SetFocus
         MSFlexGrid1.Row = 2
         pnActiveRow = 2
         detailFieldChange
         Option1(0).Value = True
         Option1(0).SetFocus
      End If
   End Select
End Sub

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      pnActiveRow = .Row
   End With

   Call detailFieldChange
   
   If pbInUpdateMode Then
      Option1(0).SetFocus
   Else
      txtSearch(0).SetFocus
   End If
End Sub

Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyDown, vbKeyUp
         If pbCtrlPress Then KeyCode = 0
   End Select
End Sub

Private Sub Option1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyControl
         If pbCtrlPress Then pbCtrlPress = False
   End Select
End Sub

Private Sub txtField_Change(Index As Integer)
   Select Case Index
      Case 0
         If Trim(txtField(Index)) <> "" Then
            oTrans.AMInxxxx(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
         Else
            oTrans.AMInxxxx(pnRow) = ""
         End If
      Case 1
         If Trim(txtField(Index)) <> "" Then
            oTrans.AMOutxxx(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
         Else
            oTrans.AMOutxxx(pnRow) = ""
         End If
      Case 2
         If Trim(txtField(Index)) <> "" Then
            oTrans.PMInxxxx(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
         Else
            oTrans.PMInxxxx(pnRow) = ""
         End If
      Case 3
         If Trim(txtField(Index)) <> "" Then
            oTrans.PMOutxxx(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
         Else
            oTrans.PMOutxxx(pnRow) = ""
         End If
      Case 4
         If Trim(txtField(Index)) <> "" Then
            oTrans.OTimeInx(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
         Else
            oTrans.OTimeInx(pnRow) = ""
         End If
      Case 5
         If Trim(txtField(Index)) <> "" Then
            oTrans.OTimeOut(pnRow) = Format(txtField(Index).Text, "HH:MM:SS AM/PM")
         Else
            oTrans.OTimeOut(pnRow) = ""
         End If
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

   pnIndex = Index
End Sub

Private Sub MSFlexGrid1_GotFocus()
   pbDetailGotFocus = True
End Sub
Private Sub MSFlexGrid1_LostFocus()
   If pbDetailGotFocus Then pbDetailGotFocus = False
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyControl
         If pbCtrlPress Then pbCtrlPress = False
   End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
   txtSearch(Index).BackColor = oApp.getColor("HT1")
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         If retreiveDetail Then
            pnActiveRow = 1
            txtSearch(0).SetFocus
         End If
      Case vbKeyTab
         If pbInUpdateMode Then cmdButton(4).SetFocus
      Case vbKeyDown, vbKeyUp
         If pbCtrlPress Then KeyCode = 0
   End Select
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtSearch_Validate(Index As Integer, Cancel As Boolean)
   With txtSearch(Index)
      .Text = Format(.Text, "MM/DD/YYYY")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub InitForm(ByVal fnEdit As Integer)
   Dim lnCtr As Integer
   Dim loTxt As TextBox

   For Each loTxt In txtField
      loTxt.BackColor = oApp.getColor("EB")
   Next

   cmdButton(2).Visible = Not (fnEdit = 0)
   cmdButton(3).Visible = Not (fnEdit = 0)

   cmdButton(0).Visible = (fnEdit = 0)
   cmdButton(1).Visible = (fnEdit = 0)
   cmdButton(4).Visible = (fnEdit = 0)
   
   cmdMove(0).Visible = Not (fnEdit = 0)
   cmdMove(1).Visible = Not (fnEdit = 0)
   
   For lnCtr = 0 To 5
      Option1(lnCtr).Visible = Not (fnEdit = 0)
   Next
   
   pbInUpdateMode = Not (fnEdit = 0)
   txtSearch(0).Enabled = Not pbInUpdateMode
End Sub

Private Sub ClearFields()
   Dim lnCtr As Integer

   For lnCtr = 0 To 5
      txtField(lnCtr) = ""
   Next
   
   lblField(0).Caption = ""
End Sub

Private Function retreiveDetail()
   retreiveDetail = False
   If IsDate(txtSearch(0)) Then
         If oTrans.loadEmpLog(txtSearch(0)) Then
            pdDate = CDate(txtSearch(0))
            LoadDetail
            retreiveDetail = True
         End If
      Else
         MsgBox "Invalid date detected!!!", vbCritical, "Warning"
   End If
End Function

Private Sub moveTime(ByVal lnDrectn As Integer)
   Dim lnCtr As Integer
   Dim lnIndex As Integer
   Dim lnHavVal As Integer
   
   'kalyptus = 2017.05.16 09:27 am
   'Allow update if log date is the current loaded log date...
   If Not (CDate(txtSearch(0).Text) = oTrans.DateTransact And oTrans.DateTransact = pdDate) Then
      MsgBox "Date is different from loaded Log Date!" & vbCrLf & _
            " Please load the correct log date before log update...."
      Exit Sub
   End If
   
   'get index of selected time
   For lnCtr = 0 To 5
      If Option1(lnCtr).Value = True Then
         lnIndex = lnCtr
         Exit For
      End If
   Next
   
   Select Case lnDrectn
      Case 0
         'check fields that have value
         lnHavVal = 0
         lnCtr = lnIndex
         Do While lnCtr >= 0
            If txtField(lnCtr).Text <> "" Then
               lnHavVal = lnHavVal + 1
               lnCtr = lnCtr - 1
            Else
               Exit Do
            End If
         Loop
         'move field values
         lnCtr = lnIndex - lnHavVal
         If lnCtr < 0 Then Exit Sub
         Do While lnCtr < lnIndex
            txtField(lnCtr).Text = txtField(lnCtr + 1)
            txtField(lnCtr + 1).Text = ""
            Option1(lnCtr).Value = True
            Option1(lnCtr).SetFocus
            lnCtr = lnCtr + 1
         Loop
      Case 1
         'check fields that have value
         lnHavVal = 0
         For lnCtr = lnIndex To 5
            If txtField(lnCtr).Text <> "" Then
               lnHavVal = lnHavVal + 1
            Else
               Exit For
            End If
         Next
         'move field values
         lnCtr = lnIndex + lnHavVal
         If lnCtr > 5 Then Exit Sub
         Do While lnCtr > lnIndex
            txtField(lnCtr).Text = txtField(lnCtr - 1)
            txtField(lnCtr - 1).Text = ""
            Option1(lnCtr).Value = True
            Option1(lnCtr).SetFocus
            lnCtr = lnCtr - 1
         Loop
   End Select
   
   Call isDFieldModified
End Sub

Private Sub SetGridRowColor(ByVal lnRow As Integer)
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .FillStyle = flexFillRepeat
      
      .Row = IIf(pnLastSelc = .Rows, pnLastSelc - 1, pnLastSelc)
      .RowSel = pnLastSelc - 1
      .Col = 2
      .ColSel = .Cols - 1
      .CellBackColor = &HFFFFFF
      
      .Row = lnRow
      .RowSel = lnRow
      .Col = 2
      .ColSel = .Cols - 1
      .CellBackColor = &HFF8080
      
      pnLastSelc = .Row
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

