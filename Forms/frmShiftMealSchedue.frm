VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmShiftMealSchedue 
   BorderStyle     =   0  'None
   Caption         =   "Meal Shift Schedule"
   ClientHeight    =   6180
   ClientLeft      =   -120
   ClientTop       =   -465
   ClientWidth     =   11355
   Icon            =   "frmShiftMealSchedue.frx":0000
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3315
      Left            =   1680
      Tag             =   "wt0;fb0"
      Top             =   600
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   5847
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Frame Frame3 
         Caption         =   "PM"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Tag             =   "wt0;fb0"
         Top             =   2280
         Width           =   9135
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   5520
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "M00111-000021"
            Top             =   280
            Width           =   2415
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   1520
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "M00111-000021"
            Top             =   280
            Width           =   2415
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Out:"
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
            Left            =   4560
            TabIndex        =   22
            Top             =   300
            Width           =   855
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time In:"
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
            Left            =   480
            TabIndex        =   21
            Top             =   300
            Width           =   690
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Break Time"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   24
         Tag             =   "wt0;fb0"
         Top             =   1440
         Width           =   9135
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   1520
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "M00111-000021"
            Top             =   280
            Width           =   2415
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
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
            Left            =   480
            TabIndex        =   20
            Top             =   300
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "AM"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   23
         Tag             =   "wt0;fb0"
         Top             =   600
         Width           =   9135
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   5520
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "M00111-000021"
            Top             =   280
            Width           =   2415
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   1520
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "8:30 AM"
            Top             =   280
            Width           =   2415
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Out:"
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
            Left            =   4560
            TabIndex        =   19
            Top             =   300
            Width           =   855
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time In:"
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
            Left            =   480
            TabIndex        =   18
            Top             =   300
            Width           =   690
         End
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Name:"
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
         Index           =   21
         Left            =   225
         TabIndex        =   9
         Top             =   210
         Width           =   1110
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmShiftMealSchedue.frx":000C
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1860
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmShiftMealSchedue.frx":0786
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1230
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1058
      Caption         =   "Searc&h"
      AccessKey       =   "h"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmShiftMealSchedue.frx":0F00
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1058
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
      Picture         =   "frmShiftMealSchedue.frx":167A
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   2490
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmShiftMealSchedue.frx":1DF4
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   1860
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmShiftMealSchedue.frx":256E
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   1230
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1058
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
      Picture         =   "frmShiftMealSchedue.frx":2C68
      CaptionAlign    =   0
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1995
      Left            =   1680
      Tag             =   "wt0;fb0"
      Top             =   3960
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   3519
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Frame Frame4 
         Caption         =   "Time"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   28
         Tag             =   "wt0;fb0"
         Top             =   600
         Width           =   9135
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   9
            Left            =   1520
            MaxLength       =   50
            TabIndex        =   8
            Text            =   "M00111-000021"
            Top             =   680
            Width           =   2415
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   1520
            MaxLength       =   50
            TabIndex        =   7
            Text            =   "M00111-000021"
            Top             =   280
            Width           =   2415
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Thru:"
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
            Left            =   480
            TabIndex        =   27
            Top             =   645
            Width           =   435
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From:"
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
            Left            =   480
            TabIndex        =   25
            Top             =   300
            Width           =   510
         End
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meal Type:"
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
         Left            =   225
         TabIndex        =   17
         Top             =   210
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmShiftMealSchedue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmShiftMealSchedule"

Private WithEvents oTrans As clsShiftMealSchedule
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim lnMsg As String

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   Select Case Index
   Case 0   'Save
      If oTrans.SaveRecord Then
         initButton (xeModeReady)
      End If
   Case 1   'search
      If pnIndex = 1 Or pnIndex = 7 Then
            oTrans.SearchMaster pnIndex, txtField(pnIndex).Text
            If txtField(pnIndex).Text <> "" Then SetNextFocus
            Else
               If txtField(pnIndex).Text <> "" Then oTrans.SearchMaster pnIndex, txtField(pnIndex).Text
      End If
   Case 2   'Browse
         If oTrans.SearchRecord = True Then
            Call LoadRecord
         End If
   Case 3   'Cancel
         lnMsg = MsgBox("Do you want to discard changes?", vbYesNo + vbQuestion, "Confirm")
         If lnMsg = vbYes Then
           If oTrans.CancelUpdate = True Then
            ClearFields
            Call initButton(xeModeReady)
            End If
         End If
   Case 4 ' New
         If oTrans.NewRecord = True Then
            Call ClearFields
            Call LoadRecord
            Call initButton(xeModeAddNew)
            txtField(1).SetFocus
         End If
   Case 5 ' Close
            Unload Me
   Case 7 ' Update
         If oTrans.UpdateRecord Then
            initButton (xeModeUpdate)
            txtField(1).Enabled = False
            txtField(7).SetFocus
         End If
   End Select

endProc:
   Exit Sub
endWithFocus:
'   txtSearch(0) = ""
'   txtSearch(1) = ""
'   txtSearch(0).SetFocus
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   'On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      bLoaded = True
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
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

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcMealVoucher.clsShiftMealSchedule
   Set oTrans.AppDriver = oApp
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   ClearFields

   oTrans.InitRecord
   oTrans.NewRecord
   Call LoadRecord
   
   initButton (xeModeAddNew)

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub LoadRecord()

   With oTrans
      txtField(1) = IFNull(.Master(1), "")
      txtField(2) = IFNull(Format(.Master(2), "hh:mm:ss AM/PM"), "")
      txtField(3) = IFNull(Format(.Master(3), "hh:mm:ss AM/PM"), "")
      txtField(4) = IFNull(Format(.Master(4), "hh:mm:ss AM/PM"), "")
      txtField(5) = IFNull(Format(.Master(5), "hh:mm:ss AM/PM"), "")
      txtField(6) = IFNull(Format(.Master(6), "hh:mm:ss AM/PM"), "")
      txtField(7) = IFNull(.Master(7), "")
      txtField(8) = IFNull(Format(.Master(8), "hh:mm:ss AM/PM"), "")
      txtField(9) = IFNull(Format(.Master(9), "hh:mm:ss AM/PM"), "")
   End With
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(2).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(7).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow

   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(3).Visible = lbShow
'
   txtField(1).Enabled = lbShow
   txtField(7).Enabled = lbShow

'   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   With oTrans
      Select Case Index
         Case 1
              txtField(Index) = IFNull(.Master(1), "")
         Case 2
              txtField(Index) = IFNull(Format(.Master(2), "hh:mm:ss AM/PM"), "")
         Case 3
              txtField(Index) = IFNull(Format(.Master(3), "hh:mm:ss AM/PM"), "")
         Case 4
              txtField(Index) = IFNull(Format(.Master(4), "hh:mm:ss AM/PM"), "")
         Case 5
              txtField(Index) = IFNull(Format(.Master(5), "hh:mm:ss AM/PM"), "")
         Case 6
              txtField(Index) = IFNull(Format(.Master(6), "hh:mm:ss AM/PM"), "")
         Case 7
              txtField(Index) = IFNull(.Master(7), "")
         Case 8
              txtField(Index) = IFNull(Format(.Master(8), "hh:mm:ss AM/PM"), "")
         Case 9
              txtField(Index) = IFNull(Format(.Master(9), "hh:mm:ss AM/PM"), "")
      End Select
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Then
            oTrans.SearchMaster Index, .Text
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then oTrans.SearchMaster Index, .Text
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



