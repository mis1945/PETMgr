VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmTimesheetConfirmation 
   BorderStyle     =   0  'None
   Caption         =   "Payroll Processing"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3165
      Left            =   150
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5583
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.ComboBox cmbEmpType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmTimesheetConfirmation.frx":0000
         Left            =   1680
         List            =   "frmTimesheetConfirmation.frx":000D
         TabIndex        =   17
         Top             =   135
         Width           =   2205
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
         Height          =   400
         Index           =   0
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   15
         Text            =   "M00111-000021"
         Top             =   1500
         Width           =   2190
      End
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
         Index           =   1
         Left            =   5355
         MaxLength       =   50
         TabIndex        =   14
         Top             =   645
         Width           =   2190
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
         Height          =   400
         Index           =   4
         Left            =   5520
         TabIndex        =   10
         Text            =   "August 28, 2011"
         Top             =   2565
         Width           =   2190
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
         Height          =   400
         Index           =   3
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "August 13, 2011"
         Top             =   2565
         Width           =   2190
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
         Height          =   400
         Index           =   2
         Left            =   5520
         TabIndex        =   6
         Text            =   "August 31, 2011"
         Top             =   2115
         Width           =   2190
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
         Height          =   400
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "August 16, 2011"
         Top             =   2115
         Width           =   2190
      End
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
         Left            =   1695
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   2205
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Type::"
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
         Index           =   2
         Left            =   150
         TabIndex        =   18
         Top             =   195
         Width           =   1590
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   405
         Left            =   1575
         Tag             =   "et0;ht2"
         Top             =   1590
         Width           =   2190
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period ID."
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
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Fr:"
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
         Left            =   4275
         TabIndex        =   13
         Top             =   690
         Width           =   750
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   7815
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5535
         TabIndex        =   2
         Tag             =   "eb0;et0"
         Top             =   1560
         Width           =   2070
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   5490
         Top             =   1545
         Width           =   2160
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   5460
         Top             =   1515
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coverage Thru"
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
         Left            =   4170
         TabIndex        =   9
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coverage From"
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
         Left            =   195
         TabIndex        =   7
         Top             =   2640
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period Thru"
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
         Index           =   0
         Left            =   4170
         TabIndex        =   5
         Top             =   2190
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period From"
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
         Left            =   195
         TabIndex        =   3
         Top             =   2190
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period ID:"
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
         Left            =   165
         TabIndex        =   0
         Top             =   705
         Width           =   945
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   8310
      TabIndex        =   11
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Confirm"
      AccessKey       =   "Confirm"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTimesheetConfirmation.frx":002C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   8310
      TabIndex        =   12
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "C&lose"
      AccessKey       =   "l"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTimesheetConfirmation.frx":07A6
   End
End
Attribute VB_Name = "frmTimesheetConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmTimesheetConfirmation"

Private oTrans As ggcPayroll.clsTimeSheetConfirmation
Private oSkin As clsFormSkin

Private Sub cmbEmpType_Click()
   If cmbEmpType.ListIndex = 0 Then
      oTrans.EmployeeType = "T"
   Else
      oTrans.EmployeeType = "R"
   End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0 'Confirm
      If LCase(oApp.ProductID) = "petmgr" Then
         If oTrans.ConfirmTransaction Then
            MsgBox "Timesheet Confirmed Successfully"
         End If
         Call LoadMaster
      End If
   Case 1 'Close/Exit Form
      Unload Me
   End Select

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_Click()

End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
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
'   'On Error GoTo errProc

   CenterChildForm mdiMain, Me
      
   Set oTrans = New ggcPayroll.clsTimeSheetConfirmation
   Set oTrans.AppDriver = oApp
   oTrans.TransStatus = 10
   oTrans.PayrollStatus = 10
   oTrans.InitTransaction
         
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   
   cmbEmpType.ListIndex = 0
   
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oTrans = Nothing
End Sub

Private Sub LoadMaster()
   
   txtField(0).Text = oTrans.Master("sPayPerID")
   txtField(1).Text = oTrans.Master("dPeriodFr")
   txtField(2).Text = oTrans.Master("dPeriodTo")
   txtField(3).Text = oTrans.Master("dCovergFm")
   txtField(4).Text = oTrans.Master("dCovergTo")
   
   txtSearch(0).Text = oTrans.Master("sPayPerID")
   txtSearch(1).Text = oTrans.Master("dPeriodFr")
   Label2 = TransStat(oTrans.Master("cPostedxx"))
   
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

Private Sub txtField_GotFocus(Index As Integer)
'   Select Case Index
'   Case 2, 4
'      txtField(Index) = Format(oTrans.Master(Index), "MM/DD/YYYY")
'   End Select
'
'   With txtField(Index)
'      .BackColor = oApp.getColor("HT1")
'      .SelStart = 0
'      .SelLength = Len(.Text)
'   End With
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   With txtField(Index)

End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With txtSearch(Index)
      If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
         oTrans.SearchTransaction txtSearch(Index).Text, False
         Call LoadMaster
      End If
   End With
End Sub
