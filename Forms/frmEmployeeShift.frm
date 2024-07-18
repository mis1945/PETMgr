VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrcontrol.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmployeeShift 
   BorderStyle     =   0  'None
   Caption         =   "Employee Shift"
   ClientHeight    =   5805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2835
      Left            =   5730
      TabIndex        =   0
      Top             =   510
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   5001
      _Version        =   393216
      Rows            =   8
      Cols            =   3
      RowHeightMin    =   338
      FocusRect       =   0
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
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2280
      Left            =   5730
      Tag             =   "wt0;fb0"
      Top             =   3360
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   4022
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox txtOthers 
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
         Index           =   82
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1725
         Width           =   4245
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2010
      End
      Begin VB.TextBox txtOthers 
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   4
         Top             =   795
         Width           =   4245
      End
      Begin VB.TextBox txtOthers 
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
         Index           =   81
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1260
         Width           =   4245
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Left            =   105
         TabIndex        =   7
         Top             =   1815
         Width           =   1005
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
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
         Left            =   105
         TabIndex        =   1
         Top             =   165
         Width           =   1200
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Left            =   105
         TabIndex        =   3
         Top             =   885
         Width           =   1440
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1725
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   2010
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
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
         TabIndex        =   5
         Top             =   1350
         Width           =   615
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1680
      Left            =   1515
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   2963
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.CheckBox chkField 
         Caption         =   "Special"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1065
         TabIndex        =   13
         Tag             =   "et0;fb0"
         Top             =   1125
         Width           =   1575
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
         Left            =   1035
         MaxLength       =   50
         TabIndex        =   12
         Top             =   645
         Width           =   3030
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
         Left            =   1035
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Saturday"
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
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
         Left            =   105
         TabIndex        =   11
         Top             =   735
         Width           =   390
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Left            =   105
         TabIndex        =   9
         Top             =   270
         Width           =   345
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   1140
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
      Picture         =   "frmEmployeeShift.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   510
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
      Picture         =   "frmEmployeeShift.frx":077A
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   3435
      Left            =   1515
      Tag             =   "wt0;fb0"
      Top             =   2220
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   6059
      BackColor       =   12632256
      Enabled         =   0   'False
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
         Index           =   9
         Left            =   2625
         TabIndex        =   21
         Top             =   945
         Width           =   1245
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
         Index           =   8
         Left            =   1305
         TabIndex        =   20
         Top             =   945
         Width           =   1245
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
         Index           =   7
         Left            =   2625
         TabIndex        =   25
         Top             =   2055
         Width           =   1245
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
         Index           =   6
         Left            =   2625
         TabIndex        =   18
         Text            =   "05:35 pm"
         Top             =   480
         Width           =   1245
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
         Index           =   5
         Left            =   1305
         TabIndex        =   17
         Text            =   "08:30 am"
         Top             =   480
         Width           =   1245
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
         Index           =   10
         Left            =   2625
         TabIndex        =   23
         Top             =   1590
         Width           =   1245
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
         Index           =   11
         Left            =   2625
         TabIndex        =   27
         Top             =   2520
         Width           =   1245
      End
      Begin VB.CheckBox chkField 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Flexi-Time"
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
         Index           =   12
         Left            =   2625
         TabIndex        =   28
         Tag             =   "et0;fb0"
         Top             =   3015
         Width           =   1335
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
         Index           =   13
         Left            =   3030
         TabIndex        =   15
         Top             =   135
         Width           =   390
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Break"
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
         Left            =   1935
         TabIndex        =   24
         Top             =   2160
         Width           =   510
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
         Index           =   3
         Left            =   1800
         TabIndex        =   14
         Top             =   135
         Width           =   180
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Work Hrs."
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
         Left            =   1560
         TabIndex        =   22
         Top             =   1695
         Width           =   885
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
         Left            =   1440
         TabIndex        =   26
         Top             =   2625
         Width           =   1005
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "1ST Qtr"
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
         Left            =   255
         TabIndex        =   16
         Tag             =   "et0;fb0"
         Top             =   540
         Width           =   705
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "2ND Qtr"
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
         Left            =   255
         TabIndex        =   19
         Tag             =   "et0;fb0"
         Top             =   990
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEmployeeShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmEmployeeShift"

Private oSkin As clsFormSkin
Private WithEvents loTrans As clsEmployeeShift
Attribute loTrans.VB_VarHelpID = -1

Property Let EmployeeID(ByVal fsEmployID As String)
   Dim lsOldProc As String

   lsOldProc = "EmployeeID"
   On Error GoTo errProc
   
   Set loTrans = New clsEmployeeShift
   Set loTrans.AppDriver = oApp
   loTrans.Branch = oApp.BranchCode
   Call loTrans.InitTransaction
   Call loTrans.NewTransaction(fsEmployID)

endProc:
   Exit Property
errProc:
   ShowError lsOldProc & "( " & " )", True

End Property

Private Sub chkField_Click(Index As Integer)
   If Index = 3 Then
      'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = chkField(3).Value
      loTrans.Detail(MSFlexGrid1.Row, 3) = chkField(3).Value
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = loTrans.Detail(MSFlexGrid1.Row, 3)
      'loTrans.Detail(MSFlexGrid1.Row - 1, 3) = chkField(3).Value
   End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0 'Cancel
      If MsgBox("Do you really want to undo the update?", vbInformation + vbYesNo, "Shift  Adjustment Confirmation") = vbYes Then
         Unload Me
      End If
   Case 2 'Save
      'loTrans.Detail(MSFlexGrid1.Row - 1, "cSpecialx") = chkField(3).Value
      loTrans.Detail(MSFlexGrid1.Row, "cSpecialx") = chkField(3).Value
      If loTrans.SaveTransaction(True) = True Then
         MsgBox "Adjustment save successfully!", vbInformation, "Shift Adjustment Confirmation"
         Unload Me
      End If
   End Select
End Sub

Private Sub Form_Activate()
   LoadData
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

   InitGrid
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Cols = 4
      .Rows = 8
      
      .Clear
      
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Row = 1
      
      'column title
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Day"
      .TextMatrix(0, 2) = "Shift"
      .TextMatrix(0, 3) = "Special"

      'column width
      .ColWidth(0) = 490
      .ColWidth(1) = 1200
      .ColWidth(2) = 3300
      .ColWidth(3) = 937

      'column allinment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      
      For lnCtr = 1 To 7
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = WeekdayName(lnCtr)
      Next
      
      'set location
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub LoadData()
   Dim loTxt As TextBox
   Dim lnCtr As Integer
   
   'Load Employee Info
   For Each loTxt In txtOthers
      loTxt = loTrans.Master(loTxt.Index)
   Next
   
   'Load the schedule in our grid
   With MSFlexGrid1
      For lnCtr = 1 To 7
         .TextMatrix(lnCtr, 2) = loTrans.Detail(lnCtr, 4)
         .TextMatrix(lnCtr, 3) = loTrans.Detail(lnCtr, 3)
      Next
   End With
   
   'Load the Shift pointed by the MSFlexGrid1.Row
   For Each loTxt In txtfield
      If loTxt.Index = 1 Then
         loTxt = WeekdayName(loTrans.Detail(MSFlexGrid1.Row, loTxt.Index))
      Else
         loTxt = loTrans.Detail(MSFlexGrid1.Row, loTxt.Index)
      End If
   Next
      
   chkField(3).Value = loTrans.Detail(MSFlexGrid1.Row, "cSpecialx")
   chkField(12).Value = loTrans.Detail(MSFlexGrid1.Row, "cFlexiTym")
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

Private Sub loTrans_DetailRetrieved(ByVal Row As Integer, ByVal Index As Variant, ByVal Value As Variant)
   Dim loTxt As TextBox
   
   If Index = 4 Then
      For Each loTxt In txtfield
         If loTxt.Index = 1 Then
            loTxt = WeekdayName(loTrans.Detail(Row, loTxt.Index))
         Else
            loTxt = loTrans.Detail(Row, loTxt.Index)
         End If
      Next
      
      chkField(3) = loTrans.Detail(Row, "cSpecialx")
      chkField(12) = loTrans.Detail(Row, "cFlexiTym")
      
      MSFlexGrid1.TextMatrix(Row, 2) = loTrans.Detail(Row, 4)
      MSFlexGrid1.TextMatrix(Row, 3) = loTrans.Detail(Row, 3)
 
   End If
End Sub

Private Sub MSFlexGrid1_Click()
   Dim loTxt As TextBox
   
   For Each loTxt In txtfield
      If loTxt.Index = 1 Then
         loTxt = WeekdayName(loTrans.Detail(MSFlexGrid1.Row, loTxt.Index))
      Else
         loTxt = loTrans.Detail(MSFlexGrid1.Row, loTxt.Index)
      End If
   Next
   
   chkField(3) = loTrans.Detail(MSFlexGrid1.Row, "cSpecialx")
   chkField(12) = loTrans.Detail(MSFlexGrid1.Row, "cFlexiTym")
   'loTrans.Detail(MSFlexGrid1.Row, "cSpecialx") = chkField(3).Value
   
   txtfield(4).SetFocus
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 4 Then
      If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
         Call loTrans.SearchDetail(MSFlexGrid1.Row, Index, txtfield(Index))
      End If
   End If
End Sub

