VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrcontrol.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm201Master 
   BorderStyle     =   0  'None
   Caption         =   "201 File"
   ClientHeight    =   7545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3150
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   5556
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
         Height          =   400
         Index           =   80
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   1230
         Width           =   1965
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
         Height          =   400
         Index           =   5
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   795
         Width           =   1965
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
         Height          =   400
         Index           =   4
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   2610
         Width           =   4515
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
         Height          =   400
         Index           =   3
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   2175
         Width           =   4515
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
         Height          =   810
         Index           =   2
         Left            =   1410
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "frm201Master.frx":0000
         Top             =   1230
         Width           =   4515
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
         Height          =   400
         Index           =   1
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   795
         Width           =   4515
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
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "August 16, 2011"
         Top             =   150
         Width           =   2190
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Year"
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
         Left            =   6255
         TabIndex        =   13
         Top             =   1310
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Hired"
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
         Left            =   6255
         TabIndex        =   11
         Top             =   870
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
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
         Left            =   240
         TabIndex        =   9
         Top             =   2690
         Width           =   705
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   2255
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1305
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   240
         TabIndex        =   3
         Top             =   875
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   240
         TabIndex        =   2
         Top             =   225
         Width           =   180
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3600
      Left            =   105
      TabIndex        =   14
      Top             =   3765
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   6350
      _Version        =   393216
      Rows            =   11
      Cols            =   3
      FocusRect       =   0
      SelectionMode   =   1
      MergeCells      =   1
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   9900
      TabIndex        =   15
      Top             =   3975
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
      Picture         =   "frm201Master.frx":0010
   End
End
Attribute VB_Name = "frm201Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frm201Master"

Private oSkin As clsFormSkin
Private p_oPayMiscx As clsPayMisc
Private oTrans As cls201Info

Private bLoaded As Boolean
Dim pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Unload Me
End Sub

Private Sub Form_Activate()

   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   On Error GoTo errProc

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
   If KeyCode = vbKeyF12 And oApp.UserLevel = xeEngineer Then
   End If
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New cls201Info
   Set oTrans.AppDriver = oApp

   oTrans.InitRecord

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance

   InitGrid

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Sub LoadMaster(ByVal fsClientID As String)
   Dim loTxt As TextBox
   Dim lnCtr As Integer

   Call oTrans.OpenRecord(fsClientID)

   For Each loTxt In txtfield
      If loTxt.Index = 80 Then
         If IsDate(oTrans.Master("dHiredxxx")) Then
            loTxt = Round(DateDiff("m", oTrans.Master("dHiredxxx"), Now()) / 12, 2)
         Else
            loTxt = "0.00"
         End If
      Else
         loTxt = IFNull(oTrans.Master(loTxt.Index))
      End If
   Next

   With MSFlexGrid1
      For lnCtr = 1 To .Cols - 1
         'Load Merit
         .TextMatrix(1, lnCtr) = Val(oTrans.Merit(.TextMatrix(0, lnCtr), "nCountxxx"))
         'Load Demerit
         .TextMatrix(2, lnCtr) = Val(oTrans.Demerit(.TextMatrix(0, lnCtr), "nCountxxx"))
         'Load IOC
         .TextMatrix(3, lnCtr) = Val(oTrans.IOC(.TextMatrix(0, lnCtr), "nCountxxx"))
         'Load Training
         .TextMatrix(4, lnCtr) = Val(oTrans.Training(.TextMatrix(0, lnCtr), "nCountxxx"))
         'Load Seminar
         .TextMatrix(5, lnCtr) = Val(oTrans.Seminar(.TextMatrix(0, lnCtr), "nCountxxx"))
         'Load Promotion
         .TextMatrix(6, lnCtr) = Val(oTrans.Promotion(.TextMatrix(0, lnCtr), "nCountxxx"))
         'Load Salary Adjustment
         .TextMatrix(7, lnCtr) = Val(oTrans.Salary(.TextMatrix(0, lnCtr), "nCountxxx"))
         'Load Absence
         .TextMatrix(8, lnCtr) = Val(IFNull(oTrans.Absence(.TextMatrix(0, lnCtr), "nCountxxx")))
         'Load Tardiness
         .TextMatrix(9, lnCtr) = Val(IFNull(oTrans.Tardiness(.TextMatrix(0, lnCtr), "nCountxxx")))
         'Load Undertime
         .TextMatrix(10, lnCtr) = Val(IFNull(oTrans.UnderTime(.TextMatrix(0, lnCtr), "nCountxxx")))
      Next
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

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Rows = 11
      .Cols = 6

      .Row = 0
      .RowHeight(0) = 320

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next

      .MergeRow(0) = True

      'Set Column Header
      .Row = 1
      .TextMatrix(0, 0) = "Particular"
      .TextMatrix(0, 1) = Year(Date)
      .TextMatrix(0, 2) = Year(Date) - 1
      .TextMatrix(0, 3) = Year(Date) - 2
      .TextMatrix(0, 4) = Year(Date) - 3
      .TextMatrix(0, 5) = Year(Date) - 4

      .RowHeightMin = 320

      'Set Column Width
      .ColWidth(0) = 3270
      For lnCtr = 1 To .Cols - 1
         .ColWidth(lnCtr) = 1250
      Next

      'column allinment
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignLeftCenter
      .ColAlignment(5) = flexAlignLeftCenter

      'set location
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1

      .TextMatrix(1, 0) = "Merit"
      .TextMatrix(2, 0) = "Demerit"
      .TextMatrix(3, 0) = "IOC"
      .TextMatrix(4, 0) = "Training"
      .TextMatrix(5, 0) = "Seminar"
      .TextMatrix(6, 0) = "Promotion"
      .TextMatrix(7, 0) = "Salary Adjustment"
      .TextMatrix(8, 0) = "Absence"
      .TextMatrix(9, 0) = "Tardiness"
      .TextMatrix(10, 0) = "Undertime"
   End With
End Sub

Private Sub MSFlexGrid1_DblClick()
   Dim loFrm As Form
   Select Case MSFlexGrid1.Row
   Case 1   'Merit
   Case 2   'Demerit
   Case 3   'IOC
      Set loFrm = New frm201IOC
      Set loFrm.Emp201 = oTrans
      loFrm.Show 1
   Case 4   'Training
      Set loFrm = New frm201Training
      Set loFrm.Emp201 = oTrans
      loFrm.Show 1
   Case 5   'Seminar
   Case 6   'Promotion
      Set loFrm = New frm201Promotion
      Set loFrm.Emp201 = oTrans
      loFrm.Show 1
   Case 7   'Salary Adjustment
      Set loFrm = New frm201Salary
      Set loFrm.Emp201 = oTrans
      loFrm.Show 1
   Case 8   'Absence
   Case 9   'Tardiness
   Case 10  'Undertime
   End Select
End Sub
