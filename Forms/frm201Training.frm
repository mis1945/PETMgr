VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrcontrol.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm201Training 
   BorderStyle     =   0  'None
   Caption         =   "Employee 201 File (Training)"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2400
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4233
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
         Height          =   400
         Index           =   80
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   1845
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   1410
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
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   975
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
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   540
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
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   105
         Width           =   4515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year of Service"
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
         Left            =   2610
         TabIndex        =   9
         Top             =   1920
         Width           =   1305
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
         Left            =   2610
         TabIndex        =   8
         Top             =   1485
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
         TabIndex        =   7
         Top             =   1050
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
         TabIndex        =   6
         Top             =   615
         Width           =   615
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
         TabIndex        =   5
         Top             =   180
         Width           =   510
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4995
      Left            =   105
      TabIndex        =   10
      Top             =   3000
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   8811
      _Version        =   393216
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   9105
      TabIndex        =   11
      Top             =   4695
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
      Picture         =   "frm201Training.frx":0000
   End
   Begin VB.Image Image2 
      Height          =   2400
      Left            =   6165
      Picture         =   "frm201Training.frx":077A
      Stretch         =   -1  'True
      Top             =   555
      Width           =   2550
   End
End
Attribute VB_Name = "frm201Training"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frm201Training"

Private oSkin As clsFormSkin
Private oTrans As cls201Info

Private bLoaded As Boolean
Dim pnIndex As Integer

Property Set Emp201(Value As cls201Info)
   Set oTrans = Value
End Property

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
      Call LoadMaster
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
   On Error GoTo errProc

   CenterChildForm mdiMain, Me

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

Sub LoadMaster()
   Dim loTxt As TextBox
   Dim lnCtr As Integer

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
      If oTrans.Training_Detail(0, "nRecCount") = 0 Then
         .Rows = 2
      Else
         .Rows = oTrans.Training_Detail(0, "nRecCount") + 1
      End If

      For lnCtr = 1 To .Rows - 1
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = IFNull(oTrans.Training_Detail(lnCtr - 1, "dDateFrom")) & _
                                 IIf(IFNull(oTrans.Training_Detail(lnCtr - 1, "dDateFrom")) = IFNull(oTrans.Training_Detail(lnCtr - 1, "dDateThru")), "", " - " & oTrans.Training_Detail(lnCtr - 1, "dDateThru"))
         .TextMatrix(lnCtr, 2) = IFNull(oTrans.Training_Detail(lnCtr - 1, "sTraining"))
         .TextMatrix(lnCtr, 3) = IFNull(oTrans.Training_Detail(lnCtr - 1, "sVenuexxx"))
      Next
   End With
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Rows = 2
      .Cols = 4

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
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Date"
      .TextMatrix(0, 2) = "Training/Seminar"
      .TextMatrix(0, 3) = "Venue"

      .RowHeightMin = 320

      'Set Column Width
      .ColWidth(0) = 520
      .ColWidth(1) = 2000
      .ColWidth(2) = 2990
      .ColWidth(3) = 2990

      'column allinment
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter

      'set location
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1

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



