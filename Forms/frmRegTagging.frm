VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmReqTagging 
   BorderStyle     =   0  'None
   Caption         =   "MC Registration Requirement Tagging"
   ClientHeight    =   7320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   570
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1005
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtSearch 
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
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   1
         Top             =   90
         Width           =   4815
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
         Index           =   25
         Left            =   105
         TabIndex        =   0
         Top             =   150
         Width           =   615
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6060
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1155
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10689
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5910
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   10425
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   8880
      TabIndex        =   3
      Top             =   3285
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
      Picture         =   "frmRegTagging.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   8880
      TabIndex        =   4
      Top             =   3915
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
      Picture         =   "frmRegTagging.frx":077A
   End
End
Attribute VB_Name = "frmReqTagging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmEmpAppraisal"
Private oDriver As clsFormDriver
Private oSkin As clsFormSkin

Private Sub Form_Load()
10   Dim lsOldProc As String

20   lsOldProc = "Form_Load"
30   On Error GoTo errProc

40   CenterChildForm mdiMain, Me

50   Set oDriver = New clsFormDriver
60   Set oDriver.AppDriver = oApp
70   Set oDriver.MainForm = Me

80   Set oSkin = New clsFormSkin
90   Set oSkin.AppDriver = oApp
100   Set oSkin.Form = Me
110   oSkin.ApplySkin xeFormTransMaintenance
   
120   Call InitForm

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitForm()
   Dim lsOldProc As String
   Dim loTxt As TextBox
   Dim loStar As StarRating
10             Dim lnCtr As Integer
   
   lsOldProc = "initForm"
   On Error GoTo errProc
   
       With MSFlexGrid1
          .Cols = 6
          .Rows = 4
      
          .Row = 0
      
      'column alignment
          For lnCtr = 0 To .Cols - 1
             .Col = lnCtr
            .CellFontBold = True
            .CellAlignment = flexAlignCenterCenter
         Next
      
         .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Account No."
         .TextMatrix(0, 2) = "Client Name"
         .TextMatrix(0, 3) = "Stencil"
         .TextMatrix(0, 4) = "SI"
        .TextMatrix(0, 5) = "ID"

         .Row = 1
      
      'column width
         .ColWidth(0) = 500
         .ColWidth(1) = 1500
         .ColWidth(2) = 2460
         .ColWidth(3) = 1300
         .ColWidth(4) = 1300
         .ColWidth(5) = 1300

      
      'column allinment
         .ColAlignment(0) = flexAlignLeftCenter
         .ColAlignment(1) = flexAlignLeftCenter
         .ColAlignment(2) = flexAlignCenterCenter
         .ColAlignment(3) = flexAlignCenterCenter
         .ColAlignment(4) = flexAlignCenterCenter
         .ColAlignment(5) = flexAlignCenterCenter

         .TextMatrix(1, 0) = "1"
         .TextMatrix(1, 1) = "M001130001"
         .TextMatrix(1, 2) = "Adversalo, Rex S."
         .TextMatrix(1, 3) = "YES"
         .TextMatrix(1, 4) = "YES"
         .TextMatrix(1, 5) = "YES"

         .TextMatrix(2, 0) = "2"
         .TextMatrix(2, 1) = "M001130002"
         .TextMatrix(2, 2) = "Yambao, Jeffrey T."
         .TextMatrix(2, 3) = "YES"
         .TextMatrix(2, 4) = "YES"
         .TextMatrix(2, 5) = "NO"
         
         .TextMatrix(3, 0) = "3"
         .TextMatrix(3, 1) = "M001130003"
         .TextMatrix(3, 2) = "Sayson, Marlon A."
         .TextMatrix(3, 3) = "NO"
         .TextMatrix(3, 4) = "NO"
         .TextMatrix(3, 5) = "NO"
         
      End With
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub



Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
10       With oApp
20          .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
30          If bEnd Then
40             .xShowError
50             End
60          Else
70             With Err
80                .Raise .Number, .Source, .Description
90             End With
100         End If
110      End With
End Sub

