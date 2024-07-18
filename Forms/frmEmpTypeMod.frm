VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmEmpTypeMod 
   BorderStyle     =   0  'None
   Caption         =   "Payroll Computation"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1770
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3122
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.CheckBox chkField 
         Caption         =   "Main Office"
         Height          =   345
         Left            =   2415
         TabIndex        =   6
         Tag             =   "wt0;fb0"
         Top             =   480
         Width           =   1665
      End
      Begin VB.OptionButton optField 
         Caption         =   "Level A "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   795
         TabIndex        =   3
         Tag             =   "wt0;fb0"
         Top             =   1260
         Width           =   1005
      End
      Begin VB.OptionButton optField 
         Caption         =   "Regular"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   795
         TabIndex        =   2
         Tag             =   "wt0;fb0"
         Top             =   870
         Width           =   1005
      End
      Begin VB.OptionButton optField 
         Caption         =   "Trainees"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   795
         TabIndex        =   1
         Tag             =   "wt0;fb0"
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Type:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         Top             =   120
         Width           =   1245
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   1
      Left            =   4695
      TabIndex        =   4
      Top             =   1005
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   741
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
      Picture         =   "frmEmpTypeMod.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   420
      Index           =   0
      Left            =   4695
      TabIndex        =   5
      Top             =   555
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   741
      Caption         =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEmpTypeMod.frx":077A
   End
End
Attribute VB_Name = "frmEmpTypeMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmEmpTypeMod"

Private loAppDrivr As clsAppDriver
Private loSkin As clsFormSkin

Private pbCancel As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set loAppDrivr = oAppDriver
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancel
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer

   Select Case Index
   Case 0, 1
      pbCancel = Index = 1
      Me.Hide
   End Select
End Sub

Private Sub Form_Load()
   Set loSkin = New clsFormSkin
   Set loSkin.AppDriver = loAppDrivr
   Set loSkin.Form = Me
   loSkin.ApplySkin xeFormTransDetail
   
   optField(0).Value = True
   chkField.Value = 1
   
   CenterChildForm mdiMain, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set loSkin = Nothing
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

Private Sub optField_Click(Index As Integer)
   Select Case Index
   Case 0
'      chkField.Enabled = True
      chkField.Enabled = False
   Case 1, 2
      chkField.Value = 1
      chkField.Enabled = False
   End Select
End Sub
