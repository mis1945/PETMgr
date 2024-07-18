VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmEmpDeductions 
   BorderStyle     =   0  'None
   Caption         =   "Employee Deductions"
   ClientHeight    =   7275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   3570
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   1260
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   6297
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.CheckBox chkField 
         Caption         =   "HOLD DEDUCTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   105
         TabIndex        =   18
         Tag             =   "wt0;fb0"
         Top             =   120
         Width           =   2055
      End
      Begin VB.CheckBox chkField 
         Caption         =   "Taxable"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   4245
         TabIndex        =   17
         Tag             =   "wt0;fb0"
         Top             =   3098
         Width           =   1560
      End
      Begin VB.ComboBox cmbField 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   9
         ItemData        =   "frmEmpDeductions.frx":0000
         Left            =   1065
         List            =   "frmEmpDeductions.frx":0013
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   3045
         Width           =   2190
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   6
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "1333.33"
         Top             =   2595
         Width           =   2190
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   5
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "1333.33"
         Top             =   2130
         Width           =   2190
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   5460
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "3"
         Top             =   2130
         Width           =   975
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
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "August 31, 2011"
         Top             =   1590
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
         Height          =   420
         Index           =   2
         Left            =   4245
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "4,000.00"
         Top             =   1590
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
         Height          =   420
         Index           =   0
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   1
         Top             =   450
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
         Height          =   420
         Index           =   1
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "Health Insurance"
         Top             =   1125
         Width           =   5370
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   135
         X2              =   6380
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   135
         X2              =   6380
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interval"
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
         Index           =   10
         Left            =   105
         TabIndex        =   16
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "POSTED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4095
         TabIndex        =   14
         Tag             =   "eb0;et0"
         Top             =   615
         Width           =   2115
      End
      Begin VB.Shape Shape3 
         Height          =   420
         Index           =   0
         Left            =   3990
         Top             =   525
         Width           =   2340
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   4035
         Top             =   570
         Width           =   2250
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1180
         Tag             =   "et0;ht2"
         Top             =   540
         Width           =   2190
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
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
         Index           =   9
         Left            =   105
         TabIndex        =   13
         Top             =   2685
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mo. Dedn"
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
         Index           =   8
         Left            =   105
         TabIndex        =   11
         Top             =   2220
         Width           =   900
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Payments"
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
         Index           =   7
         Left            =   3390
         TabIndex        =   9
         Top             =   2220
         Width           =   1980
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Pay"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   3390
         TabIndex        =   5
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Index           =   5
         Left            =   105
         TabIndex        =   3
         Top             =   540
         Width           =   960
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction"
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
         Left            =   105
         TabIndex        =   2
         Top             =   1215
         Width           =   870
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   19
      Top             =   5055
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
      Picture         =   "frmEmpDeductions.frx":0050
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   20
      Top             =   4410
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmEmpDeductions.frx":07CA
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   21
      Top             =   5055
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
      Picture         =   "frmEmpDeductions.frx":0F44
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   22
      Top             =   4410
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
      Picture         =   "frmEmpDeductions.frx":16BE
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   660
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   1164
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
         Index           =   1
         Left            =   7170
         MaxLength       =   50
         TabIndex        =   24
         Top             =   120
         Width           =   4965
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   23
         Top             =   120
         Width           =   2415
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
         Index           =   25
         Left            =   5145
         TabIndex        =   26
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Control No."
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
         Left            =   90
         TabIndex        =   25
         Top             =   225
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5895
      Left            =   8190
      TabIndex        =   27
      Top             =   1260
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   10398
      _Version        =   393216
      FocusRect       =   0
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
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   4875
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   4022
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
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
         Index           =   82
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   31
         Top             =   1725
         Width           =   4815
      End
      Begin VB.TextBox txtField 
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
         Index           =   11
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   30
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2415
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
         Index           =   80
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   29
         Top             =   795
         Width           =   4815
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
         Index           =   81
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   28
         Top             =   1260
         Width           =   4815
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         Width           =   2415
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
         TabIndex        =   32
         Top             =   1350
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmEmpDeductions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmEmpDeductions"

Private oDriver As clsFormDriver
Private WithEvents oTrans As clsEmpDeduction
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim psSelected() As String
Dim pnIndex As Integer
Dim pnRow As Integer

Private Sub chkField_Validate(Index As Integer, Cancel As Boolean)
   If Index = 7 Then
      oTrans.Detail(pnRow, "cHoldDedx") = chkField(Index).Value
   End If
End Sub

Private Sub cmbField_Change(Index As Integer)
   oTrans.Detail(pnRow, Index) = IIf(cmbField(Index).ListIndex >= 0, cmbField(Index).ListIndex, Null)
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRow As Integer
   Dim loTxt As TextBox
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc
   Select Case Index
   Case 0   'save
      If oTrans.SaveTransaction Then
         MsgBox "Transaction Successfully Updated!!!", vbInformation, "Confirm"
         oTrans.NewTransaction
         ClearFields
         InitGrid
         InitForm 0
         GoTo endWithFocus
      Else
         MsgBox "Unable to Update Transaction!!!", vbCritical, "Warning"
      End If
   Case 1   'search
      If pnIndex = 0 Or pnIndex = 1 Then
         If txtSearch(pnIndex) = "" Then
            txtSearch(pnIndex).SetFocus
            Exit Sub
         End If
         If pnIndex = 0 Then
            If oTrans.SearchTransaction(txtSearch(pnIndex).Text, False) Then
               InitForm 1
               Call LoadMaster
               Call LoadDetail
            End If
         ElseIf pnIndex = 1 Then
            If oTrans.SearchTransaction(txtSearch(pnIndex).Text) Then
               InitForm 1
               Call LoadMaster
               Call LoadDetail
            End If
         End If
         Call InitForm(1)
         SetNextFocus
      End If
   Case 2   'Close
      Unload Me
   Case 3   'cancel
      lnRep = MsgBox("Are you certain to cancel modifications?", vbQuestion + vbYesNo, "Notice")

      If lnRep = vbYes Then
         oTrans.NewTransaction
         ClearFields
         InitGrid
         InitForm 0
         GoTo endWithFocus
      End If
   End Select

endProc:
   Exit Sub
endWithFocus:
   txtSearch(0).SetFocus
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   'On Error Goto errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
   End If

   If txtSearch(0).Enabled Then txtSearch(0).SetFocus
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
   On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me

   Set oTrans = New clsEmpDeduction
   Set oTrans.AppDriver = oApp

   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   Call InitGrid
   ClearFields
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub InitForm(ByVal fnEdit As Integer)
   Dim lnCtr As Integer

   MSFlexGrid1.Enabled = Not (fnEdit = 0)
   cmdButton(0).Visible = Not (fnEdit = 0)
   cmdButton(3).Visible = Not (fnEdit = 0)
   
   txtSearch(0).Enabled = (fnEdit = 0)
   txtSearch(1).Enabled = (fnEdit = 0)
   cmdButton(1).Visible = (fnEdit = 0)
   cmdButton(2).Visible = (fnEdit = 0)
   
   chkField(7).Enabled = False
   If fnEdit = 0 Then ClearFields
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox

   For Each loTxt In txtField
      loTxt = ""
   Next

   txtSearch(0).Text = ""
   txtSearch(1).Text = ""
   With MSFlexGrid1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
   End With

   cmbField(9).ListIndex = -1
   chkField(7).Value = 0
   chkField(10).Value = 0
   Label2.Caption = ""
End Sub

Private Sub MSFlexGrid1_Click()
   Dim lnRow As Integer
   Dim loTxt As TextBox

   lnRow = MSFlexGrid1.Row - 1
   With oTrans
      For Each loTxt In txtField
         Select Case loTxt.Index
            Case 11, 80, 81, 82
            Case Else
               loTxt = IFNull(oTrans.Detail(lnRow, loTxt.Index), "")
         End Select
      Next
   
      cmbField(9).ListIndex = IIf(IsNull(.Detail(lnRow, 9)), -1, .Detail(lnRow, 9))
      chkField(10).Value = IIf(IsNull(.Detail(lnRow, 10)), 0, .Detail(lnRow, 9))
      chkField(7).Value = .Detail(lnRow, 7)
      Label2.Caption = TransStat(.Detail(lnRow, 8))
   End With
   
   chkField(7).Enabled = True
   pnRow = lnRow
   
   With MSFlexGrid1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Row As Integer, ByVal Index As Variant)
   txtField(Index) = IFNull(oTrans.Detail(pnRow, Index), "")
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant)
   Select Case Index
      Case 0
      Case 80
         txtField(Index).Text = oTrans.Master(Index)
         txtSearch(1).Text = txtField(Index).Text
      Case Else
         txtField(Index).Text = oTrans.Master(Index)
   End Select
End Sub

Private Sub LoadMaster()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   With oTrans
      txtField(11).Text = .Master(0)
      txtField(80).Text = .Master(80)
      txtField(81).Text = .Master(81)
      txtField(82).Text = .Master(82)
   End With
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer

   With MSFlexGrid1
      For lnRow = 0 To oTrans.ItemCount - 1
         .Rows = lnRow + 2
         
         If .Rows > 16 Then
            .ColWidth(1) = 3245
         End If
         
         .TextMatrix(lnRow + 1, 0) = lnRow + 1
         .TextMatrix(lnRow + 1, 1) = IFNull(oTrans.Detail(lnRow, 1), "")
         .TextMatrix(lnRow + 1, 2) = oTrans.Detail(lnRow, "nAmountxx")
      Next
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

   oDriver.ColumnIndex = Index
   pnIndex = Index
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

   oDriver.ColumnIndex = Index
   pnIndex = Index
End Sub
Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If txtSearch(pnIndex) = "" Then
         txtSearch(pnIndex).SetFocus
         Exit Sub
   End If
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 0
         If oTrans.SearchTransaction(txtSearch(Index).Text, True) Then
            'Just added InitForm 1 if like frmEmpBenefits
            InitForm 1
            LoadMaster
            LoadDetail
         End If
      Case 1
         If oTrans.SearchTransaction(txtSearch(Index).Text, False) Then
            'Just added InitForm 1 if like frmEmpBenefits
            InitForm 1
            LoadMaster
            LoadDetail
         End If
      End Select
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub
Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Cols = 3
      .Rows = 2
      
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
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Date"
      .TextMatrix(0, 2) = "Amount"
      
      'column width
      .ColWidth(0) = 550
      .ColWidth(1) = 3500
      .ColWidth(2) = 1582

      'column allinment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignRightCenter
      
      'set location
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
   pnRow = 0
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

