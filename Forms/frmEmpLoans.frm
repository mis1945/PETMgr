VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmEmpLoans 
   BorderStyle     =   0  'None
   Caption         =   "Employee Loans"
   ClientHeight    =   7275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   6297
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtAdjusted 
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
         Index           =   0
         Left            =   4245
         MaxLength       =   50
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3045
         Width           =   1620
      End
      Begin xrControl.xrButton xrButton1 
         Height          =   420
         Left            =   2775
         TabIndex        =   34
         Top             =   3045
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   741
         Caption         =   "Adjustment"
         AccessKey       =   "Adjustment"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
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
         Index           =   2
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   1125
         Width           =   1770
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
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "MC Loan"
         Top             =   1590
         Width           =   4800
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
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "M00111-000001"
         Top             =   450
         Width           =   1770
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
         Left            =   4245
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "4,000.00"
         Top             =   2055
         Width           =   1620
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
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "August 31, 2011"
         Top             =   2055
         Width           =   1620
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
         Index           =   9
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "3"
         Top             =   2580
         Width           =   555
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
         Left            =   4245
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2580
         Width           =   1620
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
         TabIndex        =   5
         Text            =   "1,333.33"
         Top             =   3045
         Width           =   1620
      End
      Begin VB.CheckBox chkField 
         Caption         =   "HOLD LOAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   105
         TabIndex        =   4
         Tag             =   "wt0;fb0"
         Top             =   120
         Width           =   1650
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   180
         X2              =   5850
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   135
         X2              =   5820
         Y1              =   390
         Y2              =   390
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
         Index           =   10
         Left            =   105
         TabIndex        =   20
         Top             =   1215
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Nm."
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
         TabIndex        =   18
         Top             =   1680
         Width           =   840
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
         TabIndex        =   17
         Top             =   540
         Width           =   960
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
         TabIndex        =   16
         Top             =   2145
         Width           =   675
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
         TabIndex        =   15
         Top             =   2145
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
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
         Left            =   105
         TabIndex        =   14
         Top             =   2700
         Width           =   495
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amort."
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
         Left            =   3390
         TabIndex        =   13
         Top             =   2700
         Width           =   630
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
         TabIndex        =   12
         Top             =   3135
         Width           =   780
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1185
         Tag             =   "et0;ht2"
         Top             =   540
         Width           =   1770
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   21
      Top             =   4425
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
      Picture         =   "frmEmpLoans.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   22
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
      Picture         =   "frmEmpLoans.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   23
      Top             =   4425
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
      Picture         =   "frmEmpLoans.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   24
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
      Picture         =   "frmEmpLoans.frx":166E
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
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   1
         Top             =   120
         Width           =   2415
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
         Left            =   7170
         MaxLength       =   50
         TabIndex        =   3
         Top             =   120
         Width           =   4965
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
         TabIndex        =   0
         Top             =   225
         Width           =   975
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
         TabIndex        =   2
         Top             =   225
         Width           =   1440
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5895
      Left            =   7590
      TabIndex        =   25
      Top             =   1260
      Width           =   6270
      _ExtentX        =   11060
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
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   4022
      BackColor       =   12632256
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
         Index           =   81
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   29
         Top             =   1260
         Width           =   4245
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
         TabIndex        =   28
         Top             =   795
         Width           =   4245
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
         Index           =   100
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   27
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2010
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
         Index           =   82
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1725
         Width           =   4245
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
         TabIndex        =   33
         Top             =   1350
         Width           =   615
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
         TabIndex        =   32
         Top             =   885
         Width           =   1440
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
         TabIndex        =   31
         Top             =   165
         Width           =   1200
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
         TabIndex        =   30
         Top             =   1815
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmEmpLoans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmEmpLoans"

Private oDriver As clsFormDriver
Private WithEvents oTrans As clsEmpLoan
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim psSelected() As String
Dim pnIndex As Integer
Dim pnRow As Integer
Dim pbSearched As Boolean

Private Sub chkField_Validate(Index As Integer, Cancel As Boolean)
   If oTrans.Detail(pnRow, "stransnox") = "" Then Exit Sub
   If Index = 7 Then
      oTrans.Detail(pnRow, "cholddedx") = chkField(Index).Value
   End If
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
      End If
      SetNextFocus
   Case 2   'Close
      Unload Me
   Case 3   'Cancel
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
   On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
   End If

   pbSearched = False
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
         If pbSearched Then SetNextFocus
      Case vbKeyUp
         If pbSearched Then SetPreviousFocus
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

   Set oTrans = New clsEmpLoan
   Set oTrans.AppDriver = oApp

   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

   Call InitGrid
   Call InitForm(0)
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
   
   pbSearched = (fnEdit <> 0)
   
   chkField(7).Enabled = False
   If fnEdit = 0 Then
      ClearFields
   Else
      MSFlexGrid1.SetFocus
   End If
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

   chkField(7).Value = 0
End Sub

Private Sub MSFlexGrid1_Click()
   Dim lnRow As Integer
   Dim loTxt As TextBox

   lnRow = MSFlexGrid1.Row - 1

   For Each loTxt In txtField
      Select Case loTxt.Index
         Case 80, 81, 82, 100
         Case 4, 5, 6
            loTxt = Format(IFNull(oTrans.Detail(lnRow, loTxt.Index), "0"), "#,##0.00")
         Case Else
            loTxt = IFNull(oTrans.Detail(lnRow, loTxt.Index), "")
      End Select
   Next
   chkField(7).Value = Val(IFNull(oTrans.Detail(lnRow, 7), 0))
   
   chkField(7).Enabled = True
   pnRow = lnRow
   
   With MSFlexGrid1
      .Col = 1
      .ColSel = .Cols - 1
   End With

   txtAdjusted(0) = Format(LoanAdjustment(oTrans.Detail(lnRow, 0)), "#,##0.00")

End Sub

Private Sub oTrans_DetailRetrieved(ByVal Row As Integer, ByVal Index As Variant)
   If Index = 7 Then
      chkField(Index) = IFNull(oTrans.Detail(Row, Index))
   Else
      txtField(Index).Text = IFNull(oTrans.Detail(Row, Index))
   End If
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant)
   Select Case Index
      Case 80
         txtField(Index).Text = oTrans.Master(Index)
         txtSearch(1).Text = txtField(Index).Text
      Case Else
         txtField(Index).Text = oTrans.Master(Index)
   End Select
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(100).Text = .Master(0)
      txtField(80).Text = .Master(80)
      txtField(81).Text = .Master(81)
      txtField(82).Text = .Master(82)
   End With
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCol As Integer
   
   With MSFlexGrid1
      For lnRow = 0 To oTrans.ItemCount - 1
         .Rows = lnRow + 2
         
         If .Rows > 16 Then
            .ColWidth(1) = 2000
         End If
         
         If IFNull(oTrans.Detail(lnRow, 7), 0) = 1 Then
            .Row = .Rows - 1
            For lnCol = 0 To .Cols - 1
               .Col = lnCol
               .CellBackColor = oApp.getColor("HT1")
            Next
         End If
         
         .TextMatrix(lnRow + 1, 0) = IFNull(oTrans.Detail(lnRow, 0), "")
         .TextMatrix(lnRow + 1, 1) = IFNull(oTrans.Detail(lnRow, 1), "")
         .TextMatrix(lnRow + 1, 2) = IFNull(oTrans.Detail(lnRow, 2), "")
         .TextMatrix(lnRow + 1, 3) = Format(oTrans.Detail(lnRow, "nLoanAmtx"), "##,###")
         
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
      .Cols = 4
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
      .TextMatrix(0, 0) = "Acct No"
      .TextMatrix(0, 1) = "Loan Desc"
      .TextMatrix(0, 2) = "Date"
      .TextMatrix(0, 3) = "Amount"
      
      'column width
      .ColWidth(0) = 1500
      .ColWidth(1) = 2300
      .ColWidth(2) = 1200
      .ColWidth(3) = 1200

      'column allinment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignRightCenter
      
      'set location
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
   pnRow = 0
End Sub

Private Function LoanAdjustment(ByVal fsAcctNmbr As String) As Currency
   Dim lsSQL As String
   Dim lors As Recordset
   
   lsSQL = "SELECT nTranAmtx" & _
          " FROM Employee_Loan_Adjustment" & _
          " WHERE sAcctNmbr = " & strParm(fsAcctNmbr) & _
            " AND cTranStat IN ('0', '1')"
   Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
   
   If Not lors.EOF Then
      LoanAdjustment = Format(lors("nTranAmtx"), "#,##0.00")
   End If
End Function

Private Sub ShowAdjustment(ByVal fsAcctNmbr As String)
   Dim loFrm As frmLoanAdjustment
   Set loFrm = New frmLoanAdjustment
   
   Load loFrm
   loFrm.AccountNumber = fsAcctNmbr
   loFrm.Show 1
End Sub

Private Sub xrButton1_Click()
   If Trim(txtField(0)) <> "" Then
      Call ShowAdjustment(txtField(0))
      txtAdjusted(0) = Format(LoanAdjustment(txtField(0)), "#,##0.00")
   End If
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

