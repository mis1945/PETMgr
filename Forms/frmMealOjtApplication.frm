VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMealOjtApplication 
   BorderStyle     =   0  'None
   Caption         =   "OJT Application"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox check 
      Caption         =   "Saturday"
      Height          =   195
      Index           =   6
      Left            =   5595
      TabIndex        =   15
      Tag             =   "wt0;fb0"
      Top             =   5460
      Width           =   1095
   End
   Begin VB.CheckBox check 
      Caption         =   "Wednesday"
      Height          =   195
      Index           =   3
      Left            =   6690
      TabIndex        =   12
      Tag             =   "wt0;fb0"
      Top             =   5145
      Width           =   1335
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   6930
      Left            =   1545
      Tag             =   "wt0;fb0"
      Top             =   495
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   12224
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.CheckBox check 
         Caption         =   "Friday"
         Height          =   195
         Index           =   5
         Left            =   2805
         TabIndex        =   14
         Tag             =   "wt0;fb0"
         Top             =   4950
         Width           =   1095
      End
      Begin VB.CheckBox check 
         Caption         =   "Thursday"
         Height          =   195
         Index           =   4
         Left            =   1635
         TabIndex        =   13
         Tag             =   "wt0;fb0"
         Top             =   4950
         Width           =   1095
      End
      Begin VB.CheckBox check 
         Caption         =   "Tuesday"
         Height          =   195
         Index           =   2
         Left            =   3990
         TabIndex        =   11
         Tag             =   "wt0;fb0"
         Top             =   4650
         Width           =   945
      End
      Begin VB.CheckBox check 
         Caption         =   "Monday"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   10
         Tag             =   "wt0;fb0"
         Top             =   4650
         Width           =   885
      End
      Begin VB.CheckBox check 
         Caption         =   "Sunday"
         Height          =   195
         Index           =   0
         Left            =   1635
         TabIndex        =   9
         Tag             =   "wt0;fb0"
         Top             =   4635
         Width           =   885
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
         Index           =   7
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   7
         Top             =   4110
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
         Index           =   5
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   5
         Top             =   3210
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
         Index           =   3
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2115
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
         Index           =   4
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2565
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
         Index           =   1
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1215
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
         Index           =   0
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "M00111-000021"
         Top             =   570
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
         Index           =   2
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1665
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
         Index           =   6
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   6
         Top             =   3660
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
         Height          =   1440
         Index           =   8
         Left            =   1620
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   5325
         Width           =   4815
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "UNKNOWN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4455
         TabIndex        =   33
         Tag             =   "eb0;et0"
         Top             =   195
         Width           =   1935
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule:"
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
         Left            =   135
         TabIndex        =   30
         Top             =   4605
         Width           =   870
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1725
         Tag             =   "et0;ht2"
         Top             =   660
         Width           =   2415
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks:"
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
         Left            =   105
         TabIndex        =   24
         Top             =   5280
         Width           =   840
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Course:"
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
         TabIndex        =   23
         Top             =   3750
         Width           =   675
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   90
         TabIndex        =   22
         Top             =   1755
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         TabIndex        =   21
         Top             =   660
         Width           =   1485
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
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
         Left            =   75
         TabIndex        =   20
         Top             =   1305
         Width           =   465
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From:"
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
         Left            =   90
         TabIndex        =   19
         Top             =   2205
         Width           =   975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Of Hour"
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
         Left            =   90
         TabIndex        =   18
         Top             =   2655
         Width           =   1020
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School:"
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
         TabIndex        =   17
         Top             =   3300
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department:"
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
         Left            =   105
         TabIndex        =   16
         Top             =   4200
         Width           =   1065
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   180
         X2              =   6300
         Y1              =   3105
         Y2              =   3105
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   180
         X2              =   6300
         Y1              =   5220
         Y2              =   5220
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   360
         Left            =   4425
         Tag             =   "et0;ht2"
         Top             =   150
         Width           =   2025
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   1110
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
      Picture         =   "frmMealOjtApplication.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   120
      TabIndex        =   26
      Top             =   495
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
      Picture         =   "frmMealOjtApplication.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   120
      TabIndex        =   27
      Top             =   4185
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
      Picture         =   "frmMealOjtApplication.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   1110
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmMealOjtApplication.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   120
      TabIndex        =   29
      Top             =   1725
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
      Picture         =   "frmMealOjtApplication.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   31
      Top             =   495
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmMealOjtApplication.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   120
      TabIndex        =   32
      Top             =   2340
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Deactvte"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMealOjtApplication.frx":2CDC
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   120
      TabIndex        =   34
      Top             =   2955
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Approve"
      AccessKey       =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMealOjtApplication.frx":3456
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   120
      TabIndex        =   35
      Top             =   3570
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&DisApprv"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMealOjtApplication.frx":3BD0
   End
End
Attribute VB_Name = "frmMealOjtApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmMealOjtApplication"

Private oDriver As clsFormDriver
Private WithEvents oTrans As clsOJTApplication
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Private pnIndex As Integer

Function getSched() As String
   Dim sched(6) As String
   sched(0) = IIf(check(0).Value = vbChecked, "1", "0")
   sched(1) = IIf(check(1).Value = vbChecked, "1", "0")
   sched(2) = IIf(check(2).Value = vbChecked, "1", "0")
   sched(3) = IIf(check(3).Value = vbChecked, "1", "0")
   sched(4) = IIf(check(4).Value = vbChecked, "1", "0")
   sched(5) = IIf(check(5).Value = vbChecked, "1", "0")
   sched(6) = IIf(check(6).Value = vbChecked, "1", "0")
   getSched = Join(sched, "")
End Function

Function loadSched() As String
   Dim sSchedule As String
   Dim splitSched() As String

   sSchedule = oTrans.Master("sSchedule")
   splitSched = Split(StrConv(sSchedule, 64), Chr(0))

   check(0).Value = IIf(splitSched(0) = "0", vbUnchecked, vbChecked)
   check(1).Value = IIf(splitSched(1) = "0", vbUnchecked, vbChecked)
   check(2).Value = IIf(splitSched(2) = "0", vbUnchecked, vbChecked)
   check(3).Value = IIf(splitSched(3) = "0", vbUnchecked, vbChecked)
   check(4).Value = IIf(splitSched(4) = "0", vbUnchecked, vbChecked)
   check(5).Value = IIf(splitSched(5) = "0", vbUnchecked, vbChecked)
   check(6).Value = IIf(splitSched(6) = "0", vbUnchecked, vbChecked)

End Function

Private Sub cmdButton_Click(Index As Integer)
Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   Select Case Index
   Case 0   'cancel
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
      If lnRep = vbYes Then
         InitForm 1
      End If
   Case 1   'Browse
      If oTrans.SearchTransaction = True Then
         InitForm 1
      End If
   Case 2   'save
      If isEntryOk = True Then
          oTrans.Master("sSchedule") = getSched
         If oTrans.SaveTransaction = True Then
            MsgBox "Record Successfully Saved!!!", vbInformation, "Confirm"
            If MsgBox("Do you want to approve this record?", vbInformation + vbYesNo, "Confirmation") = vbYes Then
               If Not oTrans.ApproveTransaction Then
                  MsgBox "Unable to approve record...", vbCritical, "Warning"
               End If
               MsgBox "Record approved successfully!", vbInformation, "Confirmation"
             End If
            oTrans.NewTransaction
            InitForm 0
         Else
            MsgBox "Unable to save Record!!!", vbCritical, "Warning"
         End If
      End If
   Case 3   'new
      oTrans.NewTransaction
      InitForm 0
      txtField(1).SetFocus
   Case 4   'close
      Unload Me
   Case 5 ' update
      If oTrans.Master("cRecdStat") <> "0" Then
         MsgBox "Unable to update transaction..." & vbCrLf & _
             "Transaction already " + lblStatus.Caption + "!!!", vbCritical, "Error"
         Exit Sub
      End If
      If oTrans.UpdateTransaction = True Then
         InitForm 0
         loadTransaction
      End If
   Case 6 'De activate
       lnRep = MsgBox("Do you want to de-activate this record?!!!", vbYesNo + vbQuestion, "Confirm")
         If oTrans.Master("cRecdStat") = "4" Then
            MsgBox "Record already deactivated!!!", vbCritical, "Warning"
         GoTo endProc
       End If
      If lnRep = vbYes Then
         If oTrans.CancelTransaction = True Then
             MsgBox "Record successfully de-activated!!!", vbInformation, "Confirm"
             oTrans.OpenTransaction (oTrans.Master("sTransNox"))
             InitForm 1
         End If
      End If
   Case 7 ' approve
      lnRep = MsgBox("Do you want to approve this record?!!!", vbYesNo + vbQuestion, "Confirm")
         If oTrans.Master("cRecdStat") <> "0" Then
            MsgBox "Unable to approve record..." + _
            vbCrLf + "Record already " + lblStatus.Caption + "!!!", vbCritical, "Warning"
         GoTo endProc
       End If
      If lnRep = vbYes Then
         If oTrans.ApproveTransaction = True Then
             MsgBox "Record successfully approved!!!", vbInformation, "Confirm"
             oTrans.OpenTransaction (oTrans.Master("sTransNox"))
             InitForm 1
          End If
      End If

   Case 8 ' disapprove
      lnRep = MsgBox("Do you want to disapproved this record?!!!", vbYesNo + vbQuestion, "Confirm")
         If oTrans.Master("cRecdStat") <> "0" Then
           MsgBox "Unable to disapproved record..." + _
            vbCrLf + "Record already " + lblStatus.Caption + "!!!", vbCritical, "Warning"
         GoTo endProc
       End If
      If lnRep = vbYes Then
         If oTrans.Disapprove = True Then
             MsgBox "Record successfully dis-approved!!!", vbInformation, "Confirm"
             oTrans.OpenTransaction (oTrans.Master("sTransNox"))
             InitForm 1
         End If
      End If

   End Select

endProc:
   Exit Sub
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

   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me

   Set oTrans = New clsOJTApplication
   Set oTrans.AppDriver = oApp

   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   oTrans.NewTransaction
   InitForm 0
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
   Dim loTxt As TextBox

   xrFrame2.Enabled = (fnEdit = 0)
   cmdButton(1).Visible = Not (fnEdit = 0)
   cmdButton(3).Visible = Not (fnEdit = 0)
   cmdButton(4).Visible = Not (fnEdit = 0)
   cmdButton(5).Visible = Not (fnEdit = 0)
   cmdButton(6).Visible = Not (fnEdit = 0)
   cmdButton(7).Visible = Not (fnEdit = 0)
   cmdButton(8).Visible = Not (fnEdit = 0)

   cmdButton(0).Visible = (fnEdit = 0)
   cmdButton(2).Visible = (fnEdit = 0)

   For Each loTxt In txtField
      loTxt = ""
   Next

   If fnEdit = 0 Then
      InitTransaction
   Else
      loadTransaction
   End If
   setTransTat (oTrans.Master("cRecdStat"))
End Sub

Private Sub InitTransaction()
   With oTrans
         txtField(0) = oTrans.Master(0)
         txtField(1) = strLongDate(oTrans.Master(1))
         txtField(3) = strLongDate(oTrans.Master(3))
         txtField(4) = oTrans.Master(4)
   End With
   loadSched
End Sub

Private Sub loadTransaction()
   With oTrans
      txtField(0) = oTrans.Master(0)
      txtField(1) = strLongDate(oTrans.Master(1))
      txtField(3) = strLongDate(oTrans.Master(3))
      txtField(4) = oTrans.Master(4)
      txtField(8) = oTrans.Master(8)
      txtField(2) = oTrans.Master(2)
      txtField(5) = oTrans.Master(5)
      txtField(6) = oTrans.Master(6)
      txtField(7) = oTrans.Master(7)
   End With
   Call loadSched

End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
      Case 0
         txtField(Index) = oTrans.Master(Index)
      Case 1, 3
         txtField(Index) = strLongDate(oTrans.Master(Index))
      Case 4
         txtField(Index) = oTrans.Master(Index)
      Case 2
         txtField(Index) = oTrans.Master(Index)
      Case 5
         txtField(Index) = oTrans.Master(Index)
      Case 6
         txtField(Index) = oTrans.Master(Index)
      Case 7
         txtField(Index) = oTrans.Master(Index)
      Case 8
         txtField(Index) = oTrans.Master(Index)
      Case Else
         txtField(Index) = oTrans.Master(Index)
   End Select
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(pnIndex)
         If KeyCode = vbKeyF3 Then
            Select Case Index
            Case Else
               If oTrans.SearchMaster(Index, .Text) = True Then
               End If
            End Select
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then
                  If oTrans.SearchMaster(Index, .Text) = True Then
               End If
            End If
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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      Select Case Index
         Case 1, 3
            If Not IsDate(.Text) Then .Text = strLongDate(oApp.ServerDate)
            oTrans.Master(Index) = .Text
         Case 4
            If Not IsNumeric(.Text) Then .Text = 0
            oTrans.Master(Index) = .Text
         Case 8
            .Text = TitleCase(.Text)
            oTrans.Master(Index) = .Text
      End Select
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
      Case 1, 3
         txtField(Index) = strShortDate(oTrans.Master(Index))
   End Select

   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

   oDriver.ColumnIndex = Index
   pnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
      Case 1, 3
         txtField(Index) = strLongDate(txtField(Index).Text)
   End Select

   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With

   pnIndex = Index
End Sub

Private Function isEntryOk() As Boolean
   If oTrans.Master("sClientID") = "" Then
      MsgBox "Empty Client Name Detected!" & vbCrLf & _
               "Pls Verify Entry Then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If

   If oTrans.Master(5) = "" Then
      MsgBox "Empty School Name Detected!" & vbCrLf & _
               "Pls Verify Entry Then Try Again!!!", vbCritical, "Warning"
      txtField(5).SetFocus
      GoTo EntryNotOK
   End If

    If oTrans.Master(6) = "" Then
      MsgBox "Empty Course Name Detected!" & vbCrLf & _
               "Pls Verify Entry Then Try Again!!!", vbCritical, "Warning"
      txtField(6).SetFocus
      GoTo EntryNotOK
   End If


   If oTrans.Master(7) = "" Then
      MsgBox "Empty Department Name Detected!" & vbCrLf & _
               "Pls Verify Entry Then Try Again!!!", vbCritical, "Warning"
      txtField(7).SetFocus
      GoTo EntryNotOK
   End If

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub setTransTat(lnStat As String)
   Select Case lnStat
   Case "0"
      lblStatus.Caption = "OPEN"
   Case "1"
      lblStatus.Caption = "APPROVED"
   Case "3"
      lblStatus.Caption = "NOT APPROVED"
   Case "4"
       lblStatus.Caption = "DONE"
   Case Else
      lblStatus.Caption = "UNKNOWN"
   End Select
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
