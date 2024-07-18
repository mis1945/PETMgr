VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmTimesheetSummary 
   BorderStyle     =   0  'None
   Caption         =   "Timesheet Summary"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1740
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   3069
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Apr 25, 2011"
         Top             =   795
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "M00111-000021"
         Top             =   135
         Width           =   1800
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "Apr 25, 2011"
         Top             =   1215
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3675
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "Apr 25, 2011"
         Top             =   1215
         Width           =   1905
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   3465
         Top             =   150
         Width           =   2115
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   3510
         Top             =   180
         Width           =   2040
      End
      Begin VB.Label Label1 
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
         Left            =   3540
         TabIndex        =   7
         Tag             =   "eb0;et0"
         Top             =   210
         Width           =   1950
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
         Index           =   0
         Left            =   75
         TabIndex        =   6
         Top             =   870
         Width           =   405
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Left            =   1530
         Tag             =   "et0;ht2"
         Top             =   225
         Width           =   1800
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans No."
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
         Left            =   75
         TabIndex        =   4
         Top             =   195
         Width           =   900
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From/Thru"
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
         TabIndex        =   2
         Top             =   1290
         Width           =   885
      End
      Begin VB.Line Line2 
         X1              =   3570
         X2              =   3435
         Y1              =   1245
         Y2              =   1560
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6915
      Left            =   7380
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   12197
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   1305
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "M00110000001"
         Top             =   165
         Width           =   1710
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3570
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   4335
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3075
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3090
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   4335
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2670
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2685
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4335
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2190
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   4335
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   1785
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4335
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   1380
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "2"
         Top             =   1380
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1305
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "GMC Dagupan - Honda"
         Top             =   975
         Width           =   4290
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1305
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Sayson, Marlon A."
         Top             =   570
         Width           =   4290
      End
      Begin VB.Label lblRemarks 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--- Will Receive Memo ---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1800
         TabIndex        =   36
         Top             =   4410
         Width           =   2160
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. ID"
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
         Left            =   135
         TabIndex        =   35
         Top             =   225
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime"
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
         Index           =   12
         Left            =   135
         TabIndex        =   29
         Top             =   3630
         Width           =   765
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   75
         X2              =   5565
         Y1              =   3510
         Y2              =   3510
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   60
         X2              =   5550
         Y1              =   2610
         Y2              =   2610
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excused"
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
         Left            =   3225
         TabIndex        =   27
         Top             =   3135
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Undertime"
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
         Left            =   135
         TabIndex        =   25
         Top             =   3150
         Width           =   885
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excused"
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
         Left            =   3225
         TabIndex        =   23
         Top             =   2730
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tardiness"
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
         Left            =   135
         TabIndex        =   21
         Top             =   2745
         Width           =   840
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suspension"
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
         Left            =   3225
         TabIndex        =   19
         Top             =   2250
         Width           =   1020
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Holiday"
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
         Left            =   3225
         TabIndex        =   17
         Top             =   1845
         Width           =   645
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "W/ Leave"
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
         Left            =   3225
         TabIndex        =   15
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Absent"
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
         Left            =   135
         TabIndex        =   13
         Top             =   1440
         Width           =   615
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
         Index           =   3
         Left            =   135
         TabIndex        =   11
         Top             =   1035
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
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
         Left            =   135
         TabIndex        =   9
         Top             =   630
         Width           =   870
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5145
      Left            =   1605
      TabIndex        =   30
      Top             =   2340
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   9075
      _Version        =   393216
      Cols            =   3
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   135
      TabIndex        =   31
      Top             =   1590
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
      Picture         =   "frmTimesheetSummary.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   135
      TabIndex        =   32
      Top             =   2850
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
      Picture         =   "frmTimesheetSummary.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   135
      TabIndex        =   33
      Top             =   2220
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Process"
      AccessKey       =   "Process"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTimesheetSummary.frx":0EF4
   End
End
Attribute VB_Name = "frmTimesheetSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmTimesheetSummary"

Private oTrans As clsSummarizedTimeshet
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private pnActiveRow As Integer

Dim poReport As clsReport
Dim pbLoad As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0 'Confirm
      If oTrans.PostTransaction Then
         Label1 = TransStat(oTrans.Master("cTranStat"))
         MsgBox "Timesheet Summary was confirmed successfully!", vbOKOnly, "Info"
      Else
         MsgBox "Unable to confirmed Timesheet Summary!", vbOKOnly, "Info"
      End If
   Case 1 'Process
      If oTrans.CloseTransaction Then
         Label1 = TransStat(oTrans.Master("cTranStat"))
         Call loadGrid
         MsgBox "Timesheet Summary was processed successfully!", vbOKOnly, "Info"
      Else
         MsgBox "Unable to process Timesheet Summary!", vbOKOnly, "Info"
      End If
   Case 2 'Close
      Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   With MSFlexGrid1
      .Refresh
   End With
      
   If Not pbLoad Then
      pbLoad = True
   End If
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
'   On Error GoTo errProc

   CenterChildForm mdiMain, Me
      
   Set oTrans = New clsSummarizedTimeshet
   Set oTrans.AppDriver = oApp
   oTrans.TransStatus = 10
   oTrans.InitTransaction
   oTrans.NewTransaction
            
   If oTrans.EditMode = xeModeAddNew Then
      oTrans.SaveTransaction
   End If
         
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   
   Set poReport = New clsReport
   
   LoadMaster
   InitGrid
   loadGrid
   ClearFields -1
      
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oTrans = Nothing
   pbLoad = False
End Sub

'----------------------------
'InitGrid()
'  'Set the property rows, columns, etc...
'----------------------------
Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Rows = 2
      .Cols = 10
      
      .Row = 0
      .RowHeight(0) = 320
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Row = 1
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "Employee"
      .TextMatrix(0, 3) = "Absences"
      .TextMatrix(0, 4) = "Tardiness"
      .TextMatrix(0, 5) = "Undertime"
      .TextMatrix(0, 6) = "Overtime"
      .TextMatrix(0, 7) = "Rec Ctr"
      
      .RowHeightMin = 320
      
      .ColWidth(0) = 635
      .ColWidth(1) = 2480
      .ColWidth(2) = 2520
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .ColWidth(7) = 0
      .ColWidth(8) = 0
      .ColWidth(9) = 0
      
      'column allinment
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      
      'set location
      .Row = 1
      .Col = 2
      .ColSel = .Cols - 1
   End With

End Sub

Private Sub loadGrid()
   Dim lnCtr As Integer
   Dim lnCtrx As Integer
   
   With MSFlexGrid1
      .Rows = 2
      .TextMatrix(lnCtr + 1, 0) = Format(lnCtr + 1, "0000")
      .TextMatrix(lnCtr + 1, 1) = ""
      .TextMatrix(lnCtr + 1, 2) = ""
      .TextMatrix(lnCtr + 1, 3) = ""
      .TextMatrix(lnCtr + 1, 4) = ""
      .TextMatrix(lnCtr + 1, 5) = ""
      .TextMatrix(lnCtr + 1, 6) = ""
      
      If oTrans.ItemCount >= 1 Then
         For lnCtrx = 0 To oTrans.ItemCount - 1
            DoEvents
            
            If oTrans.Detail(lnCtrx, "nTardyxxx") + oTrans.Detail(lnCtrx, "nTardyOms") >= CInt(oApp.getConfiguration("MATPCO")) Then
               .Rows = 2 + lnCtr
               .RowHeight(lnCtr + 1) = 338
               
               .TextMatrix(lnCtr + 1, 0) = Format(lnCtr + 1, "0000")
               .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtrx, "sBranchNm")
               .TextMatrix(lnCtr + 1, 2) = oTrans.Detail(lnCtrx, "xFullname")
               .TextMatrix(lnCtr + 1, 3) = Format(oTrans.Detail(lnCtrx, "nAbsentxx") - (oTrans.Detail(lnCtrx, "nWithLvex") + oTrans.Detail(lnCtrx, "nHolidayx") + oTrans.Detail(lnCtrx, "nSuspensn")), "#,##0.00")
               .TextMatrix(lnCtr + 1, 4) = Format(oTrans.Detail(lnCtrx, "nTardyxxx") - oTrans.Detail(lnCtrx, "nTardyOms"), "#,##0.00")
               .TextMatrix(lnCtr + 1, 5) = Format(oTrans.Detail(lnCtrx, "nUndrTime") - oTrans.Detail(lnCtrx, "nUndrOmsn"), "#,##0.00")
               .TextMatrix(lnCtr + 1, 6) = Format(oTrans.Detail(lnCtrx, "nOverTime"), "#,##0.00")
               .TextMatrix(lnCtr + 1, 7) = lnCtrx
               .TextMatrix(lnCtr + 1, 8) = oTrans.Detail(lnCtrx, "sEmpLevID")
               .TextMatrix(lnCtr + 1, 9) = oTrans.Detail(lnCtrx, "nNoDTardy")
            
               lnCtr = lnCtr + 1
            
            End If
            
            DoEvents
         Next
      End If
      
   End With
End Sub

Private Sub LoadMaster()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 1 To 3
         loTxt.Text = Format(oTrans.Master(loTxt.Index), "Mmm. DD, YYYY")
      Case Else
         loTxt.Text = oTrans.Master(loTxt.Index)
      End Select
   Next
   
   Label1 = TransStat(oTrans.Master("cTranStat"))
   
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

Private Sub MSFlexGrid1_SelChange()
   If pnActiveRow <> MSFlexGrid1.Row Then
      pnActiveRow = MSFlexGrid1.Row
            
      Call ClearFields(pnActiveRow)
   End If
End Sub

Private Sub ClearFields(ByVal fnRow As Integer)
   Dim loTxt As TextBox
   Dim lnTotal As Currency
   Dim lnRow As Integer
   
   If fnRow < 0 Then
      For Each loTxt In txtDetail
         loTxt = ""
      Next
      
      lblRemarks.Visible = False
   Else
      lnRow = Val(MSFlexGrid1.TextMatrix(fnRow, 7))
      
      For Each loTxt In txtDetail
         Select Case loTxt.Index
         Case 2 To 10
            loTxt = Format(oTrans.Detail(lnRow, loTxt.Index), "#,##0.00")
         Case Else
            loTxt = IFNull(oTrans.Detail(lnRow, loTxt.Index), "")
         End Select
      Next
      
      If Val(MSFlexGrid1.TextMatrix(fnRow, 9)) >= Val(oApp.getConfiguration("MT2HTC")) And _
         ((Val(MSFlexGrid1.TextMatrix(fnRow, 4)) > Val(oApp.getConfiguration("MATPCO")) And Val(MSFlexGrid1.TextMatrix(fnRow, 8)) = "0") Or _
          (Val(MSFlexGrid1.TextMatrix(fnRow, 4)) > Val(oApp.getConfiguration("MATPCX")) And Val(MSFlexGrid1.TextMatrix(fnRow, 8)) > "0")) Then
         lblRemarks.Visible = True
      Else
         lblRemarks.Visible = False
      End If
      
   End If
   
End Sub

