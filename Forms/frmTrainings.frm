VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTrainings 
   BorderStyle     =   0  'None
   Caption         =   "Training/Seminar"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   13110
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Height          =   6780
      Left            =   8220
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   555
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   11959
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3570
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   1710
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   6297
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
         Index           =   6
         Left            =   4155
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1635
         Width           =   2280
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   16
         Top             =   3030
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2565
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
         Height          =   800
         Index           =   1
         Left            =   1620
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   795
         Width           =   4815
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
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   1845
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
         TabIndex        =   9
         Top             =   1635
         Width           =   2280
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2100
         Width           =   4815
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   4215
         Top             =   105
         Width           =   2220
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   4245
         Top             =   135
         Width           =   2145
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
         Left            =   4290
         TabIndex        =   36
         Tag             =   "eb0;et0"
         Top             =   165
         Width           =   2055
      End
      Begin VB.Line Line1 
         X1              =   4080
         X2              =   3960
         Y1              =   1650
         Y2              =   2025
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venue"
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
         TabIndex        =   15
         Top             =   3120
         Width           =   555
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sponsor"
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
         TabIndex        =   13
         Top             =   2655
         Width           =   720
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title/Description"
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
         TabIndex        =   6
         Top             =   795
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1725
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         TabIndex        =   4
         Top             =   210
         Width           =   195
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Held"
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
         TabIndex        =   8
         Top             =   1725
         Width           =   855
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trainer"
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
         Left            =   105
         TabIndex        =   11
         Top             =   2190
         Width           =   585
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1125
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   1984
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
         Height          =   420
         Index           =   1
         Left            =   1620
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   585
         Width           =   4815
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
         Height          =   420
         Index           =   0
         Left            =   1620
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   120
         Width           =   1860
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title/Description"
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
         Index           =   15
         Left            =   105
         TabIndex        =   2
         Top             =   675
         Width           =   1395
      End
      Begin VB.Label lblField 
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
         Index           =   14
         Left            =   105
         TabIndex        =   0
         Top             =   210
         Width           =   180
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2025
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   5310
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   3572
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtDetail 
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
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox txtDetail 
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
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   990
         Width           =   4815
      End
      Begin VB.TextBox txtDetail 
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
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   540
         Width           =   4815
      End
      Begin VB.TextBox txtDetail 
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
         Index           =   0
         Left            =   1620
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   90
         Width           =   4815
      End
      Begin VB.Label lblField 
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
         Index           =   8
         Left            =   105
         TabIndex        =   23
         Top             =   1560
         Width           =   705
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
         Index           =   6
         Left            =   105
         TabIndex        =   21
         Top             =   1110
         Width           =   1005
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
         Index           =   5
         Left            =   105
         TabIndex        =   19
         Top             =   660
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
         Index           =   4
         Left            =   105
         TabIndex        =   17
         Top             =   180
         Width           =   870
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   9
      Left            =   90
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3555
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
      Picture         =   "frmTrainings.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   90
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2925
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
      Picture         =   "frmTrainings.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3555
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
      Picture         =   "frmTrainings.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   90
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2295
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
      Picture         =   "frmTrainings.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2925
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
      Picture         =   "frmTrainings.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   25
      Top             =   1050
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&ADD"
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
      Picture         =   "frmTrainings.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&DEL"
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
      Picture         =   "frmTrainings.frx":35F4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   32
      Top             =   2295
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
      Picture         =   "frmTrainings.frx":3D6E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   30
      Top             =   1050
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
      Picture         =   "frmTrainings.frx":44E8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   31
      Top             =   1680
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
      Picture         =   "frmTrainings.frx":4C62
   End
End
Attribute VB_Name = "frmTrainings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmTrainings"
Private WithEvents oTrans As clsEmpTraining
Attribute oTrans.VB_VarHelpID = -1

Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim psSelected() As String
Dim pnIndex As Integer
Dim pnCtr As Integer
Dim pnRow As Integer
Dim pbFormLoad As Boolean

Dim pbMasterGotFocus As Boolean
Dim pbGridGotFocus As Boolean
Dim pbDetailGotFocus As Boolean

Private Sub cmdButton_Click(Index As Integer)
   With MSFlexGrid1
      Select Case Index
      Case 0 ' Close
         Unload Me
      Case 1 ' New
         If oTrans.NewTransaction Then
            Call InitForm(oTrans.EditMode)
            Call LoadMaster
            Call LoadDetail
            txtField(1).SetFocus
         End If
      Case 2 'Browse
         If oTrans.SearchTransaction("", False) Then
            Call LoadMaster
            Call LoadDetail
            txtSearch(1).SetFocus
         End If
      Case 3 'Approve/Close
         If oTrans.EditMode = xeModeReady Then
            If oTrans.CloseTransaction Then
               MsgBox "Trainings/Seminars closed successfully!", vbOKOnly + vbInformation, "Confirmation"
               
               Call cmdButton_Click(1)
               
            Else
               MsgBox "Unable to closed Trainings/Seminars!", vbOKOnly + vbInformation, "Confirmation"
            End If
         End If
      Case 4 ' Update
         If oTrans.EditMode = xeModeReady Then
            If oTrans.UpdateTransaction Then
               Call InitForm(oTrans.EditMode)
               txtDetail(0).SetFocus
            End If
         End If
      Case 5 ' Add Detail
         If oTrans.EditMode = xeModeAddNew Or oTrans.EditMode = xeModeUpdate Then
            Call addDetail
         End If
      Case 6 'Delete Detail
         If oTrans.EditMode = xeModeAddNew Or oTrans.EditMode = xeModeUpdate Then
            Call deleteGridRow
         End If
      Case 7 'Search Detail
         If oTrans.EditMode = xeModeAddNew Or oTrans.EditMode = xeModeUpdate Then
            If pnIndex = 0 Then
               If oTrans.SearchDetail(.Row - 1, 0, "") Then
                  .TextMatrix(.Row, 1) = oTrans.Detail(.Row - 1, "xFullName")
               
                  If MSFlexGrid1.Rows - 1 = MSFlexGrid1.Row Then
                     Call addDetail
                  Else
                     Call LoadFromGrid(.Row - 1)
                  End If
               
               End If
            End If
         End If
      Case 8 'Save
         If .Rows > 2 Then
            pnCtr = 1
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
                  .Row = pnCtr
                  If oTrans.deleteDetail(.Row - 1) Then Call deleteGridRow
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop
         End If

         If oTrans.SaveTransaction() Then
            MsgBox "Transaction saved successfuly.", vbInformation, pxeMODULENAME
            Call InitForm(xeModeReady)
         Else
            MsgBox "Unable to save transaction.", vbCritical, pxeMODULENAME
         End If
      Case 9
         If MsgBox("Do you want to undo updates for this training?", vbOKCancel + vbInformation, "Confirmation") = vbOK Then
            If oTrans.EditMode = xeModeAddNew Then
               Call oTrans.InitTransaction
               Call InitForm(oTrans.EditMode)
               Call LoadMaster
               Call LoadDetail
            Else
               Call oTrans.OpenTransaction(oTrans.Master("sTransNox"))
               Call InitForm(oTrans.EditMode)
               Call LoadMaster
               Call LoadDetail
            End If
         End If
      End Select
   End With
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

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsEmpTraining
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   InitGrid
   InitForm oTrans.EditMode
   LoadMaster
   LoadDetail
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
   pbFormLoad = False
End Sub

Private Sub InitForm(ByVal fnEdit As xeEditMode)
   'Search frame
   xrFrame3.Enabled = (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
   
   cmdButton(0).Visible = (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
   cmdButton(1).Visible = (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
   cmdButton(2).Visible = (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
   cmdButton(3).Visible = (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
   cmdButton(4).Visible = (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)

   'Detail/Master frame
   xrFrame1.Enabled = Not (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
   xrFrame2.Enabled = Not (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
   cmdButton(5).Visible = Not (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
   cmdButton(6).Visible = Not (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
   cmdButton(7).Visible = Not (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
   cmdButton(8).Visible = Not (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
   cmdButton(9).Visible = Not (fnEdit = xeModeReady Or fnEdit = xeModeUnknown)
End Sub

Private Sub LoadMaster()
   Dim loTxt As TextBox
   
   'Load value for search fields
   txtSearch(0) = oTrans.Master(0)
   txtSearch(1) = oTrans.Master(1)
   
   'Load value to text fields
   For Each loTxt In txtField
      loTxt = oTrans.Master(loTxt.Index)
   Next
   
   Label2.Caption = TransStat(oTrans.Master(7))
   
   'Initialize detail text boxes
   For Each loTxt In txtDetail
      loTxt = ""
   Next
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Rows = 1
      For lnCtr = 0 To oTrans.ItemCount - 1
         .Rows = .Rows + 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = IFNull(oTrans.Detail(lnCtr, 1))
      Next
      
      .Row = 1
      .RowSel = 1
   End With
      
End Sub

Private Sub LoadFromGrid(ByVal fnRow As Integer)
   Dim loTxt As TextBox
   If fnRow < 0 Then
      For Each loTxt In txtDetail
         loTxt = ""
      Next
   Else
      For Each loTxt In txtDetail
         loTxt = IFNull(oTrans.Detail(fnRow, loTxt.Index))
      Next
   End If
End Sub

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      If .Row < 1 Then
         .Row = 1
         .RowSel = 1
         Call LoadFromGrid(.Row - 1)
      End If
   End With
End Sub

Private Sub MSFlexGrid1_GotFocus()
   pbGridGotFocus = True
   pbDetailGotFocus = False
   pbMasterGotFocus = False
End Sub

Private Sub MSFlexGrid1_RowColChange()
   With MSFlexGrid1
      If .Row > 0 Then
         Call LoadFromGrid(.Row - 1)
      End If
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 1
      txtField(Index) = strLongDate(Value)
   Case 7
      Label2.Caption = TransStat(oTrans.Master(Index))
   Case Else
      txtField(Index) = Value
   End Select
End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
   With txtDetail(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

   pbDetailGotFocus = True
   pbMasterGotFocus = False
   pbGridGotFocus = False

   pnIndex = Index
End Sub

Private Sub txtDetail_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      If Index = 0 Then
         With txtDetail(Index)
            If KeyCode = vbKeyF3 Then
               oTrans.SearchDetail MSFlexGrid1.Row - 1, Index, .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchDetail MSFlexGrid1.Row - 1, Index, .Text
            End If
         
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = oTrans.Detail(MSFlexGrid1.Row - 1, "xFullName")
            
            If MSFlexGrid1.Rows - 1 = MSFlexGrid1.Row Then
               Call addDetail
            Else
               Call LoadFromGrid(MSFlexGrid1.Row - 1)
            End If
         
         End With
      
      End If
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

Private Sub txtDetail_LostFocus(Index As Integer)
   With txtDetail(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtDetail_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      oTrans.Master(Index) = txtDetail(Index).Text
   End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      Select Case Index
      Case 5, 6
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")

         oTrans.Master(Index) = CDate(.Text)
      Case Else
         oTrans.Master(Index) = .Text
      End Select
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
    With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
    End With
   
   Select Case Index
   Case 5, 6
      If IsDate(oTrans.Master(Index)) Then
         txtField(Index) = strShortDate(oTrans.Master(Index))
      Else
        txtField(Index) = ""
      End If
   End Select
   
   pnIndex = Index
   
   pbMasterGotFocus = True
   pbDetailGotFocus = False
   pbGridGotFocus = False
   
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = MSFlexGrid1.hWnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
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

Private Sub InitGrid()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Cols = 2
      .Font = "MS Sans Serif"
      .RowHeight(0) = 350

      'Column Title
      .TextMatrix(0, 1) = "Employee"
      .Row = 1

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 4
      Next

      .ColWidth(0) = 450
      .ColWidth(1) = 4500

      .ColAlignment(1) = 1
   End With
End Sub

Private Sub addDetail()
   Dim lsOldProc As String

   lsOldProc = pxeMODULENAME & "addDetail"
   'On Error GoTo errProc

   With MSFlexGrid1
      If oTrans.addDetail Then
         
         'add row
         .Rows = .Rows + 1
         
         'set the last row as current row
         .Row = .Rows - 1
         
         'set the last row as current row
         .RowSel = .Row
                  
         .TextMatrix(.Row, 0) = .Row
         
         'Load details
         Call LoadFromGrid(.Row - 1)
         
         txtDetail(0).SetFocus
      Else
         MsgBox "Unable to add new employee. Please check your entry!", vbOKOnly + vbInformation, "Confirmation"
      End If

      .Refresh
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub deleteGridRow()
   Dim lnLastRow As Integer
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Rows = .Rows - 1
      If .Rows = 1 Then .Rows = .Rows + 1
      
      lnLastRow = .Row - 1
      For pnCtr = lnLastRow To oTrans.ItemCount - 1
         .TextMatrix(pnCtr + 1, 0) = pnCtr + 1
         .TextMatrix(pnCtr + 1, 1) = oTrans.Detail(pnCtr, "xFullName")
         .Row = pnCtr + 1
         If (pnCtr + 1) Mod 2 = 0 Then
            For lnCtr = 1 To .Cols - 1
               .Col = lnCtr
               .CellBackColor = oApp.getColor("fb0")
            Next
         End If
      Next
      
      .Row = lnLastRow + 1
      .RowSel = .Row
   
      Call LoadFromGrid(.Row - 1)
   End With
End Sub
