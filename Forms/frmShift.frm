VERSION 5.00
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmShift 
   BorderStyle     =   0  'None
   Caption         =   "Shifts"
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrFrame xrFrame2 
      Height          =   4830
      Left            =   5940
      Top             =   555
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8520
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   5160
         Left            =   -30
         TabIndex        =   25
         Top             =   -30
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   9102
         AllowBigSelection=   -1  'True
         AutoAdd         =   0   'False
         AutoNumber      =   0   'False
         BACKCOLOR       =   -2147483643
         BACKCOLORBKG    =   8421504
         BACKCOLORFIXED  =   -2147483633
         BACKCOLORSEL    =   -2147483635
         BORDERSTYLE     =   0
         COLS            =   2
         FILLSTYLE       =   0
         FIXEDCOLS       =   1
         FIXEDROWS       =   1
         FOCUSRECT       =   1
         EDITORBACKCOLOR =   -2147483643
         EDITORFORECOLOR =   -2147483640
         FORECOLOR       =   -2147483640
         FORECOLORFIXED  =   -2147483630
         FORECOLORSEL    =   -2147483634
         FORMATSTRING    =   ""
         Object.HEIGHT          =   5160
         GRIDCOLOR       =   12632256
         GRIDCOLORFIXED  =   0
         BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GRIDLINES       =   1
         GRIDLINESFIXED  =   2
         GRIDLINEWIDTH   =   1
         MOUSEICON       =   "frmShift.frx":0000
         MOUSEPOINTER    =   0
         REDRAW          =   -1  'True
         RIGHTTOLEFT     =   0   'False
         ROWS            =   2
         SCROLLBARS      =   3
         SCROLLTRACK     =   0   'False
         SELECTIONMODE   =   0
         Object.TOOLTIPTEXT     =   ""
         WORDWRAP        =   0   'False
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4830
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   8520
      BackColor       =   12632256
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
         Index           =   9
         Left            =   1725
         TabIndex        =   10
         Top             =   2055
         Width           =   1400
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PAID"
         Height          =   330
         Left            =   3165
         TabIndex        =   8
         Top             =   1665
         Width           =   1680
      End
      Begin VB.ComboBox cmbSchedule 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "frmShift.frx":001C
         Left            =   2460
         List            =   "frmShift.frx":0038
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   4245
         Width           =   3210
      End
      Begin VB.CheckBox chkFlexi 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "FLEXI TIME"
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
         Left            =   4140
         TabIndex        =   23
         Tag             =   "et0;fb0"
         Top             =   3945
         Width           =   1530
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
         Index           =   8
         Left            =   4275
         TabIndex        =   12
         Top             =   2055
         Width           =   1400
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
         Left            =   1725
         TabIndex        =   5
         Top             =   1155
         Width           =   1400
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   1830
         TabIndex        =   15
         Top             =   2880
         Width           =   1400
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   1830
         TabIndex        =   17
         Top             =   3330
         Width           =   1400
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
         Index           =   4
         Left            =   1725
         TabIndex        =   7
         Top             =   1605
         Width           =   1400
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   4155
         TabIndex        =   20
         Top             =   2925
         Width           =   1400
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   4155
         TabIndex        =   22
         Top             =   3375
         Width           =   1400
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
         Left            =   1155
         TabIndex        =   3
         Top             =   690
         Width           =   4500
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
         Left            =   1155
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Work Day Equivalent"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   11
         Left            =   510
         TabIndex        =   9
         Top             =   2010
         Width           =   1560
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule"
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
         Index           =   20
         Left            =   1605
         TabIndex        =   24
         Top             =   4320
         Width           =   810
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "2nd Quarter"
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
         Left            =   3585
         TabIndex        =   18
         Tag             =   "et0;fb0"
         Top             =   2595
         Width           =   1035
      End
      Begin VB.Shape Shape2 
         Height          =   1200
         Index           =   1
         Left            =   3450
         Top             =   2715
         Width           =   2220
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "1st Quarter"
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
         Left            =   1230
         TabIndex        =   13
         Tag             =   "et0;fb0"
         Top             =   2610
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         Height          =   1200
         Index           =   0
         Left            =   1140
         Top             =   2715
         Width           =   2220
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
         Left            =   3195
         TabIndex        =   11
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hours Of Work"
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
         Left            =   210
         TabIndex        =   4
         Top             =   1245
         Width           =   1290
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
         Index           =   2
         Left            =   1275
         TabIndex        =   14
         Top             =   2955
         Width           =   180
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
         Index           =   3
         Left            =   1230
         TabIndex        =   16
         Top             =   3390
         Width           =   390
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Break Time"
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
         Left            =   510
         TabIndex        =   6
         Top             =   1710
         Width           =   990
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
         Index           =   5
         Left            =   3690
         TabIndex        =   19
         Top             =   2985
         Width           =   180
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
         Index           =   6
         Left            =   3690
         TabIndex        =   21
         Top             =   3435
         Width           =   405
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1260
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   2415
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         TabIndex        =   2
         Top             =   750
         Width           =   975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift ID"
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
         Index           =   0
         Left            =   105
         TabIndex        =   0
         Top             =   210
         Width           =   690
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   13140
      TabIndex        =   33
      Top             =   5640
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmShift.frx":0085
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   12360
      TabIndex        =   32
      Top             =   5640
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmShift.frx":07FF
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   10020
      TabIndex        =   27
      Top             =   5640
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmShift.frx":0F79
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   10800
      TabIndex        =   30
      Top             =   5640
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmShift.frx":16F3
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   10020
      TabIndex        =   26
      Top             =   5640
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmShift.frx":1E6D
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   13140
      TabIndex        =   34
      Top             =   5640
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmShift.frx":25E7
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   11580
      TabIndex        =   31
      Top             =   5640
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Delete"
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
      Picture         =   "frmShift.frx":2D61
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   7
      Left            =   11580
      TabIndex        =   29
      Top             =   5640
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Del Det"
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
      Picture         =   "frmShift.frx":34DB
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   10800
      TabIndex        =   28
      Top             =   5640
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Add Det"
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
      Picture         =   "frmShift.frx":3C55
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeModuleName = "frmShift"

Private WithEvents oRecord As clsShift
Attribute oRecord.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Private p_nActiveRow As Integer

Dim pnCtr As Integer
Dim pnIndex As Integer
Dim pbLoaded As Boolean
Dim pbSearched As Boolean

Private Sub chkFlexi_Click()
   oRecord.Detail(p_nActiveRow, "cFlexiTym") = chkFlexi.Value
End Sub

Private Sub cmbSchedule_Click()
   Dim lnCtr As Integer

   If Not pbLoaded Then Exit Sub
   With cmbSchedule
      Select Case .ListIndex
      Case 0
         p_nActiveRow = .ListIndex
         
         With GridEditor1
            For lnCtr = 0 To .Cols - 1
               .Col = lnCtr
               .CellBackColor = &H80000005
            Next
          
            'set current row back color
            .Row = cmbSchedule.ListIndex + 2
            For lnCtr = 0 To .Cols - 1
               .Col = lnCtr
               .CellBackColor = &HE0E0E0
            Next
         End With
         
         Call loadCurrentDetail(p_nActiveRow)
      Case Else
         If GridEditor1.Rows - 2 > .ListIndex Then
            p_nActiveRow = .ListIndex
            oRecord.Detail(p_nActiveRow, "nDayOfWik") = .ListIndex
            
            With GridEditor1
               For lnCtr = 0 To .Cols - 1
                  .Col = lnCtr
                  .CellBackColor = &H80000005
               Next
            
               'set current row back color
               .Row = cmbSchedule.ListIndex + 2
               For lnCtr = 0 To .Cols - 1
                  .Col = lnCtr
                  .CellBackColor = &HE0E0E0
               Next
            End With
            
            Call loadCurrentDetail(p_nActiveRow)
         End If
      End Select
      
      txtField(1).Enabled = .ListIndex = 0
   End With
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsProcName As String
   Dim lnCtr As Integer
   Dim lbAdd As Boolean
   
   lsProcName = "cmdButton_Click"
   On Error GoTo errProc
   
   Select Case Index
   Case 0 'Cancel Edit Mode
      If oRecord.EditMode = xeModeAddNew Then
         oRecord.InitRecord
         Call ClearFields
         initButton xeModeReady
      ElseIf oRecord.OpenRecord(oRecord.Detail(0, "sShiftIDx")) Then
         Call LoadRecord
         initButton xeModeReady
'      Else
'         Call ClearFields
      End If
   Case 1 'Browse
      If oRecord.SearchRecord Then
         Call LoadRecord
         initButton xeModeReady
      End If
   Case 2 'Save
      If oRecord.SaveRecord Then
         MsgBox "Record Save Successfully...", vbInformation, "INFO"
         oRecord.InitRecord
         Call ClearFields
         initButton xeModeReady
      End If
   Case 3 'Update
      If oRecord.UpdateRecord Then
         initButton xeModeUpdate
         If cmbSchedule.ListIndex = 0 Then
            txtField(1).SetFocus
         Else
            txtField(2).SetFocus
         End If
      End If
   Case 4 'New
      If oRecord.NewRecord Then
         initButton xeModeAddNew
         
         If txtField(1).Enabled Then txtField(1).SetFocus
         Call ClearFields
      End If
   Case 5 'Close
      Unload Me
   Case 6 'Delete
      If oRecord.DeleteRecord Then
         MsgBox "Record Successfully deleted...", vbInformation, "INFO"
         
         Call ClearFields
      End If
   Case 7 'Delete Detail
      If oRecord.ItemCount > 1 Then
         If p_nActiveRow > 0 Then
            If oRecord.DeleteDetail(p_nActiveRow) Then
               GridEditor1.DeleteRow
               p_nActiveRow = GridEditor1.Row - 2
               Call loadCurrentDetail(p_nActiveRow)
               
               oRecord.InitRecord
               Call ClearFields
               initButton xeModeReady
            End If
         End If
      End If
   Case 8 'Add Detail
      If cmbSchedule.ListIndex > 0 Then
         If oRecord.ItemCount = 1 Then
            lbAdd = True
         Else
            lbAdd = True
            For lnCtr = 1 To oRecord.ItemCount - 1
               If oRecord.Detail(lnCtr, "nDayOfWik") = cmbSchedule.ListIndex Then
                  lbAdd = False
                  Exit For
               End If
            Next
         End If
         
         If lbAdd Then
            oRecord.addDetail
              
            With GridEditor1
               .Rows = .Rows + 1
               p_nActiveRow = .Rows - 3
               
               'set previous row to default back color
               .Row = .Rows - 2
               For lnCtr = 0 To .Cols - 1
                  .Col = lnCtr
                  .CellBackColor = &H80000005
               Next
            
               'set current row back color
               .Row = .Rows - 1
               For lnCtr = 0 To .Cols - 1
                  .Col = lnCtr
                  .CellBackColor = &HE0E0E0
               Next
               
               .TextMatrix(.Row, 1) = cmbSchedule.List(cmbSchedule.ListIndex)
            End With
            
            oRecord.Detail(p_nActiveRow, "nDayOfWik") = cmbSchedule.ListIndex
            Call loadCurrentDetail(p_nActiveRow)
            txtField(2).SetFocus
         End If
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If Not pbLoaded Then
      pbLoaded = True
   End If
   
   pbSearched = False
   
   If txtField(1).Enabled Then txtField(1).SetFocus
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   On Error GoTo errProc
   
   Call CenterChildForm(mdiMain, Me)
   
   Set oRecord = New clsShift
   Set oRecord.AppDriver = oApp
   oRecord.InitRecord
   
   oRecord.NewRecord
   initButton xeModeAddNew
   Call ClearFields
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   Call InitGrid
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   pbLoaded = False
End Sub

Private Sub oRecord_DetailRetrieved(ByVal Index As Integer, ByVal Value As Variant)
   txtField(Index) = IFNull(Value, "")
   
   With GridEditor1
      Select Case Index
      Case 2, 3
         .TextMatrix(p_nActiveRow + 2, Index) = IFNull(Value, "")
      Case 5, 6
         .TextMatrix(p_nActiveRow + 2, Index - 1) = IFNull(Value, "")
      End Select
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   pnIndex = Index
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

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_LostFocus"
   On Error GoTo errProc
   Select Case Index
   Case 1
      txtField(Index).Text = TitleCase(txtField(Index).Text)
   Case 2, 3, 5, 6
      txtField(Index).Text = getCTime(txtField(Index).Text)
   End Select
   
   oRecord.Detail(p_nActiveRow, Index) = txtField(Index).Text
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   
   With GridEditor1
      .Cols = 6
      .Rows = 3
      
'      .RowHeight(0) = 350
'      .RowHeight(1) = 400
'      .RowHeight(2) = 400
      .Font = "Arial"
      .ForeColor = oApp.getColor("ET0")
      .CellFontSize = 10
      .FontWidth = 6
      
      
      .MergeCells = flexMergeRestrictRows
      .MergeRow(0) = True
      .MergeCol(1) = True
      
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellBackColor = &H80000018
         .ColEnabled(lnCtr) = False
      Next
      
      'set backcolor in row 2
      .Row = 2
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellBackColor = &HE0E0E0
         .CellForeColor = &HFF&
      Next
            
      .TextMatrix(0, 0) = ""
      .TextMatrix(1, 1) = "DAY"
      .TextMatrix(0, 2) = "1ST QUARTER"
      .TextMatrix(0, 3) = "1ST QUARTER"
      .TextMatrix(0, 4) = "2ND QUARTER"
      .TextMatrix(0, 5) = "2ND QUARTER"
      .TextMatrix(1, 2) = "IN"
      .TextMatrix(1, 3) = "OUT"
      .TextMatrix(1, 4) = "IN"
      .TextMatrix(1, 5) = "OUT"
          
      .RowHeightMin = 350
   
      .ColWidth(0) = 0
      .ColWidth(1) = 2000
      .ColWidth(2) = 1478
      .ColWidth(3) = 1478
      .ColWidth(4) = 1478
      .ColWidth(5) = 1478
   
      .ColAlignment(1) = 4
      .ColAlignment(2) = 4
      .ColAlignment(3) = 4
      .ColAlignment(4) = 4
      .ColAlignment(5) = 4
      
      .Row = 2
      .Col = 1
   End With
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(3).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   
   cmdButton(0).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(7).Visible = lbShow
   cmdButton(8).Visible = lbShow
   
   xrFrame1.Enabled = lbShow
End Sub

Private Sub ClearFields()
   Dim lnCtr As Integer
   
   With txtField
      For lnCtr = 0 To .Count - 1
         If oRecord.EditMode = xeModeAddNew Then
            txtField(lnCtr) = oRecord.Detail(0, lnCtr)
         Else
            txtField(lnCtr) = ""
         End If
      Next
      
      If oRecord.EditMode = xeModeAddNew Then
         chkFlexi.Value = oRecord.Detail(0, "cFlexiTym")
      Else
         chkFlexi.Value = Unchecked
      End If
      cmbSchedule.ListIndex = 0
   End With
   
   With GridEditor1
      .Rows = 3
      
      .TextMatrix(2, 1) = ""
      .TextMatrix(2, 2) = ""
      .TextMatrix(2, 3) = ""
      .TextMatrix(2, 4) = ""
      .TextMatrix(2, 5) = ""
   End With
   
   p_nActiveRow = 0
End Sub

Private Sub loadCurrentDetail(ByVal nRow As Integer)
   Dim lnCtr As Integer

   With oRecord
      For lnCtr = 0 To txtField.Count - 1
         txtField(lnCtr) = IFNull(.Detail(nRow, lnCtr), "")
      Next

      chkFlexi.Value = IFNull(.Detail(nRow, "cFlexiTym"), 0)

      With GridEditor1
         .TextMatrix(nRow + 4, 0) = oRecord.Detail(nRow, "sShiftNme")
         .TextMatrix(nRow + 4, 1) = oRecord.Detail(nRow, "dTimeInAM")
         .TextMatrix(nRow + 4, 2) = oRecord.Detail(nRow, "dTmeOutAM")
         .TextMatrix(nRow + 4, 3) = oRecord.Detail(nRow, "dTimeInPM")
         .TextMatrix(nRow + 4, 4) = oRecord.Detail(nRow, "dTmeOutPM")
      End With
   End With
End Sub

Private Sub LoadRecord()
   Dim lnCtr As Integer

   With oRecord
      For lnCtr = 0 To txtField.Count - 1
         txtField(lnCtr) = IFNull(.Detail(0, lnCtr), "")
      Next

      chkFlexi.Value = IFNull(.Detail(0, "cFlexiTym"), 0)
      
      With GridEditor1
         .Rows = oRecord.ItemCount + 2
         For lnCtr = 0 To oRecord.ItemCount - 1
            .TextMatrix(lnCtr + 2, 1) = IIf(oRecord.Detail(lnCtr, "nDayOfWik") = 0, "", cmbSchedule.List(oRecord.Detail(lnCtr, "nDayOfWik")))
            .TextMatrix(lnCtr + 2, 2) = oRecord.Detail(lnCtr, "dTimeInAM")
            .TextMatrix(lnCtr + 2, 3) = oRecord.Detail(lnCtr, "dTmeOutAM")
            .TextMatrix(lnCtr + 2, 4) = oRecord.Detail(lnCtr, "dTimeInPM")
            .TextMatrix(lnCtr + 2, 5) = oRecord.Detail(lnCtr, "dTmeOutPM")
         Next
         
         'set previous row to default back color
         For lnCtr = 0 To .Cols - 1
            .Col = lnCtr
            .CellBackColor = &H80000005
         Next
            
         'set current row back color
         .Row = .Rows - 1
         For lnCtr = 0 To .Cols - 1
            .Col = lnCtr
            .CellBackColor = &HE0E0E0
         Next
         
         p_nActiveRow = .Rows - 3
      End With
      
      cmbSchedule.ListIndex = p_nActiveRow
   End With
End Sub

Private Sub ComputeShift(ByVal sLogDate, ByVal dLoginx As Variant, ByVal dLogout As Variant, ByVal nBreak As Integer, ByVal nMinWork As Integer, ByVal bPaid As Boolean, dPeriod() As Variant)
    Dim lnRemain As Integer
    
    'Initialize the Periods
    dPeriod(0) = Null
    dPeriod(1) = Null
    dPeriod(2) = Null
    dPeriod(3) = Null
    
    dPeriod(0) = CDate(sLogDate + " " + Format(dLoginx, "HH:MM"))
    
    'Get the possible immediate logout of the shift...
    If Not IsDate(dLogout) Then
        'For unspecified first period out and second period in
        'Unspecified first period out and second period in means
        ' 1. we will not monitor the exact time of break of employee
        ' 2. break could be a paid break
                
        'To get the possible logout:
        ' 1. add the total minutes of work to login
        ' 2. add the meal break if not paid break
        dPeriod(1) = DateAdd("n", nMinWork + IIf(bPaid, 0, nBreak), dPeriod(0))
        dPeriod(2) = Null
        dPeriod(3) = Null
    Else
        dPeriod(1) = CDate(sLogDate + " " + Format(dLogout, "HH:MM"))
        
        'Adjust p_dLogOutAM if p_dLoginxAM seems to be higher
        If (dPeriod(0) > dPeriod(1)) Then
            dPeriod(1) = DateAdd("D", 1, dPeriod(1))
        End If
                
                
        'Get the possible login for the second period through the logout from first period and break
        dPeriod(2) = DateAdd("n", nBreak, dPeriod(1))
        
        'Compute for the remaining work minutes...
        lnRemain = nMinWork - DateDiff("n", dPeriod(0), dPeriod(1))
        
        'Get the second period logout through the remaining work minutes...
        dPeriod(3) = DateAdd("n", lnRemain, dPeriod(2))
    End If
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oApp
      .xLogError Err.Number, Err.Description, pxeModuleName, lsProcName, Erl
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

