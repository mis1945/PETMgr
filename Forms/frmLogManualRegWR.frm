VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLogManualRegWR 
   BorderStyle     =   0  'None
   Caption         =   "Log Manual"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   13440
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   555
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
      Picture         =   "frmLogManualRegWR.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   13440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1180
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
      Picture         =   "frmLogManualRegWR.frx":077A
   End
   Begin xrControl.xrButton cmdDetail 
      Height          =   600
      Index           =   2
      Left            =   13440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   555
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Retreive"
      AccessKey       =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmLogManualRegWR.frx":0EF4
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   645
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   1138
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
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   105
         Width           =   2415
      End
      Begin VB.Shape Shape4 
         Height          =   405
         Index           =   0
         Left            =   10455
         Top             =   105
         Width           =   2475
      End
      Begin VB.Shape Shape3 
         Height          =   345
         Index           =   0
         Left            =   10485
         Top             =   135
         Width           =   2415
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
         Height          =   285
         Left            =   10515
         TabIndex        =   2
         Tag             =   "eb0;et0"
         Top             =   165
         Width           =   2355
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         Top             =   180
         Width           =   1365
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   8955
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1215
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   15796
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin xrControl.xrFrame xrFrame1 
         Height          =   1335
         Left            =   105
         Tag             =   "wt0;fb0"
         Top             =   1700
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   2355
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            Top             =   500
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   6
            Left            =   10275
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   870
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   5
            Left            =   10275
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   500
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   4
            Left            =   6135
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   9
            Top             =   870
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   3
            Left            =   6135
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   500
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   870
            Width           =   2415
         End
         Begin VB.TextBox txtOthers 
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
            Index           =   0
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   6
            Top             =   80
            Width           =   4395
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   9030
            TabIndex        =   19
            Top             =   945
            Width           =   1095
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime In"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   9030
            TabIndex        =   18
            Top             =   570
            Width           =   960
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2nd Q. Time Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   4575
            TabIndex        =   17
            Top             =   945
            Width           =   1305
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2nd Q. Time In"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   4590
            TabIndex        =   16
            Top             =   570
            Width           =   1170
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   15
            Top             =   195
            Width           =   825
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1st Q. Time In"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   105
            TabIndex        =   14
            Top             =   570
            Width           =   1125
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1st Q. Time Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   105
            TabIndex        =   13
            Top             =   945
            Width           =   1260
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   1575
         Index           =   0
         Left            =   105
         Tag             =   "wt0;fb0"
         Top             =   105
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   2778
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
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
            Index           =   0
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   120
            Width           =   2415
         End
         Begin VB.TextBox txtField 
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
            Left            =   7695
            MaxLength       =   50
            TabIndex        =   24
            Top             =   635
            Width           =   5000
         End
         Begin VB.TextBox txtField 
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
            Index           =   2
            Left            =   7695
            MaxLength       =   50
            TabIndex        =   23
            Top             =   120
            Width           =   2355
         End
         Begin VB.ComboBox cmbField 
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
            Index           =   7
            ItemData        =   "frmLogManualRegWR.frx":166E
            Left            =   7680
            List            =   "frmLogManualRegWR.frx":1681
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   635
            Width           =   2925
         End
         Begin VB.TextBox txtField 
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
            Index           =   10
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "F3 / Enter to Search "
            Top             =   635
            Width           =   3675
         End
         Begin VB.TextBox txtField 
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
            Index           =   9
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1015
            Width           =   5715
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Name"
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
            Left            =   6165
            TabIndex        =   30
            Top             =   690
            Width           =   1185
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   105
            TabIndex        =   29
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   6165
            TabIndex        =   28
            Top             =   210
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   315
            Left            =   1725
            Tag             =   "et0;ht2"
            Top             =   210
            Width           =   2415
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reason"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   11
            Left            =   120
            TabIndex        =   27
            Top             =   680
            Width           =   660
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   12
            Left            =   120
            TabIndex        =   26
            Top             =   1045
            Width           =   765
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5760
         Left            =   105
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3060
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   10160
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         WordWrap        =   -1  'True
         Enabled         =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLogManualRegWR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeModuleName = "frmLogManualReg"
Private Const pxBRANCHCODES = "M001»H001»N001»PHO1»PHO2"

Private WithEvents oTrans As clsLogManualWR
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim pnRow As Integer
Dim pnActiveRow As Integer

Dim pbCtrlPress As Boolean
Dim pbFormLoad As Boolean
Dim pbDetailGotFocus As Boolean
Dim pbUpdateMode As Boolean
Dim pbCopy2All As Boolean
Dim pnLastSelc As Integer
Dim pbMoveUpxx As Boolean
Dim pbMoveDown As Boolean

Private Sub cmdButton_Click(Index As Integer)
10       Dim lsOldProc As String
20       Dim lnRep As Integer

30       lsOldProc = "cmdButton_Click"
40       'On Error GoTo errProc

50       Select Case Index
         Case 3   'browse
320            If pnIndex > 1 Then pnIndex = 1
         
330            If oTrans.SearchTransaction(txtOthers(pnIndex), True) Then
340               Call ClearFields
350               Call InitGrid
360               Call LoadMaster
370               Call LoadDetail
380               Call detailFieldChange
390            End If
400            GoTo endWithFocus
500         Case 5   'close
510            Unload Me
520      End Select
   
endProc:
530      Exit Sub
endWithFocus:
540      If xrFrame1.Enabled = False Then
550         txtSearch(0).SetFocus
560      Else
570         txtOthers(0).SetFocus
580      End If
590      GoTo endProc
errProc:
600      ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub cmdDetail_Click(Index As Integer)
10       Dim lnCtr As Integer
20       Dim lnCtr1 As Integer
30       Dim lnRow As Integer

40       If Not pbUpdateMode Then GoTo endWithFocus
   
50       Select Case Index
      Case 0   'delete
60             If oTrans.ItemCount = 0 And oTrans.Detail(0, 0) = "" Then GoTo endProc
         
70             If pnActiveRow < 2 Then Exit Sub 'MAC(07.02.12)
80             If oTrans.DeleteDetail(pnActiveRow - 2) Then
90                pnActiveRow = pnActiveRow - 1
100               If pnActiveRow < 2 Then pnActiveRow = 2
110               LoadDetail
120               ClearOthers
130               Call detailFieldChange
140               GoTo endWithFocus
150            End If
160         Case 1   'add
170            If oTrans.ItemCount = 0 Then
180               oTrans.InitTransaction
190               pnRow = 0
200            End If
         
'         MoveToLastRec
         'Causes an error
         'kalyptus-2012.04.18
210            If oTrans.Detail(pnRow, "xfullname") = "" Then GoTo endProc
         'If oTrans.Detail(pnRow - 2, 0) = "" Then GoTo endProc
            
220            With MSFlexGrid1
         
230               .Rows = .Rows + 1
240               .Row = .Rows - 2
250               pnRow = .Row
            
260               .TextMatrix(pnRow + 1, 0) = pnRow + 1
            
270               For lnCtr = 0 To 6
280                  .TextMatrix(pnRow + 1, lnCtr + 1) = IFNull(oTrans.Detail(pnRow - 1, lnCtr), "")
290               Next
            
300               Call oTrans.addDetail
310               ClearOthers
            
320               If MSFlexGrid1.Rows > 13 Then
330                  .ColWidth(1) = 3050
'               .TopRow = pnRow
340               Else
350                  .ColWidth(1) = 3300
360               End If
               
370               .TextMatrix(pnRow + 1, 0) = pnRow
380               .Row = pnRow + 1
390               pnActiveRow = .Row
                        

400               Call detailFieldChange
410            End With
420            GoTo endWithFocus
430         Case 2
440            If oTrans.loadEmployee Then
450               lnRow = oTrans.ItemCount
            
460               With MSFlexGrid1
470                  MSFlexGrid1.Rows = lnRow + 2
'               If MSFlexGrid1.Rows > 13 Then .ColWidth(1) = 2750
               
480                  If MSFlexGrid1.Rows > 13 Then
490                     .ColWidth(1) = 3050
500                  Else
510                     .ColWidth(1) = 3300
520                  End If
               
530                  For lnCtr = 1 To lnRow
540                     MSFlexGrid1.TextMatrix(lnCtr + 1, 0) = lnCtr
550                     For lnCtr1 = 0 To 6
560                        MSFlexGrid1.TextMatrix(lnCtr + 1, lnCtr1 + 1) = IFNull(oTrans.Detail(lnCtr - 1, lnCtr1), "")
570                     Next
580                  Next
590               End With
            
600               Call detailFieldChange
610               Call flexFocus
620            Else
630               MSFlexGrid1.Rows = 3
640            End If
650            GoTo endWithFocus
660      End Select
   
endProc:
670      Exit Sub
endWithFocus:
680      If xrFrame1.Enabled Then
690         txtOthers(0).SetFocus
700      ElseIf xrFrame3(1).Enabled Then
710         txtSearch(0).SetFocus
720      End If
730      GoTo endProc
End Sub


Private Sub cmdDetail_GotFocus(Index As Integer)
10       Select Case Index
      Case 0, 1
20             MSFlexGrid1_Click
30       End Select
End Sub

Private Sub Form_Activate()
10       Dim lsOldProc As String

20       lsOldProc = "Form_Activate"
30       'On Error GoTo errProc

40       oApp.MenuName = Me.Tag
50       Me.ZOrder 0

60       If bLoaded = False Then
70          bLoaded = True
      
80          If InStr(1, pxBRANCHCODES, oApp.BranchCode) = 0 Then
'80          If oApp.BranchCode <> "M001" Then
90             txtField(1).Text = oApp.BranchName
100            cmbField(7).Visible = False
110            txtField(1).Visible = True
120            txtField(1).TabStop = False
130            txtField(1).Locked = True
140         Else
150            lblField(1) = "Branch Name"
160            cmbField(7).Visible = False
170            txtField(1).Visible = True
180         End If
190      End If

200      If Not pbFormLoad Then pbFormLoad = True
210      pnActiveRow = 2
220      pnRow = 2
   
'   txtSearch(0).SetFocus
230      pbUpdateMode = False
endProc:
240      Exit Sub
errProc:
250      ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10       With MSFlexGrid1
20          Select Case KeyCode
      Case vbKeyReturn
30             If Not pbMoveDown Then Exit Sub
40             If GetFocus = .hWnd Then Exit Sub
50             SetNextFocus
60          Case vbKeyDown
70             If pbCtrlPress Then

80                If pnActiveRow = .Rows - 1 Then Exit Sub
            
90                If Not (pnActiveRow = 0) Then
100                  pnActiveRow = pnActiveRow + 1
110               End If
            
120               .Row = pnActiveRow
            
130               If Not .RowIsVisible(.Row + IIf(.Row = .Rows - 1, 0, 1)) Then .TopRow = .Row - 1

140               Call detailFieldChange
150            Else
160               If Not pbMoveDown Then Exit Sub
170               If Not (GotFocus = cmbField(7).hWnd) Then
180                  SetNextFocus
190               End If
200            End If
210         Case vbKeyUp
220            If pbCtrlPress Then
230               If .Row > 2 Then
240                  If Not (pnActiveRow = 1) Then
250                     pnActiveRow = pnActiveRow - 1
260                  End If
               
270                  .Row = pnActiveRow
   
280                  If Not .RowIsVisible(.Row) Then .TopRow = .TopRow - 1

290                  Call detailFieldChange
300               End If
310            Else
320               If Not pbMoveUpxx Then Exit Sub
330               If Not (GotFocus = cmbField(7).hWnd) Then
340                     SetPreviousFocus
350               End If
360            End If
370         Case vbKeyControl
380            pbCtrlPress = True
390         End Select
400      End With
End Sub

Private Sub Form_Load()
10       Dim lsOldProc As String

20       lsOldProc = "Form_Load"
'   'On Error GoTo errProc

30       CenterChildForm mdiMain, Me

70       Set oSkin = New clsFormSkin
80       Set oSkin.AppDriver = oApp
90       Set oSkin.Form = Me
100      oSkin.ApplySkin xeFormTransEqualRight

110      Set oTrans = New clsLogManualWR
120      Set oTrans.AppDriver = oApp
130      oTrans.Branch = oApp.BranchCode
140      oTrans.InitTransaction

150      Call ClearFields
160      Call InitGrid
170      Call InitForm(0)
endProc:
180      Exit Sub
errProc:
190      ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
10       Set oTrans = Nothing
20       Set oSkin = Nothing

30       pnActiveRow = 0
40       pbFormLoad = False
50       pbCtrlPress = False
60       pbUpdateMode = False
70       bLoaded = False
   
End Sub

Private Sub detailFieldChange()
10       Dim lnCtr As Integer
20       Dim lnRow As Integer
   
30       lnRow = pnActiveRow
40       SetGridRowColor (lnRow)
45       pnActiveRow = MSFlexGrid1.Row
50       lnRow = pnActiveRow
   
55       With oTrans
60          For lnCtr = 0 To 6
70             txtOthers(lnCtr) = IFNull(.Detail(lnRow - 2, lnCtr), "")
71             Select Case lnCtr
               Case 1
72                txtOthers(lnCtr).Enabled = oTrans.IsActiveAMInxxxx
73             Case 2
74                txtOthers(lnCtr).Enabled = oTrans.IsActiveAMOutxxx
75             Case 3
76                txtOthers(lnCtr).Enabled = oTrans.IsActivePMInxxxx
77             Case 4
78                txtOthers(lnCtr).Enabled = oTrans.IsActivePMOutxxx
79             Case 5
80                txtOthers(lnCtr).Enabled = oTrans.IsActiveOTimeInx
81             Case 6
82                txtOthers(lnCtr).Enabled = oTrans.IsActiveOTimeOut
83             End Select
85          Next
110      End With
   
120      If txtOthers(0).Enabled Then txtOthers(0).SetFocus
End Sub

Private Sub InitGrid()
10       Dim lnCtr As Integer
20       With MSFlexGrid1
30          .Cols = 8
40          .Rows = 2
50          .MergeCells = flexMergeFree
      
60          .Clear
      
70          .Row = 0
      
      'column alignment
80          For lnCtr = 0 To .Cols - 1
90             .Col = lnCtr
100            .CellFontBold = True
110            .CellAlignment = flexAlignCenterCenter
120         Next
      
130         .MergeRow(0) = True
140         .TextMatrix(0, 0) = ""
150         .TextMatrix(0, 1) = ""
160         .TextMatrix(0, 2) = "1ST QUARTER"
170         .TextMatrix(0, 3) = "1ST QUARTER"
180         .TextMatrix(0, 4) = "2ND QUARTER"
190         .TextMatrix(0, 5) = "2ND QUARTER"
200         .TextMatrix(0, 6) = "OVERTIME"
210         .TextMatrix(0, 7) = "OVERTIME"

220         .Row = 1
      
      'column alignment
230         For lnCtr = 0 To .Cols - 1
240            .Col = lnCtr
250            .CellFontBold = True
260            .CellAlignment = flexAlignCenterCenter
270         Next
      
      'column title
280         .TextMatrix(1, 0) = "No"
290         .TextMatrix(1, 1) = "Employee Name"
300         .TextMatrix(1, 2) = "In"
310         .TextMatrix(1, 3) = "Out"
320         .TextMatrix(1, 4) = "In"
330         .TextMatrix(1, 5) = "Out"
340         .TextMatrix(1, 6) = "In"
350         .TextMatrix(1, 7) = "Out"
      
360         .RowHeightMin = 338
      
      'column width
370         .ColWidth(0) = 500
380         .ColWidth(1) = 3300
390         .ColWidth(2) = 1500
400         .ColWidth(3) = 1500
410         .ColWidth(4) = 1500
420         .ColWidth(5) = 1500
430         .ColWidth(6) = 1500
440         .ColWidth(7) = 1500
      
      'column allinment
450         .ColAlignment(0) = flexAlignLeftCenter
460         .ColAlignment(1) = flexAlignLeftCenter
470         .ColAlignment(2) = flexAlignCenterCenter
480         .ColAlignment(3) = flexAlignCenterCenter
490         .ColAlignment(4) = flexAlignCenterCenter
500         .ColAlignment(5) = flexAlignCenterCenter
510         .ColAlignment(6) = flexAlignCenterCenter
520         .ColAlignment(7) = flexAlignCenterCenter
            
530         .Rows = 3
540         .TextMatrix(2, 0) = "1"
      
550         .Row = 2
560         pnLastSelc = .Row
570         SetGridRowColor (.Row)
580      End With
End Sub

Private Sub InitForm(lnStat As Integer)
10       Dim lbShow As Boolean
20       Dim lnCtr As Integer

30       lbShow = IIf(lnStat = 0, False, True)
40       cmdButton(3).Visible = Not lbShow

60       cmdButton(5).Visible = Not lbShow



   
90       txtSearch(0).Enabled = Not lbShow

110      txtField(1).Enabled = lbShow
120      txtField(2).Enabled = lbShow
130      cmbField(7).Enabled = lbShow
   
140      If lbShow Then
150         xrFrame1.Enabled = lbShow
220      Else
230         For lnCtr = 1 To 6
240            txtOthers(lnCtr).Enabled = False
250         Next
260         xrFrame1.Enabled = lbShow
270      End If
End Sub
Private Sub LoadMaster()
      Dim lnCtr As Integer
   
   For lnCtr = 0 To 2
      If lnCtr = 1 Then
         If Len(oTrans.Master(7)) = 4 Then
            txtField(lnCtr) = oTrans.Master(lnCtr)
            txtField(1).Visible = True
            cmbField(7).Visible = False
            lblField(1) = "Branch Name"
         Else
            cmbField(7).ListIndex = oTrans.Master(7)
            txtField(1).Visible = False
            cmbField(7).Visible = True
            lblField(1) = "Employee Level"
         End If
      ElseIf lnCtr = 2 Then
         txtField(lnCtr) = strLongDate(oTrans.Master(lnCtr))
      Else
         txtField(lnCtr) = oTrans.Master(lnCtr)
      End If
   Next
   
   txtSearch(0) = txtField(0)
   txtField(2) = Format(txtField(2), "MMMM DD, YYYY")
   txtField(9) = IFNull(oTrans.Master(9), "")
   txtField(10) = IFNull(oTrans.Master(10), "")

   If oTrans.Master("cTranStat") = "4" Then
      Label2.Caption = "APPLIED"
   Else
      Label2.Caption = TransStat(CInt(oTrans.Master("cTranStat")))
   End If

End Sub

Private Sub LoadDetail()
       Dim lnRow As Integer
       Dim lnCtr As Integer

       If oTrans.ItemCount = 0 Then oTrans.InitTransaction
       lnRow = oTrans.ItemCount

       With MSFlexGrid1
          .Rows = 3
          .Rows = lnRow + 2
      
          If MSFlexGrid1.Rows > 13 Then
             .ColWidth(1) = 3050
         Else
            .ColWidth(1) = 3300
         End If
      
         For lnCtr = 0 To lnRow - 1
            .TextMatrix(lnCtr + 2, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 2, 1) = IFNull(oTrans.Detail(lnCtr, 0), "")
            .TextMatrix(lnCtr + 2, 2) = IFNull(oTrans.Detail(lnCtr, 1), "")
            .TextMatrix(lnCtr + 2, 3) = IFNull(oTrans.Detail(lnCtr, 2), "")
            .TextMatrix(lnCtr + 2, 4) = IFNull(oTrans.Detail(lnCtr, 3), "")
            .TextMatrix(lnCtr + 2, 5) = IFNull(oTrans.Detail(lnCtr, 4), "")
            .TextMatrix(lnCtr + 2, 6) = IFNull(oTrans.Detail(lnCtr, 5), "")
            .TextMatrix(lnCtr + 2, 7) = IFNull(oTrans.Detail(lnCtr, 6), "")
         Next
      
'      .Row = 2
'      pnActiveRow = .Row
'      pnRow = pnActiveRow - 2
      
           Call detailFieldChange 'set info into textbox
      End With
   
End Sub

Private Sub MSFlexGrid1_Click()
10       Dim lnCtr As Integer
   
20       With oTrans
30          pnActiveRow = MSFlexGrid1.Row
40          pnRow = pnActiveRow - 2

      
'      If pnActiveRow > oTrans.ItemCount Then GoTo endProc
      
50          For lnCtr = 0 To 6
60             txtOthers(lnCtr) = Format(IFNull(.Detail(pnRow, lnCtr), ""), "HH:MM AM/PM")
70          Next
80       End With
   
90       Call detailFieldChange
endProc:
100      If xrFrame1.Enabled = False Then
110         txtSearch(0).SetFocus
120      Else
130         txtOthers(0).SetFocus
140      End If
End Sub

Private Sub MSFlexGrid1_GotFocus()
10       pbDetailGotFocus = True
20       If xrFrame1.Enabled = False Then Exit Sub
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
10       If pbCtrlPress Then
20          If KeyCode = vbKeyControl Then pbCtrlPress = False
30       End If
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
10       If Index = 6 Then Label2.Caption = TransStat(CInt(Value))
End Sub

Private Sub txtField_GotFocus(Index As Integer)
10       With txtField(Index)
20          .BackColor = oApp.getColor("HT1")
30          .SelStart = 0
40          .SelLength = Len(.Text)
50       End With
   
60       pbMoveDown = True
70       pbMoveUpxx = True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
10       With txtField(Index)
20          .BackColor = oApp.getColor("EB")
30       End With
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
10       Select Case KeyCode
      Case vbKeyF3, vbKeyReturn
20             If oTrans.SearchTransaction(txtSearch(Index).Text, True) Then
30                ClearFields
40                InitGrid
50                LoadMaster
60                LoadDetail
70             End If
80        End Select
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
10       With txtSearch(Index)
20          .BackColor = oApp.getColor("EB")
30       End With

40       pnIndex = Index
End Sub
Private Sub txtSearch_GotFocus(Index As Integer)
10       With txtSearch(Index)
20          .BackColor = oApp.getColor("HT1")
30       End With

50       pnIndex = Index
End Sub
Private Sub ClearFields()
10       Dim loTxt As TextBox
   
20       For Each loTxt In txtSearch
30          loTxt.Text = ""
40       Next
50       For Each loTxt In txtField
60          loTxt.Text = ""
70       Next
80       For Each loTxt In txtOthers
90          loTxt.Text = ""
100      Next
End Sub
Private Sub txtOthers_GotFocus(Index As Integer)
10       With txtOthers(Index)
20          .BackColor = oApp.getColor("HT1")
30       End With
   
40       Select Case Index
      Case 0
50             pbMoveUpxx = False
60             pbMoveDown = True
70          Case 6
80             pbMoveUpxx = True
90             pbMoveDown = False
100         Case Else
110            pbMoveDown = True
120            pbMoveUpxx = True
130      End Select
   
140      pnIndex = Index
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
10       Dim lnRow As Integer
20       Dim lbCancel As Boolean
   
30       pnRow = MSFlexGrid1.Row
   
40       Select Case Index
      Case 0
50             Select Case KeyCode
            Case vbKeyReturn
60                   If oTrans.Detail(IIf(pnActiveRow > 0, pnActiveRow - 1, pnActiveRow), Index) = "" Then
70                         If oTrans.SearchDetail(pnRow - 2, Index, txtOthers(Index)) Then
80                            txtOthers(Index).SetFocus
90                            Call LoadDetail
100                           Call MoveToLastRec
110                        End If
120                  Else
130                     MsgBox "Unable to overwrite row data." + vbCrLf + _
                     "Delete the row if incorrect entry.", vbInformation
140                     txtOthers(Index).SetFocus
150                  End If
160               Case vbKeyF3
170                  If oTrans.Detail(IIf(pnActiveRow > 0, pnActiveRow - 1, pnActiveRow), Index) = "" Then
180                     If oTrans.SearchDetail(pnRow - 2, Index, txtOthers(Index)) Then
190                        Call LoadDetail
200                        Call MoveToLastRec
210                     End If
220                  Else
230                     MsgBox "Unable to overwrite row data." + vbCrLf + _
                     "Delete the row if incorrect entry.", vbInformation
240                     txtOthers(Index).SetFocus
250                  End If
260               Case vbKeyAdd
270                  If pbCtrlPress Then Call cmdDetail_Click(1)
280               Case vbKeySubtract
290                  If pbCtrlPress Then Call cmdDetail_Click(0)
300            End Select
'Mac PH 08.11.12
'      Case 4
'         If KeyCode = vbKeyReturn Then
'            Call txtOthers_Validate(4, lbCancel)
'            If Not lbCancel Then
'               If MSFlexGrid1.Row + 1 < MSFlexGrid1.Rows Then
'                  MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'                  MSFlexGrid1_Click
'               End If
'            End If
'         End If
310         Case Else
320            Select Case KeyCode
            Case vbKeyF12
330                  If pbCtrlPress Then Call CopyT2AllEmp
340            End Select
350      End Select
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
10       With txtOthers(Index)
20          .BackColor = oApp.getColor("EB")
30       End With
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
10       If Index = 0 Or (pnActiveRow < 1) Then Exit Sub
   
20       Dim lsTime As String
   
   'dont allow to pass time parameter when Employee Name is null.
30       If oTrans.Detail(pnActiveRow - 2, 0) = "" Then
40          ClearOthers
50          Exit Sub
60       End If
   
'65       If oTrans.Detail(pnActiveRow - 2, Index) <> "" Then
'66          txtOthers(Index) = oTrans.Detail(pnActiveRow - 2, Index)
'67          Exit Sub
'68       End If
   
70       Select Case Index
      Case 1, 2
80             oTrans.Detail(pnActiveRow - 2, Index) = getCTime(txtOthers(Index))
90          Case 3, 4, 5, 6
100            If InStr(txtOthers(Index), "p") = True Or InStr(txtOthers(Index), "pm") Then
110               lsTime = txtOthers(Index) & "p"
120               oTrans.Detail(pnActiveRow - 2, Index) = getCTime(lsTime)
130            Else
140               oTrans.Detail(pnActiveRow - 2, Index) = getCTime(txtOthers(Index))
150            End If
         
         'Mac PH (08.23.12)
         'Moves down when pm out is validated.
160            If Index = 4 Then
170               pbCtrlPress = True
180               Call Form_KeyDown(vbKeyDown, False)
190               pbCtrlPress = False
200            End If
210      End Select
End Sub

Private Sub CopyT2AllEmp()
10       Dim lnCtr As Integer
20       pbCopy2All = True
30       For lnCtr = 0 To oTrans.ItemCount - 1
40          With oTrans
50             .Detail(lnCtr, pnIndex) = txtOthers(pnIndex)
         
60             MSFlexGrid1.TextMatrix(lnCtr + 2, pnIndex + 1) = .Detail(lnCtr, pnIndex)
70          End With
80       Next
90       pbCopy2All = False
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer, ByVal Value As Variant)
10       With txtOthers(Index)
20          .Text = IFNull(Value, "")
30       End With

'   If pnActiveRow < 1 Or pbCopy2All Or (MSFlexGrid1.Row = MSFlexGrid1.Rows - 1) Then Exit Sub
40       If pnActiveRow < 1 Or pbCopy2All Then Exit Sub
   
50       MSFlexGrid1.TextMatrix(pnActiveRow, Index + 1) = IFNull(Value, "")
End Sub

Private Sub ClearOthers()
10       Dim lnCtr As Integer
   
20       For lnCtr = 0 To 6
30          txtOthers(lnCtr) = ""
40       Next
   
'   pnRow = oTrans.ItemCount - 1
'   MoveToLastRec
End Sub
Private Sub flexFocus()
10       With MSFlexGrid1
20          If .Row > 15 Then .TopRow = 1
30          .Row = 2
40          pnActiveRow = .Row
50          pbDetailGotFocus = True
60       End With
End Sub

Private Sub SetGridRowColor(ByVal lnRow As Integer)
10       Dim lnCtr As Integer
   
20       With MSFlexGrid1
30          .FillStyle = flexFillRepeat
      
40          .Row = IIf(pnLastSelc = .Rows, pnLastSelc - 1, pnLastSelc)
50          .RowSel = pnLastSelc - 1
60          .Col = 1
70          .ColSel = .Cols - 1
80          .CellBackColor = &HFFFFFF

90          .Row = lnRow
100         .RowSel = lnRow
110         .Col = 1
120         .ColSel = .Cols - 1
130         .CellBackColor = &HFF8080
      
140         pnLastSelc = .Row
150         If Not .RowIsVisible(.Row + IIf(.Row = .Rows - 1, 0, 1)) Then .TopRow = .Row
160      End With
End Sub

Private Sub MoveToLastRec()
10       With MSFlexGrid1
20          .Row = .Rows - 1
30          pnActiveRow = .Row
40          pnRow = pnActiveRow - 2
      
50          If Not .RowIsVisible(.Row) Then
60             .TopRow = .Row - 10
70          End If
      
80          detailFieldChange
90       End With
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
10       With oApp
20          .xLogError Err.Number, Err.Description, pxeModuleName, lsProcName, Erl
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


