VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmBatchShftReg 
   BorderStyle     =   0  'None
   Caption         =   "Batch Shift Change Register"
   ClientHeight    =   9855
   ClientLeft      =   0
   ClientTop       =   4320
   ClientWidth     =   20415
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   19084.29
   ScaleMode       =   0  'User
   ScaleWidth      =   31352.44
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4395
      Left            =   120
      TabIndex        =   0
      Top             =   5344
      Width           =   18705
      _ExtentX        =   32994
      _ExtentY        =   7752
      _Version        =   393216
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   19067
      TabIndex        =   1
      Top             =   1187
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
      Picture         =   "frmBatchShftReg.frx":0000
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1125
      Index           =   1
      Left            =   95
      Tag             =   "wt0;fb0"
      Top             =   542
      Width           =   18735
      _ExtentX        =   33046
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
         Text            =   "GMC Dagupan - Honda"
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
         TabIndex        =   2
         Text            =   "M00122-000015"
         Top             =   120
         Width           =   2415
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
         Index           =   15
         Left            =   105
         TabIndex        =   5
         Top             =   675
         Width           =   1185
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
         TabIndex        =   4
         Top             =   210
         Width           =   1365
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   8010
      Index           =   2
      Left            =   94
      Tag             =   "wt0;fb0"
      Top             =   1699
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   14129
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin xrControl.xrFrame xrFrame2 
         Height          =   3615
         Index           =   0
         Left            =   0
         Tag             =   "wt0;fb0"
         Top             =   0
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   6376
         BackColor       =   12632256
         Enabled         =   0   'False
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Index           =   10
            Left            =   1680
            TabIndex        =   32
            Text            =   "May 02, 2022"
            Top             =   1365
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
            Height          =   1155
            Index           =   3
            Left            =   1680
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   2235
            Width           =   5055
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   8
            Top             =   810
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   7
            Text            =   "M00122-000015"
            Top             =   240
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
            Index           =   2
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1800
            Width           =   5055
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Effectivity"
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
            Left            =   120
            TabIndex        =   33
            Top             =   1455
            Width           =   825
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            Left            =   120
            TabIndex        =   14
            Top             =   2280
            Width           =   780
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
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   900
            Width           =   405
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
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1485
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   420
            Index           =   0
            Left            =   1800
            Tag             =   "et0;ht2"
            Top             =   360
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
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   1890
            Width           =   615
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
            Left            =   6870
            TabIndex        =   10
            Tag             =   "eb0;et0"
            Top             =   300
            Width           =   2070
         End
         Begin VB.Shape Shape4 
            Height          =   420
            Index           =   0
            Left            =   6800
            Top             =   240
            Width           =   2220
         End
         Begin VB.Shape Shape3 
            Height          =   360
            Index           =   0
            Left            =   6820
            Top             =   270
            Width           =   2160
         End
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   3615
         Index           =   1
         Left            =   9720
         Tag             =   "wt0;fb0"
         Top             =   0
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   6376
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   1200
            TabIndex        =   30
            Text            =   "Day 7"
            Top             =   3015
            Width           =   6375
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   1200
            TabIndex        =   29
            Text            =   "Day 6"
            Top             =   2625
            Width           =   6375
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   1200
            TabIndex        =   28
            Text            =   "Day 5"
            Top             =   2235
            Width           =   6375
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   1200
            TabIndex        =   27
            Text            =   "Day 4"
            Top             =   1845
            Width           =   6375
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1200
            TabIndex        =   26
            Text            =   "Day 3"
            Top             =   1455
            Width           =   6375
         End
         Begin VB.TextBox txtOthers 
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
            Left            =   1200
            TabIndex        =   17
            Text            =   "Sayson, Marlon"
            Top             =   240
            Width           =   6375
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "Day 1"
            Top             =   675
            Width           =   6375
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "Day 2"
            Top             =   1065
            Width           =   6375
         End
         Begin VB.Label lblField 
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
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   330
            Width           =   510
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sunday"
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
            Left            =   120
            TabIndex        =   24
            Top             =   742
            Width           =   660
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monday"
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
            Left            =   120
            TabIndex        =   23
            Top             =   1132
            Width           =   690
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tueday"
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
            Left            =   120
            TabIndex        =   22
            Top             =   1522
            Width           =   630
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wednesday"
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
            Left            =   120
            TabIndex        =   21
            Top             =   1912
            Width           =   1035
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Thursday"
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
            Left            =   120
            TabIndex        =   20
            Top             =   2302
            Width           =   795
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Friday"
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
            Left            =   120
            TabIndex        =   19
            Top             =   2692
            Width           =   660
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saturday"
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
            Left            =   120
            TabIndex        =   18
            Top             =   3082
            Width           =   900
         End
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   19067
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   542
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
      Picture         =   "frmBatchShftReg.frx":077A
   End
End
Attribute VB_Name = "frmBatchShftReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeModuleName = "frmBatchShftReg"
Private Const pxBRANCHCODES = "M001»H001»N001"

Private WithEvents oTrans As clsBatchShiftSched
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
Dim pbDtlLoaded As Boolean


Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   Select Case Index
      
      Case 3   'browse
         If pnIndex = 0 Or pnIndex = 1 Then
          pbDtlLoaded = False
         If oTrans.SearchTransaction(txtSearch(pnIndex).Text, pnIndex = 0) Then
         
            Call ClearFields
            Call InitGrid
            Call LoadMaster
            Call LoadDetail
            Call detailFieldChange
            txtSearch(pnIndex).SetFocus
         End If
         End If
         
         
      Case 5   'close
         Unload Me
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
      
'      If Not (oApp.BranchCode = "M001" Or oApp.BranchCode = "N001" Or oApp.BranchCode = "H001") Then
        If InStr(1, pxBRANCHCODES, oApp.BranchCode) = 0 Then
         txtField(1).Visible = True
         txtField(1).TabStop = False
         txtField(1).Locked = True

      End If
   End If

   If Not pbFormLoad Then pbFormLoad = True
   pnActiveRow = 1
   pnRow = 0
   pbDtlLoaded = False
   txtSearch(0).SetFocus

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   With MSFlexGrid1
      Select Case KeyCode
      Case vbKeyReturn
         If Not pbMoveDown Then Exit Sub
         If GetFocus = .hWnd Then Exit Sub
         SetNextFocus
      Case vbKeyDown
         If pbCtrlPress Then

            If pnActiveRow = .Rows - 1 Then Exit Sub
            
            If Not (pnActiveRow = 0) Then
               pnActiveRow = pnActiveRow + 1
            End If
            
            .Row = pnActiveRow
            
            If Not .RowIsVisible(.Row + IIf(.Row = .Rows - 1, 0, 1)) Then .TopRow = .Row - 1

            Call detailFieldChange
         Else
            If Not pbMoveDown Then Exit Sub
            
            SetNextFocus
         End If
      Case vbKeyUp
         If pbCtrlPress Then
            If .Row > 2 Then
               If Not (pnActiveRow = 1) Then
                  pnActiveRow = pnActiveRow - 1
               End If
               
               .Row = pnActiveRow
   
               If Not .RowIsVisible(.Row) Then .TopRow = .TopRow - 1

               Call detailFieldChange
            End If
         Else
            If Not pbMoveUpxx Then Exit Sub

            SetPreviousFocus
         End If
      Case vbKeyControl
         pbCtrlPress = True
      End Select
   End With
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
'   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   Set oTrans = New clsBatchShiftSched
   Set oTrans.AppDriver = oApp
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Call ClearFields
   Call InitGrid
   Call InitForm(0)
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing

   pnActiveRow = 0
   pbFormLoad = False
   pbCtrlPress = False
   pbUpdateMode = False
   bLoaded = False
   
End Sub

Private Sub detailFieldChange()
   Dim lnCtr As Integer
   Dim lnRow As Integer
 
   
   xrFrame2(1).Enabled = False
   lnRow = pnActiveRow - 1
   txtOthers(0).Text = oTrans.Detail(lnRow, "xFullName")
   For lnCtr = 1 To 7
      txtOthers(lnCtr).Text = IFNull(oTrans.Detail(lnRow, "sShiftNm" & lnCtr), "")
   Next
   
   SetGridRowColor (lnRow + 1)
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Cols = 9
      .Rows = oTrans.ItemCount
      .Clear
      
      .Row = 0
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      

      'column title
      .TextMatrix(0, 0) = " "
      .TextMatrix(0, 1) = "Employee"
      .TextMatrix(0, 2) = "Sunday"
      .TextMatrix(0, 3) = "Monday"
      .TextMatrix(0, 4) = "Tuesday"
      .TextMatrix(0, 5) = "Wednesday"
      .TextMatrix(0, 6) = "Thursday"
      .TextMatrix(0, 7) = "Friday"
      .TextMatrix(0, 8) = "Saturday"
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      
      'column width
      .ColWidth(0) = 340
      .ColWidth(1) = 2815
      .ColWidth(2) = 2200
      .ColWidth(3) = 2200
      .ColWidth(4) = 2200
      .ColWidth(5) = 2200
      .ColWidth(6) = 2200
      .ColWidth(7) = 2200
      .ColWidth(8) = 2200
      
      'column allinment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignLeftCenter
      .ColAlignment(5) = flexAlignLeftCenter
      .ColAlignment(6) = flexAlignLeftCenter
      .ColAlignment(7) = flexAlignLeftCenter
      .ColAlignment(8) = flexAlignLeftCenter
            
      .Rows = 2
      .TextMatrix(1, 0) = "1"
      
      .Row = 1
      pnLastSelc = .Row
      SetGridRowColor (.Row)
   End With
End Sub


Private Sub InitForm(ByVal fnEdit As Integer)
   Dim lnCtr As Integer
   
   xrFrame2(1).Enabled = (fnEdit = 0)
   xrFrame2(0).Enabled = (fnEdit = 0)
   
   If fnEdit = 1 Then
      pnActiveRow = 1
      pnRow = 0
   End If
End Sub

Private Sub LoadMaster()
   Dim lnCtr As Integer
   
   For lnCtr = 0 To 3
     If lnCtr = 1 Then
         txtField(lnCtr) = strLongDate(oTrans.Master(lnCtr))
         
         ElseIf lnCtr = 3 Then
         txtField(lnCtr) = strLongDate(oTrans.Master(4))
      Else
         txtField(lnCtr) = oTrans.Master(lnCtr)
      End If
   Next
   
   txtSearch(0) = txtField(0)
   txtSearch(1) = txtField(2)
   txtField(1) = strLongDate(txtField(1))
   txtField(10) = strLongDate(oTrans.Master("dEffectve"))

   If oTrans.Master("cTranStat") = "4" Then
      Label2.Caption = "APPLIED"
   Else
      Label2.Caption = TransStat(CInt(oTrans.Master("cTranStat")))
   End If
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCtr As Integer

   lnRow = oTrans.ItemCount

   With MSFlexGrid1
      .Rows = 2
      .Rows = lnRow + 1
                If MSFlexGrid1.Rows > 18 Then
             .ColWidth(1) = 2565
         Else
           .ColWidth(1) = 2815
         End If
      For lnCtr = 0 To lnRow - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = IFNull(oTrans.Detail(lnCtr, "xFullName"), "")
         .TextMatrix(lnCtr + 1, 2) = IFNull(oTrans.Detail(lnCtr, "sShiftNm1"), "")
         .TextMatrix(lnCtr + 1, 3) = IFNull(oTrans.Detail(lnCtr, "sShiftNm2"), "")
         .TextMatrix(lnCtr + 1, 4) = IFNull(oTrans.Detail(lnCtr, "sShiftNm3"), "")
         .TextMatrix(lnCtr + 1, 5) = IFNull(oTrans.Detail(lnCtr, "sShiftNm4"), "")
         .TextMatrix(lnCtr + 1, 6) = IFNull(oTrans.Detail(lnCtr, "sShiftNm5"), "")
         .TextMatrix(lnCtr + 1, 7) = IFNull(oTrans.Detail(lnCtr, "sShiftNm6"), "")
         .TextMatrix(lnCtr + 1, 8) = IFNull(oTrans.Detail(lnCtr, "sShiftNm7"), "")
      Next
      
      pnLastSelc = .Row
      SetGridRowColor (.Row)
   End With
   
   pbDtlLoaded = True
End Sub

Private Sub MSFlexGrid1_GotFocus()
   pbDetailGotFocus = True
   If xrFrame2(1).Enabled = False Then Exit Sub
End Sub
Private Sub MSFlexGrid1_Click()
   Dim lnCtr As Integer

   If oTrans.ItemCount = 1 Then Exit Sub

   With oTrans
      pnActiveRow = MSFlexGrid1.Row

      For lnCtr = 1 To 7
         txtOthers(lnCtr - 1) = Format(IFNull(.Detail(pnActiveRow - 1, lnCtr - 1), ""), "HH:MM:SS AM/PM")
      Next
      
   Call detailFieldChange
   End With
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If pbCtrlPress Then
      If KeyCode = vbKeyControl Then pbCtrlPress = False
   End If
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   If Index = 6 Then Label2.Caption = TransStat(CInt(Value))
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   pbMoveDown = True
   pbMoveUpxx = True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF3, vbKeyReturn
         If oTrans.SearchTransaction(txtSearch(Index).Text, True) Then
            ClearFields
            InitGrid
            LoadMaster
            LoadDetail
         End If
    End Select
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("EB")
   End With

   pnIndex = Index
End Sub
Private Sub txtSearch_GotFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
End Sub
Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtSearch
      loTxt.Text = ""
   Next
   For Each loTxt In txtField
      loTxt.Text = ""
   Next
   For Each loTxt In txtOthers
      loTxt.Text = ""
   Next
End Sub
Private Sub txtOthers_GotFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("HT1")
   End With
   
   Select Case Index
      Case 0
         pbMoveUpxx = False
         pbMoveDown = True
      Case 6
         pbMoveUpxx = True
         pbMoveDown = False
      Case Else
         pbMoveDown = True
         pbMoveUpxx = True
   End Select
   
   pnIndex = Index
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Or (pnActiveRow < 1) Then Exit Sub
   
   Dim lsTime As String
   
   'dont allow to pass time parameter when Employee Name is null.
   If oTrans.Detail(pnActiveRow - 2, 0) = "" Then
      ClearOthers
      Exit Sub
   End If
   
   Select Case Index
      Case 1, 2
         oTrans.Detail(pnActiveRow - 2, Index) = getCTime(txtOthers(Index))
      Case 3, 4, 5, 6
         If InStr(txtOthers(Index), "p") = True Or InStr(txtOthers(Index), "pm") Then
            lsTime = txtOthers(Index) & "p"
            oTrans.Detail(pnActiveRow - 2, Index) = getCTime(lsTime)
         Else
            oTrans.Detail(pnActiveRow - 2, Index) = getCTime(txtOthers(Index))
         End If
         
         'Mac PH (08.23.12)
         'Moves down when pm out is validated.
         If Index = 4 Then
            pbCtrlPress = True
            Call Form_KeyDown(vbKeyDown, False)
            pbCtrlPress = False
         End If
   End Select
End Sub

Private Sub CopyT2AllEmp()
   Dim lnCtr As Integer
   pbCopy2All = True
   For lnCtr = 0 To oTrans.ItemCount - 1
      With oTrans
         .Detail(lnCtr, pnIndex) = txtOthers(pnIndex)
         
         MSFlexGrid1.TextMatrix(lnCtr + 2, pnIndex + 1) = .Detail(lnCtr, pnIndex)
      End With
   Next
   pbCopy2All = False
End Sub

Private Sub ClearOthers()
   Dim lnCtr As Integer
   
   For lnCtr = 0 To 6
      txtOthers(lnCtr) = ""
   Next
   
'   pnRow = oTrans.ItemCount - 1
'   MoveToLastRec
End Sub
Private Sub flexFocus()
   With MSFlexGrid1
      If .Row > 15 Then .TopRow = 1
      .Row = 2
      pnActiveRow = .Row
      pbDetailGotFocus = True
   End With
End Sub

Private Sub SetGridRowColor(ByVal lnRow As Integer)
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .FillStyle = flexFillRepeat
      
      .Row = IIf(pnLastSelc = .Rows, pnLastSelc - 1, pnLastSelc)
      .RowSel = pnLastSelc - 1
      .Col = 1
      .ColSel = .Cols - 1
      .CellBackColor = &HFFFFFF
      
      .Row = lnRow
      .RowSel = lnRow
      .Col = 1
      .ColSel = .Cols - 1
      .CellBackColor = &HFF8080
      
      pnLastSelc = .Row
      If Not .RowIsVisible(.Row + IIf(.Row = .Rows - 1, 0, 1)) Then .TopRow = .Row
   End With
End Sub

Private Sub MoveToLastRec()
   With MSFlexGrid1
      .Row = .Rows - 1
      pnActiveRow = .Row
      pnRow = pnActiveRow - 2
      
      If Not .RowIsVisible(.Row) Then
         .TopRow = .Row - 10
      End If
      
      detailFieldChange
   End With
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






