VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmBatchShiftChange 
   BorderStyle     =   0  'None
   Caption         =   "Batch Shift Change"
   ClientHeight    =   9855
   ClientLeft      =   0
   ClientTop       =   4320
   ClientWidth     =   20400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9736.444
   ScaleMode       =   0  'User
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   526
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
      Picture         =   "frmBatchShiftChange.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1770
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
      Picture         =   "frmBatchShiftChange.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   525
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
      Picture         =   "frmBatchShiftChange.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2385
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
      Picture         =   "frmBatchShiftChange.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1155
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
      Picture         =   "frmBatchShiftChange.frx":1DE8
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   9221
      Index           =   3
      Left            =   1580
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   18720
      _ExtentX        =   33020
      _ExtentY        =   16272
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5695
         Left            =   0
         TabIndex        =   30
         Top             =   3500
         Width           =   18705
         _ExtentX        =   32994
         _ExtentY        =   10054
         _Version        =   393216
         AllowBigSelection=   0   'False
         ScrollBars      =   2
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   3495
         Index           =   0
         Left            =   0
         Tag             =   "wt0;fb0"
         Top             =   0
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   6165
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.CheckBox Check1 
            Caption         =   "Load Employees"
            Height          =   195
            Left            =   1680
            TabIndex        =   6
            Tag             =   "wt0;fb0"
            Top             =   1470
            Width           =   1470
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1725
            Width           =   5295
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
            TabIndex        =   1
            Top             =   165
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
            Index           =   1
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   3
            Top             =   750
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
            Height          =   1155
            Index           =   4
            Left            =   1680
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   2160
            Width           =   5295
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
            TabIndex        =   4
            Top             =   1815
            Width           =   615
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   420
            Index           =   0
            Left            =   1800
            Tag             =   "et0;ht2"
            Top             =   285
            Width           =   2415
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
            TabIndex        =   0
            Top             =   255
            Width           =   1485
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
            TabIndex        =   2
            Top             =   840
            Width           =   405
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
            TabIndex        =   7
            Top             =   2250
            Width           =   780
         End
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   3495
         Index           =   1
         Left            =   9480
         Tag             =   "wt0;fb0"
         Top             =   0
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   6165
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtOthers 
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
            Height          =   375
            Index           =   5
            Left            =   1200
            TabIndex        =   20
            Text            =   "sShftDay5"
            Top             =   2205
            Width           =   5295
         End
         Begin VB.TextBox txtOthers 
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
            Height          =   375
            Index           =   3
            Left            =   1200
            TabIndex        =   16
            Text            =   "sShftDay3"
            Top             =   1425
            Width           =   5295
         End
         Begin VB.TextBox txtOthers 
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
            Height          =   375
            Index           =   7
            Left            =   1200
            TabIndex        =   24
            Text            =   "sShftDay7"
            Top             =   2985
            Width           =   5295
         End
         Begin VB.TextBox txtOthers 
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
            Height          =   375
            Index           =   6
            Left            =   1200
            TabIndex        =   22
            Text            =   "sShftDay6"
            Top             =   2595
            Width           =   5295
         End
         Begin VB.TextBox txtOthers 
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
            Height          =   375
            Index           =   4
            Left            =   1200
            TabIndex        =   18
            Text            =   "sShftDay4"
            Top             =   1815
            Width           =   5295
         End
         Begin VB.TextBox txtOthers 
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
            ForeColor       =   &H80000009&
            Height          =   420
            Index           =   0
            Left            =   1200
            TabIndex        =   10
            Text            =   "Company Name"
            Top             =   210
            Width           =   5295
         End
         Begin VB.TextBox txtOthers 
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
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   12
            Text            =   "sShftDay1"
            Top             =   645
            Width           =   5295
         End
         Begin VB.TextBox txtOthers 
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
            Height          =   375
            Index           =   2
            Left            =   1200
            TabIndex        =   14
            Text            =   "sShftDay2"
            Top             =   1035
            Width           =   5295
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   300
            Width           =   555
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
            TabIndex        =   11
            Top             =   712
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
            TabIndex        =   13
            Top             =   1102
            Width           =   690
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tuesday"
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
            TabIndex        =   15
            Top             =   1492
            Width           =   735
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
            TabIndex        =   17
            Top             =   1882
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
            TabIndex        =   19
            Top             =   2272
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
            TabIndex        =   21
            Top             =   2662
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
            TabIndex        =   23
            Top             =   3052
            Width           =   900
         End
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1155
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
      Picture         =   "frmBatchShiftChange.frx":2562
   End
End
Attribute VB_Name = "frmBatchShiftChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeModuleName = "frmBatchShiftChange"
Private Const pxBRANCHCODES = "M001»A001»N001»H001"
Private WithEvents oTrans As clsBatchShiftSched
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim psSelected() As String
Dim pnIndex As Integer
Dim pnLastSelc As Integer
Dim pnActiveRow As Integer
Dim pnRow As Integer

Dim pbFormLoad As Boolean
Dim pbCtrlPress As Boolean
Private bLoaded As Boolean
Private pbByBranch As Boolean
Dim pbCopy2All As Boolean
Dim pbDetailGotFocus As Boolean

'controlling get/next focus on detail
Dim pbMoveUpxx As Boolean
Dim pbMoveDown As Boolean
Property Let ByBranch(vValue As Boolean)
   pbByBranch = vValue
End Property

Private Sub Check1_Click()
   oTrans.AutoLoadEmployees = Check1.Value = 1
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lnRep As Integer
   Debug.Print Index
   Select Case Index
      Case 0 'cancel
         lnRep = MsgBox("Are you certain in aborting the current transaction?", vbQuestion + vbYesNo, "Confirm")
         
         If lnRep = vbYes Then
            If oTrans.InitTransaction Then
               Call ClearFields
               Call ClearOthers
               Call InitGrid
               Call InitForm(0)
            End If
         End If
      Case 1 'search
         If Index = 1 Then
            If oTrans.SearchMaster(2, txtField(2).Text) Then
               Call InitGrid
               Call LoadMaster
               Call LoadDetail
            End If
         End If
      Case 2 'save
         If txtField(2).Text <> "" Then
         
            If oTrans.SaveTransaction() Then
               MsgBox "Transaction Saved Successfully!!!", vbInformation, "Confirm"
               oTrans.InitTransaction
               
               Call InitForm(0)
               Call InitGrid
               Call ClearFields
               Call ClearOthers
            Else
               MsgBox "Unable to Saved Transaction!!!", vbCritical, "Warning"
            End If
         Else
               MsgBox "Invalid Branch Code Detected!!!", vbCritical, "Warning"
         End If
      Case 3 'new
         Call ClearFields
         Call ClearOthers
         Call InitGrid
         Call InitForm(1)
          
         oTrans.NewTransaction
         
         txtField(2).SetFocus
         LoadMaster
      Case 4 'close
         Unload Me
      Case 5 'delete detail
         If oTrans.DeleteDetail(pnRow) Then LoadDetail
   End Select
   
endProc:
   Exit Sub
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   'On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then

   Else
      'Set the default value to 1
      oTrans.HasParent = True
      oTrans.Master(7) = 1
   End If

   bLoaded = True

   If InStr(1, pxBRANCHCODES, oApp.BranchCode) = 0 Then
      txtField(1).Visible = True
      txtField(1).TabStop = False
      txtField(1).Locked = True
   End If

   If Not pbFormLoad Then pbFormLoad = True
   pnActiveRow = 2
   pnRow = 2
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
         If GetFocus = txtOthers(7).hWnd Then Exit Sub
         
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   Set oTrans = New clsBatchShiftSched
   Set oTrans.AppDriver = oApp
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
    
   Call InitGrid
   Call InitForm(0)
   Call ClearFields
   Call ClearOthers
   
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
   bLoaded = False

End Sub


Private Sub detailFieldChange()
   Dim lnCtr As Integer
   Dim lnRow As Integer
   Dim psOthers As Boolean
   
   lnRow = pnActiveRow - 1
   txtOthers(0).Text = IFNull(oTrans.Detail(lnRow, "xFullName"))
   
   For lnCtr = 1 To 7
      txtOthers(lnCtr).Text = IFNull(oTrans.Detail(lnRow, "sShiftNm" & lnCtr))
   Next
   
   SetGridRowColor (lnRow + 1)
   
   'txtOthers(0).SetFocus
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

Private Sub LoadMaster()
   txtField(0) = oTrans.Master("sTransNox")
   txtField(1) = strLongDate(oTrans.Master("dTransact"))
   txtField(4) = strLongDate(oTrans.Master("sRemarksx"))
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCtr As Integer
   
   lnRow = oTrans.ItemCount
   
   With MSFlexGrid1
      .Rows = 2
      .Rows = lnRow + 1
      
      If MSFlexGrid1.Rows > 23 Then
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
           
      'pnActiveRow = 1
      detailFieldChange
   End With
   
   Check1.Enabled = False
   txtField(2).Enabled = False
   txtOthers(0).Enabled = Check1.Value = 0
End Sub
Private Sub MSFlexGrid1_Click()
   Dim lnCtr As Integer
 
   With oTrans
      pnActiveRow = MSFlexGrid1.Row
      pnRow = pnActiveRow - 1
      
      For lnCtr = 1 To 7
         txtOthers(lnCtr) = Format(IFNull(.Detail(pnRow, lnCtr), ""), "")
      Next
      
      Call detailFieldChange
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Row As Integer, ByVal Index As Integer, ByVal Value As Variant)
   Dim lnCtr As Integer
   If txtField(2).Text <> "" Then
      With txtOthers(Index)
         .Text = IFNull(Value, "")
      End With

      If pnActiveRow < 1 Or pbCopy2All Then Exit Sub
      
      LoadDetail
   End If
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Debug.Print ("Index = " & Index & " Value = " & Value)
   With txtField(Index)
      Select Case Index
         Case 1
            .Text = strLongDate(IFNull(Value, ""))
         Case 2
            If oTrans.Master("sBranchCd") <> "" Then
               oTrans.loadEmployee
               .Text = IFNull(Value, "")
            End If
         Case Else
            .Text = IFNull(Value, "")
      End Select
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   pbMoveUpxx = True
   pbMoveDown = True
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF3
      If Index = 2 Then
         If oTrans.SearchMaster(Index, txtField(Index).Text) Then
            Call InitGrid
            Call LoadMaster
            Call LoadDetail
         End If
      End If
   Case vbKeyReturn
      If Index = 2 Then
         If oTrans.SearchMaster(Index, txtField(Index).Text) Then
            Call InitGrid
            Call LoadMaster
            Call LoadDetail
         End If
      End If
   End Select
End Sub

Private Sub MSFlexGrid1_GotFocus()
   pbDetailGotFocus = True
   If xrFrame2(1).Enabled = False Then Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If pbCtrlPress Then
      If KeyCode = vbKeyControl Then pbCtrlPress = False
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      If Index <> 2 Then oTrans.Master(Index) = .Text
   End With
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("HT1")
   End With
   
   Select Case Index
      Case 1
         pbMoveUpxx = False
         pbMoveDown = True
      Case 7
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
   If Check1.Value = 1 Then If Index = 0 Or (pnActiveRow < 2) Then Exit Sub

   With txtOthers(Index)
'      Select Case Index
'         Case 1 To 7
'            oTrans.Detail(pnActiveRow - 1, Index) = .Text
'      End Select
   End With
End Sub



Private Sub InitForm(ByVal fnEdit As Integer)
   Dim lnCtr As Integer
   
   xrFrame2(1).Enabled = Not (fnEdit = 0)
   xrFrame2(3).Enabled = Not (fnEdit = 0)
   cmdButton(3).Visible = (fnEdit = 0)
   cmdButton(4).Visible = (fnEdit = 0)

   cmdButton(0).Visible = Not (fnEdit = 0)
   cmdButton(1).Visible = Not (fnEdit = 0)
   cmdButton(2).Visible = Not (fnEdit = 0)
   cmdButton(5).Visible = Not (fnEdit = 0)
   
   If fnEdit = 1 Then
      pnActiveRow = 1
      pnRow = 0
   End If
End Sub

Private Sub ClearFields()
   Dim lnCtr As Integer
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt.Text = ""
      loTxt.BackColor = oApp.getColor("EB")
   Next
   
   Check1.Value = 1
   oTrans.AutoLoadEmployees = True
   
   Check1.Enabled = True
   txtField(2).Enabled = True
   
   pnActiveRow = 1
   pnRow = 0
End Sub

Private Sub ClearOthers()
   Dim lnCtr As Integer
   Dim loTxt As TextBox
   
   For lnCtr = 0 To 7
      txtOthers(lnCtr) = ""
   Next
   
   For Each loTxt In txtOthers
      loTxt.BackColor = oApp.getColor("EB")
   Next
   
   pnRow = oTrans.ItemCount - 1
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

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim txtIndex As Integer
   pnRow = MSFlexGrid1.Row
   
   Select Case KeyCode
   Case vbKeyF3, vbKeyReturn
      Select Case Index
      Case 0
         oTrans.Detail(pnActiveRow - 1, Index) = txtOthers(Index)
      Case 7
         If oTrans.SearchDetail(pnRow - 1, Index, txtOthers(Index).Text) Then
            With MSFlexGrid1
               If .Row < .Rows - 1 Then .Row = .Row + 1
               MSFlexGrid1_Click
            End With
         End If
      Case Else
         oTrans.SearchDetail pnRow - 1, Index, txtOthers(Index).Text
      End Select
   End Select
End Sub
