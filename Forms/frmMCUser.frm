VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMCUser 
   BorderStyle     =   0  'None
   Caption         =   "Motorcycle User Maintenance"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4500
      Left            =   4935
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1065
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   7938
      _Version        =   393216
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   480
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   847
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   4815
         TabIndex        =   3
         Text            =   "Cuison, Michael Torres"
         Top             =   75
         Width           =   4260
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1365
         TabIndex        =   1
         Text            =   "May 16, 2014"
         Top             =   75
         Width           =   1530
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   10
         Left            =   3900
         TabIndex        =   2
         Top             =   135
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   0
         Top             =   135
         Width           =   900
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3135
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1050
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      BackColor       =   12632256
      Enabled         =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1380
         TabIndex        =   22
         Text            =   "May 16, 2014"
         Top             =   1590
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   555
         Index           =   2
         Left            =   1380
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmMCUser.frx":0000
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1380
         TabIndex        =   15
         Text            =   "Macarling, Federico Jr."
         Top             =   2670
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   5
         Text            =   "M001140001"
         Top             =   150
         Width           =   1530
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   7
         Text            =   "Cuison, Michael Torres"
         Top             =   630
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1380
         TabIndex        =   11
         Text            =   "May 16, 2014"
         Top             =   1920
         Width           =   1530
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1380
         TabIndex        =   13
         Text            =   "Macarling, Federico Jr."
         Top             =   2340
         Width           =   3255
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENGINE NO."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   105
         TabIndex        =   23
         Top             =   1650
         Width           =   1110
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   5
         Left            =   645
         TabIndex        =   8
         Top             =   975
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Co Owner #2"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   14
         Top             =   2730
         Width           =   945
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Account No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Name"
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   6
         Top             =   690
         Width           =   825
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Purc."
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   10
         Top             =   1980
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Co Owner #1"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   12
         Top             =   2400
         Width           =   945
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1470
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   1530
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1380
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   4200
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2434
      BackColor       =   12632256
      Enabled         =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   3105
         TabIndex        =   20
         Text            =   "May 16, 2014"
         Top             =   900
         Width           =   1530
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   1380
         TabIndex        =   18
         Text            =   "Cuison, Michael Torres"
         Top             =   555
         Width           =   3255
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Started using As Of"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   1365
         TabIndex        =   19
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOTORCYCLE USER INFO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   16
         Top             =   210
         Width           =   2340
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   195
         Index           =   9
         Left            =   510
         TabIndex        =   17
         Top             =   615
         Width           =   705
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9615
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   975
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
      Picture         =   "frmMCUser.frx":0028
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   9615
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1605
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
      Picture         =   "frmMCUser.frx":07A2
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   9615
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2235
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
      Picture         =   "frmMCUser.frx":0F1C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9615
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1605
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
      Picture         =   "frmMCUser.frx":1696
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   4
      Left            =   9615
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2235
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
      Picture         =   "frmMCUser.frx":1E10
   End
End
Attribute VB_Name = "frmMCUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmMCUser"

Private oSkin As clsFormSkin
Private bLoaded As Boolean

Private p_sSQLRecord As String
Private p_sSQLBrowse As String
Private p_sSQLMCUser As String

Private p_oRSRecord As Recordset
Private p_oRSUser As Recordset
Private pnUserCount As Integer

Private p_oClient As ggcClients.clsClient

Dim pnIndex As Integer
Dim pnEditMode As Integer
Dim psClientNm As String

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   Select Case Index
      Case 0
         If pnIndex = 0 Then
            If BrowseRecord(txtSearch(pnIndex), True) Then LoadRecord
         Else
            If BrowseRecord(txtSearch(pnIndex)) Then LoadRecord
         End If
      Case 1 'save
      Case 2 'update
         If pnEditMode <> xeModeReady Then Exit Sub
         
         Call addUser
         xrFrame3.Enabled = True
         
         txtField(7).SetFocus
         pnEditMode = xeModeUpdate
      Case 3 'del row
      Case 4 'cancel
         If pnEditMode <> xeModeUpdate Then Exit Sub
         
         xrFrame3.Enabled = False
      Case 5 'close
         Unload Me
   End Select

endProc:
   Exit Sub
endWithFocus:
   GoTo endProc
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
         If GetFocus = txtField(8).hwnd Then Exit Sub
      
         SetNextFocus
      Case vbKeyUp
         If GetFocus = txtField(7).hwnd Then Exit Sub
      
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
   Set p_oRSRecord = New Recordset
   Set p_oRSUser = New Recordset
   
   Set p_oClient = New ggcClients.clsClient
   Set p_oClient.AppDriver = oApp
   p_oClient.BranchCode = oApp.BranchCode
   If p_oClient.InitClient() = False Then GoTo endProc
   p_oClient.ShowEntry = True
   
   initGrid
   initSQL
   ClearFields
   
   pnEditMode = xeModeUnknown
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set p_oRSRecord = Nothing
   Set p_oRSUser = Nothing
   
   psClientNm = ""
   pnUserCount = ""
   bLoaded = False
End Sub

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   If Index = 7 Then
      pnIndex = Index
      
      With MSFlexGrid1
         .Row = .Rows - 1
         
         .Col = 0
         .ColSel = .Cols - 1
      End With
   ElseIf Index = 8 Then
      pnIndex = Index
      txtField(Index) = strShortDate(txtField(Index))
   End If
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyF3
      If pnIndex = 7 Then
         Call getClient(txtField(pnIndex), False)
      ElseIf Index = 8 Then
         If Not IsDate(txtField(pnIndex)) Then Exit Sub
         
         p_oRSUser("dStartUse") = CDate(txtField(pnIndex))
         
         If psClientNm = "" Then
            txtField(7).SetFocus
         Else
            Call addUserOnGrid
         End If
      End If
   End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn, vbKeyF3
         If pnIndex = 0 Then
            If BrowseRecord(txtSearch(pnIndex), True) Then LoadRecord
         Else
            If BrowseRecord(txtSearch(pnIndex)) Then LoadRecord
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
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   pnIndex = Index
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

Private Function SaveRecord() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "BrowseRecord"
   'On Error GoTo errProc
   
   With p_oRSUser
      If pnUserCount = p_oRSUser.RecordCount Then
         MsgBox "Detail Not Modified. Unable to Update Record.", vbCritical, "Warning"
         GoTo endProc
      End If
      
      
      .Move pnUserCount, adBookmarkFirst
      oApp.BeginTrans
      Do Until .EOF
         lsSQL = ADO2SQL(p_oRSUser, "MC_AR_User", , Encrypt(oApp.UserName), oApp.ServerDate, "xClientNm")
         Debug.Print lsSQL
      
         oApp.Execute lsSQL, "MC_AR_User"
         .MoveNext
      Loop
      oApp.CommitTrans
   End With
   
   SaveRecord = OpenRecord(p_oRSRecord("sAcctNmbr"))
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function


Private Function BrowseRecord(ByVal fsValue As String, _
                              Optional ByVal fbCode As Boolean = False) As Boolean
   Dim lsOldProc As String
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lasDetail() As String

   lsOldProc = "BrowseRecord"
   'On Error GoTo errProc
   
   If fbCode Then
      If pnEditMode = xeModeReady Then If p_oRSRecord("sAcctNmbr") = fsValue Then GoTo endProc
   
      lsSQL = AddCondition(p_sSQLBrowse, "a.sAcctNmbr = " & strParm(fsValue))
   Else
      If pnEditMode = xeModeReady Then If p_oRSRecord("xCustName") = fsValue Then GoTo endProc
   
      lsSQL = AddCondition(p_sSQLBrowse, _
                  "CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) LIKE " & strParm(fsValue & "%"))
   End If
   
   Set loRS = New Recordset
   
   Debug.Print lsSQL
   loRS.Open lsSQL, oApp.Connection, , , adCmdText
   
   With loRS
      Set .ActiveConnection = Nothing
      
      If .RecordCount = 0 Then
         MsgBox "No Record Found.", vbInformation, "Notice"
         GoTo endProc
      ElseIf .RecordCount = 1 Then
         BrowseRecord = OpenRecord(loRS("sAcctNmbr"))
      Else
         lsSQL = KwikBrowse(oApp, loRS, _
                           "sAcctNmbr»xCustName»dPurchase", _
                           "Account No»Full Name»Date Purchase", _
                           "@»@»Mmm dd, yyyy", _
                           "a.sAcctNmbr»CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)»a.dPurchase")
         If lsSQL = "" Then GoTo endProc
         
         lasDetail = Split(lsSQL, "»")
         
         BrowseRecord = OpenRecord(lasDetail(0))
      End If
   End With
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & fsValue & " )", True
End Function

Private Function OpenRecord(ByVal fsAcctNmbr As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "OpenRecord"
   'On Error GoTo errProc
   
   lsSQL = AddCondition(p_sSQLRecord, "a.sAcctNmbr = " & strParm(fsAcctNmbr))
   Debug.Print lsSQL
   
   If p_oRSRecord.State = adStateOpen Then p_oRSRecord.Close
   
   p_oRSRecord.Open lsSQL, oApp.Connection, , , adCmdText
   
   With p_oRSRecord
      Set .ActiveConnection = Nothing
      
      If .RecordCount <> 1 Then GoTo endProc
   End With

   OpenRecord = getUsers(fsAcctNmbr)
   pnEditMode = xeModeReady
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & fsAcctNmbr & " )", True
End Function

Private Function getUsers(ByVal fsAcctNmbr As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "OpenRecord"
   'On Error GoTo errProc

   lsSQL = AddCondition(p_sSQLMCUser, "a.sAcctNmbr = " & strParm(fsAcctNmbr))
   Debug.Print lsSQL
   
   If p_oRSUser.State = adStateOpen Then p_oRSUser.Close
   
   p_oRSUser.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set p_oRSUser.ActiveConnection = Nothing
   
   pnUserCount = p_oRSUser.RecordCount
   
   getUsers = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & fsAcctNmbr & " )", True
End Function

Private Function addUser() As Boolean
   With p_oRSUser
      p_oClient.InitClient
   
      .AddNew
      .Fields("sAcctNmbr") = p_oRSRecord("sAcctNmbr")
      .Fields("sClientID") = ""
      .Fields("dStartUse") = oApp.ServerDate
      
      psClientNm = ""
      txtField(7) = ""
      txtField(8) = strLongDate(.Fields("dStartUse"))
      
      If xrFrame3.Enabled Then txtField(7).SetFocus
   End With
End Function

Private Sub addUserOnGrid()
   With MSFlexGrid1
      .TextMatrix(.Rows - 1, 0) = .Row
      .TextMatrix(.Rows - 1, 1) = psClientNm
      .TextMatrix(.Rows - 1, 2) = strLongDate(p_oRSUser("dStartUse"))
      
      .Rows = .Rows + 1
   End With
   
   Call addUser
End Sub

Private Function getClient(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsOldProc As String

   lsOldProc = "getClient"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   On Error GoTo errProc
   
   getClient = False
   
   If lsValue <> "" Then
      If Trim(LCase(lsValue)) = Trim(LCase(psClientNm)) Then
         getClient = True
         Exit Function
      End If
      
      If p_oClient.SearchClient(lsValue, lbSearch) = False Then GoTo endProc
   Else
      GoTo endWithClear
   End If
   
   p_oRSUser("sClientID") = p_oClient.ClientID
   psClientNm = p_oClient.FullName
   
   getClient = True
   
endProc:
   txtField(7) = psClientNm

   Exit Function
endWithClear:
   p_oRSUser("sClientID") = ""
   psClientNm = ""
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Sub LoadRecord()
   Dim loTxt As TextBox
   Dim lnIndex As Integer
   
   initGrid
   ClearFields
   
   For Each loTxt In txtField
      lnIndex = loTxt.Index
      
      Select Case lnIndex
         Case 3
            txtField(lnIndex) = strLongDate(p_oRSRecord(lnIndex))
         Case 7, 8
         Case Else
            txtField(lnIndex) = IFNull(p_oRSRecord(lnIndex))
      End Select
   Next
   
   txtSearch(0) = txtField(0)
   txtSearch(1) = txtField(1)
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
   txtSearch(0) = ""
   txtSearch(1) = ""
   
   psClientNm = ""
   pnUserCount = ""
   
   Call initGrid
   If bLoaded Then txtSearch(0).SetFocus
End Sub

Private Sub initGrid()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Rows = 2
      .Cols = 3
      
      .Row = 0
      
      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = "Name"
      .TextMatrix(0, 2) = "Date"
      
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .ColWidth(0) = 400
      .ColWidth(1) = 2500
      .ColWidth(2) = 1400
      
      .Row = 1
      
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub initSQL()
   
   p_sSQLRecord = "SELECT" & _
                     "  a.sAcctNmbr" & _
                     ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xCustName" & _
                     ", CONCAT(b.sHouseNox, ' ', b.sAddressx, ' ', e.sTownName, ', ', f.sProvName) xAddressx" & _
                     ", a.dPurchase" & _
                     ", CONCAT(c.sLastName, ', ', c.sFrstName, ' ', c.sMiddName) xCoCltID1" & _
                     ", CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName) xCoCltID2" & _
                     ", g.sEngineNo" & _
                  " FROM MC_AR_Master a" & _
                           " LEFT JOIN Client_Master b" & _
                                 " ON a.sClientID = b.sClientID" & _
                           " LEFT JOIN Client_Master c" & _
                                 " ON a.sCoCltID1 = c.sClientID" & _
                           " LEFT JOIN Client_Master d" & _
                                 " ON a.sCoCltID2 = d.sClientID" & _
                        ", TownCity e" & _
                           " LEFT JOIN Province f" & _
                              " ON e.sProvIDxx = f.sProvIDxx" & _
                        ", MC_Serial g" & _
                  " WHERE a.sSerialID = g.sSerialID" & _
                     " AND b.sTownIDxx = e.sTownIDxx"
   
   p_sSQLMCUser = "SELECT" & _
                     "  a.sAcctNmbr" & _
                     ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xClientNm" & _
                     ", a.dStartUse" & _
                     ", a.sClientID" & _
                  " FROM MC_AR_User a" & _
                           " LEFT JOIN Client_Master b" & _
                               " ON a.sClientID = b.sClientID"
   
   p_sSQLBrowse = "SELECT" & _
                     "  a.sAcctNmbr" & _
                     ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xCustName" & _
                     ", a.dPurchase" & _
                  " FROM MC_AR_Master a" & _
                        " LEFT JOIN Client_Master b" & _
                           " ON a.sClientID = b.sClientID" & _
                        ", TownCity c" & _
                           " LEFT JOIN Province d" & _
                              " ON c.sProvIDxx = d.sProvIDxx" & _
                  " WHERE b.sTownIDxx = c.sTownIDxx"
End Sub
