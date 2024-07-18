VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSelfieLogCriteria 
   BorderStyle     =   0  'None
   Caption         =   "Transaction Summary"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Tag             =   "et0;eb0;et0;bc2"
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1890
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   3334
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1230
         Width           =   2640
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   345
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   675
         Width           =   1890
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   2
         Left            =   285
         TabIndex        =   4
         Top             =   1275
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   2
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   0
         Top             =   375
         Width           =   810
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   1170
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
      Picture         =   "frmSelfieLogCriteria.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   4440
      TabIndex        =   6
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSelfieLogCriteria.frx":077A
   End
End
Attribute VB_Name = "frmSelfieLogCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmRouteDateCriteria"

Private p_oAppDrivr As clsAppDriver
Private oSkin As clsFormSkin

Dim pnOkFocus As Long
Dim pnCancel As Long
Dim pdDateFrom  As Date
Dim pdDateThru As Date
Dim pbCancel As Boolean
Dim psEmployID As String
Dim oForm As Object

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancel
End Property

Property Set FormDateCriteria(ByRef Form As Object)
   Set oForm = Form
End Property

Property Get DateFrom()
   DateFrom = pdDateFrom
End Property

Property Get DateThru()
   DateThru = pdDateThru
End Property

Property Get Employee()
   Employee = psEmployID
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc
   
   pbCancel = True
   Select Case Index
   Case 0
      pdDateFrom = CDate(txtField(0).Text)
      pdDateThru = CDate(txtField(1).Text)
      pbCancel = False
      Me.Hide
   Case 1
      Unload Me
   End Select
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   On Error GoTo errProc
   
   CenterChildForm mdiMain, Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransDetail
   
   txtField(0).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   txtField(1).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   txtField(2).Text = ""
   txtField(2).Tag = ""
   
   pnOkFocus = cmdButton(0).hWnd
   pnCancel = cmdButton(1).hWnd
   psEmployID = ""

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   If Index < 2 Then txtField(Index).Text = Format(txtField(Index).Text, "MM/DD/YY")
   If txtField(Index).Text <> "" Then
      txtField(Index).SelStart = 0
      txtField(Index).SelLength = Len(txtField(Index).Text)
   End If
End Sub


Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField
         Select Case Index
         
         Case 2
            Call SearchEmployee(.Item(Index))
         End Select
      End With
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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   On Error GoTo errProc
   
   Select Case Index
   Case 0, 1
      If Not IsDate(txtField(Index).Text) Then
         txtField(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Else
         txtField(Index).Text = Format(txtField(Index).Text, "MMMM DD, YYYY")
      End If
   Case 2
      If GetFocus = pnCancel Or GetFocus = pnOkFocus Then Exit Sub
      If txtField(Index).Text <> "" Then
         If txtField(Index).Text <> txtField(Index).Tag Then SearchEmployee txtField(Index).Text
      Else
         psEmployID = ""
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Sub SearchEmployee(ByVal Value As String)
   Dim lsBrowse As String, lsSelected() As String, lsSQL As String
   Dim lrs As ADODB.Recordset
   Dim lasMaster() As String
   If Value = txtField(2).Tag Then Exit Sub

   lasMaster = GetSplitedName(Value)
   
       lsSQL = " SELECT " & _
                 "  a.sEmployID" & _
                 ", CONCAT(b.sLastName, ', ', b.sFrstName, IF(TRIM(IFNULL(b.sSuffixNm, '')) = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName) sEmployNm" & _
                 ", c.sBranchNm" & _
           " FROM Employee_Master001 a" & _
               " LEFT JOIN Client_Master b ON a.sEmployID = b.sClientID" & _
               " LEFT JOIN Branch c ON a.sBranchCD = c.sBranchCD" & _
           " WHERE b.sLastName LIKE " & strParm(lasMaster(0) & "%") & _
               IIf(UBound(lasMaster) > 0, " AND b.sFrstName LIKE " & strParm(lasMaster(1) & "%"), "")
               

       Set lrs = New ADODB.Recordset
   lrs.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If lrs.EOF Then
      psEmployID = ""
      txtField(2) = ""
   ElseIf lrs.RecordCount = 1 Then
      psEmployID = lrs("sEmployID")
      txtField(2) = lrs("sEmployNm")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, lrs _
                              , "sEmployID»sEmployNm»sBranchNm" _
                              , "ID»Employee»Branch" _
                              , "@»@»@", False)

      If lsBrowse <> "" Then
         lsSelected = Split(lsBrowse, "»")
         psEmployID = lsSelected(0)
         txtField(2) = lsSelected(1)
      Else
         psEmployID = ""
         txtField(2) = ""
      End If
   End If
   
   txtField(2).Tag = txtField(2).Text
   
   lrs.Close

   Set lrs = Nothing
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

