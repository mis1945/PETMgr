VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmRepPreview 
   BorderStyle     =   0  'None
   Caption         =   "Print Preview"
   ClientHeight    =   8760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12390
      Top             =   2895
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   12390
      TabIndex        =   0
      Top             =   2460
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
      Picture         =   "frmRepPreview.frx":0000
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   8100
      Index           =   0
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   14288
      BackColor       =   12632256
      BorderStyle     =   1
      Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
         Height          =   7980
         Left            =   45
         TabIndex        =   2
         Top             =   30
         Width           =   11955
         lastProp        =   600
         _cx             =   21087
         _cy             =   14076
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   -1  'True
         EnableStopButton=   0   'False
         EnablePrintButton=   0   'False
         EnableZoomControl=   -1  'True
         EnableCloseButton=   0   'False
         EnableProgressControl=   0   'False
         EnableSearchControl=   0   'False
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   0   'False
         EnableAnimationControl=   0   'False
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   0   'False
         DisplayTabs     =   0   'False
         DisplayBackgroundEdge=   0   'False
         SelectionFormula=   ""
         EnablePopupMenu =   0   'False
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
         LaunchHTTPHyperlinksInNewBrowser=   0   'False
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   12390
      TabIndex        =   1
      Top             =   525
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Print"
      AccessKey       =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmRepPreview.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   12390
      TabIndex        =   4
      Top             =   1815
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
      Picture         =   "frmRepPreview.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   12390
      TabIndex        =   3
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Set&up"
      AccessKey       =   "u"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmRepPreview.frx":166E
   End
End
Attribute VB_Name = "frmRepPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmRepPreview"
'Ok
Private p_oSkin As clsFormSkin
Private p_bRepPreview As Boolean

Public Event PrintReport()
Public Event BrowseReport()
Public Event PrintSetup()

Property Let AllowBrowse(ByVal Value As Boolean)
   p_bRepPreview = True
End Property

Private Sub Form_Activate()
   If p_bRepPreview = True Then
'      cmdButton(2).Caption = "&Browse"
'      cmdButton(3).Caption = "&Close"
      cmdButton(2).Visible = True
      cmdButton(3).Visible = True
      cmdButton(3).Top = 2460
'      cmdButton(3).Left = 10590
   Else
'      cmdButton(2).Caption = "&Close"
'      cmdButton(3).Visible = False
      cmdButton(2).Visible = False
      cmdButton(3).Top = 1815
'      cmdButton(3).Left = 10590
   End If
End Sub

Private Sub Form_Load()
   CenterChildForm oApp.mdiMain, Me

   Set p_oSkin = New clsFormSkin
   Set p_oSkin.AppDriver = oApp
   Set p_oSkin.Form = Me
   p_oSkin.DisableClose = True
   p_oSkin.ApplySkin xeFormTransMaintenance
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set p_oSkin = Nothing
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   
   Select Case Index
   Case 0
      RaiseEvent PrintReport
   Case 1
      RaiseEvent PrintSetup
   Case 2
      If cmdButton(Index).Caption = "&Browse" Then
         RaiseEvent BrowseReport
'      Else
'         Unload Me
      End If
   Case 3
      Unload Me
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
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
