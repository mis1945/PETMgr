VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrcontrol.ocx"
Begin VB.Form frmImage 
   BorderStyle     =   0  'None
   Caption         =   "Employee Image"
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4800
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   4400
      _ExtentX        =   7752
      _ExtentY        =   8467
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4500
         Left            =   120
         Top             =   120
         Width           =   4100
      End
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmImage"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean


Dim pnCtr As Integer
Dim pnIndex As Integer
Dim p_oSource As String


Function ImageSource(ByVal Value As String)
   p_oSource = Value
   Call EmployeeImage
End Function

Private Sub Form_Load()
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   On Error GoTo errProc
   
   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormMessageBox
   
   Call EmployeeImage
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub
Private Sub EmployeeImage()

                      If Dir(p_oSource) <> "" Then
                Image1.Picture = LoadPicture(p_oSource)
            Else
                Image1.Picture = Nothing

                    End If

        
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oDriver = Nothing
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


