VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLeaveResetPosting 
   BorderStyle     =   0  'None
   Caption         =   "Leave Reset Posting"
   ClientHeight    =   9525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmLeaveResetPosting.frx":0000
   ScaleHeight     =   9525
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   1590
      TabIndex        =   22
      Top             =   5415
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   7011
      _Version        =   393216
      BorderStyle     =   0
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
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1665
      Index           =   0
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   1185
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   2937
      BackColor       =   12632256
      Enabled         =   0   'False
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
         Height          =   390
         Index           =   3
         Left            =   4695
         MaxLength       =   50
         TabIndex        =   7
         Top             =   720
         Width           =   1965
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
         Height          =   390
         Index           =   2
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "Jan 01, 20102"
         Top             =   720
         Width           =   1965
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
         Height          =   390
         Index           =   4
         Left            =   1275
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1140
         Width           =   5385
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
         Height          =   390
         Index           =   1
         Left            =   4695
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Test"
         Top             =   105
         Width           =   1965
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
         Height          =   390
         Index           =   0
         Left            =   1275
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "M00111-000021"
         Top             =   105
         Width           =   1965
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Thru"
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
         Left            =   3630
         TabIndex        =   6
         Top             =   795
         Width           =   930
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
         Left            =   330
         TabIndex        =   8
         Top             =   1215
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leave From"
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
         TabIndex        =   4
         Top             =   795
         Width           =   1005
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
         Left            =   4155
         TabIndex        =   2
         Top             =   180
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Left            =   150
         TabIndex        =   0
         Top             =   180
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   390
         Left            =   1380
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   1965
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1470
      Index           =   1
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   2865
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   2593
      BackColor       =   12632256
      Enabled         =   0   'False
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
         Height          =   390
         Index           =   1
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Supervisor"
         Top             =   525
         Width           =   2505
      End
      Begin VB.TextBox txtOthers 
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
         Height          =   390
         Index           =   7
         Left            =   5565
         MaxLength       =   50
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   525
         Width           =   1095
      End
      Begin VB.TextBox txtOthers 
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
         Height          =   390
         Index           =   2
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   945
         Width           =   1095
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
         Height          =   390
         Index           =   0
         Left            =   1290
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   90
         Width           =   5385
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
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
         Left            =   705
         TabIndex        =   12
         Top             =   600
         Width           =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dflt Lv Cr"
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
         Left            =   4650
         TabIndex        =   14
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rem Leave"
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
         Left            =   180
         TabIndex        =   16
         Top             =   1020
         Width           =   960
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
         Left            =   270
         TabIndex        =   10
         Top             =   180
         Width           =   870
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1020
      Index           =   2
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   4350
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   1799
      BackColor       =   12632256
      Enabled         =   0   'False
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
         Height          =   390
         Index           =   4
         Left            =   1275
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   510
         Width           =   5385
      End
      Begin VB.TextBox txtOthers 
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
         Height          =   390
         Index           =   3
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   19
         Text            =   "00"
         Top             =   105
         Width           =   1140
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
         Index           =   11
         Left            =   270
         TabIndex        =   20
         Top             =   585
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asgd Lv Cr"
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
         Left            =   270
         TabIndex        =   18
         Top             =   180
         Width           =   945
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   23
      Top             =   7320
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
      Picture         =   "frmLeaveResetPosting.frx":1643F2
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   24
      Top             =   5430
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
      Picture         =   "frmLeaveResetPosting.frx":164B6C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   25
      Top             =   6060
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
      Picture         =   "frmLeaveResetPosting.frx":1652E6
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   26
      Top             =   6690
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&DisApprv"
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
      Picture         =   "frmLeaveResetPosting.frx":165A60
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   615
      Index           =   3
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   1085
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   1
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
         Height          =   390
         Index           =   9
         Left            =   1275
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "M00111-000021"
         Top             =   105
         Width           =   1965
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Index           =   14
         Left            =   150
         TabIndex        =   28
         Top             =   180
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmLeaveResetPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Employee Leave Reset Form
'
' Copyright 2011 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Mac PH (08.17.12)
'     Created this form.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME = "frmLeaveResetPosting"

Private WithEvents oTrans As clsBatchLeave
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim psSelected() As String
Dim pnIndex As Integer
Dim pnActiveRow As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc
   With oTrans
      Select Case Index
         Case 0 'close
            Unload Me
         Case 1 'browse
         Case 2 'approve
            If .PostTransaction Then
               .InitTransaction
            End If
         Case 3 'cancel
            If .CancelTransaction Then
               .InitTransaction
            End If
      End Select
   End With
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   On Error GoTo errProc
   
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
   On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   Set oTrans = New clsBatchLeave
   Set oTrans.AppDriver = oApp
   With oTrans
      .Branch = oApp.BranchCode
      .InitTransaction
   End With
   
   ClearFields
   InitGrid
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
   bLoaded = False
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Cols = 5
      .Rows = 3
      
      .FixedRows = 2
      .Clear
      
      .Row = 0
      .TextMatrix(0, 0) = " "
      .TextMatrix(0, 1) = " "
      .TextMatrix(0, 2) = "Leave"
      .TextMatrix(0, 3) = "Leave"
      .TextMatrix(0, 4) = "Leave"
      
      'Row 1
      .TextMatrix(1, 0) = " "
      .TextMatrix(1, 1) = "Employee"
      .TextMatrix(1, 2) = "Dflt."
      .TextMatrix(1, 3) = "Rem."
      .TextMatrix(1, 4) = "Asgd."
      
      .MergeCells = flexMergeFree 'disables colsel procedure
      .MergeRow(0) = True

      .Row = 0
      'Column Width
      .ColWidth(0) = 500
      .ColWidth(1) = 3700
      .ColWidth(2) = 850
      .ColWidth(3) = 850
      .ColWidth(4) = 850
      
      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1
      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      
      .Row = 2
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignRightCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
   End With
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0).Text = .Master("sTransNox")
      txtField(1).Text = strLongDate(.Master("dTransact"))
      If bLoaded Then txtField(2).SetFocus
   End With
End Sub
Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCtr As Integer
   
   lnRow = oTrans.ItemCount
   
   With MSFlexGrid1
      .Rows = lnRow + 2
      If .Rows > 16 Then
         .ColWidth(1) = 3450
      Else
         .ColWidth(1) = 3700
      End If
      For lnCtr = 0 To lnRow - 1
         .TextMatrix(lnCtr + 2, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 2, 1) = oTrans.Detail(lnCtr, "sEmployNm")
         .TextMatrix(lnCtr + 2, 2) = oTrans.Detail(lnCtr, "nLeaveCrd")
         .TextMatrix(lnCtr + 2, 3) = oTrans.Detail(lnCtr, "nPrevLeav")
         .TextMatrix(lnCtr + 2, 4) = oTrans.Detail(lnCtr, "nNoOfLeav")
      Next
      
      .Row = 2
      pnActiveRow = 0
      
      Call SetFieldInfo
   End With
End Sub

Private Sub SetFieldInfo()
   With MSFlexGrid1
      Call HiglightRow(Me.MSFlexGrid1, .Row, 1)
      pnActiveRow = .Row - 2
      
      If Not .RowIsVisible(.Row) Then .TopRow = .TopRow + 1
   End With
   With oTrans
      txtOthers(0) = .Detail(pnActiveRow, 0)
      txtOthers(1) = .Detail(pnActiveRow, 1)
      txtOthers(2) = .Detail(pnActiveRow, 2)
      txtOthers(3) = .Detail(pnActiveRow, 3)
      txtOthers(4) = .Detail(pnActiveRow, 4)
      txtOthers(7) = .Detail(pnActiveRow, 7)
   End With
End Sub

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      If .Row < 2 Then Exit Sub
      
'      Call SetFieldInfo
'      txtOthers(3).SetFocus
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Row As Integer, ByVal Index As Variant)
   With oTrans
      Select Case Index
         Case 3
            txtOthers(3) = .Detail(Row, Index)
            MSFlexGrid1.TextMatrix(Row + 2, 4) = txtOthers(3)
         Case 4
            txtOthers(Index) = .Detail(Row, 4)
      End Select
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant)
   Select Case Index
      Case 1, 2, 3
         txtField(Index) = strLongDate(oTrans.Master(Index))
      Case Else
         txtField(Index) = oTrans.Master(Index)
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   oTrans.Master(Index) = txtField(Index)
End Sub
Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
         Case 2, 3
            .Text = strShortDate(.Text)
      End Select
      
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3
   Case vbKeyReturn
   End Select
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
         If GetFocus <> txtOthers(4).hWnd Then SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   For Each loTxt In txtOthers
      loTxt = ""
   Next
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

Private Sub txtOthers_GotFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         If Index = 4 Then
            With MSFlexGrid1
               If .Row <> .Rows - 1 Then
                  Call txtOthers_Validate(4, False)
                  .Row = .Row + 1
                  Call MSFlexGrid1_Click
               End If
            End With
         End If
   End Select
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub
Private Sub HiglightRow(ByVal loGrid As MSFlexGrid, _
                           ByVal lnRow As Integer, _
                           ByVal lnColStart As Integer)
   Dim lnCtr As Integer
   
   With loGrid
      If lnRow < 2 Then Exit Sub
      
      If lnRow <> pnActiveRow Then
         .Row = pnActiveRow + 2
         For lnCtr = lnColStart To .Cols - 1
            .Col = lnCtr
            .CellBackColor = &HFFFFFF
         Next
         
         .Row = lnRow
         For lnCtr = lnColStart To .Cols - 1
            .Col = lnCtr
            .CellBackColor = &H8000000D
         Next
      End If
   End With
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 3, 4
         oTrans.Detail(pnActiveRow, Index) = txtOthers(Index)
   End Select
End Sub
