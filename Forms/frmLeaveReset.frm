VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLeaveReset 
   BorderStyle     =   0  'None
   Caption         =   "Leave Reset"
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5970
      Left            =   8265
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   570
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10530
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
      Height          =   2415
      Index           =   0
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   4260
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
         Height          =   390
         Index           =   3
         Left            =   4545
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1185
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
         Top             =   1185
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
         Height          =   645
         Index           =   4
         Left            =   1275
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmLeaveReset.frx":0000
         Top             =   1605
         Width           =   5235
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
         Left            =   1275
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Test"
         Top             =   765
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
         Top             =   180
         Width           =   1965
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   4290
         Top             =   165
         Width           =   2220
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   4320
         Top             =   195
         Width           =   2160
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
         Left            =   4365
         TabIndex        =   28
         Tag             =   "eb0;et0"
         Top             =   225
         Width           =   2070
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
         Left            =   3495
         TabIndex        =   6
         Top             =   1260
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
         Top             =   1605
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
         Top             =   1260
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
         Left            =   705
         TabIndex        =   2
         Top             =   825
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
         Top             =   255
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   390
         Left            =   1380
         Tag             =   "et0;ht2"
         Top             =   285
         Width           =   1965
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3855
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
      Picture         =   "frmLeaveReset.frx":003A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4485
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
      Picture         =   "frmLeaveReset.frx":07B4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2595
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
      Picture         =   "frmLeaveReset.frx":0F2E
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1590
      Index           =   1
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   4965
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   2805
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
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Supervisor"
         Top             =   630
         Width           =   1965
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
         Left            =   5415
         MaxLength       =   50
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   630
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
         Left            =   5415
         MaxLength       =   50
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1050
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
         Left            =   1275
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Width           =   5235
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
         Left            =   690
         TabIndex        =   18
         Top             =   705
         Width           =   435
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Leave Crdt."
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
         Left            =   3645
         TabIndex        =   20
         Top             =   705
         Width           =   1635
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining Leave"
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
         Left            =   3810
         TabIndex        =   22
         Top             =   1125
         Width           =   1470
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
         Left            =   255
         TabIndex        =   16
         Top             =   210
         Width           =   870
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3225
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
      Picture         =   "frmLeaveReset.frx":16A8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3855
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
      Picture         =   "frmLeaveReset.frx":1E22
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   630
      Index           =   2
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   2985
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   1111
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
         Height          =   390
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   11
         Top             =   105
         Width           =   5235
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
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
         Index           =   15
         Left            =   465
         TabIndex        =   10
         Top             =   180
         Width           =   675
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1320
      Index           =   3
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   3630
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   2328
      BackColor       =   12632256
      ClipControls    =   0   'False
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
         Left            =   5415
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "00"
         Top             =   120
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
         Height          =   645
         Index           =   4
         Left            =   1275
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   540
         Width           =   5235
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asgd. Leave Crdt."
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
         Index           =   9
         Left            =   3600
         TabIndex        =   12
         Top             =   195
         Width           =   1680
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
         Left            =   345
         TabIndex        =   14
         Top             =   540
         Width           =   780
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4485
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
      Picture         =   "frmLeaveReset.frx":259C
   End
End
Attribute VB_Name = "frmLeaveReset"
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
'  Mac PH (08.10.12)
'     Created this form.
'  Mac PH (08.21.12)
'     Combines entry and posting of leave reset in this form.
'     Added Search On function.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeModuleName = "frmLeaveReset"
Private Const pxeVisibleRow = 17

Private WithEvents oTrans As clsBatchLeave
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim psSelected() As String
Dim pnIndex As Integer
Dim pnActiveRow As Integer
Dim pbControl As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc
   With oTrans
      Select Case Index
         Case 0 'browse
         Case 1 'new/update
            InitForm 1
         Case 2 'save
            If .SaveTransaction Then
               MsgBox "Transaction was saved successfully.", vbInformation, "Notice"
               .NewTransaction
               InitForm .EditMode
               LoadMaster
               LoadDetail
            End If
         Case 3 'approve
            If .EditMode = xeModeUnknown Then
               MsgBox "No transaction is active.", vbInformation, "Notice"
               Exit Sub
            End If
            
            If .PostTransaction Then
               Label2.Caption = TransStat(.Master("cTranstat"))
               MsgBox "Transaction was posted successfully.", vbInformation, "Notice"
               Unload Me
            End If
         Case 4 'disaaprove
            If .EditMode = xeModeUnknown Then
               MsgBox "No transaction is active.", vbInformation, "Notice"
               Exit Sub
            End If
            
            If .CancelTransaction Then
               Label2.Caption = TransStat(.Master("cTranstat"))
               MsgBox "Transaction was cancelled successfully.", vbInformation, "Notice"
               .NewTransaction
               InitForm .EditMode
            Else
               MsgBox "Unable to cancel transaction.", vbInformation, "Notice"
            End If
         Case 5 'close
            Unload Me
         Case 6 'cancel
            lnRep = MsgBox("This will disregard any changes in this transaction." + _
                        vbCrLf + vbCrLf + "Do you want to proceed?", vbQuestion + vbYesNo, "Confirm")
                        
            If lnRep = vbYes Then
               If .EditMode = xeModeAddNew Then
                  Unload Me
               ElseIf .EditMode = xeModeUpdate Then
                  .InitTransaction
                  .NewTransaction
                  InitForm .EditMode
                  LoadMaster
                  LoadDetail
               End If
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
   'On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      bLoaded = True
   End If

   pbControl = False
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyControl Then pbControl = False
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

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
      .NewTransaction
      
      InitForm .EditMode
   End With
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
      .ColWidth(0) = 550
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

Private Sub InitForm(ByVal fnEdit As Integer)
   'buttons : 0 browse, 1 new, 2 save, 3 approve, 4 disapprove, 5 close, 6 cancel
   'fnEdit : 0 ready mode, 1 entry mode, 2 update
'   Dim loFrame As xrFrame
  
   cmdButton(1).Visible = (fnEdit = 2)
   cmdButton(3).Visible = (fnEdit = 2)
   cmdButton(4).Visible = (fnEdit = 2)
   cmdButton(5).Visible = (fnEdit = 2)
   
   cmdButton(2).Visible = (fnEdit = 1)
   cmdButton(6).Visible = Not (fnEdit = 0)
   
   MSFlexGrid1.Enabled = Not (fnEdit = 0)
   xrFrame2(0).Enabled = (fnEdit = 1)
   xrFrame2(3).Enabled = (fnEdit = 1)
   
   InitGrid
   ClearFields
   
   Select Case fnEdit
      Case 0
      Case 1, 2
         LoadMaster
         LoadDetail
   End Select
End Sub

Private Sub LoadMaster()
   Dim loTxt As TextBox
   With oTrans
      For Each loTxt In txtField
         Select Case loTxt.Index
         Case 1 To 3
            loTxt = Format(.Master(loTxt.Index), "Mmm DD, YYYY")
         Case Else
            loTxt = IFNull(.Master(loTxt.Index))
         End Select
      Next
      
      Label2.Caption = IIf(.EditMode = xeModeAddNew, "NEW", TransStat(.Master("cTranstat")))
      If bLoaded = True And .EditMode <> xeModeUpdate Then
         txtField(2).SetFocus
      End If
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
         .TextMatrix(lnCtr + 2, 3) = IFNull(oTrans.Detail(lnCtr, "nPrevLeav"), 0)
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
      
      If Not .RowIsVisible(.Row) Then .TopRow = .Row
   End With
   With oTrans
      txtOthers(0) = .Detail(pnActiveRow, 0)
      txtOthers(1) = IFNull(.Detail(pnActiveRow, 1))
      txtOthers(2) = .Detail(pnActiveRow, 2)
      txtOthers(3) = .Detail(pnActiveRow, 3)
      txtOthers(4) = .Detail(pnActiveRow, 4)
      txtOthers(7) = .Detail(pnActiveRow, 7)
   End With
End Sub

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      If .Row < 2 Then Exit Sub
      
      Call SetFieldInfo
      txtSearch.SetFocus
   End With
End Sub

Private Sub MSFlexGrid1_GotFocus()
   txtSearch.SetFocus
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Row As Integer, ByVal Index As Variant)
   With oTrans
      If .EditMode = xeModeUnknown Then Exit Sub
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
   If oTrans.EditMode = xeModeUnknown Then Exit Sub
   Select Case Index
      Case 1 To 3
         txtField(Index) = Format(oTrans.Master(Index), "Mmm DD, YYYY")
      Case Else
         txtField(Index) = oTrans.Master(Index)
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   oTrans.Master(Index) = txtField(Index)
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With oTrans
      If .EditMode = xeModeUnknown Or _
         .EditMode = xeModeUpdate Then Exit Sub
   End With
   
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

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   With MSFlexGrid1
      Select Case KeyCode
         Case vbKeyReturn, vbKeyUp, vbKeyDown
            Select Case KeyCode
            Case vbKeyReturn, vbKeyDown
               If pbControl Then
                  If GetFocus <> txtSearch.hWnd Then Exit Sub
                  If .Row = .Rows - 1 Then Exit Sub
                  
                  .Row = .Row + 1
                  SetFieldInfo
               Else
                  If GetFocus = txtSearch.hWnd Then Exit Sub
                  If GetFocus <> txtOthers(4).hWnd Then SetNextFocus
               End If
            Case vbKeyUp
               If pbControl Then
                  If GetFocus <> txtSearch.hWnd Then Exit Sub
                  If .Row = 2 Then Exit Sub
                  
                  .Row = .Row - 1
                  SetFieldInfo
               Else
                  SetPreviousFocus
               End If
            End Select
         Case vbKeyControl
            pbControl = True
         Case vbKeyTab
            If GetFocus = txtOthers(4).hWnd Then
               txtSearch.SetFocus
               KeyCode = 0
            End If
      End Select
   End With
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
      loTxt.BackColor = oApp.getColor("EB")
   Next
   For Each loTxt In txtOthers
      loTxt = ""
      loTxt.BackColor = oApp.getColor("EB")
   Next
   
   txtSearch = ""
   txtSearch.BackColor = oApp.getColor("EB")
   
   Label2.Caption = ""
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
                  txtSearch.SetFocus
'                  .Row = .Row + 1
'                  Call MSFlexGrid1_Click
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
'------ SEARCH ON ------'
Private Function ResultingText(iKeyAscii%) As String
   Dim sLeft As String             ' string element
   Dim sSel As String              ' selected string element
   Dim sRight As String            ' string element
   Dim sResult As String           ' what well return
   
   On Error Resume Next
   
   With txtSearch
      sLeft = Left$(.Text, .SelStart)         ' SelStart is 0-based
      sSel = Mid$(.Text, .SelStart + 1, .SelLength)
      sRight = Mid$(.Text, .SelStart + .SelLength + 1)
   End With
   
   Select Case iKeyAscii
      Case vbKeyBack             'Backspace Key
         If Len(sSel) = 0 Then   'Nothing selected
            sResult = MinusRightChar(sLeft) & sRight  'Del first char on the left
         Else                    'Selection exists
            sResult = sLeft & sRight   'Delete selected text only
         End If
         
      Case vbKeyDelete           'Delete key
         If Len(sSel) = 0 Then   'Nothing selected
            sResult = sLeft & MinusLeftChar(sRight)    'Del first char on the right
         Else
            sResult = sLeft & sRight    'Delete selected text only
         End If
         
      Case Else         'an ordinary character
         sResult = sLeft & Chr$(iKeyAscii) & sRight
   End Select
   ResultingText = sResult
End Function

Private Function MinusLeftChar(ByVal sGiven As String) As String

   'Purpose: Returns <sGiven> with the leftmost character removed, or "" if
   '         <sGiven> was empty.
   '
   'Returns: The trimmed string
   '
   'Remarks: Just a safe wrapper for Mid$()
   On Error Resume Next
   
   If Len(sGiven) = 0 Then
      MinusLeftChar = ""
   Else
      MinusLeftChar = Mid$(sGiven, 2)
   End If
End Function

Private Function MinusRightChar(ByVal sGiven As String) As String

   'Purpose: Returns <sGiven> with the rightmost character removed, or "" if
   '         <sGiven> was empty.
   '
   'Returns: The trimmed string
   '
   'Remarks: Just a safe wrapper for Left$()
   On Error Resume Next
   
   If Len(sGiven) = 0 Then
      MinusRightChar = ""
   Else
      MinusRightChar = Left$(sGiven, Len(sGiven) - 1)
   End If
End Function

Private Function SearchOn(ByVal lsSeek) As Boolean
   Dim lnCtr As Long
   Dim lnRow As Integer
   
   With MSFlexGrid1
      For lnCtr = 2 To .Rows - 1
         Debug.Print lnCtr
         If StrComp(Left(.TextMatrix(lnCtr, 1), Len(lsSeek)), lsSeek, vbTextCompare) >= 0 Then
            If StrComp(Left(.TextMatrix(lnCtr, 1), Len(lsSeek)), lsSeek, vbTextCompare) = 0 Then
               lnRow = lnCtr
            End If
            Exit For
         End If
      Next
      If lnRow = 0 Then
         lnRow = 2
      End If
      
      .Row = lnRow
      Call MSFlexGrid1_Click
      SearchOn = True
   End With
End Function

Private Sub txtSearch_GotFocus()
   With txtSearch
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   'Remarks: This procedure only exists to trap a delete key, which irritatingly,
   '         does not trigger a KeyPress event
   '
   Dim lsSearchOn As String          'current string to search on

   On Error Resume Next
   
   'Check if we're dealing with a Delete key
   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or _
          KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
      MSFlexGrid1.SetFocus
      Exit Sub
   ElseIf KeyCode = vbKeyReturn Then
      KeyCode = 0
      Exit Sub
   ElseIf KeyCode <> vbKeyDelete Then
      Exit Sub
   End If
   
   'The delete key was pressed; decide what to search on
   lsSearchOn = ResultingText(KeyCode)
   SearchOn lsSearchOn
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
   Dim lsSearchOn As String             'current string to search on
   
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      Exit Sub
   End If
   'A content-modifying key was entered; decide what to search on
   lsSearchOn = ResultingText(KeyAscii)
   If SearchOn(lsSearchOn) = False Then KeyAscii = 0
End Sub

Private Sub txtSearch_LostFocus()
   With txtSearch
      .BackColor = oApp.getColor("EB")
   End With
End Sub
