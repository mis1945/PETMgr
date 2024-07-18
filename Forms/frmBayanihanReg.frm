VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmBayanihanReg 
   BorderStyle     =   0  'None
   Caption         =   "Bayanihan Program Reg"
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   8430
      TabIndex        =   5
      Top             =   1515
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
      Picture         =   "frmBayanihanReg.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   8430
      TabIndex        =   4
      Top             =   885
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
      Picture         =   "frmBayanihanReg.frx":077A
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1125
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   8010
      _ExtentX        =   14129
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
         Index           =   0
         Left            =   1620
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   120
         Width           =   2415
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
         Index           =   1
         Left            =   1620
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   585
         Width           =   4815
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
         Top             =   210
         Width           =   1365
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
         Index           =   15
         Left            =   105
         TabIndex        =   2
         Top             =   675
         Width           =   870
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   7095
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1740
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   12515
      BackColor       =   12632256
      ClipControls    =   0   'False
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   28
         Text            =   "M00111-000021"
         Top             =   120
         Width           =   2190
      End
      Begin VB.TextBox txtOther 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   27
         Top             =   6525
         Width           =   1650
      End
      Begin VB.TextBox txtOther 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   26
         Top             =   6525
         Width           =   1650
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
         Index           =   11
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   25
         Top             =   5130
         Width           =   1650
      End
      Begin VB.TextBox txtOther 
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
         Index           =   2
         Left            =   3810
         MaxLength       =   50
         TabIndex        =   24
         Top             =   5130
         Width           =   1650
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
         Index           =   10
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   23
         Top             =   5130
         Width           =   1650
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
         Index           =   7
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   22
         Top             =   6060
         Width           =   1650
      End
      Begin VB.TextBox txtOther 
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
         Index           =   4
         Left            =   3810
         MaxLength       =   50
         TabIndex        =   21
         Top             =   6060
         Width           =   1650
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   20
         Top             =   6060
         Width           =   1650
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
         Index           =   9
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   19
         Top             =   5595
         Width           =   1650
      End
      Begin VB.TextBox txtOther 
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
         Index           =   3
         Left            =   3810
         MaxLength       =   50
         TabIndex        =   18
         Top             =   5595
         Width           =   1650
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
         Index           =   8
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   17
         Top             =   5595
         Width           =   1650
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
         Index           =   13
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   16
         Top             =   4665
         Width           =   1650
      End
      Begin VB.TextBox txtOther 
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
         Index           =   1
         Left            =   3810
         MaxLength       =   50
         TabIndex        =   15
         Top             =   4665
         Width           =   1650
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
         Index           =   12
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   14
         Top             =   4665
         Width           =   1650
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
         Index           =   15
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   13
         Top             =   4200
         Width           =   1650
      End
      Begin VB.TextBox txtOther 
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
         Index           =   0
         Left            =   3810
         MaxLength       =   50
         TabIndex        =   12
         Top             =   4200
         Width           =   1650
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
         Index           =   14
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   11
         Top             =   4200
         Width           =   1650
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
         Index           =   5
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3180
         Width           =   2415
      End
      Begin VB.ComboBox cmbField 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         ItemData        =   "frmBayanihanReg.frx":0EF4
         Left            =   1620
         List            =   "frmBayanihanReg.frx":0EFE
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1860
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
         Height          =   800
         Index           =   4
         Left            =   1620
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2340
         Width           =   6240
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   7
         Top             =   915
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
         Index           =   80
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1395
         Width           =   6240
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
         Left            =   105
         TabIndex        =   43
         Top             =   210
         Width           =   1485
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1725
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   2175
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   4395
         Top             =   120
         Width           =   2220
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   4425
         Top             =   150
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
         Left            =   4470
         TabIndex        =   42
         Tag             =   "eb0;et0"
         Top             =   180
         Width           =   2070
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   6615
         TabIndex        =   41
         Top             =   3900
         Width           =   420
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contribution"
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
         Left            =   4125
         TabIndex        =   40
         Top             =   3900
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Member Count"
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
         Left            =   1815
         TabIndex        =   39
         Top             =   3900
         Width           =   1275
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Associate"
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
         Left            =   375
         TabIndex        =   38
         Top             =   6150
         Width           =   870
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisor"
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
         Left            =   375
         TabIndex        =   37
         Top             =   5685
         Width           =   900
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Head"
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
         Left            =   375
         TabIndex        =   36
         Top             =   5220
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dept Head"
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
         Left            =   375
         TabIndex        =   35
         Top             =   4755
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area Head"
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
         Left            =   375
         TabIndex        =   34
         Top             =   4290
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claim Date"
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
         TabIndex        =   33
         Top             =   3270
         Width           =   960
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relationship"
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
         Left            =   105
         TabIndex        =   32
         Top             =   1950
         Width           =   1185
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee "
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
         Left            =   105
         TabIndex        =   31
         Top             =   1470
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
         Left            =   105
         TabIndex        =   30
         Top             =   2340
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
         Left            =   105
         TabIndex        =   29
         Top             =   1005
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmBayanihanReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeModuleName = "frmBayanihanReg"

Private oDriver As clsFormDriver
Private WithEvents oTrans As clsBayanihanProgram
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim psSelected() As String
Dim pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   Select Case Index
   Case 3   'browse
       If pnIndex = 0 Or pnIndex = 1 Then
         If pnIndex = 0 Then
            If oTrans.SearchTransaction(txtSearch(pnIndex).Text, False) Then
               ClearFields
               LoadMaster
            End If
         Else
            If oTrans.SearchTransaction(txtSearch(pnIndex).Text) Then
               ClearFields
               LoadMaster
            End If
         End If
         pnIndex = 3
      Else
         If oTrans.SearchTransaction("") Then
            ClearFields
            LoadMaster
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
   End If

'   pbSearched = False
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me

   Set oTrans = New clsBayanihanProgram
   Set oTrans.AppDriver = oApp
   
'   oTrans.TransStatus = 10
   
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
'   If psTransNox <> "" Then
'      Call oTrans.OpenTransaction(psTransNox)
'   End If

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
   ClearFields
'   InitFields
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   txtSearch(0) = ""
   txtSearch(1) = ""
   Label2.Caption = ""
End Sub

Private Sub InitForm(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   cmdButton(3).Visible = Not lbShow
'   cmdButton(2).Visible = Not lbShow
'   cmdButton(5).Visible = Not lbShow
'   txtField(1).Enabled = Not lbShow
'   xrFrame2.Enabled = lbShow
'
'   cmdButton(0).Visible = lbShow
'   cmdButton(4).Visible = lbShow
'
'   If lbShow Then
'      txtField(2).SetFocus
'   End If
End Sub

Private Sub LoadMaster()
   Dim lnCtr As Integer
   Dim loTxt As TextBox
   
   txtField(0).Text = oTrans.Master(0)
   txtField(1).Text = strLongDate(oTrans.Master(1))
   txtField(5).Text = strLongDate(oTrans.Master(5))
      
   For Each loTxt In txtField
      lnCtr = loTxt.Index
   
      Select Case lnCtr
         Case 1, 5
            txtField(lnCtr).Text = strLongDate(oTrans.Master(lnCtr))
         Case 7, 9, 11, 13, 15
            txtField(lnCtr).Text = Format(oTrans.Master(lnCtr), "#,##0.00")
         Case Else
            txtField(lnCtr).Text = oTrans.Master(lnCtr)
      End Select
   Next
   
   If txtField(15).Text = 0# Or txtField(15).Text = "" Then
      txtOther(0) = 0#
   Else
      txtOther(0) = txtField(15) / txtField(14)
   End If
   txtOther(1) = txtField(13) / txtField(12)
   txtOther(2) = txtField(11) / txtField(10)
   txtOther(3) = txtField(9) / txtField(8)
   txtOther(4) = txtField(7) / txtField(6)
   txtOther(5) = CInt(txtField(6)) + CInt(txtField(8)) + CInt(txtField(10)) + CInt(txtField(12)) + CInt(txtField(14))
   txtOther(6) = CDec(txtField(7)) + CDec(txtField(9)) + CDec(txtField(11)) + CDec(txtField(13)) + CDec(txtField(15))
   txtOther(6) = Format(txtOther(6), "#,##0.00")
      
   txtSearch(0) = txtField(0)
   txtSearch(1) = txtField(80)
   
   cmbField(3).ListIndex = oTrans.Master(3)
   
   If oTrans.Master("cTranStat") = "4" Then
      Label2.Caption = "APPLIED"
   Else
      Label2.Caption = TransStat(CInt(oTrans.Master("cTranStat")))
   End If
   
'   pbSearched = True
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
      Case 1, 5
         txtField(Index) = strLongDate(Value)
      Case 7, 9, 11, 13, 15
         txtField(Index) = Format(Value, "#,##0.00")
         Select Case Index
         Case 15
            txtOther(0) = txtField(15) / txtField(14)
         Case 13
            txtOther(1) = txtField(13) / txtField(12)
         Case 11
            txtOther(2) = txtField(11) / txtField(10)
         Case 9
            txtOther(3) = txtField(9) / txtField(8)
         Case 7
            txtOther(4) = txtField(7) / txtField(6)
         End Select
      Case 16
         Label2.Caption = TransStat(CInt(Value))
      Case Else
         txtField(Index) = Value
   End Select
End Sub

'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   With txtField(Index)
'      If Index = 2 Then Exit Sub
'      oTrans.Master(Index) = .Text
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   Select Case Index
'      Case 1, 3, 4
'         txtField(Index).Text = Format(oTrans.Master(Index), "MM-DD-YYYY")
'   End Select
'
'   With txtField(Index)
'      .BackColor = oApp.getColor("HT1")
'      .SelStart = 0
'      .SelLength = Len(.Text)
'   End With
'
'   oDriver.ColumnIndex = Index
'   pnIndex = Index
'End Sub

'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyF3
'      If Index = 2 Then
'         Call oTrans.SearchMaster(2, txtField(Index))
'      End If
'   Case vbKeyReturn
'      If Index = 2 Then
'         If txtField(Index) <> "" Then
'            Call oTrans.SearchMaster(2, txtField(Index))
'         Else
'            oTrans.Master(Index) = txtField(Index)
'         End If
'      End If
'   End Select
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub

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
   
   oDriver.ColumnIndex = Index
   pnIndex = Index
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 0
         If oTrans.SearchTransaction(txtSearch(Index).Text, False) Then
            ClearFields
            LoadMaster
         End If
      Case 1
         If oTrans.SearchTransaction(txtSearch(Index).Text) Then
            ClearFields
            LoadMaster
         End If
      End Select
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


