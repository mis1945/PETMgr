VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWTaxChart 
   BorderStyle     =   0  'None
   Caption         =   "WTax Chart"
   ClientHeight    =   9285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   2910
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   5265
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   5133
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
         Height          =   400
         Index           =   2
         Left            =   1785
         TabIndex        =   7
         Text            =   "25,500.00"
         Top             =   255
         Width           =   1410
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
         Height          =   400
         Index           =   3
         Left            =   1785
         TabIndex        =   9
         Text            =   ".05"
         Top             =   705
         Width           =   1410
      End
      Begin VB.TextBox txtAmntx 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5895
         MaxLength       =   50
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   675
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   9960
         MaxLength       =   50
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "7, 200.00"
         Top             =   675
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5895
         MaxLength       =   50
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "10,000.00"
         Top             =   1095
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   5895
         MaxLength       =   50
         TabIndex        =   19
         Text            =   "10,000.00"
         Top             =   1515
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   5895
         MaxLength       =   50
         TabIndex        =   21
         Text            =   "10,000.00"
         Top             =   1935
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5895
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "0"
         Top             =   255
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   5895
         MaxLength       =   50
         TabIndex        =   23
         Text            =   "10,000.00"
         Top             =   2355
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   9960
         MaxLength       =   50
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "10,000.00"
         Top             =   1095
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   9960
         MaxLength       =   50
         TabIndex        =   29
         Text            =   "10,000.00"
         Top             =   1515
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   9960
         MaxLength       =   50
         TabIndex        =   31
         Text            =   "10,000.00"
         Top             =   1935
         Width           =   1755
      End
      Begin VB.TextBox txtAmntx 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   9960
         MaxLength       =   50
         TabIndex        =   33
         Text            =   "10,000.00"
         Top             =   2355
         Width           =   1755
      End
      Begin xrControl.xrButton cmdButton 
         Height          =   600
         Index           =   6
         Left            =   1785
         TabIndex        =   10
         Top             =   1170
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1058
         Caption         =   "&Process"
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
         Picture         =   "frmWTaxChart.frx":0000
      End
      Begin xrControl.xrButton cmdButton 
         Height          =   600
         Index           =   7
         Left            =   1785
         TabIndex        =   11
         Top             =   1785
         Width           =   1410
         _ExtentX        =   2487
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
         Picture         =   "frmWTaxChart.frx":077A
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Amount"
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
         Left            =   195
         TabIndex        =   6
         Top             =   335
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "In Excess"
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
         Left            =   195
         TabIndex        =   8
         Top             =   780
         Width           =   870
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single"
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
         Left            =   4380
         TabIndex        =   14
         Top             =   735
         Width           =   540
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Married"
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
         Left            =   8745
         TabIndex        =   24
         Top             =   750
         Width           =   645
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S1"
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
         Left            =   4380
         TabIndex        =   16
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S2"
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
         Left            =   4380
         TabIndex        =   18
         Top             =   1575
         Width           =   240
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S3"
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
         Left            =   4380
         TabIndex        =   20
         Top             =   1995
         Width           =   240
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Exemption"
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
         Left            =   4380
         TabIndex        =   12
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S4"
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
         Left            =   4380
         TabIndex        =   22
         Top             =   2415
         Width           =   240
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ME1"
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
         Left            =   8745
         TabIndex        =   26
         Top             =   1185
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ME2"
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
         Left            =   8745
         TabIndex        =   28
         Top             =   1590
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ME3"
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
         Left            =   8745
         TabIndex        =   30
         Top             =   2010
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ME4"
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
         Left            =   8745
         TabIndex        =   32
         Top             =   2415
         Width           =   405
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   11400
      TabIndex        =   37
      Top             =   8475
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmWTaxChart.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   10620
      TabIndex        =   36
      Top             =   8475
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Add"
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
      Picture         =   "frmWTaxChart.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   11400
      TabIndex        =   38
      Top             =   8475
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmWTaxChart.frx":2700
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   10620
      TabIndex        =   39
      Top             =   8475
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmWTaxChart.frx":2E7A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   9825
      TabIndex        =   35
      Top             =   8475
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmWTaxChart.frx":35F4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   9060
      TabIndex        =   34
      Top             =   8475
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmWTaxChart.frx":3D6E
      PicturePos      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3930
      Left            =   105
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1275
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   6932
      _Version        =   393216
      Rows            =   12
      Cols            =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   675
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   1191
      BackColor       =   12632256
      ClipControls    =   0   'False
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
         Index           =   1
         ItemData        =   "frmWTaxChart.frx":44E8
         Left            =   6300
         List            =   "frmWTaxChart.frx":44F8
         TabIndex        =   3
         Text            =   "Semi-Monthly"
         Top             =   105
         Width           =   2190
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
         Height          =   400
         Index           =   0
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "August 16, 2011"
         Top             =   105
         Width           =   2190
      End
      Begin xrControl.xrButton cmdButton 
         Height          =   480
         Index           =   0
         Left            =   8940
         TabIndex        =   4
         Top             =   105
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   847
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
         Picture         =   "frmWTaxChart.frx":4522
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table Type"
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
         Left            =   4725
         TabIndex        =   2
         Top             =   185
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effectivity Date"
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
         Left            =   240
         TabIndex        =   0
         Top             =   185
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmWTaxChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmWTaxChart"

Private WithEvents oTrans As clsWTaxChart
Attribute oTrans.VB_VarHelpID = -1
Private poRS As Recordset
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnCtr As Integer
Dim pnIndex As Integer

Dim pcSalTypex As String
Dim pdWTaxEfct As Date

Dim pnTaxColmn As Integer

Property Let Period(ByVal Value As String)
   If InStr("D«W«S«M", Value) > 0 Then
      pcSalTypex = Value
   End If
End Property

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0 'Retrieve
      'We don't have anything to do yet...
   Case 4 'New
      'We don't have anything to do yet...
   Case 6 'Process
      If oTrans.OpenTransaction(pcSalTypex, txtField(2).Text, txtField(3).Text) Then
         If oTrans.UpdateTransaction Then
            Call InitForm(xeModeUpdate)
            txtAmntx(0).SetFocus
            pnTaxColmn = -1
         End If
      End If
   Case 7 'Process-Cancel
      Call loadChart
      
      'Reload the contents of txtAmntx fields
      txtField(2).Text = "0.00"
      txtField(3).Text = "0.00"
      Call loadTaxDetail
      
      Call InitForm(xeModeReady)
      cmdButton(8).SetFocus
   Case 3 'Save
      If oTrans.SaveTransaction Then
         MsgBox "WTax Chart was save successfully!", vbOKOnly + vbInformation, "WTax Maintenance"
         
         Call loadChart
         Call InitForm(xeModeReady)
         cmdButton(5).SetFocus
      Else
         MsgBox "Unable to save WTax Chart!", vbOKOnly + vbCritical, "WTax Maintenance"
      End If
   Case 2 ' cancel
      If pnTaxColmn = -1 Then
         txtField(2).Text = "0.00"
         txtField(3).Text = "0.00"
         Call loadTaxDetail
         Call InitForm(xeModeReady)
         cmdButton(8).SetFocus
      Else
         Call loadChart
         'Reload WTax Chart
         Call loadTaxDetail
         Call InitForm(xeModeReady)
         cmdButton(5).SetFocus
      End If
   Case 5 'Update
      If oTrans.UpdateTransaction Then
         Call InitForm(xeModeUpdate)
         txtAmntx(1).SetFocus
      End If
   Case 8 'Add New Tax Column
      txtField(2).Text = "0.00"
      txtField(3).Text = "0.00"
      Call loadTaxDetail
      Call InitForm(xeModeAddNew)
      txtField(2).SetFocus
   Case 1  'Close
      Unload Me
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   Set oTrans = New clsWTaxChart
   Set oTrans.AppDriver = oApp
   
   oTrans.InitTransaction

   pcSalTypex = "S"
   pdWTaxEfct = oApp.getConfiguration("WTaxDte")
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin

   Call InitGrid

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      Call loadChart

      txtField(0).Text = Format(pdWTaxEfct, "Mmmm dd, yyyy")
      
      Select Case pcSalTypex
      Case "M"
         cmbField(1).Text = "Monthly"
      Case "S"
         cmbField(1).Text = "Semi-Monthly"
      Case "W"
         cmbField(1).Text = "Weekly"
      Case "D"
         cmbField(1).Text = "Daily"
      End Select
      
      txtField(2).Text = "0.00"
      txtField(3).Text = "0.00"
      
      Call loadTaxDetail
      Call InitForm(xeModeReady)
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
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   bLoaded = False
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub


Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSHFlexGrid1
      
      Set poRS = getTaxCol()
      
      .Cols = poRS.RecordCount + 1
      .MergeCells = flexMergeFree
      
      .Clear
      
      .Row = 0
      
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         '.CellAlignment = flexAlignCenterCenter
         .WordWrap = True
      Next
               
      .Row = 0
      .RowHeight(0) = 500
      .MergeRow(0) = True
      .TextMatrix(0, 0) = "Exemption"
                         
      lnCtr = 1
      Do Until poRS.EOF
         .TextMatrix(0, lnCtr) = Format(poRS("nTaxAmntx"), "#,##0.00") & vbCrLf & "+ " & poRS("nTaxOverx") * 100 & "% Over"
         lnCtr = lnCtr + 1
         poRS.MoveNext
      Loop
      
      .TextMatrix(1, 0) = "Zero Exemption(Z)"
      .TextMatrix(2, 0) = "Single(S)"
      .TextMatrix(3, 0) = "Married(M)"
      .TextMatrix(4, 0) = "Single w/ 1 Qualified Dependent(S1)"
      .TextMatrix(5, 0) = "Single w/ 2 Qualified Dependent(S2)"
      .TextMatrix(6, 0) = "Single w/ 3 Qualified Dependent(S3)"
      .TextMatrix(7, 0) = "Single w/ 4 Qualified Dependent(S4)"
      .TextMatrix(8, 0) = "Married w/ 1 Qualified Dependent(ME1)"
      .TextMatrix(9, 0) = "Married w/ 2 Qualified Dependent(ME2)"
      .TextMatrix(10, 0) = "Married w/ 3 Qualified Dependent(ME3)"
      .TextMatrix(11, 0) = "Married w/ 4 Qualified Dependent(ME4)"

      'column width
      .ColWidth(0) = 3550
      .ColWidth(1) = 1200
      .ColWidth(2) = 1200
      .ColWidth(3) = 1200
      .ColWidth(4) = 1200
      .ColWidth(5) = 1200
      .ColWidth(6) = 1200
      .ColWidth(7) = 1200
      
      .Rows = 12
   End With
   
End Sub

Private Sub loadChart()
   Dim lnColCtr As Integer
   Dim lnRowCtr As Integer
   Dim lors As Recordset
   
   poRS.MoveFirst
   lnColCtr = 1
   Do Until poRS.EOF
      Set lors = LoadDetail(poRS("nTaxAmntx"), poRS("nTaxOverx"))
      lnRowCtr = 1
      Do Until lors.EOF
         MSHFlexGrid1.TextMatrix(lnRowCtr, lnColCtr) = Format(IFNull(lors("nRangeFrm"), 0), "#,##0.00")
         lnRowCtr = lnRowCtr + 1
         lors.MoveNext
      Loop
      lnColCtr = lnColCtr + 1
      poRS.MoveNext
   Loop
   
End Sub

Private Function getTaxCol() As Recordset
   Dim lsSQL As String
   
   lsSQL = "SELECT DISTINCT a.nTaxAmntx, a.nTaxOverx" & _
          " FROM WTAX_Chart a" & _
          " WHERE a.dFromDate = " & dateParm(pdWTaxEfct) & _
            " AND a.cSalTypex = " & strParm(pcSalTypex) & _
          " ORDER BY a.nTaxAmntx, a.nTaxOverx"

   Set getTaxCol = oApp.Connection.Execute(lsSQL, , adCmdText)
End Function

Private Function LoadDetail(ByVal fnTaxAmntx As Currency, ByVal fnTaxOverx As Single) As Recordset
   Dim lsSQL As String
   
   lsSQL = "SELECT" & _
                  "  b.sExemptNm" & _
                  ", a.nRangeFrm" & _
                  ", a.nTaxAmntx" & _
                  ", a.nTaxOverx" & _
                  ", a.dFromDate" & _
                  ", a.cSalTypex" & _
                  ", b.sExemptID" & _
            " FROM Exemption b" & _
               " LEFT JOIN WTAX_Chart a" & _
                  " ON a.sExemptID = b.sExemptID" & _
                 " AND a.dFromDate = " & dateParm(pdWTaxEfct) & _
                 " AND a.cSalTypex = " & strParm(pcSalTypex) & _
                 " AND a.nTaxAmntx = " & fnTaxAmntx & _
                 " AND a.nTaxOverx = " & fnTaxOverx & _
            " ORDER BY a.nTaxAmntx, a.nTaxOverx, b.sExemptID"
   
   Set LoadDetail = oApp.Connection.Execute(lsSQL, , adCmdText)
End Function

Private Sub loadTaxDetail()
   Dim loTxt As TextBox
   
   oTrans.InitTransaction
   
   If Not (txtField(2).Text = 0 And txtField(3).Text = 0) Then
      Call oTrans.OpenTransaction(pcSalTypex, txtField(2).Text, txtField(3).Text)
   End If
   
   For Each loTxt In txtAmntx
      If oTrans.ItemCount > 0 Then
         loTxt.Text = Format(oTrans.Detail(loTxt.Index - 1, "nRangeFrm"), "#,##0.00")
      Else
         loTxt.Text = "0.00"
      End If
   Next
End Sub

Private Sub InitForm(ByVal pnEditMode As xeEditMode)
   Dim lnCtr As Integer
   
   'Retrieve is not yet needed so we need to hide from the eyes of the user....
   cmdButton(0).Visible = False
   cmdButton(4).Visible = False
   
'   Select Case pnEditMode
'   Case xeEditMode.xeModeAddNew
'      txtField(2).Enabled = True
'      txtField(3).Enabled = True
'      cmdButton(6).Enabled = True
'      cmdButton(7).Enabled = True
'      For lnCtr = 1 To 11
'         txtAmntx(lnCtr).Enabled = False
'      Next
'   Case xeEditMode.xeModeReady
'      txtField(2).Enabled = False
'      txtField(3).Enabled = False
'      cmdButton(6).Enabled = False
'      cmdButton(7).Enabled = False
'
'      For lnCtr = 1 To 11
'         txtAmntx(lnCtr).Enabled = False
'      Next
'   Case xeEditMode.xeModeUpdate
'      txtField(2).Enabled = False
'      txtField(3).Enabled = False
'      cmdButton(6).Enabled = False
'      cmdButton(7).Enabled = False
'
'      For lnCtr = 1 To 11
'         txtAmntx(lnCtr).Enabled = True
'      Next
'   End Select
   
   'Enable only if in xeModeAddNew Mode
   txtField(2).Enabled = pnEditMode = xeModeAddNew
   txtField(3).Enabled = pnEditMode = xeModeAddNew
   cmdButton(6).Enabled = pnEditMode = xeModeAddNew
   cmdButton(7).Enabled = pnEditMode = xeModeAddNew
   
   'Enable only during xeModeUpdate Mode
   For lnCtr = 1 To 11
      txtAmntx(lnCtr).Enabled = pnEditMode = xeModeUpdate
   Next
   
   'Save and Cancel button should be visible during update mode only
   cmdButton(2).Visible = pnEditMode = xeModeUpdate
   cmdButton(3).Visible = pnEditMode = xeModeUpdate
   
   '1-Close;4-New;5-Update;8-Add
   'These buttons should be visible if not in xeModeUpdate Mode
   cmdButton(1).Visible = Not (pnEditMode = xeModeUpdate)
   'cmdButton(4).Visible = Not (pnEditMode = xeModeUpdate)
   cmdButton(5).Visible = Not (pnEditMode = xeModeUpdate)
   cmdButton(8).Visible = Not (pnEditMode = xeModeUpdate)
   
   'Allow Event for these button if in xeModeReady Mode
   cmdButton(1).Enabled = (pnEditMode = xeModeReady)
   'cmdButton(4).Enabled = (pnEditMode = xeModeReady)
   cmdButton(5).Enabled = (pnEditMode = xeModeReady)
   cmdButton(8).Enabled = (pnEditMode = xeModeReady)
End Sub


Private Sub MSHFlexGrid1_DblClick()
   pnTaxColmn = MSHFlexGrid1.ColSel
   poRS.MoveFirst
   If pnTaxColmn > 1 Then
      poRS.Move pnTaxColmn - 1
   End If
   txtField(2).Text = poRS("nTaxAmntx")
   txtField(3).Text = poRS("nTaxOverx")
   
   Call loadTaxDetail
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Row As Integer, ByVal Index As Integer, ByVal Value As Variant)
   txtAmntx(Row + 1) = Format(Value, "#,##0.00")
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
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

Private Sub txtAmntx_GotFocus(Index As Integer)
   With txtAmntx(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
      
      pnIndex = Index
   End With
End Sub

Private Sub txtAmntx_LostFocus(Index As Integer)
   With txtAmntx(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtAmntx_Validate(Index As Integer, Cancel As Boolean)
   oTrans.Detail(Index - 1, "nRangeFrm") = txtAmntx(Index).Text
   txtAmntx(Index) = Format(oTrans.Detail(Index - 1, "nRangeFrm"), "#,##0.00")
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

