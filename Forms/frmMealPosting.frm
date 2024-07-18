VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMealPosting 
   BorderStyle     =   0  'None
   Caption         =   "Meal Posting"
   ClientHeight    =   9660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtField 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   3600
      MaxLength       =   50
      TabIndex        =   6
      Text            =   "50"
      Top             =   4335
      Width           =   1245
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   9060
      Left            =   5640
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   15981
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   8370
         Left            =   30
         TabIndex        =   17
         Top             =   570
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   14764
         _Version        =   393216
         Rows            =   8
         Cols            =   3
         RowHeightMin    =   338
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
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
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "UNKNOWN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3855
         TabIndex        =   47
         Tag             =   "eb0;et0"
         Top             =   165
         Width           =   1995
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   360
         Left            =   3735
         Tag             =   "et0;ht2"
         Top             =   105
         Width           =   2205
      End
   End
   Begin xrControl.xrFrame xrFrame4 
      Height          =   9060
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   15981
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "50"
         Top             =   5910
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2010
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "50"
         Top             =   1680
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   11
         Text            =   "50"
         Top             =   6285
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "40"
         Top             =   6660
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "20"
         Top             =   7035
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "50"
         Top             =   4170
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "40"
         Top             =   4545
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "20"
         Top             =   4920
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   2010
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "50"
         Top             =   2055
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2010
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "40"
         Top             =   2430
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   2010
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "20"
         Top             =   2805
         Width           =   1245
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
         Height          =   390
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "Saturday"
         Top             =   180
         Width           =   2385
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
         Left            =   1170
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "Saturday"
         Top             =   855
         Width           =   2415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   14
         Left            =   1020
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "frmMealPosting.frx":0000
         Top             =   8265
         Width           =   2565
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Graveyard:"
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
         Index           =   25
         Left            =   795
         TabIndex        =   50
         Top             =   5955
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Graveyard:"
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
         Index           =   24
         Left            =   795
         TabIndex        =   49
         Top             =   3810
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Graveyard:"
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
         Index           =   23
         Left            =   795
         TabIndex        =   48
         Top             =   1695
         Width           =   930
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks:"
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
         Index           =   22
         Left            =   165
         TabIndex        =   44
         Top             =   8205
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   0
         X2              =   6120
         Y1              =   7755
         Y2              =   7755
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "110"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   21
         Left            =   2100
         TabIndex        =   43
         Top             =   7845
         Width           =   1185
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total No. Of Meal Serve"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   20
         Left            =   135
         TabIndex        =   42
         Top             =   7830
         Width           =   2100
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "110"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   19
         Left            =   2085
         TabIndex        =   41
         Top             =   7410
         Width           =   1185
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   18
         Left            =   780
         TabIndex        =   40
         Top             =   7350
         Width           =   465
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "110"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   17
         Left            =   2085
         TabIndex        =   39
         Top             =   5265
         Width           =   1185
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   16
         Left            =   180
         TabIndex        =   38
         Top             =   5265
         Width           =   465
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "110"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   15
         Left            =   2025
         TabIndex        =   37
         Top             =   3150
         Width           =   1185
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   120
         TabIndex        =   36
         Top             =   3150
         Width           =   465
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Guess"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   210
         TabIndex        =   29
         Top             =   5640
         Width           =   1155
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OJT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   165
         TabIndex        =   28
         Top             =   3495
         Width           =   375
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   120
         TabIndex        =   27
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Break Fast:"
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
         Left            =   795
         TabIndex        =   26
         Top             =   6345
         Width           =   1020
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lunch:"
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
         Left            =   795
         TabIndex        =   25
         Top             =   6690
         Width           =   585
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dinner:"
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
         Left            =   795
         TabIndex        =   24
         Top             =   7065
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Break Fast:"
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
         Left            =   795
         TabIndex        =   23
         Top             =   4185
         Width           =   1020
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lunch:"
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
         Left            =   795
         TabIndex        =   22
         Top             =   4560
         Width           =   585
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dinner:"
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
         Left            =   795
         TabIndex        =   21
         Top             =   4935
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Break Fast:"
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
         Left            =   795
         TabIndex        =   20
         Top             =   2085
         Width           =   1020
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lunch:"
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
         Left            =   795
         TabIndex        =   19
         Top             =   2460
         Width           =   585
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dinner:"
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
         Left            =   795
         TabIndex        =   18
         Top             =   2835
         Width           =   615
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   0
         X2              =   6120
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   0
         X2              =   6120
         Y1              =   5610
         Y2              =   5610
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   0
         X2              =   6120
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   360
         Left            =   1290
         Tag             =   "et0;ht2"
         Top             =   315
         Width           =   2430
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans No:"
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
         Left            =   165
         TabIndex        =   16
         Top             =   270
         Width           =   840
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
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
         Left            =   150
         TabIndex        =   15
         Top             =   960
         Width           =   465
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   1125
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
      Picture         =   "frmMealPosting.frx":0003
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   120
      TabIndex        =   31
      Top             =   510
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
      Picture         =   "frmMealPosting.frx":077D
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   120
      TabIndex        =   32
      Top             =   3585
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
      Picture         =   "frmMealPosting.frx":0EF7
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   1740
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
      Picture         =   "frmMealPosting.frx":1671
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   120
      TabIndex        =   34
      Top             =   1125
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
      Picture         =   "frmMealPosting.frx":1DEB
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   35
      Top             =   510
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
      Picture         =   "frmMealPosting.frx":2565
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   120
      TabIndex        =   45
      Top             =   2970
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Void"
      AccessKey       =   "V"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMealPosting.frx":2CDF
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   120
      TabIndex        =   46
      Top             =   2355
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Post"
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
      Picture         =   "frmMealPosting.frx":3459
   End
End
Attribute VB_Name = "frmMealPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmMealPosting"

Private oSkin As clsFormSkin
Private oDriver As clsFormDriver
Private WithEvents loTrans As clsMealSummary
Attribute loTrans.VB_VarHelpID = -1
Private p_oResult As Recordset

Private Sub ClearFields()
   Dim loTxt As TextBox
      For Each loTxt In txtField
         loTxt = ""
      Next
End Sub

Private Function isEntryOk() As Boolean
   If lblField(21).Caption = CInt("0") Then
      MsgBox "Unable to save zero total meal voucher serve!!!" & vbCrLf & _
               "Pls Verify Entry Then Try Again!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      GoTo EntryNotOK
   End If

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub sendOtherIfo()
   Dim lnCtr As Integer
   For lnCtr = 2 To 5
      Select Case lnCtr
      Case 2
         loTrans.Master(lnCtr) = CInt(txtField(lnCtr).Text)
      Case 3
         loTrans.Master(lnCtr) = CInt(txtField(lnCtr).Text)
      Case 4
         loTrans.Master(lnCtr) = CInt(txtField(lnCtr).Text)
      Case 5
         loTrans.Master(lnCtr) = CInt(txtField(lnCtr).Text)
      End Select
   Next
End Sub

Private Sub setTransTat(lnStat As String)
   Select Case lnStat
   Case "0"
      lblStatus.Caption = "OPEN"
   Case "1"
      lblStatus.Caption = "CLOSED"
   Case "2"
      lblStatus.Caption = "POSTED"
   Case "3"
      lblStatus.Caption = "VOIDED"
   Case Else
      lblStatus.Caption = "UNKNOWN"
   End Select
End Sub

Private Sub EmptyResult()
   lblField(15).Caption = 0
   lblField(17).Caption = 0
   lblField(19).Caption = 0
   lblField(21).Caption = 0
   txtField(2).Text = 0
   txtField(3).Text = 0
   txtField(4).Text = 0
   txtField(5).Text = 0
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
      Case 4 'close
            Unload Me
      Case 2 'Post
         If isEntryOk = True Then
            sendOtherIfo
            If loTrans.SaveTransaction = True Then
               MsgBox "Transaction save successfully!", vbInformation, "Meal Posting Confirmation"
               If MsgBox("Do you want to post this transaction?", vbInformation + vbYesNo, "Meal Posting Confirmation") = vbYes Then
                  If Not loTrans.PostTransaction Then
                   MsgBox "Unable to post transaction...", vbCritical, "Warning"
                  End If
                  MsgBox "Transaction posted successfully!", vbInformation, "Meal Posting Confirmation"
                End If
                If loTrans.NewTransaction = True Then
                  InitForm 1
                  InitGrid
                  EmptyResult
               End If
            Else
               MsgBox "Transaction unable to save!", vbCritical, "Meal Posting Confirmation"
            End If
         End If
      Case 3 ' New
         If loTrans.NewTransaction = True Then
         InitForm 1
         InitGrid
         EmptyResult
         End If
      Case 0 ' Cancel
         If MsgBox("Do you really want to cancel update?", vbInformation + vbYesNo, "Confirmation") = vbYes Then
            InitForm 0
            InitGrid
            EmptyResult
            setTransTat ("-1")
         End If
      Case 5 ' Update
         If loTrans.Master("cTranStat") <> "0" Then
            MsgBox "Unable to update transaction..." & vbCrLf & _
                "Transaction already " + lblStatus.Caption + "!!!", vbCritical, "Error"
            Exit Sub
         End If
         If loTrans.UpdateTransaction = True Then
            InitForm 1
         End If
      Case 1 ' Browse
         If loTrans.SearchTransaction = True Then
            InitForm 0
         End If
      Case 7 'post
            If MsgBox("Do you want to post this transaction?", vbInformation + vbYesNo, "Meal Posting Confirmation") = vbYes Then
               If loTrans.Master("cTranStat") <> "0" Then
                 MsgBox "Unable to post transaction..." & vbCrLf & _
                     "Transaction already " + lblStatus.Caption + "!!!", vbCritical, "Error"
                  Exit Sub
               End If
               If loTrans.PostTransaction Then
                  MsgBox "Transaction posted successfully!", vbInformation, "Meal Posting Confirmation"
                  loTrans.OpenTransaction (loTrans.Master("sTransNox"))
                  InitForm 0
               End If
             End If
      Case 6 'void
            If MsgBox("Do you want to void this transaction?", vbInformation + vbYesNo, "Meal Posting Confirmation") = vbYes Then
               If loTrans.Master("cTranStat") <> "0" Then
                 MsgBox "Unable to void transaction..." & vbCrLf & _
                     "Transaction already " + lblStatus.Caption + "!!!", vbCritical, "Error"
                  Exit Sub
               End If
               If loTrans.CancelTransaction Then
                  MsgBox "Transaction voided successfully!", vbInformation, "Meal Posting Confirmation"
                  loTrans.OpenTransaction (loTrans.Master("sTransNox"))
                  InitForm 0
               End If
             End If
   End Select
End Sub

Private Sub Form_Activate()
   EmptyResult
   InitForm 1
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

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me

   Set loTrans = New clsMealSummary
   Set loTrans.AppDriver = oApp

   loTrans.Branch = oApp.BranchCode
   loTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   InitGrid
   loTrans.InitTransaction
   loTrans.NewTransaction

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Cols = 4
      .Rows = 8

      .Clear

      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next

      .Row = 1

      'column title
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Employee Name"
      .TextMatrix(0, 2) = "Voucher"
      .TextMatrix(0, 3) = "Meal Type"

      'column width
      .ColWidth(0) = 400
      .ColWidth(1) = 2900
      .ColWidth(2) = 1400
      .ColWidth(3) = 1240

      'column allinment
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter

      'set location
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub InitForm(ByVal fnEdit As Integer)
   Dim lnCtr As Integer
   Dim loTxt As TextBox

   xrFrame1.Enabled = Not (fnEdit = 0)
   xrFrame4.Enabled = Not (fnEdit = 0)

   cmdButton(0).Visible = Not (fnEdit = 0)
   cmdButton(2).Visible = Not (fnEdit = 0)

   cmdButton(1).Visible = (fnEdit = 0)
   cmdButton(3).Visible = (fnEdit = 0)
   cmdButton(4).Visible = (fnEdit = 0)
   cmdButton(5).Visible = (fnEdit = 0)
   cmdButton(6).Visible = (fnEdit = 0)
   cmdButton(7).Visible = (fnEdit = 0)

   For Each loTxt In txtField
      loTxt = ""
   Next

   If Not fnEdit = 0 Then
      InitTransaction
   Else
      LoadData
   End If
      setTransTat (loTrans.Master("cTranStat"))
End Sub

Private Sub InitTransaction()
Dim lnCtr As Integer
For lnCtr = 0 To 14
   Select Case lnCtr
   Case 0
      txtField(lnCtr).Text = loTrans.Master(lnCtr)
   Case 1
      txtField(lnCtr).Text = strLongDate(loTrans.Master(lnCtr))
   Case Else
      txtField(lnCtr).Text = loTrans.Master(lnCtr)
   End Select
Next

End Sub

Private Sub LoadData()
  Dim lnCtr As Integer

   For lnCtr = 0 To 14
      Select Case lnCtr
         Case 1
            txtField(lnCtr) = strLongDate(loTrans.Master(lnCtr))
         Case Else
            txtField(lnCtr) = IFNull(loTrans.Master(lnCtr), "")
      End Select
   Next

   ShowResult
   lblField(17).Caption = CInt(txtField(6).Text) + CInt(txtField(7).Text) + CInt(txtField(8).Text) + CInt(txtField(9).Text)
   lblField(19).Caption = CInt(txtField(10).Text) + CInt(txtField(11).Text) + CInt(txtField(12).Text) + CInt(txtField(13).Text)
   lblField(21).Caption = CInt(lblField(15).Caption) + CInt(lblField(17).Caption) + CInt(lblField(19).Caption)

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

Private Sub loTrans_MasterRetrieved(ByVal Index As Integer)
   Dim lnCtr As Integer

   For lnCtr = 0 To 14
      Select Case lnCtr
         Case Else
         txtField(lnCtr) = loTrans.Master(lnCtr)
      End Select
   Next

End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
      Case 1
         txtField(Index) = strShortDate(loTrans.Master(Index))
   End Select

   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

   oDriver.ColumnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
      Case 1
         txtField(Index) = strLongDate(txtField(Index).Text)
   End Select

   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With

   lblField(17).Caption = CInt(txtField(6).Text) + CInt(txtField(7).Text) + CInt(txtField(8).Text) + CInt(txtField(9).Text)
   lblField(19).Caption = CInt(txtField(10).Text) + CInt(txtField(11).Text) + CInt(txtField(12).Text) + CInt(txtField(13).Text)
   lblField(21).Caption = CInt(lblField(15).Caption) + CInt(lblField(17).Caption) + CInt(lblField(19).Caption)

End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
With txtField(Index)
      Select Case Index
         Case 1
            If Not IsDate(.Text) Then .Text = strLongDate(oApp.ServerDate)
            loTrans.Master(Index) = .Text
            ShowResult
         Case 14
            .Text = TitleCase(.Text)
            loTrans.Master(Index) = .Text
         Case Else
            If Not IsNumeric(.Text) Then .Text = 0
            loTrans.Master(Index) = CInt(.Text)
         End Select
   End With
End Sub

Private Sub ShowResult()
   Dim lnCtr As Integer
   Set p_oResult = loTrans.loadEmployee(loTrans.Master(1))
   If p_oResult.EOF Then
         InitGrid
         EmptyResult
         GoTo endProc
   End If
   With MSFlexGrid1
      p_oResult.MoveFirst
      For lnCtr = 0 To p_oResult.RecordCount - 1
         .Rows = .Rows + 1
         .TextMatrix(lnCtr + 1, 0) = Format(lnCtr + 1, "00")
         .TextMatrix(lnCtr + 1, 1) = p_oResult("sCompnyNm")
         .TextMatrix(lnCtr + 1, 2) = IFNull(p_oResult("sVouchrNo"))
         .TextMatrix(lnCtr + 1, 3) = p_oResult("sMealDesc")
          p_oResult.MoveNext
      Next
      .Rows = .Rows - 1
   End With

   p_oResult.Filter = "sMealCode = " & strParm("04")
   txtField(2) = p_oResult.RecordCount

   p_oResult.Filter = adFilterNone
   p_oResult.Filter = "sMealCode = " & strParm("01")
   txtField(3) = p_oResult.RecordCount

   p_oResult.Filter = adFilterNone
   p_oResult.Filter = "sMealCode = " & strParm("03")
   txtField(4) = p_oResult.RecordCount

   p_oResult.Filter = adFilterNone
   p_oResult.Filter = "sMealCode = " & strParm("02")
   txtField(5) = p_oResult.RecordCount

   p_oResult.Filter = adFilterNone
   lblField(15).Caption = p_oResult.RecordCount
   lblField(21).Caption = p_oResult.RecordCount

endProc:
   Exit Sub
End Sub
