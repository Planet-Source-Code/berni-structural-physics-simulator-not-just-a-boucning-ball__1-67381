VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Model editor"
   ClientHeight    =   12270
   ClientLeft      =   405
   ClientTop       =   615
   ClientWidth     =   15420
   ControlBox      =   0   'False
   DrawWidth       =   3
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   MousePointer    =   2  'Cross
   ScaleHeight     =   12270
   ScaleWidth      =   15420
   Begin VB.Frame FpropE 
      BackColor       =   &H00000000&
      Caption         =   "Propertys>"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   51
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "No Selection"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame FpropN 
      BackColor       =   &H00000000&
      Caption         =   "Propertys>"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   43
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox PropMass 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   46
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox PropBounce 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   45
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox PropLock 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   44
         Top             =   600
         Width           =   255
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000000&
         Caption         =   "Mass"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Bouce"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Locked"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   975
      End
      Begin VB.Label PropNodeA 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   975
         Width           =   1815
      End
   End
   Begin VB.Frame FpropL 
      BackColor       =   &H00000000&
      Caption         =   "Propertys>"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox PropFlex 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   37
         Text            =   "10"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox PropStrenth 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   36
         Text            =   "10000"
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox PropNoBreak 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check2"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1680
         MaskColor       =   &H00FFFF00&
         TabIndex        =   35
         Top             =   960
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox PropRope 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check2"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1680
         MaskColor       =   &H00FFFF00&
         TabIndex        =   34
         Top             =   1320
         Width           =   255
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label PropLinkA 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1695
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "Flex"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "Strenth"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "Unbreakable"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Rope"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1320
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   3840
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "PhisSim Model Files (*.phs)|*.phs"
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Enviorment>"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2280
      MousePointer    =   1  'Arrow
      TabIndex        =   26
      Top             =   120
      Width           =   2175
      Begin VB.TextBox Tgrav 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   29
         Text            =   "3"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Tair 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   28
         Text            =   "0.99"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Gravity"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "AirDensity"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Timer MenuTimerO 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   2520
   End
   Begin VB.Timer MenuTimerC 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   2520
   End
   Begin VB.Frame MenuF 
      BackColor       =   &H00000000&
      Caption         =   "Menu>"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1300
      Left            =   4560
      MousePointer    =   1  'Arrow
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2160
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2160
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label BTNclear 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "OCR-A BT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label BTNload 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   " Load model"
         BeginProperty Font 
            Name            =   "OCR-A BT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label BTNsave 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Save model"
         BeginProperty Font 
            Name            =   "OCR-A BT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label BTNbtsim 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   " Back to sim"
         BeginProperty Font 
            Name            =   "OCR-A BT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   4560
      Top             =   2520
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Tools>"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   14
      Top             =   120
      Width           =   2055
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         Top             =   720
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label BTNaddnode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add Node"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label BTNaddlink 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add Link"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label BTNmove 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Move"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label BTNdel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Frame AddLinkF 
      BackColor       =   &H00000000&
      Caption         =   "Add Link>"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CheckBox Crope 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check2"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1680
         MaskColor       =   &H00FFFF00&
         TabIndex        =   32
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Cnobreak 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check2"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1680
         MaskColor       =   &H00FFFF00&
         TabIndex        =   13
         Top             =   960
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox Tbreak 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Text            =   "10000"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Tflex 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Text            =   "10"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Rope"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Unbreakable"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Strenth"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Flex"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame AddNodeF 
      BackColor       =   &H00000000&
      Caption         =   "Add Node>"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
      Begin VB.CheckBox Clock 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Tmass 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Tbouce 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Locked"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Mass"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Bouce"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   375
      Left            =   4560
      TabIndex        =   24
      Top             =   120
      Width           =   1095
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const DragColor = &HFFFFFF    ' Moseover button color
Const NormColor = &HC0C0C0    ' Normal button color
Const SelBackColor = &H808080
Const UnSelBackColor = &H404040


Dim SelNode As Integer
Dim a As Integer
Dim SelE
Dim Mode As Byte '0 = Add Node  1 = Add Link   2 = Del   4 = Move

Private Sub BTNaddlink_Click()
AddLinkF.Visible = True
AddNodeF.Visible = False
FpropE.Visible = False
FpropN.Visible = False
FpropL.Visible = False
BTNaddnode.BackColor = UnSelBackColor
BTNdel.BackColor = UnSelBackColor
BTNaddlink.BackColor = SelBackColor
BTNmove.BackColor = UnSelBackColor

Mode = 1
LinkHandleDrawEnable False
End Sub

Private Sub BTNaddlink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BTNaddlink.ForeColor = DragColor
BTNaddnode.ForeColor = NormColor
BTNdel.ForeColor = NormColor
BTNmove.ForeColor = NormColor
End Sub

Private Sub BTNaddnode_Click()
SelectNode 0, 0
AddNodeF.Visible = True
AddLinkF.Visible = False
FpropE.Visible = False
FpropN.Visible = False
FpropL.Visible = False
BTNaddnode.BackColor = SelBackColor
BTNdel.BackColor = UnSelBackColor
BTNaddlink.BackColor = UnSelBackColor
BTNmove.BackColor = UnSelBackColor
Mode = 0
LinkHandleDrawEnable False
End Sub

Private Sub BTNaddnode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BTNaddnode.ForeColor = DragColor
BTNaddlink.ForeColor = NormColor
BTNdel.ForeColor = NormColor
BTNmove.ForeColor = NormColor
End Sub

Private Sub BTNbtsim_Click()
TrasferModelToSim
Timer1.Enabled = False
UpdateStat
Form1.Show
Form2.Hide
SetEnviroment Tair, Tgrav
Form1.Cls
Render Form1
'SaveModelEdit App.Path & "\LastEdited.phs"
End Sub

Private Sub BTNclear_Click()
ClearModel
End Sub

Private Sub BTNdel_Click()
AddLinkF.Visible = False
AddNodeF.Visible = False
FpropE.Visible = False
FpropN.Visible = False
FpropL.Visible = False
BTNaddnode.BackColor = UnSelBackColor
BTNdel.BackColor = SelBackColor
BTNaddlink.BackColor = UnSelBackColor
BTNmove.BackColor = UnSelBackColor

Mode = 3
LinkHandleDrawEnable True

End Sub

Private Sub BTNdel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BTNdel.ForeColor = DragColor
BTNaddnode.ForeColor = NormColor
BTNaddlink.ForeColor = NormColor
BTNmove.ForeColor = NormColor
End Sub

Private Sub BTNload_Click()
On Error GoTo konec
FileDialog.ShowOpen
ClearModel
LoadModelEdit FileDialog.FileName
konec:
End Sub

Private Sub BTNmove_Click()
AddLinkF.Visible = False
AddNodeF.Visible = False
BTNaddnode.BackColor = UnSelBackColor
BTNdel.BackColor = UnSelBackColor
BTNaddlink.BackColor = UnSelBackColor
BTNmove.BackColor = SelBackColor

Mode = 4
LinkHandleDrawEnable True

FpropE.Visible = True
FpropN.Visible = False
FpropL.Visible = False
End Sub

Private Sub BTNmove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BTNmove.ForeColor = DragColor
BTNaddnode.ForeColor = NormColor
BTNdel.ForeColor = NormColor
BTNaddlink.ForeColor = NormColor
End Sub

Private Sub BTNsave_Click()
On Error GoTo konec
FileDialog.ShowSave
SaveModelEdit FileDialog.FileName
konec:
End Sub


Private Sub Form_Load()
SelectNode 0, 0
ReLenLinkEnable True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Val(PropFlex) < 10 Then PropFlex = 10
If Val(Tflex) < 10 Then Tflex = 10
Select Case Mode
Case 0: 'Add node
    AddNode X, Y, Tmass, Tbouce, Clock.Value
    LinkDrawEnable False
Case 1: 'Add Link part1
    SelNode = SelectNode(X, Y)
    LinkDrawEnable True
    If SelNode < 1001 Then Mode = 2
Case 2: 'Add Link part2
    SelNodeN = SelectNode(X, Y)
    If SelNodeN > 1001 Then AddNode X, Y, 1, 1, False
    SelNodeN = SelectNode(X, Y)
        AddLink SelNode, SelNodeN, Tflex, Tbreak, Cnobreak, Crope
        Mode = 1
        LinkDrawEnable False
    SelectNode 0, 0
Case 3: 'Delete
    DeleteElement X, Y
    LinkDrawEnable False
Case 4: 'Move
LinkDrawEnable False
    If SelectNode(X, Y) < 1000 Then
        t = SelectNode(X, Y)
        SelE = t
        StartMov t
        PropMass = GetNodeDat(t).mass
        PropBounce = GetNodeDat(t).Bouce
        PropLock.Value = OneOrZero(GetNodeDat(t).locked)
        FpropN.Visible = True
        FpropL.Visible = False
        FpropE.Visible = False
        ReLenLinkEnable True
        If Button = 2 Then ReLenLinkEnable False
    ElseIf GetLinkHandle(X, Y) < 1000 Then
        t = GetLinkHandle(X, Y)
        SelE = t
        PropStrenth = GetLinkDat(t).breakpoint
        PropFlex = GetLinkDat(t).flex
        PropRope.Value = OneOrZero(GetLinkDat(t).rope)
        PropNoBreak.Value = OneOrZero(GetLinkDat(t).Indestuctable)
        FpropL.Visible = True
        FpropN.Visible = False
        FpropE.Visible = False
    Else
        FpropE.Visible = True
        FpropN.Visible = False
        FpropL.Visible = False
    End If
End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BTNaddnode.ForeColor = NormColor
BTNdel.ForeColor = NormColor
BTNaddlink.ForeColor = NormColor
BTNmove.ForeColor = NormColor

UpdateCursor X, Y

If MenuF.Visible = True And MenuTimerC.Enabled = False Then
MenuTimerC.Enabled = True
a = 13
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
EndMov
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BTNaddnode.ForeColor = NormColor
BTNdel.ForeColor = NormColor
BTNaddlink.ForeColor = NormColor
BTNmove.ForeColor = NormColor
End Sub


Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MenuTimerO.Enabled = True
MenuF.Height = 0
MenuF.Visible = True
a = 0
End Sub

Private Sub MenuTimerC_Timer()
On Error Resume Next
a = a - 1
MenuF.Height = a * 100
If a < 1 Then
MenuTimerC.Enabled = False
MenuF.Visible = False
End If
End Sub

Private Sub MenuTimerO_Timer()
a = a + 1
MenuF.Height = a * 100 + 500
If a > 8 Then MenuTimerO.Enabled = False
End Sub

Private Sub PropLinkA_Click()
If Val(PropFlex) < 10 Then PropFlex = 10
If Val(Tflex) < 10 Then Tflex = 10
SetLinkDat SelE, PropFlex, PropStrenth, TrueOrFalse(PropNoBreak.Value), TrueOrFalse(PropRope.Value)
End Sub

Private Sub PropNodeA_Click()
SetNodeDat SelE, PropMass, PropBounce, TrueOrFalse(PropLock)
End Sub


Private Sub Tair_LostFocus()
If Val(Tair) > 1 Then Tair = 1
If Val(Tair) < 0.001 Then Tair = 0
End Sub

Private Sub Timer1_Timer()
RenderEdit Form2
End Sub

