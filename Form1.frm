VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "PhisSim 0.1    Simulation Paused"
   ClientHeight    =   12360
   ClientLeft      =   225
   ClientTop       =   615
   ClientWidth     =   15510
   ControlBox      =   0   'False
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12360
   ScaleWidth      =   15510
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3240
      Top             =   600
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   13095
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         Top             =   0
         Width           =   975
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         Top             =   0
         Width           =   975
      End
      Begin VB.Label BtnP 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Stop III"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   10
         Width           =   975
      End
      Begin VB.Label Btng 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004000&
         Caption         =   "Run >>>"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   10
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   9120
         Top             =   0
         Width           =   1815
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   7080
         Top             =   0
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   5040
         Top             =   0
         Width           =   1815
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   11160
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Help"
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
         Left            =   11160
         TabIndex        =   7
         Top             =   30
         Width           =   1815
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Show Menu"
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
         Left            =   0
         TabIndex        =   6
         Top             =   30
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Reset Model"
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
         Left            =   5040
         TabIndex        =   8
         Top             =   30
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Edit Model"
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
         Left            =   7080
         TabIndex        =   9
         Top             =   30
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Exit"
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
         Left            =   9120
         TabIndex        =   10
         Top             =   30
         Width           =   1815
      End
   End
   Begin VB.Frame MenuF 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "OCR-B 10 BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4500
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox Cnodeind 
         BackColor       =   &H00000000&
         Caption         =   "Show Node index"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Cnode 
         BackColor       =   &H00000000&
         Caption         =   "Show Nodes"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Cstress 
         BackColor       =   &H00000000&
         Caption         =   "Show Stress"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00000000&
         Caption         =   "Vrey Slow"
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
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "Slow"
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
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Normal"
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
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Fast"
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
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label SimT 
         BackColor       =   &H00000000&
         Caption         =   "0 s"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Sim Time:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label StatDisp 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   1335
         Left            =   75
         Top             =   195
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000008&
         Caption         =   "Simulation Speed:"
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
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4920
      Top             =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btng_Click()
Timer1.Enabled = True
Btng.BackColor = &HC000&
BtnP.BackColor = &H40&
Form1.Caption = "PhisSim 0.1    Simulation Running"
End Sub


Private Sub BtnP_Click()
Timer1.Enabled = False
Btng.BackColor = &H4000&
BtnP.BackColor = &HC0&
Form1.Caption = "PhisSim 0.1    Simulation Stoped"
End Sub

Private Sub Command1_Click()
Timer1.Interval = 10
End Sub

Private Sub Command2_Click()
Timer1.Interval = 50
End Sub

Private Sub Command3_Click()
Timer1.Interval = 1
End Sub

Private Sub Command4_Click()
Timer1.Interval = 0
End Sub

Private Sub Command5_Click()
Timer1.Interval = 100
End Sub

Private Sub Form_Load()
UpdateStat
Render Form1
Timer1.Interval = 10
MenuF.Height = 0
End Sub



Private Sub Label1_Click()
SimTime = 0
TrasferModelToSim
Form1.Cls
Render Form1
End Sub

Private Sub Label2_Click()
SimTime = 0
BtnP_Click
Form2.Timer1.Enabled = True
Form2.Show
Form1.Hide
End Sub

Private Sub Label3_Click()

Unload Form1
Unload Form2
Unload Help
End
End Sub

Private Sub Label4_Click()
Timer2.Enabled = True
End Sub

Private Sub Label5_Click()
Help.Show
End Sub

Private Sub Option1_Click()
Timer1.Interval = 1
End Sub

Private Sub Option2_Click()
Timer1.Interval = 10
End Sub

Private Sub Option3_Click()
Timer1.Interval = 50
End Sub

Private Sub Option4_Click()
Timer1.Interval = 100
End Sub

Private Sub Timer1_Timer()
UpdateStat
SimulateFrame
If Timer1.Interval = 1 Then SimulateFrame
Form1.Cls
'DebugPrint
Render Form1
End Sub

Private Sub Timer2_Timer()
If Label4 = "Hide Menu" Then
    MenuF.Height = MenuF.Height - 450
    If MenuF.Height = 15 Then
        Timer2.Enabled = False
        MenuF.Visible = False
        Label4 = "Show Menu"
    End If
Else
    MenuF.Height = MenuF.Height + 450
    MenuF.Visible = True
    If MenuF.Height > 4500 Then
        Timer2.Enabled = False
        Label4 = "Hide Menu"
    End If
End If
End Sub
