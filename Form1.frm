VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SBYC TIME KEEPING UTILITY (-FOR VERIFICATION-)"
   ClientHeight    =   12165
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11835
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12165
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   52
      Top             =   600
      Width           =   5415
      Begin VB.OptionButton Emp_status 
         Appearance      =   0  'Flat
         Caption         =   "BREAK IN"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   56
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Emp_status 
         Appearance      =   0  'Flat
         Caption         =   "BREAK OUT"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   55
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Emp_status 
         Appearance      =   0  'Flat
         Caption         =   "TIME IN"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Emp_status 
         Appearance      =   0  'Flat
         Caption         =   "TIME OUT"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   53
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   50
      Top             =   11895
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame frerror 
      Height          =   2415
      Left            =   2520
      TabIndex        =   8
      Top             =   4440
      Width           =   6615
      Begin VB.CommandButton abort 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Abort"
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdok 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&OK"
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblerror 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ERROR: SYSTEM ERROR DETECTED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Ok"
      Height          =   600
      Left            =   11880
      TabIndex        =   1
      Top             =   8880
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SWITCH"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2760
      Width           =   855
   End
   Begin VB.Timer ltimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10440
      Top             =   5280
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "T"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8880
      Width           =   615
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   120
      TabIndex        =   43
      Top             =   9480
      Width           =   8175
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Time Keeping V 5.01"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1800
         TabIndex        =   46
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   120
      TabIndex        =   33
      Top             =   8760
      Width           =   10935
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sunday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   9480
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saturday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   7920
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Friday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   6360
         TabIndex        =   38
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Thursday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wednesday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tuesday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer timex 
      Interval        =   1000
      Left            =   10440
      Top             =   4080
   End
   Begin VB.TextBox txtdpt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   7560
      TabIndex        =   25
      Text            =   "0000"
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "EVENING:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5160
      TabIndex        =   19
      Top             =   8760
      Width           =   2415
      Begin VB.OptionButton OPTDL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIMEOUT"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OPTDL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIMEIN"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "AFTERNOON:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   2640
      TabIndex        =   18
      Top             =   8760
      Width           =   2415
      Begin VB.OptionButton OPTDL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIMEOUT"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OPTDL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIMEIN"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "MORNING:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   8760
      Width           =   2415
      Begin VB.OptionButton OPTDL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIMEIN "
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OPTDL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIMEOUT"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtEMP_NO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Text            =   "0000"
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CheckBox SWITCH 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "SCANNER ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9480
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1935
      Left            =   7200
      ScaleHeight     =   1875
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Frame frameinfo 
      Height          =   3855
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   11415
      Begin VB.Timer vtime 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   10200
         Top             =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Employee No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label EID 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbln1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   2880
         TabIndex        =   42
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label lblcaps 
         Appearance      =   0  'Flat
         Caption         =   "Last Known Logged Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   41
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label lblcaps 
         Appearance      =   0  'Flat
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   32
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblcaps 
         Appearance      =   0  'Flat
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   31
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblcaps 
         Appearance      =   0  'Flat
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   30
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lbln1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   2880
         TabIndex        =   29
         Top             =   2280
         Width           =   5655
      End
      Begin VB.Label lbln1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   5880
         TabIndex        =   28
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label lbln1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2880
         TabIndex        =   27
         Top             =   1320
         Width           =   2775
      End
   End
   Begin MSComctlLib.ListView viewer 
      Height          =   1215
      Left            =   120
      TabIndex        =   51
      Top             =   10200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   2143
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIMEKEEPER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   11895
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblf 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   49
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6720
      MouseIcon       =   "Form1.frx":1472
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":177C
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "ID No:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   24
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblID 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee No: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label display 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   7800
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CURRENT TIME IS:"
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
      Height          =   255
      Index           =   1
      Left            =   8880
      TabIndex        =   5
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CURRENT DATE IS:"
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
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   4
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Label lbltime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Image Image3 
      Height          =   1500
      Left            =   120
      Picture         =   "Form1.frx":26D2
      Stretch         =   -1  'True
      Top             =   -1680
      Width           =   7500
   End
   Begin VB.Image Image4 
      Height          =   9735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   -960
      Width           =   11895
   End
   Begin VB.Menu APP 
      Caption         =   "Application"
      Begin VB.Menu SO 
         Caption         =   "Scanner On"
      End
      Begin VB.Menu LE 
         Caption         =   "Login Employees"
      End
      Begin VB.Menu LOE 
         Caption         =   "Logout Employees"
      End
      Begin VB.Menu rpt 
         Caption         =   "Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu EX 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Vw 
      Caption         =   "View"
      Begin VB.Menu MOE 
         Caption         =   "View MOE"
      End
      Begin VB.Menu VEL 
         Caption         =   "View Entry List"
      End
      Begin VB.Menu CEL 
         Caption         =   "View Information"
      End
      Begin VB.Menu EME 
         Caption         =   "Enable Manual Entry"
      End
      Begin VB.Menu DME 
         Caption         =   "Disable Manual Entry"
      End
      Begin VB.Menu Srch 
         Caption         =   "Search"
      End
   End
   Begin VB.Menu DEV 
      Caption         =   "Device"
      Begin VB.Menu DD 
         Caption         =   "&Detect Device"
      End
   End
   Begin VB.Menu Hlp 
      Caption         =   "&Help"
      Begin VB.Menu Inf 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbx As New clsDBCode
Dim data As New clsDBCode
Private Const MB_RETRYCANCEL = &H5&
Private Const MB_ICONEXCLAMATION = &H30&
Private Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long

Const MON = 1
Const TUE = 2
Const WED = 3
Const THU = 4
Const FRI = 5
Const SAT = 6
Const SUN = 7

Const XIN = "IN"
Const XOUT = "OUT"

Const GENQRY = "SELECT * FROM DB_FINGER_PRINT INNER JOIN TK_TIMEKEEPING ON DB_FINGER_PRINT.EMPNO = TK_TIMEKEEPING.empID"
Const GENQRY2 = "SELECT * FROM DB_FINGER_PRINT INNER JOIN TK_TIMEKEEPING ON DB_FINGER_PRINT.EMPNO = TK_TIMEKEEPING.EMPID"
Const UNIONQRY = "SELECT * FROM DB_FINGER_PRINT UNION ALL SELECT * FROM TK_TIMEKEEPING"
Const IVAL = 1


Const EIN = 1
Const EOUT = 0

Dim lv As ListItem
Dim LAPSECOUNT As Integer
Dim TDOWN, TUP
Dim G_FNAME As String
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim data_source, uid, pwd
Dim dbuser As New DPUsersDB
Dim Usr As DPUser

Dim ACTIVE As Boolean

Dim fptemp As FPTemplate
Dim WithEvents FPGTEMP As FPGetTemplate
Attribute FPGTEMP.VB_VarHelpID = -1
Dim WithEvents FPDEVICES As FPDEVICES
Attribute FPDEVICES.VB_VarHelpID = -1
Dim WithEvents fpdev As FPDevice
Attribute fpdev.VB_VarHelpID = -1
Dim FPver As New FPVerify

Dim id As Variant

Private Sub abort_Click()

modular.XEND = 0
    abort.Visible = False
    frerror.Visible = False
    lblerror.Alignment = 0
    Exit Sub
    
End Sub

Private Sub CEL_Click()
frameinfo.Visible = True
End Sub

Private Sub Check1_Click()


If Check1.value = 1 Then

  frmlogin.Show 1
  
End If

'If Check1.value = 1 Then
 '   Check1.value = 0
'End If


End Sub

Private Sub cmdok_Click()

If modular.XEND = 1 Then
    Set FPGTEMP = Nothing
    Set fptemp = Nothing
    Set FPDEVICES = Nothing
    Set FPver = Nothing
    Set db = Nothing
    Set rs = Nothing
    End
Else
    modular.XEND = 0
    frerror.Visible = False
End If
End Sub



Private Sub Command34_Click()
Dim c As Long

c = MessageBoxEx(Main.hwnd, Err.Description, "Time Keeping", MB_RETRYCANCEL + MB_ICONEXCLAMATION, 0)

If c = 4 Then

ElseIf c = 2 Then
End
End If

End Sub

Private Sub Command1_Click()
data.MDB_Access_Connect
data.Sql_SELECT_QUERY_Execute "select top 40 * from db_finger_print"

While Not data.EOF
x = data.GetFields("empno")
If Len(x) = 1 Then x = "000" & x
If Len(x) = 2 Then x = "00" & x
    txtEMP_NO = x
    data.MoveNext
Wend
End Sub

Private Sub Command3_Click()
' Call VERIFY_IF_INOUT(3122, 1)

Dim t1 As Date
Dim t2 As Date

t1 = "13:00"
t2 = "14:15"

tx = TimeValue(t1) - TimeValue(t2)

MsgBox Hour(tx) & ":" & Minute(tx)

End Sub

Private Sub Command2_Click()
'EME_Click
txtEMP_NO.Enabled = True
End Sub

Private Sub DD_Click()
Call DETECT_DEVICE
End Sub

Private Sub DME_Click()
On Error Resume Next
txtEMP_NO.Visible = True
txtEMP_NO.Enabled = False
lblID(2).Visible = False
lblf.Alignment = 2
txtdpt.Enabled = True
txtdpt.SetFocus
End Sub

Private Sub EME_Click()
frmlogin.Show
End Sub

Private Sub EX_Click()

    modular.XEND = 1
    If modular.XEND = 1 Then
        lblerror.Alignment = 2
        frerror.Visible = True
        lblerror.Visible = True
        lblerror.Caption = "THIS WILL END YOUR SESSION IF YOU ARE SURE PRESS OK OTHERWISE PRESS ABORT"
        abort.Visible = True
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
rpt.Enabled = False '''''''''''''''

End Sub


Function Ex_Test()

Dim db1 As New clsDBCode

db1.MDB_Access_Connect


End Function

Private Sub Form_Load()

DME_Click

Dim lines As Control

For Each lines In Me.Controls

    If TypeOf lines Is Line Then
         lines.BorderWidth = 1
          ' lines.BackColor = vbBlack
          lines.BorderColor = vbBlack
    End If
    
    
Next
'dbcode.MDB_Access_Connect
'Call DISPLAY_DATA(F, D, 0)

frameinfo.Visible = True
txtdpt.MaxLength = 4
frerror.Left = 2520
frerror.Top = 4680
SWITCH.Caption = "VERIFIER OFF"
SWITCH.ForeColor = RGB(255, 0, 0)
Command1.Enabled = False
'SaveSetting VB.APP.EXEName, "CONFIG", "DAY", 3
GS = GetSetting(VB.APP.EXEName, "CONFIG", "DAY")
If GS = 1 Then
Label2(0).BackColor = RGB(255, 0, 0)
ElseIf GS = 2 Then
Label2(1).BackColor = RGB(255, 0, 0)
ElseIf GS = 3 Then
Label2(2).BackColor = RGB(255, 0, 0)
ElseIf GS = 4 Then
Label2(3).BackColor = RGB(255, 0, 0)
ElseIf GS = 5 Then
Label2(4).BackColor = RGB(255, 0, 0)
ElseIf GS = 6 Then
Label2(5).BackColor = RGB(255, 0, 0)
ElseIf GS = 7 Then
Label2(6).BackColor = RGB(255, 0, 0)
End If
With viewer
    .ColumnHeaders.Add , , "#", 400, 0
    .ColumnHeaders.Add , , "ID", 800
    .ColumnHeaders.Add , , "DATE", 1000, 0
    .ColumnHeaders.Add , , "TIME IN", 1200, 0
    .ColumnHeaders.Add , , "TIME OUT", 1200, 0
    .ColumnHeaders.Add , , "TIME IN", 1200, 0
    .ColumnHeaders.Add , , "TIME OUT", 1200, 0
    .ColumnHeaders.Add , , "TIME IN", 1200, 0
    .ColumnHeaders.Add , , "TIME OUT", 1200, 0
    .ColumnHeaders.Add , , "EMPLOYEE NAME", 2500, 0
    .ColumnHeaders.Add , , "MOE", 1500, 0
    .ColumnHeaders.Add , , "Department", 2000, 0
    .ColumnHeaders.Add , , "Total Hours", 2000, 0
End With
sbar.Panels(1).Text = Main.Caption
sbar.SimpleText = Main.Caption & Space(20) & "M I S TEAM"
abort.Visible = False
frerror.Visible = False
    Call DETECT_DEVICE
    txtEMP_NO.MaxLength = 4
    SWITCH.value = 1
    Main.Show
    viewer.Visible = True
   ' Call SetWindowPos(Main.hwnd, -1, 0, 0, 0, 0, FLAGS)

End Sub


Function DETECT_DEVICE()

On Error Resume Next

Dim DEVICE_COUNT As Integer
Dim fpdev As FPDevice
DEVICE_COUNT = 0
frerror.Visible = True
Set FPDEVICES = New FPDEVICES
If FPDEVICES.Count = 0 Then: lblerror.Caption = "NO OF DEVICE DETECTED: " & DEVICE_COUNT & vbCr & _
                                                "DEVICE ACTIVE: " & ACTIVE & vbCr & vbCr & _
                                                "CONNECT THE USB FINGER SCANNER "
For Each fpdev In FPDEVICES
    DEVICE_COUNT = FPDEVICES.Count
    If DEVICE_COUNT <> 0 Then
    frerror.Visible = True
    ACTIVE = True
    lblerror.Caption = "NO OF DEVICE DETECTED: " & DEVICE_COUNT & vbCr & _
                        "SERIAL CODE: " & fpdev.SerialNumber & vbCr & _
                        "DEVICE ACTIVE: " & ACTIVE & vbCr & _
                        "PRODUCT NAME: " & fpdev.Product
    End If
Next


End Function

Private Sub Form_LostFocus()
txtdpt.SetFocus
End Sub

Private Sub fpdev_SampleAcquired(ByVal pRawSample As Object)
MsgBox ""
End Sub

Private Sub FPDEVICES_DeviceDisconnected(ByVal serNum As String)

modular.XEND = 0
frerror.Visible = True
lblerror.Caption = "[DEVICE FINGER PRINT SCANNER HAS BEEN DISCONNECTED OR NOT CONNECTED PROPERLY]"

End Sub

Private Sub FPGTEMP_Done(ByVal pTemplate As Object)

' EVENT THAT TRIGER AFTER THE SCANNER CAPTURE THE FINGER PRINT SUCCESSFULLY

'On Error GoTo EXCEPTION_THROWN
Dim GTIME As Variant
Dim FNAME As String
Dim FBITSTR As String
Dim AUTHENTICATE As Boolean
Dim LTHUMBFP() As Byte
Dim LINDEXFP() As Byte
Dim RTHUMBFP() As Byte
Dim RINDEXFP() As Byte
Dim DT As Date


Set fptemp = New FPTemplate
Set FPver = New FPVerify
Set FPGTEMP = New FPGetTemplate

GTIME = Time

If db.State = 1 Then
    db.Close
End If

If txtdpt.Text = "" Then txtdpt.Text = "0000"
    

'dbcode.SqlConnect "DBEMPLOYMENT"
dbcode.MDB_Access_Connect

If txtdpt <> "" Then dbcode.Sql_SELECT_QUERY_Execute "SELECT * FROM DB_FINGER_PRINT WHERE EMPNO = " & txtdpt ', db, adOpenKeyset, adLockOptimistic

'If txtdptx <> "" Then dbcode.Sql_SELECT_QUERY_Execute "SELECT * FROM DB_FINGER_PRINT WHERE depcode = " & txtdptx & " ',db, adOpenKeyset, adLockOptimistic"
Do While Not dbcode.EOF
DoEvents
  
    If IsNull(dbcode.GetFields("RTHUMB")) = True Then
    Else
    RTHUMBFP = dbcode.GetFields("RTHUMB")
   ' ReDim BINARYA(0 To Len(RTHUMBFP) / 2) As Byte
   '     For I = 1 To Len(RTHUMBFP) Step 2
   '         HX = Mid(retval, I, 2)
   '         BINARYA(((I + 1) / 2) - 1) = Val("&H" + Mid(RTHUMBFP, I, 2))
   '      Next I
    End If
    
    Set fptemp = New FPTemplate
    Set FPver = New FPVerify
    fptemp.Import RTHUMBFP ' RIGHT THUMB
    FPver.Compare fptemp, pTemplate, AUTHENTICATE, SCORE, THRESHOLD, False, Sm_None
    If AUTHENTICATE = True Then
    SaveSetting VB.APP.EXEName, "CONFIG", "T", 0
    UNI_BOOL = "True right thumb" & dbcode.GetFields("FNAME")
    FNAME = dbcode.GetFields("FNAME")
    LNAME = dbcode.GetFields("LNAME")
    empname = LNAME & "," & FNAME
    id = dbcode.GetFields("EMPNO")
    EID(1).Caption = dbcode.GetFields("EMPNO")
    display.ForeColor = RGB(0, 0, 255)
    display.BackColor = RGB(255, 255, 255)
   ' display.Caption = dbcode.GetFields("FNAME") & " " & dbcode.GetFields("LNAME")
    id = dbcode.GetFields("EMPNO")
    lbln1(0) = FNAME
    lbln1(1) = LNAME
    lbln1(2) = dbcode.GetFields("deptname")
    EID_X = id
   ' method = "SCANNER"
    'mee = 1
'==============================================================================================RIGHT THUMB VALIDATION
            
' If txtdpt <> "" Then
'     Call DATA_ENTRY(EID_X)
' End If

'open vb.App.Path + "\logentry.txt"

If Dir(VB.APP.Path + "\logentry.txt") <> "" Then
Open VB.APP.Path + "\logentry.txt" For Append As #1
Print #1, Now, empname
Close #1
Else
Open VB.APP.Path + "\logentry.txt" For Output As #1
Print #1, Now, empname
Close #1
End If

' PASS THE VALUE OF ID TO TEXTEMPNO OBJECT TO TRIGER THE ENTRY EVENT
' This is the method i prefer rather than creating a whole new function with the same code
txtEMP_NO = id

'===============================================================================================RIGHT THUMB VALIDATION
        FPGTEMP.Run
        Exit Do
    ElseIf AUTHENTICATE = False Then
    display.ForeColor = RGB(255, 0, 0)
    display.BackColor = vbWhite
    display.Caption = "ENTRY DENIED"
    EID(1).Caption = ""
    End If
 '===================================================EOF RIGHT THUMB============================

        If IsNull(dbcode.GetFields("RINDEX")) = True Then
        Else
            RINDEXFP = dbcode.GetFields("RINDEX")
           ' ReDim BINARYB(0 To Len(RINDEXFP) / 2) As Byte
           ' For I = 1 To Len(RINDEXFP) Step 2
           '     HX = Mid(retval, I, 2)
           '     BINARYB(((I + 1) / 2) - 1) = Val("&H" + Mid(RINDEXFP, I, 2))
           ' Next I
        End If


    Set fptemp = New FPTemplate
    Set FPver = New FPVerify
    fptemp.Import RINDEXFP ' RIGHT INDEX
    FPver.Compare fptemp, pTemplate, AUTHENTICATE, SCORE, THRESHOLD, False, Sm_None
    If AUTHENTICATE = True Then
    UNI_BOOL = "True RIGHT INDEX" & dbcode.GetFields("FNAME")
    FNAME = dbcode.GetFields("FNAME")
    LNAME = dbcode.GetFields("LNAME")
     id = dbcode.GetFields("EMPNO")
     empname = LNAME & "," & FNAME
    EID(1).Caption = dbcode.GetFields("EMPNO")
        display.ForeColor = RGB(0, 0, 255)
        display.BackColor = vbWhite
       ' display.Caption = dbcode.GetFields("FNAME") & " " & dbcode.GetFields("LNAME")
        lbln1(0) = FNAME
        lbln1(1) = LNAME
        lbln1(2) = dbcode.GetFields("deptname")
       ' method = "SCANNER"
       ' mee = 1
'===============================================================================================RIGHT INDEX VALIDATION
 'If txtdpt <> "" Then
 '   Call DATA_ENTRY(ID)
' End If


txtEMP_NO = id
'===============================================================================================RIGHT INDEX VALIDATION
        FPGTEMP.Run
        Exit Do
      ElseIf AUTHENTICATE = False Then
       display.ForeColor = RGB(255, 0, 0)
       display.BackColor = vbWhite
       display.Caption = "ENTRY DENIED"
       EID(1).Caption = ""
    End If
    dbcode.MoveNext
    
Loop
FPGTEMP.Run
Exit Sub
EXCEPTION_THROWN:
Call dbcode.WriteErrorLog(Err)
'========================EOF RIGHT INDEX========================================================

End Sub

' sample 1

Private Sub FPGTEMP_SampleReady(ByVal pSample As Object)

Dim samp As FPSample

              Set samp = pSample
                samp.PictureOrientation = Or_Portrait
                samp.PictureWidth = Picture2.Width / Screen.TwipsPerPixelX
                samp.PictureHeight = Picture2.Height / Screen.TwipsPerPixelY
                Picture2.Picture = samp.Picture
              ' SavePicture Picture2.Picture, txtEMP_NO.Text & ".jpg"
                SavePicture Picture2.Picture, txtdpt.Text & ".jpg"
End Sub

Private Sub Label2_Click(Index As Integer)

On Error Resume Next
Select Case Index

Case 0

For I = 0 To 6

    Label2(Index).BackColor = RGB(255, 0, 0)
    
    SaveSetting VB.APP.EXEName, "CONFIG", "DAY", MON
    
    Label2(I).BackColor = RGB(0, 0, 255)
    
Next I

Case 1


For I = 0 To 6
    Label2(Index).BackColor = RGB(255, 0, 0)
    
    Label2(I).BackColor = RGB(0, 0, 255)
    
    SaveSetting VB.APP.EXEName, "CONFIG", "DAY", TUE

    
Next I

Case 2

For I = 0 To 6

    Label2(Index).BackColor = RGB(255, 0, 0)
    
    Label2(I).BackColor = RGB(0, 0, 255)
    
    SaveSetting VB.APP.EXEName, "CONFIG", "DAY", WED

    
Next I

Case 3

For I = 0 To 6

    Label2(Index).BackColor = RGB(255, 0, 0)
    
    Label2(I).BackColor = RGB(0, 0, 255)
    
    SaveSetting VB.APP.EXEName, "CONFIG", "DAY", THU

    
Next I

Case 4

For I = 0 To 6

    Label2(Index).BackColor = RGB(255, 0, 0)
    
    Label2(I).BackColor = RGB(0, 0, 255)
    
    SaveSetting VB.APP.EXEName, "CONFIG", "DAY", FRI

    
Next I

Case 5

For I = 0 To 6

    Label2(Index).BackColor = RGB(255, 0, 0)
    
    Label2(I).BackColor = RGB(0, 0, 255)
    
    SaveSetting VB.APP.EXEName, "CONFIG", "DAY", SAT

    
Next I

Case 6

For I = 0 To 6 - 1

    Label2(Index).BackColor = RGB(255, 0, 0)
    
    Label2(I).BackColor = RGB(0, 0, 255)
    
    SaveSetting VB.APP.EXEName, "CONFIG", "DAY", SUN

    
Next I
    


End Select



End Sub

Private Sub List1_Click(Index As Integer)

End Sub

Private Sub List1_LostFocus(Index As Integer)

End Sub

Private Sub LE_Click()

modular.OnTop False
frmReport.Label1.Caption = "List Of Loged In Employees"
'query = "select * from tk_timekeeping where td = 1 AND TODAY = " & "#" & Date & "#" & " order by a_timein desc"
'query = GENQRY & " where td = 1 AND TODAY BETWEEN '" & Date - 1 & "' AND '" & Date & "' order by a_timein ASC"
query = GENQRY & " where td = 1 and today = #" & Date & "#" & " order by A_TIMEIN OR P_TIMEIN OR N_TIMEIN"

status = 1
frmReport.Show 1
End Sub

Private Sub LOE_Click()

Call OnTop(False)
frmReport.Label1.Caption = "List Of Loged Out Employees"
'query = "select * from tk_timekeeping where td = 0 AND TODAY = " & "#" & Date & "#" & " order by a_timeOUT desc "

query = GENQRY & " where td = 0 and today = #" & Date & "#" & " order by A_TIMEOUT OR P_TIMEOUT OR N_TIMEOUT DESC"
'query = GENQRY & " where td = 0 AND TODAY   BETWEEN '" & Date - 1 & "' AND '" & Date & "' order by a_timeOUT ASC"
status = 0
frmReport.Show 1

End Sub

Private Sub ltimer_Timer()

Static cout As Integer

cout = cout + 1

If cout = 10 Then
    'DME_Click
   ' ltimer.Enabled = False
    txtEMP_NO.Text = "0000"
    cout = 0
End If
End Sub

Private Sub MOE_Click()
Call OnTop(False)
frmMoe.Show 1
End Sub

Private Sub OPTDL_Click(Index As Integer)

Select Case Index

    
    Case 0
    
      OPTDL(1).value = False
      OPTDL(2).value = False
      OPTDL(3).value = False
      OPTDL(4).value = False
      OPTDL(5).value = False
      
    Case 1
    
      OPTDL(0).value = False
      OPTDL(2).value = False
      OPTDL(3).value = False
      OPTDL(4).value = False
      OPTDL(5).value = False
      
    Case 2
    
        
      OPTDL(0).value = False
      OPTDL(1).value = False
      OPTDL(3).value = False
      OPTDL(4).value = False
      OPTDL(5).value = False
      
    Case 3
    
      OPTDL(0).value = False
      OPTDL(1).value = False
      OPTDL(2).value = False
      OPTDL(4).value = False
      OPTDL(5).value = False
      
   Case 4
   
      OPTDL(0).value = False
      OPTDL(1).value = False
      OPTDL(2).value = False
      OPTDL(3).value = False
      OPTDL(5).value = False
      
  Case 5
  
      OPTDL(0).value = False
      OPTDL(1).value = False
      OPTDL(2).value = False
      OPTDL(3).value = False
      OPTDL(4).value = False

End Select

End Sub

Private Sub rpt_Click()
frmrpt.Show
End Sub

Private Sub SO_Click()

    If SWITCH.value = 0 Then
    
    SWITCH.value = 1
    SO.Caption = "Scanner is ON"
    
       ElseIf SWITCH.value = 1 Then
    
    SWITCH.value = 0
    SO.Caption = "Scanner is OFF"
    End If
    
End Sub

Private Sub Srch_Click()

input1$ = InputBox("Enter Last Name or Employee ID", "Time Keeper")
input1$ = UCase(input1$)
input1$ = UCase(input1$)

If IsNumeric(input1$) Then

For I = 1 To viewer.ListItems.Count
    If Left(viewer.ListItems(I).SubItems(1), Len(input1$)) = input1$ Then
    viewer.ListItems(I).Selected = True
    viewer.ListItems(I).EnsureVisible
    viewer.SetFocus
    Else
     viewer.ListItems(I).Selected = False
    End If
Next I
    viewer.Refresh
Else

For I = 1 To viewer.ListItems.Count
    If Left(viewer.ListItems(I).SubItems(9), Len(input1$)) = input1$ Then
    viewer.ListItems(I).Selected = True
    viewer.ListItems(I).EnsureVisible
    viewer.SetFocus
    Else
     viewer.ListItems(I).Selected = False
    End If
Next I
    viewer.Refresh
End If
    
End Sub

Private Sub SWITCH_Click()
On Error Resume Next
Dim DPLAY As New clsDBCode


dbcode.ReleaseConnection
dbx.ReleaseConnection
dby.ReleaseConnection

dbcode.MDB_Access_Connect
Call DISPLAY_DATA(F, D, 0)
Dim BOOL_A(0 To 5) As String

   Select Case SWITCH.value
    Case 0
        SWITCH.Caption = "VERIFIER OFF"
        Set FPGTEMP = Nothing
        Set fptemp = Nothing
        For I = 0 To OPTDL.Count - 1
           OPTDL(I).value = False
        Next I
        SWITCH.ForeColor = RGB(255, 0, 0)
        Exit Sub
    Case 1
        For I = 0 To 5
            BOOL_A(I) = OPTDL(I).value
            If BOOL_A(I) = False Then x = x + 1
                If x = OPTDL.Count Then
                modular.XEND = 0
                SWITCH.Caption = "VERIFIER OFF"
            End If
        Next I
        SWITCH.Caption = "VERIFIER ON"
        SWITCH.ForeColor = RGB(0, 0, 255)
        Set FPGTEMP = New FPGetTemplate
        FPGTEMP.Run
        
  End Select
End Sub

Private Sub timex_Timer()

On Error Resume Next

Dim GS As Integer

GS = GetSetting(VB.APP.EXEName, "CONFIG", "DAY")

LAPSECOUNT = LAPSECOUNT + 1


If Time = "12:00:00 AM" Then
    GS = GS + 1
    
    If GS = 7 Then SaveSetting VB.APP.EXEName, "CONFIG", "DAY", 0
    
    If GS = 1 Then
    Label2(0).BackColor = RGB(255, 0, 0)
    Else
    Label2(0).BackColor = RGB(0, 0, 255)
    End If
    
    If GS = 2 Then
    Label2(1).BackColor = RGB(255, 0, 0)
    Else
    Label2(1).BackColor = RGB(0, 0, 255)
    End If
   
    If GS = 3 Then
    Label2(2).BackColor = RGB(255, 0, 0)
    Else
    Label2(2).BackColor = RGB(0, 0, 255)
    End If
    
    If GS = 4 Then
    Label2(3).BackColor = RGB(255, 0, 0)
    Else
    Label2(3).BackColor = RGB(0, 0, 255)
    End If
    
    If GS = 5 Then
    Label2(4).BackColor = RGB(255, 0, 0)
    Else
    Label2(4).BackColor = RGB(0, 0, 255)
    End If
    
    If GS = 6 Then
    Label2(5).BackColor = RGB(255, 0, 0)
    Else
    Label2(5).BackColor = RGB(0, 0, 255)
    End If
    
    If GS = 7 Then
    Label2(6).BackColor = RGB(255, 0, 0)
    Else
    Label2(6).BackColor = RGB(0, 0, 255)
    End If
    
End If

If LAPSECOUNT = 10 Then
    Image1.Picture = LoadPicture(VB.APP.Path + "\0.JPG")
    LAPSECOUNT = 0
    'lbln1(0) = ""
    'lbln1(1) = ""
    'lbln1(2) = ""
    'lbln1(3) = ""
   ' EID(1) = ""
    Picture2.Picture = LoadPicture
    If txtEMP_NO.Enabled = True Then txtEMP_NO.SetFocus
   ' txtdpt.SetFocus
   
End If


   lbltime = Format(Time, "hh:mm:ss")
   lbldate = Date
   
   
   If Time = "11:59:00 AM" Then
    Open VB.APP.Path + "\YDATE" For Output As #1
     Print #1, Date
    Close #1
   End If
    
    
End Sub


Sub DISPLAY_DATA(ByVal FNAME As String, ByVal LNAME As String, Optional DT As Integer)
Dim MOE As String
dby.ReleaseConnection
dbx.ReleaseConnection

viewer.Refresh
viewer.ListItems.Clear
dbcode.MDB_Access_Connect
If DT = 0 Then

    'dbcode.Sql_SELECT_QUERY_Execute GENQRY & " WHERE today BETWEEN " & "#" & Date - 1 & "#" & " AND " & "#" & Date & "#" & " AND TD = 1 ORDER BY A_TIMEIN DESC", ""
     dbcode.Sql_SELECT_QUERY_Execute GENQRY & " WHERE today = " & "#" & Date & "#" & " ORDER BY A_TIMEIN DESC", ""
ElseIf DT = 1 Then
    dbcode.Sql_SELECT_QUERY_Execute GENQRY & " WHERE today = #" & Date - 1 & "#" & " ORDER BY A_TIMEIN DESC", ""
End If
    
    With dbcode
    While Not .EOF
    DoEvents
    I = I + 1
            id = .GetFields("EMPID")
            TDAY = .GetFields("TODAY")
            A_IN = .GetFields("A_TIMEIN")
            A_OUT = .GetFields("A_TIMEOUT")
            P_IN = .GetFields("P_TIMEIN")
            P_OUT = .GetFields("P_TIMEOUT")
            N_IN = .GetFields("N_TIMEIN")
            N_OUT = .GetFields("N_TIMEOUT")
            EMP = .GetFields("EMPLOYEE")
            MOE = .GetFields("MOE")
            dept = .GetFields("deptname")
            
            Set lv = viewer.ListItems.Add(, , I)
               lv.SubItems(1) = id
               lv.SubItems(2) = TDAY
               lv.SubItems(3) = A_IN
               lv.SubItems(4) = A_OUT
               lv.SubItems(5) = P_IN
               lv.SubItems(6) = P_OUT
               lv.SubItems(7) = N_IN
               lv.SubItems(8) = N_OUT
               lv.SubItems(9) = EMP
               lv.SubItems(10) = MOE
               lv.SubItems(11) = dept
     .MoveNext
     Wend
    End With
End Sub
Sub DISPLAY_IMAGE(ByVal IDX As String)

With dbcode
    .OUT_IMAGE_SZ IDX
End With

    If Dir("temp") <> "" Then
        Image1.Picture = LoadPicture("TEMP")
        Kill "TEMP"
    End If

End Sub

Private Sub txtdpt_Change()
    If Not IsNumeric(txtdpt) Then txtdpt = ""
End Sub

Private Sub txtdpt_GotFocus()
txtdpt = ""
lblf.Caption = "MOE: Scanner"
End Sub

Private Sub txtdptx_Change()
If Not IsNumeric(txtdptx) Then txtdptx = ""

End Sub

Private Sub txtdptx_GotFocus()
txtdpt = ""
End Sub

Private Sub txtEMP_NO_Change()

If txtEMP_NO.Text = "0000" Then ltimer.Enabled = False

If Len(txtEMP_NO) = 4 Then

'If mee = 1 Then method = lblf
'If mee <> 1 Then method = lblf
    method = lblf
    KEY = txtEMP_NO
    Call TIME_KEEPER(KEY, method)
End If

End Sub


Sub Get_TotalTime()


Dim dbs As New clsDBCode


End Sub


Sub TIME_KEEPER(ByVal KEY As Variant, Optional MOE As String)

' INSERT METHOD
Dim dbref As New clsDBCode
Dim ints As Integer
dby.MDB_Access_Connect
dbref.MDB_Access_Connect
dbref.Sql_SELECT_QUERY_Execute "select * from tk_timekeeping where empid = " & KEY & " and TD = 1"

If dby.Sql_SELECT_QUERY_Execute("select * from db_finger_print where empno = " & KEY) = True Then
frerror.Visible = True: lblerror.Caption = "ID Number Not Found"
Exit Sub
Else

dbx.MDB_Access_Connect
If dbx.Sql_SELECT_QUERY_Execute("select * from tk_timekeeping where empid = " & KEY & "and today = #" & Date & "#") = False Then '{


If dbx.Sql_SELECT_QUERY_Execute("select * from tk_timekeeping where empid = " & KEY & "and today = #" & Date & "#", "A_TIMEOUT") = True Then
  dbx.SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set A_timeout = '" & Format(Time, "hh:mm") & "',td = " & EOUT & "  where empid = " & KEY & " and today = #" & Date & "#", ""
  Call dbx.Sql_SELECT_QUERY_Execute("select * from tk_timekeeping where empid = " & KEY & " and today = #" & Date & "#")
  
 
  Call VERIFY_IF_INOUT(KEY, 2)
  Call ALTERLOG(dby.GetFields("empno") & " " & Space(5) & dby.GetFields("lname") & "," & dby.GetFields("fname"), "OUT", 0)
ElseIf dbx.Sql_SELECT_QUERY_Execute("select * from tk_timekeeping where empid = " & KEY & "and today = #" & Date & "#", "P_timein") = True Then
  dbx.SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set p_timein = '" & Format(Time, "hh:mm") & "',td = " & EIN & " where empid = " & KEY & " and today = #" & Date & "#", ""
  Call VERIFY_IF_INOUT(KEY, 3)
  Call ALTERLOG(dby.GetFields("empno") & " " & Space(5) & dby.GetFields("lname") & "," & dby.GetFields("fname"), "IN", 1)

ElseIf dbx.Sql_SELECT_QUERY_Execute("select * from tk_timekeeping where empid = " & KEY & "and today = #" & Date & "#", "p_timeout") = True Then
  dbx.SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set p_timeout = '" & Format(Time, "hh:mm") & "',td = " & EOUT & " where empid = " & KEY & " and today = #" & Date & "#", ""
  Call VERIFY_IF_INOUT(KEY, 4)
  Call ALTERLOG(dby.GetFields("empno") & " " & Space(5) & dby.GetFields("lname") & "," & dby.GetFields("fname"), "OUT", 0)

ElseIf dbx.Sql_SELECT_QUERY_Execute("select * from tk_timekeeping where empid = " & KEY & "and today = #" & Date & "#", "n_timein") = True Then
  dbx.SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set n_timein = '" & Format(Time, "hh:mm") & "',td = " & EIN & " where empid = " & KEY & " and today = #" & Date & "#", ""
  Call VERIFY_IF_INOUT(KEY, 5)
  Call ALTERLOG(dby.GetFields("empno") & " " & Space(5) & dby.GetFields("lname") & "," & dby.GetFields("fname"), "IN", 1)

ElseIf dbx.Sql_SELECT_QUERY_Execute("select * from tk_timekeeping where empid = " & KEY & "and  today = #" & Date & "#", "n_timeout") = True Then
  dbx.SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set n_timeout = '" & Format(Time, "hh:mm") & "',td = " & EOUT & " where empid = " & KEY & " and today = #" & Date & "#", ""
  Call VERIFY_IF_INOUT(KEY, 6)
  Call ALTERLOG(dby.GetFields("empno") & " " & Space(5) & dby.GetFields("lname") & "," & dby.GetFields("fname"), "OUT", 0)
End If

Else
'--------------------------------------------------------------------------------------------------------------------------------------------
If dbx.Sql_SELECT_QUERY_Execute("select * from tk_timekeeping where empid =" & KEY & " AND TD = 1", "") = False Then

If dbref.GetFields("a_timeout") = "" Then
  dbx.SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set A_timeout = '" & Format(Time, "hh:mm") & "',td = " & EOUT & " where empid = " & KEY & "and td = 1", "" 'and today = #" & Date - 1 & "#", ""
End If

If dbref.GetFields("p_timeout") = "" Then
  dbx.SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set p_timeout = '" & Format(Time, "hh:mm") & "',td = " & EOUT & " where empid = " & KEY & "and td = 1", "" 'and today = #" & Date - 1 & "#", ""
End If

If dbref.GetFields("n_timeout") = "" Then
  dbx.SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set n_timeout = '" & Format(Time, "hh:mm") & "',td = " & EOUT & " where empid = " & KEY & "and td = 1", "" 'and today = #" & Date - 1 & "#", ""
End If

  Call ALTERLOG(dby.GetFields("empno") & " " & Space(5) & dby.GetFields("lname") & "," & dby.GetFields("fname"), "OUT", -1)
  Call VERIFY_IF_INOUT(KEY, 7)
  vdis = 1
  GoTo displaydata1
End If

'If dbx.Sql_SELECT_QUERY_Execute("select * from tk_timekeeping where empid =" & KEY & " AND TD = 1", "p_timeout") = False Then
 ' dbx.SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set p_timeout = '" & Format(Time, "hh:mm") & "',td = " & EOUT & " where empid = " & KEY & "and td = 1", "" 'and today = #" & Date - 1 & "#", ""
 ' Call ALTERLOG(dby.GetFields("empno") & " " & Space(5) & dby.GetFields("lname") & "," & dby.GetFields("fname"), "OUT", -1)
 ' Call VERIFY_IF_INOUT(KEY, 7)
 ' vdis = 1
 ' GoTo displaydata1
'End If

'If dbx.Sql_SELECT_QUERY_Execute("select * from tk_timekeeping where empid =" & KEY & " AND TD = 1", "n_timeout") = False Then
'  dbx.SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set n_timeout = '" & Format(Time, "hh:mm") & "',td = " & EOUT & " where empid = " & KEY & "and td = 1", "" 'and today = #" & Date - 1 & "#", ""
'  Call ALTERLOG(dby.GetFields("empno") & " " & Space(5) & dby.GetFields("lname") & "," & dby.GetFields("fname"), "OUT", -1)
'  Call VERIFY_IF_INOUT(KEY, 7)
 ' vdis = 1
'  GoTo displaydata1
'End If

Dim x1() As Byte
Dim x2() As Byte
Dim ipix As StdPicture
Dim tempB() As Byte

method = lblf
    dby.Sql_SELECT_QUERY_Execute "select * from db_finger_print where empno = " & KEY
    empname = dby.GetFields("lname") & "," & dby.GetFields("fname")
    dept = dby.GetFields("deptname")
    stat = dby.GetFields("empstatus")
    FNAME = dby.GetFields("fname")
    LNAME = dby.GetFields("lname")
        
    dbx.SqlCommand_NON_SELECT_QUERY_Execute "insert into tk_timekeeping (empid,today,a_timein,employee,td,deptname,empstat,MOE) values ('" & KEY & "','" & Date & "','" & Format(Time, "hh:mm") & "','" & empname & "','" & EIN & "','" & dept & "','" & stat & "','" & method & "')", ""
    Call VERIFY_IF_INOUT(KEY, 1)
    Call ALTERLOG(dby.GetFields("empno") & " " & Space(5) & dby.GetFields("lname") & "," & dby.GetFields("fname"), "IN", 1)
    vdis = 0
End If '}
End If
displaydata1:


   ' x1 = dby.GetFields("rthumb")
    'x2 = dby.GetFields("rindex")
    '
    'Set fptemp = New FPTemplate
    'fptemp.Import x1
    
    
    On Error Resume Next
    
   fptemp.Export tempB()
 
 'Open "d.jpg" For Binary Access Write As #1
 'Debug.Print tempB
 '   Put #1, , tempB()
 'Close #1
  
          If vdis = 0 Then
              ints = 0
           Else
              ints = 1
          End If
         ' namex = lv.SubItems(9)
          
          splname = mi
    On Error Resume Next
    
    Call DISPLAY_DATA(FNAME, LNAME, ints)
    Call DISPLAY_IMAGE(KEY)
    Call search_hi_lite(KEY)
    
    Dim fp As FPSample
    Dim j As New FPRawSamplePro


    EID(1) = KEY
      lbln1(0) = Mid(viewer.SelectedItem.SubItems(9), InStr(1, viewer.SelectedItem.SubItems(9), ",") + 1)
      lbln1(1) = Mid(viewer.SelectedItem.SubItems(9), 1, InStr(1, viewer.SelectedItem.SubItems(9), ",") - 1)
      lbln1(2) = viewer.SelectedItem.SubItems(11)
      lbln1(3) = Time
On Local Error Resume Next

mee = 0

End Sub


Sub VERIFY_IF_INOUT(KEY As Variant, ByVal x As Integer)


Dim dbcode As New clsDBCode
display.Caption = "VERIFYING..., YOU JUST WAIT THERE..!"
For I = 1 To 50
DoEvents
dbcode.MDB_Access_Connect
If x = 1 Then
  If dbcode.Sql_SELECT_QUERY_Execute("SELECT  * FROM TK_TIMEKEEPING WHERE EMPID = " & KEY & " AND TODAY = #" & Date & "#", "A_TIMEIN") = False Then
    display.Caption = dbcode.GetFields("EMPLOYEE") '& " " & "IS ALREADY TIMED IN"
    lblerror.Caption = dbcode.GetFields("EMPLOYEE") & " " & "TIMED-IN AT " & dbcode.GetFields("A_TIMEIN")
    lblerror.FontSize = 8
    lblerror.FontBold = True
    lblerror.ForeColor = vbBlack
    frerror.Visible = True
    vtime.Enabled = True
    txtdpt.Text = ""
   'SWITCH.value = 0
    Else
    'display.Caption = "TRY AGAIN"
  End If
ElseIf x = 2 Then
  If dbcode.Sql_SELECT_QUERY_Execute("SELECT * FROM TK_TIMEKEEPING WHERE EMPID = " & KEY & " AND TODAY = #" & Date & "#", "A_TIMEOUT") = False Then
    display.Caption = dbcode.GetFields("EMPLOYEE") '& " " & "IS ALREADY TIMED OUT"
    lblerror.Caption = dbcode.GetFields("EMPLOYEE") & " " & "TIMED-OUT AT " & dbcode.GetFields("A_TIMEOUT")
    lblerror.FontSize = 8
    lblerror.ForeColor = vbBlack
    lblerror.FontBold = True
    frerror.Visible = True
    vtime.Enabled = True
    txtdpt.Text = "0000"

    'SWITCH.value = 0

    Else
   ' display.Caption = "TRY AGAIN"
  End If
ElseIf x = 3 Then
  If dbcode.Sql_SELECT_QUERY_Execute("SELECT * FROM TK_TIMEKEEPING WHERE EMPID = " & KEY & " AND TODAY = #" & Date & "#", "P_TIMEIN") = False Then
    display.Caption = dbcode.GetFields("EMPLOYEE") '& " " & "IS ALREADY TIMED IN"
    lblerror.Caption = dbcode.GetFields("EMPLOYEE") & " " & "TIMED-IN AT " & dbcode.GetFields("P_TIMEIN")
    lblerror.FontSize = 8
     lblerror.FontBold = True
    lblerror.ForeColor = vbBlack
    frerror.Visible = True
    vtime.Enabled = True
    txtdpt.Text = "0000"

    'SWITCH.value = 0

    Else
    'display.Caption = "TRY AGAIN"
  End If
ElseIf x = 4 Then
  If dbcode.Sql_SELECT_QUERY_Execute("SELECT * FROM TK_TIMEKEEPING WHERE EMPID = " & KEY & " AND TODAY = #" & Date & "#", "P_TIMEOUT") = False Then
    display.Caption = dbcode.GetFields("EMPLOYEE") '& " " & "IS ALREADY TIMED OUT"
    lblerror.Caption = dbcode.GetFields("EMPLOYEE") & " " & "TIMED-OUT AT " & dbcode.GetFields("P_TIMEOUT")
    lblerror.FontSize = 8
     lblerror.FontBold = True
    lblerror.ForeColor = vbBlack
    frerror.Visible = True
    vtime.Enabled = True
     txtdpt.Text = "0000"

   ' SWITCH.value = 0

    Else
   ' display.Caption = "TRY AGAIN"
  End If
ElseIf x = 5 Then
  If dbcode.Sql_SELECT_QUERY_Execute("SELECT * FROM TK_TIMEKEEPING WHERE EMPID = " & KEY & " AND TODAY = #" & Date & "#", "N_TIMEIN") = False Then
    display.Caption = dbcode.GetFields("EMPLOYEE") '& " " & "IS ALREADY TIMED IN"
    lblerror.Caption = dbcode.GetFields("EMPLOYEE") & " " & "TIMED-IN AT " & dbcode.GetFields("N_TIMEIN")
    lblerror.FontSize = 8
     lblerror.FontBold = True
    lblerror.ForeColor = vbBlack
    frerror.Visible = True
     vtime.Enabled = True
      txtdpt.Text = "0000"

    ' SWITCH.value = 0
   Else
   ' display.Caption = "TRY AGAIN"
  End If
ElseIf x = 6 Then
  If dbcode.Sql_SELECT_QUERY_Execute("SELECT * FROM TK_TIMEKEEPING WHERE EMPID = " & KEY & " AND TODAY = #" & Date & "#", "N_TIMEOUT") = False Then
    display.Caption = dbcode.GetFields("EMPLOYEE") '& " " & "IS ALREADY TIMED OUT"
    lblerror.Caption = dbcode.GetFields("EMPLOYEE") & " " & "TIMED-OUT AT " & dbcode.GetFields("N_TIMEOUT")
    lblerror.FontSize = 8
     lblerror.FontBold = True
    lblerror.ForeColor = vbBlack
    frerror.Visible = True
     vtime.Enabled = True
        txtdpt.Text = "0000"
     '  SWITCH.value = 0

    Else
   ' display.Caption = "TRY AGAIN"
  End If
  
ElseIf x = 7 Then
  If dbcode.Sql_SELECT_QUERY_Execute("SELECT * FROM TK_TIMEKEEPING WHERE EMPID = " & KEY & " AND TODAY = #" & Date - 1 & "#", "A_TIMEOUT") = False Then
    display.Caption = dbcode.GetFields("EMPLOYEE") '& " " & "IS ALREADY TIMED OUT"
    lblerror.Caption = dbcode.GetFields("EMPLOYEE") & " " & "TIMED-OUT AT " & dbcode.GetFields("A_TIMEOUT")
    lblerror.FontSize = 8
     lblerror.FontBold = True
    lblerror.ForeColor = vbBlack
    frerror.Visible = True
     vtime.Enabled = True
        txtdpt.Text = "0000"
End If

End If

Next I

display.Caption = ""

End Sub


Sub ALTERLOG(ByVal ENAME As String, stat, int1 As Integer)
Dim DXT As String
DXT = Date

DXT = Replace(DXT, "/", ".")
If Dir(VB.APP.Path + "\LOGS\" & DXT & ".txt") <> "" Then
Open VB.APP.Path + "\LOGS\" & DXT & ".txt" For Append As #1
    Print #1, Now, ENAME & Space(10), int1, stat
Close #1
    Else
    'MsgBox VB.APP.Path + "\LOGS\" & DXT & ".txt"
Open VB.APP.Path + "\LOGS\" & DXT & ".txt" For Output As #1
    Print #1, Now, ENAME & Space(10), int1, stat
Close #1
End If

End Sub

Sub search_hi_lite(ByVal EID As Variant)

Dim lv As ListItem
Debug.Print EID
input1$ = UCase(Trim(EID))
If IsNumeric(input1$) Then
For I = 1 To Main.viewer.ListItems.Count
    If Left(Main.viewer.ListItems(I).SubItems(1), Len(input1$)) = input1$ Then
    Main.viewer.ListItems(I).EnsureVisible
    Main.viewer.ListItems(I).Selected = True
    Main.viewer.SetFocus
    Else
    Main.viewer.ListItems(I).Selected = False
    End If
Next I
End If
End Sub


Private Sub txtEMP_NO_GotFocus()
txtEMP_NO = ""
lblf.Caption = "MOE: Manual"
End Sub

Private Sub VEL_Click()
frameinfo.Visible = False
End Sub


Sub savesettings()
    SaveSetting VB.APP.EXEName, "config", "T", 1
End Sub



Sub DATA_ENTRY(ByVal ID_X As String)

' the main data entry method same as with the txtemp_no_change event

With dbcode
   .SqlConnect "dbemployment"
   
If .Sql_SELECT_QUERY_Execute(GENQRY + " where empno = '" & ID_X & "'") = True Then
If .Sql_SELECT_QUERY_Execute("SELECT * FROM DB_FINGER_PRINT WHERE EMPNO = '" & ID_X & "'", "") = True Then frerror.Visible = True: lblerror.Caption = "ID Number Not Found": Exit Sub
       FNAME = .GetFields("fname")
       LNAME = .GetFields("lname")
       IDX = .GetFields("EMPNO")
       empname = LNAME + "," + FNAME
       dept = .GetFields("DEPTNAME")
       Else
 .Sql_SELECT_QUERY_Execute "SELECT * FROM DB_FINGER_PRINT WHERE EMPNO ='" & ID_X & "'"
       FNAME = .GetFields("fname")
       LNAME = .GetFields("lname")
       IDX = .GetFields("EMPNO")
       empname = LNAME + "," + FNAME
       dept = .GetFields("DEPTNAME")
       stat = .GetFields("EMPSTATUS")
End If

If .Sql_SELECT_QUERY_Execute(GENQRY & " where empid = '" & ID_X & "' and today = '" & Date & "'") = True Then
         
            If .GetFields("TDOWN") = "" Then
             TDOWN = 1
            End If
        
 ' check for the employee's grave yard shift timeout
        
 If .Sql_SELECT_QUERY_Execute(GENQRY + " where empid = '" & ID_X & "' and td = '" & TDOWN & "'") = False Then
 
 If .GetFields("a_timeout") = "" Then
        .SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set a_timeout = '" & Time & "',td = '" & EOUT & "'  where empid = '" & IDX & "' and td = '1'", ""
        GoTo displaydata1
 ElseIf .GetFields("p_timein") = "" Then
       .SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set p_timein = '" & Time & "',td = '" & EIN & "'  where empid = '" & IDX & "' and td = '0'", ""
        GoTo displaydata1
 ElseIf .GetFields("p_timeout") = "" Then
       .SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set p_timeout = '" & Time & "',td = '" & EOUT & "'  where empid = '" & IDX & "' and td = '1'", ""
        GoTo displaydata1
 End If
 
 End If
       .SqlCommand_NON_SELECT_QUERY_Execute "insert into tk_timekeeping (empid,today,a_timein,employee,td,deptname,empstat) values ('" & ID_X & "','" & Date & "','" & Time & "','" & empname & "','" & EIN & "','" & dept & "','" & stat & "')", ""
  
Else
            TDOWN = .GetFields("td")
  If .Sql_SELECT_QUERY_Execute(GENQRY & " where empid = '" & ID_X & "' and today = '" & Date & "'", "a_timeout") = True Then
      TDOWN = TDOWN + 1
     .SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set a_timeout = '" & Time & "',td = '" & EOUT & "' where empid = '" & ID_X & "' and today = '" & Date & "'", ""
          GoTo displaydata1
  ElseIf .Sql_SELECT_QUERY_Execute(GENQRY & " where empid = '" & ID_X & "' and today = '" & Date & "'", "p_timein") = True Then
           TDOWN = TDOWN + 1
         .SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set p_timein = '" & Time & "',td = '" & EIN & "' where empid = '" & ID_X & "' and today = '" & Date & "'", ""
         GoTo displaydata1
   ElseIf .Sql_SELECT_QUERY_Execute(GENQRY & " where empid = '" & ID_X & "' and today = '" & Date & "'", "p_timeout") = True Then
           TDOWN = TDOWN + 1
          .SqlCommand_NON_SELECT_QUERY_Execute "update tk_timekeeping set p_timeout = '" & Time & "',td = '" & EOUT & "' where empid = '" & ID_X & "' and today = '" & Date & "'", ""
          GoTo displaydata1
  End If
  End If
  
  
  '===============================================================================================================
  
' If .Sql_SELECT_QUERY_Execute(GENQRY & " where empid = " & txtEMP_NO & " and today = " & "#" & Date - 1 & "#") = True Then ' T1
      
displaydata1:
        display.Caption = FNAME & " " & LNAME
        lbln1(0) = FNAME
        lbln1(1) = LNAME
        lbln1(2) = dept
        lbln1(3) = Time
        txtEMP_NO = ""
        txtEMP_NO.SetFocus
       Call DISPLAY_DATA(FNAME, LNAME)
       Call DISPLAY_IMAGE(IDX)
      ' On Error Resume Next
       Call search_hi_lite(IDX)
End With





End Sub

Private Sub vtime_Timer()

Static x As Integer

t1 = Format(Time, "ss")

x = x + 1
Debug.Print x
If x = 5 Then
x = 0
    frerror.Visible = False
    txtdpt.Locked = False
    txtdpt.Enabled = True
    vtime.Enabled = False
    lblerror.FontBold = False
   ' lblerror.ForeColor = RGB(0, 0, 0)
    lblerror.FontSize = 8
   txtdpt = ""
   txtEMP_NO = ""

End If




End Sub
