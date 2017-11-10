VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   3735
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   3300
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   3300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   1
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Ok"
         Height          =   375
         Index           =   0
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         Picture         =   "frmlogin.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         Index           =   0
         X1              =   240
         X2              =   4680
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIME KEEPING UTILITY"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbcode As New clsDBCode

Private Sub cmd_Click(Index As Integer)

Select Case Index

Case 0

If InStr(Text1(0), "'") <> 0 Then Text1(0) = "": Exit Sub
If InStr(Text1(1), "'") <> 0 Then Text1(1) = "": Exit Sub

dbcode.MDB_Access_Connect
dbcode.Sql_SELECT_QUERY_Execute "SELECT * FROM TK_MELOGIN WHERE UNAME = '" & Text1(0) & "' AND PWORD = '" & Text1(1) & "'"

If dbcode.EOF = True Then
    dbcode.message "Invalid Login"
Else

un = dbcode.GetFields("uname")
pw = dbcode.GetFields("pword")

Label3 = "Username is: " + un
Label4 = "Password is: " + pw

    Main.txtEMP_NO.Enabled = True
    Main.lblID(2).Visible = True
   ' Main.txtEMP_NO.Enabled = True
   ' Main.lblID(2).Enabled = True
   ' Main.lblf.Caption = ""
   Main.txtdpt.Enabled = False
    Main.ltimer.Enabled = True
    Unload Me
   ' Set frmlogin = Nothing
End If

Case 1
Call OnTop(True)
Unload Me

End Select

End Sub

Private Sub Command1_Click()
'Me.Width = 8985

Dim s As Integer



End Sub

Private Sub Form_Load()
Text1(0).Text = "admin"
Text1(1).PasswordChar = "*"
End Sub

Private Sub Text1_GotFocus(Index As Integer)

Select Case Index
Case 0
Text1(0).Text = ""
Case 1
Text1(1).Text = ""
End Select
End Sub
