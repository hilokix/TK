VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmrpt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   4785
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   840
      TabIndex        =   14
      Top             =   3360
      Width           =   3135
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   " E&xit"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print Report &All"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print Report By &Status"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2040
         Width           =   2895
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Printing"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   4575
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   4335
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2160
            TabIndex        =   13
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox stat 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dept Code"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   4335
         Begin MSComCtl2.DTPicker DTX 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            Format          =   190775297
            CurrentDate     =   39925
         End
         Begin VB.TextBox txtid 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2160
            TabIndex        =   10
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox stat 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Employees ID"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker DTX 
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   24
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            Format          =   190775297
            CurrentDate     =   39925
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TO"
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   23
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "FROM"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "For Viewing"
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   4335
         Begin VB.CheckBox stat 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Employees Status"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report Options"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "R E P O R T S"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      X1              =   0
      X2              =   4800
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   600
      Picture         =   "frmrpt.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "frmrpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2  As New ADODB.Recordset



Const uid = ""
Const pwd = ""
Const data_source = "mis33"
Const database = "dbemployment"

Dim OBJ As PageSet.PrinterControl

Private Sub Command1_Click(Index As Integer)

Select Case Index

Case 0


If stat(0).value = 1 Then
    If Combo1.Text <> " " Then
    dbcode.SqlConnect "DBEMPLOYMENT"
    dbcode.Sql_SELECT_QUERY_Report "select * from db_finger_print WHERE EMPSTATUS = '" & Combo1.Text & "' ORDER BY LNAME", 2
    While Not dbcode.EOF
   ' DataReport2.Sections(3).Controls("TEXT1").DataField = "FNAME"
   ' DataReport2.Sections(2).Controls("label5").cap = "LNAME"
    dbcode.MoveNext
    Wend
    DataReport2.Show
    End If
ElseIf stat(1).value = 1 Then

    If txtid <> "" Then
    dbcode.ReleaseConnection
    'DATA_SOURCE = "dbemployment"
    
    If db1.State = 1 Then
        db1.Close
    End If
    
    If InStr(1, txtid, "'") = 1 Then Exit Sub
    
   ' db1.Open "Driver={SQL Server};Server=" & data_source & ";Database=" & database & ";Uid=" & uid & ";Pwd=" & pwd
   db1.Sql_SELECT_QUERY_Execute "select * from tk_timekeeping where MOE = 'MOE: Manual'"
   ' db1.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" + DATA_SOURCE + ";Uid=admin;Pwd="
    
    Set rs1 = db1.Execute("select empid,today,a_timein,a_timeout,p_timein,p_timeout,employee FROM tk_timekeeping where empid = '" & txtid.Text & "' AND TODAY between '" & DTX(0).value & "' and '" & DTX(1).value & "' order by today")
    If rs1.BOF <> True Or rs1.EOF <> True Then
    fullname = rs1!employee
    Set DataReport1.DataSource = rs1
    DataReport1.Sections(2).Controls("label8").Caption = fullname
    DataReport1.Sections(2).Controls("label9").Caption = rs1!EMPID
    Set OBJ = New PrinterControl
    OBJ.ChngOrientationPortrait
    DataReport1.Show
   ' DataReport1.PrintReport , rptRangeFromTo, 1, 1
    Else
    MsgBox "No Record Found", vbInformation
    
    End If
    End If
ElseIf stat(2).value = 1 Then
    If Text2 <> "" Then
    dbcode.SqlConnect "dbemployment"
    dbcode.Sql_SELECT_QUERY_Execute GENQRY & " where empid = '" & Text2 & "'"
    End If
End If

Case 1

Frame5.Visible = True
Combo2.Clear
Combo2.AddItem "CASUAL"
Combo2.AddItem "REGULAR"
Combo2.AddItem "OJT"
Combo2.AddItem " "

Combo2.Text = " "

Case 2

Set frmrpt = Nothing
Unload Me

End Select
End Sub

Private Sub Command2_Click(Index As Integer)


Select Case Index


Case 0
'On Error Resume Next


ans = MsgBox("This Will Print All Employees Attendance Record!, Do You Wish To Continue?", vbQuestion + vbYesNo, "Time Keeper")

If ans = vbYes Then

'DATA_SOURcE = "MISSERVER"
'Data_base = "dbemployment"
'UID = "sa"
'PWD = "gama27"
    

'=================================================================
    'db1.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" + DATA_SOURCE + ";Uid=admin;Pwd="
    db1.Open "Driver={SQL Server};Server=" & data_source & ";Database=" & data_base & ";Uid=" & uid & ";Pwd=" & pwd
    
    rs1.Open "select * from db_finger_print", db1, adOpenKeyset, adLockOptimistic
    
    frmAttd.pbar.Min = 0
    
    frmAttd.Show
    
    frmAttd.pbar.Max = rs1.RecordCount

    While Not rs1.EOF
    
    DoEvents
    
    id = rs1!empno
    
    frmAttd.pbar.value = frmAttd.pbar + 1
        Set rs2 = db1.Execute("select * from tk_timekeeping where empid = '" & id & "'")
        Set DataReport1.DataSource = rs2
        DataReport1.Sections(2).Controls("label8").Caption = rs2!employee
        DataReport1.Sections(2).Controls("label9").Caption = rs2!EMPID
       ' DataReport1.PrintReport , rptRangeFromTo, 1, 1
        DataReport1.Show
        rs1.MoveNext
        Unload DataReport1
    Wend
       db1.Close
       
ElseIf ans = vbNo Then
Exit Sub

End If
       
Case 1

If Combo2.Text = " " Then MsgBox "Select From Employees Status List", vbInformation, "Time Keeper": Exit Sub
data_base = "dbemployment"
'uid = "sa"
'pwd = "gama27"

    db1.Open "Driver={SQL Server};Server=" & data_source & ";Database=" & data_base & ";Uid=" & uid & ";Pwd=" & pwd
    rs1.Open "select * from db_finger_print WHERE EMPSTATUS = '" & Combo2.Text & "'", db1, adOpenKeyset, adLockOptimistic
    frmAttd.pbar.Min = 0
    frmAttd.Show
    frmAttd.pbar.Max = rs1.RecordCount
    While Not rs1.EOF
    DoEvents
    
    frmAttd.pbar.value = frmAttd.pbar + 1
        Set rs2 = db1.Execute("SELECT * FROM TK_TIMEKEEPING where estat = '" & Combo2.Text & "' order by EMPLOYEE")
         'Set rs2 = db1.Execute("SELECT * FROM TK_TIMEKEEPING WHERE EMPSTATUS  = '" & Combo2.Text & "' order by lname")

        Set DataReport1.DataSource = rs2
        DataReport1.Sections(2).Controls("label8").Caption = rs2!employee
        DataReport1.Sections(2).Controls("label9").Caption = rs2!EMPID
        'DataReport1.PrintReport , rptRangeFromTo, 1, 1
        DataReport1.Show
        rs1.MoveNext
        Unload DataReport1
    Wend
       db.Close
       
    
Case 2

Frame5.Visible = False


End Select

End Sub

Private Sub Form_Load()

Frame5.Visible = False
With Combo1
    .AddItem "CASUAL"
    .AddItem "REGULAR"
    .AddItem "OJT"
    .AddItem " "
End With

Combo1.Enabled = False
End Sub

Private Sub stat_Click(Index As Integer)

Select Case Index

Case 0

If stat(0).value = 1 Then
Combo1.Enabled = True
stat(1).value = 0
Else
Combo1.Enabled = False
Combo1.Text = " "
End If

Case 1

If stat(1).value = 1 Then
    txtid.Enabled = True
    txtid = ""
    stat(0).value = 0
Else
    txtid.Enabled = False
    
End If


Case 2

If stat(2).value = 1 Then
    Text2.Enabled = True
    Text2 = ""
Else
    Text2.Enabled = False
    
End If



End Select


End Sub
