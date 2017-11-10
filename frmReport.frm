VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EMPLOYEE REPORT"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search"
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   9000
      Width           =   4815
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         CheckBox        =   -1  'True
         Format          =   108789761
         CurrentDate     =   39932
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SHOW LIST"
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
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   9615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   9900
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   3240
      Picture         =   "frmReport.frx":0000
      ScaleHeight     =   1935
      ScaleWidth      =   3855
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin MSComctlLib.ListView viewer 
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   9615
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   9960
      X2              =   120
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim valx As Integer


Dim dbcode As New clsDBCode
Dim lv As ListItem


Private Sub Command1_Click()
Call OnTop(True)
Set frmReport = Nothing
Unload Me

End Sub

Private Sub Command2_Click()
VIEWER.ListItems.Clear
VIEWER.Refresh

'dbcode.SqlConnect "dbemployment"
dbcode.MDB_Access_Connect
dbcode.Sql_SELECT_QUERY_Execute query


While Not dbcode.EOF
DoEvents

    Set lv = VIEWER.ListItems.Add(, , dbcode.GetFields("today"))
    
     lv.SubItems(1) = dbcode.GetFields("empid")
     lv.SubItems(2) = dbcode.GetFields("EMPLOYEE")
     
     If status = 1 Then
     lv.SubItems(3) = "IN"
     ElseIf status = 0 Then
     lv.SubItems(3) = "OUT"
     End If
     
     lv.SubItems(4) = dbcode.GetFields("DEPTNAME")
     
     If dbcode.GetFields("A_TIMEOUT") <> "" Then
     lv.SubItems(5) = dbcode.GetFields("A_TIMEOUT")
     End If
     
     If dbcode.GetFields("P_TIMEOUT") <> "" Then
     lv.SubItems(5) = dbcode.GetFields("P_TIMEOUT")
     End If
     
     If dbcode.GetFields("N_TIMEOUT") <> "" Then
     lv.SubItems(5) = dbcode.GetFields("N_TIMEOUT")
     End If
   
      dbcode.MoveNext
Wend
    
dbcode.ReleaseConnection

End Sub

Private Sub Command3_Click()



If Len(Text1) = 4 And IsNumeric(Text1) Then

Text1 = UCase(Text1)

For I = 1 To VIEWER.ListItems.Count
    If Left(VIEWER.ListItems(I).SubItems(1), Len(Text1)) = Text1 Then
    VIEWER.ListItems(I).Selected = True
    VIEWER.ListItems(I).EnsureVisible
    VIEWER.SetFocus
    Else
     VIEWER.ListItems(I).Selected = False
    End If
Next I
    VIEWER.Refresh
Else

For I = 1 To VIEWER.ListItems.Count
    If Left(VIEWER.ListItems(I).SubItems(2), Len(Text1)) = UCase(Text1) Then
    VIEWER.ListItems(I).Selected = True
    VIEWER.ListItems(I).EnsureVisible
    VIEWER.SetFocus
    Else
     VIEWER.ListItems(I).Selected = False
    End If
Next I
    VIEWER.Refresh
End If

End Sub

Private Sub DT_Change()
Dim EN As Integer

VIEWER.ListItems.Clear
VIEWER.Refresh

'dbcode.SqlConnect "dbemployment"
dbcode.MDB_Access_Connect
dbcode.Sql_SELECT_QUERY_Execute GENQRY & " where td = 0 AND TODAY = #" & DT.value & "#" & "order by a_timeOUT desc"


While Not dbcode.EOF
DoEvents

    Set lv = VIEWER.ListItems.Add(, , dbcode.GetFields("today"))
     lv.SubItems(1) = dbcode.GetFields("EMPID")
     
    lv.SubItems(2) = dbcode.GetFields("EMPLOYEE")

     
     If status = 1 Then
     lv.SubItems(3) = "IN"
     ElseIf status = 0 Then
     lv.SubItems(3) = "OUT"
     End If
     
     lv.SubItems(4) = dbcode.GetFields("DEPTNAME")
     lv.SubItems(5) = dbcode.GetFields("A_TIMEOUT")
     
      
      dbcode.MoveNext
      
Wend
    
dbcode.ReleaseConnection

End Sub

Private Sub Form_Activate()
DT.value = Date


With frmReport.VIEWER

    .ColumnHeaders.Add , , "DATE", 1000, 0
    '.ColumnHeaders.Add , , "TIME IN", 1500, 0
    '.ColumnHeaders.Add , , "TIME OUT", 1500, 0
    '.ColumnHeaders.Add , , "TIME IN", 1500, 0
    '.ColumnHeaders.Add , , "TIME OUT", 1500, 0
    '.ColumnHeaders.Add , , "OVERTIME OUT", 1500, 0
    .ColumnHeaders.Add , , "ID", 900
    .ColumnHeaders.Add , , "EMPLOYEE NAME", 3000
    .ColumnHeaders.Add , , "STATUS", 900
    .ColumnHeaders.Add , , "DEPARTMENT", 2000
    .ColumnHeaders.Add , , "TIME", 3000
End With




End Sub

Private Sub optd_Click(Index As Integer)

End Sub

