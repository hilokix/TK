VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMoe 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MOE"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMoe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Exit"
      Height          =   600
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   360
      Picture         =   "frmMoe.frx":000C
      ScaleHeight     =   2055
      ScaleWidth      =   3855
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmMoe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lst As ListItem


Dim dbx As New clsDBCode

Private Sub Command1_Click()
dbx.ReleaseConnection
Call OnTop(True)
Unload Me
End Sub

Private Sub Form_Load()

With ListView1

    .ColumnHeaders.Add , , "EmpID", 1000
    .ColumnHeaders.Add , , "MOE", 1500
    .ColumnHeaders.Add , , "Name", 2500

End With

dbx.MDB_Access_Connect

dbx.Sql_SELECT_QUERY_Execute "select * from tk_timekeeping where MOE = 'MOE: Manual'"

While Not dbx.EOF

Set lst = ListView1.ListItems.Add(, , dbx.GetFields("empid"))

    lst.SubItems(1) = dbx.GetFields("MOE")
    lst.SubItems(2) = dbx.GetFields("employee")

dbx.MoveNext


Wend

End Sub



