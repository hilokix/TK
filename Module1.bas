Attribute VB_Name = "modular"
Public XEND As Integer
Public dbcode As New clsDBCode
Public dbx As New clsDBCode
Public dby As New clsDBCode
Public query As String
Public status As Integer
Public method As String
Public mee As Integer

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const GENQRY = "SELECT * FROM DB_FINGER_PRINT INNER JOIN TK_TIMEKEEPING ON DB_FINGER_PRINT.EMPNO = TK_TIMEKEEPING.empID"

'Public Const GENQRY = "SELECT distinct TODAY,A_TIMEIN,A_TIMEOUT,P_TIMEIN,P_TIMEOUT FROM DB_FINGER_PRINT INNER JOIN TK_TIMEKEEPING ON DB_FINGER_PRINT.EMPNO = TK_TIMEKEEPING.empID"
Public Const UNIONQRY = "SELECT * FROM DB_FINGER_PRINT UNION ALL SELECT * FROM TK_TIMEKEEPING"
Public Const IVAL = 1

Public Sub OnTop(bOpt As Boolean)

    If bOpt = True Then Call SetWindowPos(Main.hWnd, -1, 0, 0, 0, 0, FLAGS)
    
    If bOpt = False Then Call SetWindowPos(Main.hWnd, -2, 0, 0, 0, 0, FLAGS)

End Sub

Public Sub INITIALIZED()

'If Emp_status(0).Value = True Then
           ' ID = DBCODE.GetFields("EMPNO")
           ' X = GENQRY & " WHERE FNAME = '" & FNAME & "'"
           'DBCODE.Sql_SELECT_QUERY_Execute GENQRY & " WHERE FNAME = '" & FNAME & "' AND TODAY = '" & Date & "'"
           
           
          ' If DBCODE.GetFields("TIMEIN") = "" Then
            ' DBCODE.SqlCommand_NON_SELECT_QUERY_Execute "INSERT INTO TK_TIMEKEEPING (EMPID,TODAY,TIMEIN) VALUES ( '" & ID & "','" & Date & "','" & Time & "')", ""
           'Else
          '  display.Caption = "YOU ARE ALREADY IN " & DBCODE.GetFields("FNAME")
         ' End If
      ' End If
       
       'If Emp_status(1).Value = True Then
        '    ID = DBCODE.GetFields("EMPNO")
        '    DBCODE.Sql_SELECT_QUERY_Execute GENQRY & " WHERE FNAME = '" & FNAME & "' AND TODAY = '" & Date & "'"

        '   If DBCODE.GetFields("BREAKOUT") = "" Then
        '     DBCODE.SqlCommand_NON_SELECT_QUERY_Execute "UPDATE TK_TIMEKEEPING SET BREAKOUT = '" & Time & "' WHERE FNAME = '" & G_FNAME & "'", ""
         '  Else
         '   display.Caption = "YOU ARE ALREADY IN " & DBCODE.GetFields("FNAME")
         ' End If
      ' End If
      
      

With Main.VIEWER
    .ColumnHeaders.Add , , "TIME IN", 2000, 0
    .ColumnHeaders.Add , , "BREAK OUT OUT", 2000, 0
    .ColumnHeaders.Add , , "BREAK IN", 2000, 0
    .ColumnHeaders.Add , , "TIME OUT", 2000, 0
    .ColumnHeaders.Add , , "EMPLOYEE NAME", 5000, 0
End With
'MsgBox Main.Caption
        

'=======================================================LTHUMB


 
End Sub
