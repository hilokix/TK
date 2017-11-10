VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 1


Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_OK = &H0&
Private Const MB_RETRYCANCEL = &H5&
Private Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
Private Declare Sub GetLocalTime Lib "KERNEL32" (lpSystemTime As SYSTEMTIME)
Private db As New ADODB.Connection
Private rs As New ADODB.Recordset
Private cmd As New ADODB.Command


Public Function MDB_Access_Connect() As String


On Error GoTo ExceptionHandler

Dim fileread As Integer
Dim chars As String
    fileread = FreeFile()
    Open App.Path + "\DBaccess.INI" For Input As fileread
       ' chars = Input(LOF(fileread), fileread)
       Line Input #fileread, Header
       Line Input #fileread, NullChar
       Line Input #fileread, data_source
       Line Input #fileread, uid
       Line Input #fileread, pwd
    Close fileread
    intdatasource = InStr(data_source, ":")
    data_source = Mid(data_source, intdatasource + 1)
    intId = InStr(uid, ":")
    uid = Mid(uid, intId + 1)
    intpwd = InStr(pwd, ":")
    pwd = Mid(pwd, intpwd + 1)
If db.State = 1 Then
    db.Close
End If
db.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" + data_source + ";Uid=admin;Pwd=" & pwd
     
      Exit Function
ExceptionHandler:
MsgBox Err.Description
modular.XEND = 0
   Main.frerror.Caption = "MESSAGE"
   Main.frerror.Visible = True
   Main.lblerror.Caption = Err.Description & vbCr & _
   "Call MIS Support"
dbcode.ReleaseConnection


End Function

Public Function message(msg)
Call MessageBoxEx(Main.hwnd, msg, "Time Keeping", MB_OK + MB_ICONEXCLAMATION, 0)
End Function




Public Function SqlConnect(database As String) As String
'Provider=SQLOLEDB.1;Password=sultrec;Persist Security Info=True;User ID=kode;Initial Catalog=DBTIMEKEEPING;Data Source=underground
'ADODB_Connection.Open "driver={SQL Server};server=underground;uid=;pwd=;database=dbbilling"

On Error GoTo ExceptionHandler

Dim fileread As Integer
Dim chars As String
    fileread = FreeFile()
    Open App.Path + "\DB.INI" For Input As fileread
       ' chars = Input(LOF(fileread), fileread)
       Line Input #fileread, Header
       Line Input #fileread, NullChar
       Line Input #fileread, data_source
       Line Input #fileread, uid
       Line Input #fileread, pwd
    Close fileread
    
    intdatasource = InStr(data_source, ":")
    data_source = Mid(data_source, intdatasource + 1)
    intId = InStr(uid, ":")
    uid = Mid(uid, intId + 1)
    intpwd = InStr(pwd, ":")
    pwd = Mid(pwd, intpwd + 1)
   
If db.State = 1 Then
    db.Close
End If

      db.Open "Driver={SQL Server};Server=" & data_source & ";Database=" & database & ";Uid=" & uid & ";Pwd=" & pwd
      SqlConnect = db
      Exit Function
ExceptionHandler:
modular.XEND = 0
   Main.frerror.Caption = "MESSAGE"
   Main.frerror.Visible = True
   Main.lblerror.Caption = Err.Description & vbCr & _
   "Call MIS Support" & " SqlConnect Function Error"
  
End Function


Public Function Sql_SELECT_QUERY_Execute(query As String, Optional fld1 As String) As Boolean

On Error GoTo handler
If rs.State = 1 Then
    rs.Close
End If
If db.State = 0 Then Exit Function
Debug.Print query
1:    rs.Open query, db, adOpenKeyset, adLockOptimistic
2:    If rs.BOF <> True Or rs.EOF <> True Then
        Sql_SELECT_QUERY_Execute = False
        Debug.Print query
        Else
        Sql_SELECT_QUERY_Execute = True
        Located = Sql_SELECT_QUERY_Execute
      End If
      
    If fld1 <> "" Then
    If Located <> True Then
        If IsNull(rs.Fields(fld1).value) Then
            Sql_SELECT_QUERY_Execute = True
        Else
            Sql_SELECT_QUERY_Execute = False
        End If
    End If
    End If
    Exit Function
    
handler:
    Select Case Err.Number
        Case 3021
        MsgBox Cerr
            SqlExecute = False
            modular.XEND = 0
            Main.frerror.Caption = "MESSAGE"
            Main.frerror.Visible = True
            Main.lblerror.Caption = Err.Description & " " & Erl & " Sql_SELECT_QUERY_Execute Error"
            Call WriteErrorLog(Err)
        Case Else
            SqlExecute = False
            modular.XEND = 0
            Main.frerror.Caption = "MESSAGE"
            Main.frerror.Visible = True
            Main.lblerror.Caption = Err.Description & " " & Erl & " Sql_SELECT_QUERY_Execute Error"
            Call WriteErrorLog(Err)
    End Select
End Function


Public Function GetTimeLocal() As String
Dim ltime As SYSTEMTIME
Dim timex As String
     GetLocalTime ltime
     timex = ltime.wHour & ":" & ltime.wMinute
     GetTimeLocal = timex
End Function


Public Function GetDateLocal() As String
Dim days As String
Dim ltime As SYSTEMTIME
Dim timex As String
intdays = Array(1, 2, 3, 4, 5, 6, 7)
strdays = Array("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday")
GetLocalTime ltime
intemp = ltime.wDayOfWeek
    datex = ltime.wMonth & "\" & ltime.wDay & "\" & ltime.wYear
    GetDateLocal = datex
End Function

Public Function GetFields(FIELD As String) As String

On Error GoTo errhandler

If rs.EOF = True Or rs.BOF = True Then GetFields = "": Exit Function
   If IsNumeric(FIELD) Then
    If IsNull(rs.Fields(FIELD).value) Then
        Else
            GetFields = rs.Fields(FIELD).value
    End If
        Else
    If IsNull(rs.Fields(FIELD).value) Or rs.EOF Then
        Else
            GetFields = rs.Fields(FIELD).value
            SZ = rs.Fields(FIELD).ActualSize
    End If
   End If
   Exit Function
errhandler:
modular.XEND = 0
   Main.frerror.Caption = "MESSAGE"
   Main.frerror.Visible = True
   Main.lblerror.Caption = Err.Description & vbCr & _
   "Call MIS Support" & "GetfieldsFunction Error"
   Call WriteErrorLog(Err)
End Function

Public Function GetSingleFields(FIELD As String) As String

On Error GoTo errhandler
'If rs.EOF = True Or rs.BOF = True Then Exit Function
   If IsNumeric(FIELD) Then
    If IsNull(rs.Fields(FIELD).value) Then
        Else
            GetSingleFields = rs.Fields(FIELD).value
    End If
        Else
    If IsNull(rs.Fields(FIELD).value) Then
        Else
            GetSingleFields = rs.Fields(FIELD).value
    End If
   End If
   Exit Function
errhandler:
        Call errorhandler(Err)
End Function


Public Function EOF() As Boolean
    
    On Error GoTo ERRHAND
    EOF = rs.EOF
    Exit Function
ERRHAND:
    Call errorhandler(Err)

End Function

Public Function MoveNext()

On Error GoTo ERRHAND
    rs.MoveNext
Exit Function
ERRHAND:
    Call errorhandler(Err)
    
End Function

Public Function MoveFirst()

On Error GoTo ERRHAND
    rs.MoveFirst
Exit Function
ERRHAND:
    Call errorhandler(Err)
End Function

Public Function MoveLast()

On Error GoTo ERRHAND
    rs.MoveLast
    Exit Function
ERRHAND:
   Call errorhandler(Err)
    
End Function

Public Function MovePrev()

On Error GoTo ERRHAND
    rs.MovePrevious
    Exit Function
ERRHAND:
 Call errorhandler(Err)
    
End Function
 
Public Function CountAllRecords() As Long

On Error GoTo ERRHAND
Dim ctr As Long
If rs.EOF = True Then rs.MoveFirst
    Intrecords = 0
    While Not rs.EOF
        Intrecords = Intrecords + 1
        rs.MoveNext
    Wend
      CountAllRecords = Intrecords
    Exit Function
ERRHAND:
    Call errorhandler(Err)
End Function

Private Function errorhandler(er As ErrObject)

Dim intFile As Integer
    Select Case er.Number
    
       Case 3704
       Call WriteErrorLog(er)
           
       Case -2147217900
       Call WriteErrorLog(er)
    
       Case Else
       Call WriteErrorLog(er)
    End Select

End Function


Public Function WriteErrorLog(er As ErrObject)

Dim retval As Long
Dim intFile As Integer
retval = MessageBoxEx(Main.hwnd, er.Description + " Call Mis Support!", "Time Keeping", MB_OK + MB_ICONEXCLAMATION, 0)
If retval = 1 Then

intFile = FreeFile()
 If Dir(App.Path + "\log.txt") <> "" Then
            Open App.Path + "\log.txt" For Append As intFile
            Print #intFile, er.Number, er.Description
            Close #intFile
  Else
            Open App.Path + "\log.txt" For Output As intFile
            Print #intFile, er.Number, er.Description
            Close #intFile
  End If
   End
 End If
End Function


Public Function RecToString() As String

currpos = rs.AbsolutePosition
If currpos = -3 Then
    rs.MoveFirst
    RecToString = rs.GetString
    Else
    RecToString = rs.GetString
End If
End Function

'rem THIS FUNCTION IS MADE FOR SQL CREATE VIEW, INSERT, UPDATE, DELETE AND OTHER NON SELECT QUERY FUNCTION

Public Function SqlCommand_NON_SELECT_QUERY_Execute(query As String, strView As String) As Boolean

Debug.Print query
On Error GoTo ERRHAND
With cmd
        .ActiveConnection = db
        .CommandType = adCmdText
        .CommandText = query
        Set rs = .Execute(.CommandText)
       
End With
Exit Function
ERRHAND:
modular.XEND = 0
   Main.frerror.Caption = "MESSAGE"
   Main.frerror.Visible = True
   Main.lblerror.Caption = Err.Description & vbCr & _
   "Call MIS Support" & " SqlCommand_NON_SELECT_QUERY_Execute Function Error"
   Call WriteErrorLog(Err)
End Function

Public Function SqlDropViewCommand(query As String) As Boolean

Dim droped As Boolean
On Error GoTo ERRHAND
    With cmd
        .ActiveConnection = db
        .CommandType = adCmdText
        .CommandText = query
        .Execute .CommandText
         droped = True
         SqlDropViewCommand = droped
    End With
    Exit Function
ERRHAND:
droped = False
SqlDropViewCommand = droped
modular.XEND = 0
   Main.frerror.Caption = "MESSAGE"
   Main.frerror.Visible = True
   Main.lblerror.Caption = Err.Description & vbCr & _
   "Call MIS Support" & " SqlDropViewCommand Function Error"
   Call WriteErrorLog(Err)
End Function


Public Function ReleaseConnection()

On Error GoTo ERRHAND
If rs.State = 1 Then
    rs.Close
End If

If db.State = 1 Then
    db.Close
End If

    Set rs = Nothing
    Set db = Nothing
Exit Function
ERRHAND:
   modular.XEND = 0
   Main.frerror.Caption = "MESSAGE"
   Main.frerror.Visible = True
   Main.lblerror.Caption = Err.Description & vbCr & _
   "Call MIS Support" & " ReleaseConnection() Function Error"
End Function



Public Function OUT_IMAGE_SZ(IDX As String) As Boolean

On Error GoTo errhandler
Dim BITS() As Byte
Dim SZ As Long
If rs.State = 1 Then
rs.Close
End If
rs.Open "select PIX from db_finger_print where empno = " & IDX & ""
    SZ = rs.Fields("PIX").ActualSize
    If SZ = 0 Then Exit Function
    BITS() = rs.Fields("PIX").GetChunk(SZ)
    Open App.Path + "\TEMP" For Binary Access Write As #1
        Put #1, , BITS()
    Close #1
    Exit Function
errhandler:
   modular.XEND = 1
   Main.frerror.Caption = "MESSAGE"
   Main.frerror.Visible = True
   Main.lblerror.Caption = Err.Description & vbCr & _
   "Call MIS Support" & "  OUT_IMAGE_SZ() Function Error"
   Call WriteErrorLog(Err)


End Function

Public Function Sql_SELECT_QUERY_Report(query As String, view As Integer)

Debug.Print query
On Error GoTo handler
If rs.State = 1 Then
    rs.Close
End If
If db.State = 0 Then Exit Function
If view = 1 Then
1:    rs.Open query, db, adOpenKeyset, adLockOptimistic
2:    If rs.BOF <> True Or rs.EOF <> True Then
      Set DataReport1.DataSource = rs
      End If
End If
If view = 2 Then
3:    rs.Open query, db, adOpenKeyset, adLockOptimistic
4:    If rs.BOF <> True Or rs.EOF <> True Then
      Set DataReport2.DataSource = rs
      End If
      
End If
    Exit Function
handler:
    Select Case Err.Number
    
        Case 3021
            SqlExecute = False
            modular.XEND = 0
            Main.frerror.Caption = "MESSAGE"
            Main.frerror.Visible = True
            Main.lblerror.Caption = Err.Description & " " & Erl & " Sql_SELECT_QUERY_Execute Error"
            dbcode.ReleaseConnection
            Call WriteErrorLog(Err)
        Case Else
            SqlExecute = False
            modular.XEND = 0
            Main.frerror.Caption = "MESSAGE"
            Main.frerror.Visible = True
            Main.lblerror.Caption = Err.Description & " " & Erl & " Sql_SELECT_QUERY_Execute Error"
            Debug.Print query
            Call WriteErrorLog(Err)

    End Select


End Function




