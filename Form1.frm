VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ToDos"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   360
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get ToDo list"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Simple todo list app
' Works best for many users; for trainers with trainees



Private Sub Command1_Click()

Dim s1 As String
Dim todos() As String

todos = GetToDo(2, "dddd, mmmm d hh:mm:ss")
If UBound(todos) >= LBound(todos) Then
    For i = LBound(todos) To UBound(todos)
        List1.AddItem todos(i)
    Next i
End If

'MsgBox todos(0)
'Debug.Print (todos(1))

'test
's1 = InputBox("What is your name?", "INVOICER")
'Debug.Print s1


End Sub


Sub test()


    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connString As String
    Dim sqlStr As String
    

    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' Define your connection string (Modify with your server and DB details)
    connString = "Provider=SQLOLEDB;Data Source=TELLI;Initial Catalog=TODO;Integrated Security=SSPI;"
    
    On Error GoTo ErrorHandler
    
    ' Open the connection
    conn.ConnectionString = connString
    conn.Open
    
    MsgBox "Connection Successful!", vbInformation, "Connected"
    
    ' Example: Fetching data from a table
    sqlStr = "SELECT * FROM TODO"
    rs.Open sqlStr, conn, adOpenKeyset, adLockOptimistic
    
    ' Loop through the results if records exist
    Do While Not rs.EOF
        Debug.Print rs.Fields("Name").Value
        rs.MoveNext
    Loop
    
    ' Clean up Recordset
    rs.Close
    Set rs = Nothing
    
    ' Clean up Connection
    conn.Close
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Connection Failed"
    
    ' Ensure objects are released in case of failure
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Sub

Function GetToDo(intTodoID As Integer, dateFormat As String) As String()
    Dim todos() As String

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connString As String
    Dim sqlStr As String
    
    Dim users() As String
    Dim todorec() As TodoRecord
    Dim dataMatrix As Variant
    Dim rowIdx As Long
    Dim totalRows As Long
    Dim listdelim As String
    listdelim = vbTab
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    connString = "Provider=SQLOLEDB;Data Source=TELLI;Initial Catalog=TODO;Integrated Security=SSPI;"
    
    On Error GoTo ErrorHandler
    
    ' Open the connection
    conn.ConnectionString = connString
    conn.Open
    
    sqlStr = "SELECT * FROM Todo WHERE userid = " & CStr(intTodoID)
    
    ' Use forward-only, read-only for maximum speed
    rs.Open sqlStr, conn, adOpenForwardOnly, adLockReadOnly
    
    Dim rows As Variant
    
    rows = rs.GetRows
    totalRows = UBound(rows, 2)
    ReDim todorec(0 To totalRows)

    'ReDim userrec(rs.GetRows)
    rs.MoveFirst
     
    ' Loop through the results if records exist
    'Do While Not rs.EOF
        dataMatrix = rs.GetRows()
        
        totalRows = UBound(dataMatrix, 2) ' Second dimension is Rows (0 to N)
        Debug.Print totalRows
        'ReDim tuples(0 To totalRows)
        
        ' 4. Loop through the rows matrix to fill the structured string array
        For i = 0 To totalRows
            ' Note: The first parameter represents the column index (0, 1, 2). Append "" to handle NULL entries safely without runtime crashes
            todorec(i).TodoID = dataMatrix(0, i)
            todorec(i).TodoName = dataMatrix(1, i)
            todorec(i).TodoDetails = dataMatrix(2, i) & ""
            todorec(i).TodoDate = dataMatrix(3, i)
            todorec(i).UserID = dataMatrix(4, i)
        Next i
        
        'rs.MoveNext
    'Loop
           

    

    rs.Close
    Set rs = Nothing
    
    ' Clean up Connection
    conn.Close
    Set conn = Nothing
    
    totalRows = UBound(dataMatrix, 2)
    ReDim todos(0 To totalRows)
    
    For i = 0 To totalRows
        todos(i) = todorec(i).TodoName & listdelim & CStr(format$(todorec(i).TodoDate, dateFormat))
            ' Note: The first parameter represents the column index (0, 1, 2). Append "" to handle NULL entries safely without runtime crashes
            'todorec(i).TodoID = dataMatrix(0, i)
            'todorec(i).TodoName = dataMatrix(1, i)
            'todorec(i).TodoDetails = dataMatrix(2, i) & ""
            'todorec(i).TodoDate = dataMatrix(3, i)
            'todorec(i).UserID = dataMatrix(4, i)
    Next i
        
    GetToDo = todos
    
    Exit Function

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "DB"
    'MsgBox Err.Description
    
    ' Ensure objects are released in case of failure
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Function




