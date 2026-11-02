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
' ADO to SQL Server dependency. Needs two tables for UserRecord and TodoRecord


' Fill combo with all usernames
Private Sub Form_Load()
    Dim s1 As String
    Dim users() As String
    
    users = GetAllUserNames()
    
    
    Combo1.Clear
    If UBound(users) >= LBound(users) Then
        For i = LBound(users) To UBound(users)
            Combo1.AddItem users(i)
        Next i
    End If

End Sub

' Change userid => fill List1 with Todos
Private Sub Combo1_Click()

    Dim selectedID As Long
        
    If Combo1.ListIndex <> -1 Then
        ' Extract the hidden tracking index mapping
        selectedID = Combo1.ItemData(Combo1.ListIndex)
        'MsgBox "Selected Database Primary Key: " & selectedID & Combo1.ListIndex
        'Debug.Print "Selected Database Primary Key: " & selectedID
        GetTodos (Combo1.ListIndex + 1)
        
    End If
End Sub


' Fills List1 with Users Todos
Private Sub GetTodos(UserID As Long)
    Dim todos() As String
    Dim users() As String
    
    todos = GetToDo(UserID, "dddd, mmmm d hh:mm:ss")
    
    If IsArray(todos) Then
        If UBound(todos) >= LBound(todos) Then
            List1.Clear
            For i = LBound(todos) To UBound(todos)
                List1.AddItem todos(i)
            Next i
        End If
    End If

End Sub

' Get Todo list per User Id
Function GetToDo(intTodoID As Long, dateFormat As String) As String()
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
    
    On Error GoTo ErrorHandler
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' TODO: Move cs!
    connString = "Provider=SQLOLEDB;Data Source=TELLI;Initial Catalog=TODO;Integrated Security=SSPI;"
    
    conn.ConnectionString = connString
    conn.Open
    
    sqlStr = "SELECT * FROM Todo WHERE userid = " & CStr(intTodoID)
    
    ' Use forward-only, read-only for maximum speed
    rs.Open sqlStr, conn, adOpenForwardOnly, adLockReadOnly
    
    Dim rows As Variant
    
    rows = rs.GetRows
    totalRows = UBound(rows, 2)
    ReDim todorec(0 To totalRows)

    rs.MoveFirst
     
          
    dataMatrix = rs.GetRows()
    
    totalRows = UBound(dataMatrix, 2) ' Second dimension is Rows (0 to N)
    'Debug.Print totalRows
    
    ' Loop through the rows matrix to fill the structured string array
    For i = 0 To totalRows
        ' Note: The first parameter represents the column index (0, 1, 2). Append "" to handle NULL entries safely without runtime crashes
        todorec(i).TodoID = dataMatrix(0, i)
        todorec(i).TodoName = dataMatrix(1, i)
        todorec(i).TodoDetails = dataMatrix(2, i) & ""
        todorec(i).TodoDate = dataMatrix(3, i)
        todorec(i).UserID = dataMatrix(4, i)
    Next i
    

    rs.Close
    Set rs = Nothing
    

    conn.Close
    Set conn = Nothing
    
    totalRows = UBound(dataMatrix, 2)
    ReDim todos(0 To totalRows)
    
    ' create string from db row
    For i = 0 To totalRows
        todos(i) = todorec(i).TodoName & listdelim & CStr(Format$(todorec(i).TodoDate, dateFormat))
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


' Get Users names
Function GetAllUserNames() As String()
    Dim todos() As String

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connString As String
    Dim sqlStr As String
    Dim i As Long
    
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
    
    sqlStr = "SELECT UserName, Userid FROM Users"
    
    ' Use forward-only, read-only for maximum speed
    rs.Open sqlStr, conn, adOpenForwardOnly, adLockReadOnly
    
    Dim rows As Variant
    
    rows = rs.GetRows
    totalRows = UBound(rows, 2)
    ReDim users(0 To totalRows)

    'ReDim userrec(rs.GetRows)
    rs.MoveFirst
    i = 0
     
    ' Loop through the results if records exist
    Do While Not rs.EOF
            
            users(i) = rs.Fields(0).Value
            Combo1.AddItem rs.Fields(0).Value & ""
            Combo1.ItemData(Combo1.NewIndex) = rs.Fields(1).Value & ""
            MsgBox Combo1.ItemData(Combo1.NewIndex)
            i = i + 1
            rs.MoveNext
    Loop
           
    GetAllUserNames = users
    

    rs.Close
    Set rs = Nothing
    
    ' Clean up Connection
    conn.Close
    Set conn = Nothing
   
   
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


' TESTER: Get Todos
Private Sub Command1_Click()

Dim s1 As String
Dim todos() As String
Dim users() As String

todos = GetToDo(2, "dddd, mmmm d hh:mm:ss")
If UBound(todos) >= LBound(todos) Then
    For i = LBound(todos) To UBound(todos)
        List1.AddItem todos(i)
    Next i
End If

users = GetAllUserNames()

Debug.Print (users(1))

'Combo1.Clear
If UBound(users) >= LBound(users) Then
    For i = LBound(users) To UBound(users)
        'Combo1.AddItem users(i)
    Next i
End If


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



