Attribute VB_Name = "Module1"
Public Type UserRecord
    UserID As Long
    UserName As String
    UserType As String
End Type

Public Type TodoRecord
    TodoID As Long
    TodoName As String
    TodoDetails As String
    TodoDate As Date
    UserID As Long
End Type
