Attribute VB_Name = "Module1"
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Sub CheckSeats()
Dim Rs As Recordset
With DataEnvironment1.Connection2
    .Open
    Set Rs = .Execute("SELECT * FROM SEATS WHERE SEAT_ALLOC<=0")
    
    If Rs.RecordCount > 0 Then
        .Execute "DELETE FROM SEATS WHERE SEAT_ALLOC <= 0"
    End If
    .Close
End With

End Sub

Function getCollegeNo(ColName As String) As String
On Error Resume Next
Dim findStr As String
With DataEnvironment1.rsCMD_COLLEGE
    .Open
     findStr = "col_name='" & ColName & "'"
    .Find findStr
    getCollegeNo = .Fields("col_no")
    .Close
End With
End Function
Function getCourseNo(CourseName As String) As String
On Error Resume Next
Dim findStr As String

With DataEnvironment1.rsCMD_COURSE
    .Open
    findStr = "course_name LIKE '" & CourseName & "'"
    .Find findStr
    getCourseNo = .Fields("course_no")
    .Close
End With
End Function

Function getResNo(ResName As String) As String
On Error Resume Next
Dim findStr As String

With DataEnvironment1.rsCMD_SPLRES
    .Open
    findStr = "res_name LIKE '" & ResName & "'"
    .Find findStr
    getResNo = .Fields("res_no")
    .Close
End With

End Function

Function getCourseFromCollege(ColName As String) As Recordset

Dim Q As String, Rs As Recordset

Q = "select * from course where course_no in " & _
    "( select course_no from seats where col_no like '" & getCollegeNo(ColName) & "')"

With DataEnvironment1.Connection2
    .Open
    Set getCourseFromCollege = .Execute(Q)
End With

End Function
Property Get IsWin95() As Boolean
    If Not mIsWin95Initialized Then
        Dim Os As OSVERSIONINFO, Ret As Long
        Os.dwOSVersionInfoSize = Len(Os)
        
        Ret = GetVersionEx(Os)
        
        mIsWin95 = Ret = 0 Or (Os.dwMajorVersion = 4 And Os.dwMinorVersion = 0)
    End If
    IsWin95 = mIsWin95
End Property

Function ToBounds(ByVal v As Long, ByVal min As Long, ByVal Max As Long) As Long
    If v < min Then
        ToBounds = min
    ElseIf v > Max Then
        ToBounds = Max
    Else
        ToBounds = v
    End If
End Function

