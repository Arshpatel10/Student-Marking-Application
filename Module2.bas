Attribute VB_Name = "Module2"
Option Explicit
Public Sub showForm()
    displayStudentEnrollment.Show
End Sub

Public Sub Astronomy(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'adjust column width for labels
    Worksheets("Students").Range("A1").ColumnWidth = 10
    Worksheets("Students").Range("B1").ColumnWidth = 10
    Worksheets("Students").Range("C1").ColumnWidth = 10
    Cells(1, 1) = "FirstName"
    Cells(1, 2) = "LastName"
    Cells(1, 3) = "StudentID"
    
    'select data
    SQL = "SELECT students.FirstName,students.LastName,grades.studentID FROM grades INNER JOIN students ON students.studentID=grades.studentID WHERE (grades.course='AS101')"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Students").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Public Sub Info(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'adjust column width for labels
    Worksheets("Students").Range("A1").ColumnWidth = 10
    Worksheets("Students").Range("B1").ColumnWidth = 10
    Worksheets("Students").Range("C1").ColumnWidth = 10
    Cells(1, 1) = "FirstName"
    Cells(1, 2) = "LastName"
    Cells(1, 3) = "StudentID"
    
    'select data
    SQL = "SELECT students.FirstName,students.LastName,grades.studentID FROM grades INNER JOIN students ON students.studentID=grades.studentID WHERE (grades.course='CP102')"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Students").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Public Sub introProg(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'adjust column width for labels
    Worksheets("Students").Range("A1").ColumnWidth = 10
    Worksheets("Students").Range("B1").ColumnWidth = 10
    Worksheets("Students").Range("C1").ColumnWidth = 10
    Cells(1, 1) = "FirstName"
    Cells(1, 2) = "LastName"
    Cells(1, 3) = "StudentID"
    
    'select data
    SQL = "SELECT students.FirstName,students.LastName,grades.studentID FROM grades INNER JOIN students ON students.studentID=grades.studentID WHERE (grades.course='CP104')"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Students").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Public Sub WAP(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'adjust column width for labels
    Worksheets("Students").Range("A1").ColumnWidth = 10
    Worksheets("Students").Range("B1").ColumnWidth = 10
    Worksheets("Students").Range("C1").ColumnWidth = 10
    Cells(1, 1) = "FirstName"
    Cells(1, 2) = "LastName"
    Cells(1, 3) = "StudentID"
    
    'select data
    SQL = "SELECT students.FirstName,students.LastName,grades.studentID FROM grades INNER JOIN students ON students.studentID=grades.studentID WHERE (grades.course='CP212')"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Students").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Public Sub compGraph(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'adjust column width for labels
    Worksheets("Students").Range("A1").ColumnWidth = 10
    Worksheets("Students").Range("B1").ColumnWidth = 10
    Worksheets("Students").Range("C1").ColumnWidth = 10
    Cells(1, 1) = "FirstName"
    Cells(1, 2) = "LastName"
    Cells(1, 3) = "StudentID"
    
    'select data
    SQL = "SELECT students.FirstName,students.LastName,grades.studentID FROM grades INNER JOIN students ON students.studentID=grades.studentID WHERE (grades.course='CP411')"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Students").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub
Public Sub digEc(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'adjust column width for labels
    Worksheets("Students").Range("A1").ColumnWidth = 10
    Worksheets("Students").Range("B1").ColumnWidth = 10
    Worksheets("Students").Range("C1").ColumnWidth = 10
    Cells(1, 1) = "FirstName"
    Cells(1, 2) = "LastName"
    Cells(1, 3) = "StudentID"
    
    'select data
    SQL = "SELECT students.FirstName,students.LastName,grades.studentID FROM grades INNER JOIN students ON students.studentID=grades.studentID WHERE (grades.course='PC120')"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Students").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Public Sub Mech(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    'adjust column width for labels
    Worksheets("Students").Range("A1").ColumnWidth = 10
    Worksheets("Students").Range("B1").ColumnWidth = 10
    Worksheets("Students").Range("C1").ColumnWidth = 10
    Cells(1, 1) = "FirstName"
    Cells(1, 2) = "LastName"
    Cells(1, 3) = "StudentID"
    
    'select data
    SQL = "SELECT students.FirstName,students.LastName,grades.studentID FROM grades INNER JOIN students ON students.studentID=grades.studentID WHERE (grades.course='PC131')"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Students").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub
Public Sub lifeSci(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'adjust column width for labels
    Worksheets("Students").Range("A1").ColumnWidth = 10
    Worksheets("Students").Range("B1").ColumnWidth = 10
    Worksheets("Students").Range("C1").ColumnWidth = 10
    Cells(1, 1) = "FirstName"
    Cells(1, 2) = "LastName"
    Cells(1, 3) = "StudentID"
    
    'select data
    SQL = "SELECT students.FirstName,students.LastName,grades.studentID FROM grades INNER JOIN students ON students.studentID=grades.studentID WHERE (grades.course='PC141')"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Students").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


Sub displayUserForm2(control As IRibbonControl)
    displayStudentEnrollment.Show
End Sub
