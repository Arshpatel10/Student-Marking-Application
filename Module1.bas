Attribute VB_Name = "Module1"
Option Explicit
Public Sub showForm()
    displayAccessData.Show
End Sub

Public Sub data(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'format columns to fit data
    Worksheets("Data").Range("A1").ColumnWidth = 5
    Worksheets("Data").Range("B1").ColumnWidth = 10
    Worksheets("Data").Range("C1").ColumnWidth = 10
    Worksheets("Data").Range("D1").ColumnWidth = 5
    Worksheets("Data").Range("E1").ColumnWidth = 5
    Worksheets("Data").Range("F1").ColumnWidth = 5
    Worksheets("Data").Range("G1").ColumnWidth = 5
    Worksheets("Data").Range("H1").ColumnWidth = 10
    Worksheets("Data").Range("I1").ColumnWidth = 10
    Cells(1, 1) = "ID"
    Cells(1, 2) = "StudentID"
    Cells(1, 3) = "Course"
    Cells(1, 4) = "A1"
    Cells(1, 5) = "A2"
    Cells(1, 6) = "A3"
    Cells(1, 7) = "A4"
    Cells(1, 8) = "Midterm"
    Cells(1, 9) = "Final"
    
    'select data
    SQL = "SELECT * FROM grades"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Data").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub

Public Sub courseInfo(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'format columns to fit data
    Worksheets("Data").Range("A1").ColumnWidth = 5
    Worksheets("Data").Range("B1").ColumnWidth = 15
    Worksheets("Data").Range("C1").ColumnWidth = 15
    Cells(1, 1) = "ID"
    Cells(1, 2) = "CourseCode"
    Cells(1, 3) = "CourseName"
    
    'select data
    SQL = "SELECT * FROM courses"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Data").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub

Public Sub studentInfo(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'format columns to fit data
    Worksheets("Data").Range("A1").ColumnWidth = 15
    Worksheets("Data").Range("B1").ColumnWidth = 15
    Worksheets("Data").Range("C1").ColumnWidth = 15
    Cells(1, 1) = "FirstName"
    Cells(1, 2) = "LastName"
    Cells(1, 3) = "studentID"
    
    'select data
    SQL = "SELECT * FROM students"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Data").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub
Public Sub std(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL1, SQL2, SQL3, SQL4, SQLmid, SQLfinal As String
    
    'format columns to fit data
    Worksheets("Data").Range("A1").ColumnWidth = 10
    Worksheets("Data").Range("B1").ColumnWidth = 10
    Worksheets("Data").Range("C1").ColumnWidth = 10
    Worksheets("Data").Range("D1").ColumnWidth = 10
    Worksheets("Data").Range("E1").ColumnWidth = 10
    Worksheets("Data").Range("F1").ColumnWidth = 10
    Cells(408, 1) = "Standard Deviations"
    Cells(409, 1) = "A1"
    Cells(409, 2) = "A2"
    Cells(409, 3) = "A3"
    Cells(409, 4) = "A4"
    Cells(409, 5) = "Midterm"
    Cells(409, 6) = "Final"
    
    'select data
    SQL1 = "SELECT STDEV(A1) FROM grades"
    SQL2 = "SELECT STDEV(A2) FROM grades"
    SQL3 = "SELECT STDEV(A3) FROM grades"
    SQL4 = "SELECT STDEV(A4) FROM grades"
    SQLmid = "SELECT STDEV(MidTerm) FROM grades"
    SQLfinal = "SELECT STDEV(Exam) FROM grades"
    
    'write data
    With rs
        .Open SQL1, cn
        Do While Not .EOF
            Worksheets("Data").Range("A410").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL2, cn
        Do While Not .EOF
            Worksheets("Data").Range("B410").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL3, cn
        Do While Not .EOF
            Worksheets("Data").Range("C410").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL4, cn
        Do While Not .EOF
            Worksheets("Data").Range("D410").CopyFromRecordset rs
        Loop
        .Close
        .Open SQLmid, cn
        Do While Not .EOF
            Worksheets("Data").Range("E410").CopyFromRecordset rs
        Loop
        .Close
        .Open SQLfinal, cn
        Do While Not .EOF
            Worksheets("Data").Range("F410").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub

Public Sub studAverage(cn As ADODB.Connection)
    Dim A1 As Double
    Dim A2 As Double
    Dim A3 As Double
    Dim A4 As Double
    Dim Amid As Double
    Dim Afinal As Double
    Dim finalGrade As Double
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'format columns to fit data
    Worksheets("Data").Range("A1").ColumnWidth = 5
    Worksheets("Data").Range("B1").ColumnWidth = 10
    Worksheets("Data").Range("C1").ColumnWidth = 10
    Worksheets("Data").Range("D1").ColumnWidth = 5
    Worksheets("Data").Range("E1").ColumnWidth = 5
    Worksheets("Data").Range("F1").ColumnWidth = 5
    Worksheets("Data").Range("G1").ColumnWidth = 5
    Worksheets("Data").Range("H1").ColumnWidth = 10
    Worksheets("Data").Range("I1").ColumnWidth = 10
    
    Cells(1, 1) = "ID"
    Cells(1, 2) = "StudentID"
    Cells(1, 3) = "Course"
    Cells(1, 4) = "A1"
    Cells(1, 5) = "A2"
    Cells(1, 6) = "A3"
    Cells(1, 7) = "A4"
    Cells(1, 8) = "Midterm"
    Cells(1, 9) = "Final"
    
    'select data
    SQL = "SELECT * FROM grades"
    
    'write data
    With rs
        .Open SQL, cn
        Do While Not .EOF
            Worksheets("Data").Range("A2").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
    Worksheets("Data").Range("J1").Value = "Student Averages"
    
    ' calculate averages, with proper weighting
    For i = 2 To 51
        A1 = Range("D" & i).Value * 0.05
        A2 = Range("E" & i).Value * 0.05
        A3 = Range("F" & i).Value * 0.05
        A4 = Range("G" & i).Value * 0.05
        Amid = Range("H" & i).Value * 0.3
        Afinal = Range("I" & i).Value * 0.5

        finalGrade = A1 + A2 + A3 + A4 + Amid + Afinal
        Range("J" & i).Value = finalGrade
        Next

End Sub

Public Sub minMax(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL1, SQL2, SQL3, SQL4, SQL5, SQL6, SQL7, SQL8, SQL9, SQL10, SQL11, SQL12 As String
    
    'format columns to fit data
    Worksheets("Data").Range("A1").ColumnWidth = 10
    Worksheets("Data").Range("B1").ColumnWidth = 10
    Worksheets("Data").Range("C1").ColumnWidth = 10
    Worksheets("Data").Range("D1").ColumnWidth = 10
    Worksheets("Data").Range("E1").ColumnWidth = 10
    Worksheets("Data").Range("F1").ColumnWidth = 10
    
    Cells(402, 1) = "MIN"
    Cells(403, 1) = "A1"
    Cells(403, 2) = "A2"
    Cells(403, 3) = "A3"
    Cells(403, 4) = "A4"
    Cells(403, 5) = "Midterm"
    Cells(403, 6) = "Final"
    
    'select min data
    SQL1 = "SELECT MIN(A1) FROM grades"
    SQL2 = "SELECT MIN(A2) FROM grades"
    SQL3 = "SELECT MIN(A3) FROM grades"
    SQL4 = "SELECT MIN(A4) FROM grades"
    SQL5 = "SELECT MIN(MidTerm) FROM grades"
    SQL6 = "SELECT MIN(Exam) FROM grades"
    
    'write min data
    With rs
        .Open SQL1, cn
        Do While Not .EOF
            Worksheets("Data").Range("A404").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL2, cn
        Do While Not .EOF
            Worksheets("Data").Range("B404").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL3, cn
        Do While Not .EOF
            Worksheets("Data").Range("C404").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL4, cn
        Do While Not .EOF
            Worksheets("Data").Range("D404").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL5, cn
        Do While Not .EOF
            Worksheets("Data").Range("E404").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL6, cn
        Do While Not .EOF
            Worksheets("Data").Range("F404").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
    Cells(405, 1) = "MAX"
    Cells(406, 1) = "A1"
    Cells(406, 2) = "A2"
    Cells(406, 3) = "A3"
    Cells(406, 4) = "A4"
    Cells(406, 5) = "Midterm"
    Cells(406, 6) = "Final"
    
    'select max data
    SQL7 = "SELECT MAX(A1) FROM grades"
    SQL8 = "SELECT MAX(A2) FROM grades"
    SQL9 = "SELECT MAX(A3) FROM grades"
    SQL10 = "SELECT MAX(A4) FROM grades"
    SQL11 = "SELECT MAX(MidTerm) FROM grades"
    SQL12 = "SELECT MAX(Exam) FROM grades"
    
    'write max data
    With rs
        .Open SQL7, cn
        Do While Not .EOF
            Worksheets("Data").Range("A407").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL8, cn
        Do While Not .EOF
            Worksheets("Data").Range("B407").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL9, cn
        Do While Not .EOF
            Worksheets("Data").Range("C407").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL10, cn
        Do While Not .EOF
            Worksheets("Data").Range("D407").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL11, cn
        Do While Not .EOF
            Worksheets("Data").Range("E407").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL12, cn
        Do While Not .EOF
            Worksheets("Data").Range("F407").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub
Public Sub average(cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    Dim SQL1, SQL2, SQL3, SQL4, SQLmid, SQLfinal As String
    
    'format columns to fit data
    Worksheets("Data").Range("A1").ColumnWidth = 10
    Worksheets("Data").Range("B1").ColumnWidth = 10
    Worksheets("Data").Range("C1").ColumnWidth = 10
    Worksheets("Data").Range("D1").ColumnWidth = 10
    Worksheets("Data").Range("E1").ColumnWidth = 10
    Worksheets("Data").Range("F1").ColumnWidth = 10
    
    Range("A411").Value = "Assignment Averages"
    Cells(412, 1) = "A1"
    Cells(412, 2) = "A2"
    Cells(412, 3) = "A3"
    Cells(412, 4) = "A4"
    Cells(412, 5) = "Midterm"
    Cells(412, 6) = "Final"
    
    'select data
    SQL1 = "SELECT AVG(A1) FROM grades"
    SQL2 = "SELECT AVG(A2) FROM grades"
    SQL3 = "SELECT AVG(A3) FROM grades"
    SQL4 = "SELECT AVG(A4) FROM grades"
    SQLmid = "SELECT AVG(MidTerm) FROM grades"
    SQLfinal = "SELECT AVG(Exam) FROM grades"
    
    'write data
    With rs
        .Open SQL1, cn
        Do While Not .EOF
            Worksheets("Data").Range("A413").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL2, cn
        Do While Not .EOF
            Worksheets("Data").Range("B413").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL3, cn
        Do While Not .EOF
            Worksheets("Data").Range("C413").CopyFromRecordset rs
        Loop
        .Close
        .Open SQL4, cn
        Do While Not .EOF
            Worksheets("Data").Range("D413").CopyFromRecordset rs
        Loop
        .Close
        .Open SQLmid, cn
        Do While Not .EOF
            Worksheets("Data").Range("E413").CopyFromRecordset rs
        Loop
        .Close
        .Open SQLfinal, cn
        Do While Not .EOF
            Worksheets("Data").Range("F413").CopyFromRecordset rs
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub

Sub displayUserForm(control As IRibbonControl)
    displayAccessData.Show
End Sub
