VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} displayAccessData 
   Caption         =   "Student Grades"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   4880
   OleObjectBlob   =   "displayAccessData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "displayAccessData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Aaverage_Click()
'open the selected file
    Dim cn As New ADODB.Connection
    Dim fn As String
    
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call average(cn)
    cn.Close
    Set cn = Nothing
End Sub

Private Sub standardDeviation_Click()
'open selected file
    Dim cn As New ADODB.Connection
    Dim fn As String
    
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call std(cn)
    cn.Close
    Set cn = Nothing
End Sub

Private Sub continue_Click()
'open selected file
    Dim cn As New ADODB.Connection
    Dim fn As String
'check which option was selected and run the subroutine
    If OptionButton1.Value = True Then
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call data(cn)
    cn.Close
    Set cn = Nothing
    
    ElseIf OptionButton2.Value = True Then
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call courseInfo(cn)
    cn.Close
    Set cn = Nothing
    
    ElseIf OptionButton4.Value = True Then
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call studentInfo(cn)
    cn.Close
    Set cn = Nothing
    End If
End Sub


Private Sub cancel_Click()
    displayAccessData.Hide
End Sub

Private Sub clear_Click()
    Range("A1:O500").ClearContents
End Sub




Private Sub word_Click()
    Dim word As word.Application
    Dim avg As Range
    Dim data As Range
    Dim chart As ChartObject
    Set word = New word.Application
    
'create document
    With word
        .Visible = True
        .Activate
        .Documents.Add

'write to document
    With .Selection
        .BoldRun
        .Font.Size = 14
        .TypeText ("Report")
        .TypeParagraph
         .BoldRun
        .TypeText ("The following report displays the student grades as well as the assignment averages, minimums, maximums and standard deviations")
        .TypeParagraph
        .Font.Size = 11
    Set avg = Worksheets("Data").Range("J1:J51")
    Set data = Worksheets("Data").Range("A402:F414")
    avg.Copy
    .Paste
    data.Copy
    .Paste
    For Each chart In ActiveSheet.ChartObjects
        chart.chart.ChartArea.Copy
        .Paste
        Next
    End With
End With
    word.ActiveDocument.SaveAs2 Environ("Application Report") & "\Desktop\" & ActiveSheet.Name & Format(Now, "yyyy-mm-dd hh-mm-ss") & ".docx"
    word.ActiveDocument.Close
    word.Quit
    
    
End Sub

Private Sub minMax_Click()
    Dim cn As New ADODB.Connection
    Dim fn As String
    
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call minMax(cn)
    cn.Close
    Set cn = Nothing
End Sub

Private Sub AssignmentAvg_Click()
    Dim cn As New ADODB.Connection
    Dim fn As String
    
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call average(cn)
    cn.Close
    Set cn = Nothing
End Sub

Private Sub StudentAvg_Click()
    Dim cn As New ADODB.Connection
    Dim fn As String

    
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & file
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call studAverage(cn)
    cn.Close
    Set cn = Nothing
    
End Sub
