VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} displayStudentEnrollment 
   Caption         =   "Courses"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   4980
   OleObjectBlob   =   "displayStudentEnrollment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "displayStudentEnrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub run_Click()
    Dim cn As New ADODB.Connection
    Dim fn As String
    
    ' check which category is selected and perform sub based on choice
    If OptionButton1.Value = True Then
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call Astronomy(cn)
    cn.Close
    Set cn = Nothing
    
    ElseIf OptionButton2.Value = True Then
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call Info(cn)
    cn.Close
    Set cn = Nothing
    
    ElseIf OptionButton3.Value = True Then
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call introProg(cn)
    cn.Close
    Set cn = Nothing
    
    ElseIf OptionButton4.Value = True Then
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call WAP(cn)
    cn.Close
    Set cn = Nothing
    
    ElseIf OptionButton5.Value = True Then
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call compGraph(cn)
    cn.Close
    Set cn = Nothing
    
    ElseIf OptionButton6.Value = True Then
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call digEc(cn)
    cn.Close
    Set cn = Nothing
    
    ElseIf OptionButton7.Value = True Then
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call Mech(cn)
    cn.Close
    Set cn = Nothing
    
    ElseIf OptionButton8.Value = True Then
    fn = Application.GetOpenFilename
    With cn
        .ConnectionString = "Data Source=" & fn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    Call lifeSci(cn)
    cn.Close
    Set cn = Nothing
    End If
  displayStudentEnrollment.Hide
End Sub

Private Sub clear_Click()
    Worksheets("Students").Range("A1:I400").ClearContents
End Sub

Private Sub cancel_Click()
    displayStudentEnrollment.Hide
End Sub

