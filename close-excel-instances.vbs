Do While  True  
    Dim objExcel
    On Error Resume Next
    Set objExcel = GetObject(,"Excel.Application")
    If Err.Number <> 0 Then
        Exit Do
    End If
    On Error GoTo 0
    objExcel.DisplayAlerts = False
    objExcel.Quit
    Set objExcel = nothing
Loop 
