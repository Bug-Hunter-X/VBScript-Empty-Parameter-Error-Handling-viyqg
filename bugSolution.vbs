Function MyFunction(param1)
  On Error Resume Next
  If IsEmpty(param1) Then
    Err.Raise 5, , "Parameter cannot be empty"
  End If
  If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Description, vbCritical
    Err.Clear
    Exit Function ' Or handle the error appropriately
  End If
  ' ... rest of the function
End Function