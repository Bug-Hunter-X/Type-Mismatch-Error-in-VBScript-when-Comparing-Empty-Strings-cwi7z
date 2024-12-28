Function MyFunction(param1, param2)
  'Improved code to handle potential type mismatch
  If Len(param1) = 0 Then
    Err.Raise 13, , "Parameter 1 cannot be empty"
  End If
End Function