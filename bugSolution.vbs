Function IsWithinRange(number, min, max)
  If IsNumeric(number) Then
    If number >= min And number <= max Then
      IsWithinRange = True
    Else
      IsWithinRange = False
    End If
  Else
    IsWithinRange = False ' Handle non-numeric input
  End If
End Function

' Example usage
Dim num
num = 5
If IsWithinRange(num, 1, 10) Then
  MsgBox num & " is within the range"
Else
  MsgBox num & " is not within the range"
End If

num = "abc"
If IsWithinRange(num, 1, 10) Then
  MsgBox num & " is within the range"
Else
  MsgBox num & " is not within the range"
End If