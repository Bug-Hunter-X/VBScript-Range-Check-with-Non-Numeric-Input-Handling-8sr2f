Function to check if a number is within a range. 

```vbscript
Function IsWithinRange(number, min, max)
  If number >= min And number <= max Then
    IsWithinRange = True
  Else
    IsWithinRange = False
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
```

The issue is that the function doesn't handle non-numeric inputs gracefully. If you pass a string or other non-numeric value to `IsWithinRange`, it will cause a type mismatch error.