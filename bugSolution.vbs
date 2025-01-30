Function calculateSum(a, b)
  'Explicitly convert inputs to numbers before adding to prevent string concatenation
  Dim numA, numB
  numA = CDbl(a)
  numB = CDbl(b)
  calculateSum = numA + numB
End Function