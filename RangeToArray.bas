Sub myArrayRange()

Dim iAmount() As Variant
Dim iNum As Integer

iAmount = Range("A1:A11")

For iNum = 1 To UBound(iAmount)
    Debug.Print iAmount(iNum, 1)
Next iNum

End Sub
