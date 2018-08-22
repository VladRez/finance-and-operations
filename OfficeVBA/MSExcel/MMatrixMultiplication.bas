Public Sub MatrixMultiplication()
Dim qtyMatrix(1, 1) As Integer
priceVector = Array(62, 49)
Dim costVector(2) As Integer


qtyMatrix(0, 0) = (100)
qtyMatrix(0, 1) = (50)
qtyMatrix(1, 0) = (200)
qtyMatrix(1, 1) = (20)

Debug.Print priceVector.Length
For i = 0 To 1
costVector(i) = 0
    For j = 0 To 1
        costVector(i) = costVector(i) + qtyMatrix(i, j) * priceVector(j)
        
    Next j
Debug.Print costVector(i)
Next i


'[8650, 13380]

End Sub