Option Explicit Off 'Para no tener que declarar i, j, k, etc dentro de ciclos for.

Module VBModule
    Sub Main()
        
        Dim numeros As Double() = { 8, 4, 1, 2, 9, 7, 3 }
        
        Dim bs As New BubbleSort()
        bs.Sort(numeros)

        Console.ReadKey()
        
    End Sub
End Module

Module AnotherModule
    Public Class BubbleSort
        Public Sub Sort(ByRef myArray() As Double)
            
            Dim temp As Double
            
            For i = LBound(myArray) + 1 To UBound(myArray) Step 1
                
                For j = LBound(myArray) To UBound(myArray) - i Step 1
                    
                    If myArray(j) > myArray(j + 1) Then
                        temp = myArray(j)
                        myArray(j) = myArray(j + 1)
                        myArray(j + 1) = temp
                    End If
                    
                Next j
            
            Next i
            
            'Impimir colección.
            For k = LBound(myArray) To UBound(myArray) Step 1
                
                Console.Write(myArray(k) & " ")
                
            Next k

        End Sub
    End Class
End Module
