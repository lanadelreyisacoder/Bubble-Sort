Option Explicit Off 'Para no tener que declarar i, j, k, etc dentro de ciclos for.
Imports System.Threading

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
			
			Console.Clear() 'Esto es porque ejecutamos desde el cmd y para que se borren los comandos que hacemos y la ruta en donde estamos ubicados.
			
			'---Animación---.
			For k = LBound(myArray) To UBound(myArray) Step 1
				
				Console.Write(myArray(k))
			
			Next k
			'------.
            
            Dim temp As Double
            
			'Estos son los pases.
            For i = LBound(myArray) + 1 To UBound(myArray) Step 1
                'Comparaciones.
                For j = LBound(myArray) To UBound(myArray) - i Step 1
                    
					'---Animación---.
					Console.ForegroundColor = ConsoleColor.DarkYellow
					Console.SetCursorPosition(j, 0)
					Console.Write(myArray(j))
					Console.SetCursorPosition(j + 1, 0)
					Console.Write(myArray(j + 1))
					
					Thread.Sleep(1000)
					'------.
					
                    If myArray(j) > myArray(j + 1) Then
						
                        temp = myArray(j)
                        myArray(j) = myArray(j + 1)
                        myArray(j + 1) = temp
						
						'---Animación---.
						Console.ForegroundColor = ConsoleColor.Green
						Console.SetCursorPosition(j, 0)
						Console.Write(myArray(j))
						Console.SetCursorPosition(j + 1, 0)
						Console.Write(myArray(j + 1))
						
						Thread.Sleep(1000)
						
						Console.ResetColor()
						Console.SetCursorPosition(j, 0)
						Console.Write(myArray(j))
						Console.SetCursorPosition(j + 1, 0)
						Console.Write(myArray(j + 1))
						'------.
						
                    Else
						
						'---Animación---.
						Console.ResetColor()
						Console.SetCursorPosition(j, 0)
						Console.Write(myArray(j))
						Console.SetCursorPosition(j + 1, 0)
						Console.Write(myArray(j + 1))
						'------.
					
					End If
                    
                Next j
            
            Next i

        End Sub
		
    End Class
	
End Module