# Excel_VBA_prime_evenodd_check


#Simple but often asked in the interview - A excel VBA program to test if a number is odd OR even Or Prime

Sub poe()
                
    i = InputBox("Enter the Number", "Enter value")
      
          a = i Mod 2 ' check a = 0 if i is divisble 2
                                       
           'condition for odd/even/prime check
          
            If a > 0 Then
                MsgBox i & " is odd number"
                For j = 2 To i
                   x = i \ j   'quotient
                   Z = i Mod j 'remainder
                   'if quotient is > or = to j and z = 0 will exit for and no. is not prime
                   If x >= j And Z = 0 Then
                    Exit For
                   End If
                   'if quotient = 1 and z = 0 the number is prime
                   If x = 1 And Z = 0 Then MsgBox i & " is prime number"
                Next
            ElseIf a = 0 Then
                MsgBox i & " is even number"
            End If
            'finally let's test 1 & 2 as well
            If i = 1 Or i = 2 Then MsgBox i & " is prime number"            
End Sub
