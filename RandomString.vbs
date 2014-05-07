Function RandomString(n)
 Dim i,s
 For i = 1 to n
    Randomize:s = s & Chr(Int((26 * Rnd) + 1) + 64)
 Next
 RandomString = s
End Function
