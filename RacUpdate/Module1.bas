Attribute VB_Name = "Module1"
Function GetPath(File As String)



    For a = 1 To Len(File$) 'begin Loop
        b = Right(File$, a) 'get the last letter/s
        If Left(b, 1) = "\" Then Exit For        'check if the begining of the String above is a slash then exit the Loop
    Next a
    
    c = Left(File$, Len(File$) - Len(b) + 1) 'get the letters from the begining of the String To the final '\' in the string
    GetPath = c
End Function

Function GetPathFileN(File As String)



    For a = 1 To Len(File$) 'begin Loop
        b = Right(File$, a) 'get the last letter/s
        If Left(b, 1) = "\" Then
        b = Right(File$, a - 1)
        Exit For 'check if the begining of the String above is a slash then exit the Loop
        End If
    Next a

    GetPathFileN = b
End Function
