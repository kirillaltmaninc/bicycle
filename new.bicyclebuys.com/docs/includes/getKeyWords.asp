<%     
Dim words(1001)
words(1000) = 0

Function addWord(astr)
    dim x, found
    astr=replace(replace(astr,"'",""),"""","")
    If words(1000) >= 999 Then Exit Function
    found = False
    For x = 1 To words(1000)
        If words(x) = astr Then
            found = True
            Exit For
        End If
    Next
    
    If Not found Then
        words(1000) = words(1000) + 1
        words(words(1000)) = astr
    End If

End Function


Sub myParseWords(str)
    dim s,e, astr
    str = replace(str,"/", " ")
    If words(1000) >= 999 Then Exit Sub
    s = 1
    e = 1
    e = InStr(s, str, " ")
    While e > 0
        
        astr = lcase(Mid(str, s, e - s))
        addWord (astr)
        s = e + 1
        e = InStr(s, str, " ")
    Wend
    e = Len(str) + 1
    astr = lcase(Mid(str, s, e - s))
    addWord (astr)
End Sub

Function clearWords()
    dim x
    For x = 1 To words(999)
        words(x) = ""
    Next
    words(1000) = 0
End Function

Function getWords()
    dim x
    getWords = ""
    For x = 1 To words(1000)
        getWords = getWords & words(x) & " "
    Next

End Function


%>