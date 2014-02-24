hello
<%
	Dim g_KeyLocation, g_Key, g_KeyString
	g_KeyLocation = "D:\root\new.bicyclebuys.com\crypt\crypt_key.txt"
	g_KeyString = ReadKeyFromFile(g_KeyLocation)

Dim RegEx
Set RegEx = New regexp
RegEx.Pattern = "[0-9]{3}"
RegEx.Global = True
RegEx.IgnoreCase = True

SSNum = "abbb077-33-49s99"
if RegEx.Test(SSNum) = FALSE then
	response.write ("NO")
else
	response.write ("YES")
end if
Set RegEx = NOTHING

	DIM CRYPT
	CRYPT = request.form("Crypt")

	if (CRYPT <> "") then
		g_Key = mid(g_KeyString,1,Len(CRYPT))

	%>

	<p>Decrypted Value: <%= DeCrypt(CRYPT) %></p>

<%
	end if

Function EnCrypt(strCryptThis)
   Dim strChar, iKeyChar, iStringChar, i, iCryptChar, strEncrypted
   for i = 1 to Len(strCryptThis)
      iKeyChar = Asc(mid(g_Key,i,1))
      iStringChar = Asc(mid(strCryptThis,i,1))
      iCryptChar = iKeyChar Xor iStringChar
      strEncrypted =  strEncrypted & Chr(iCryptChar)
   next
   EnCrypt = strEncrypted
End Function

Function DeCrypt(strEncrypted)
Dim strChar, iKeyChar, iStringChar, i, iDeCryptChar, strDecrypted
   for i = 1 to Len(strEncrypted)
      iKeyChar = (Asc(mid(g_Key,i,1)))
      iStringChar = Asc(mid(strEncrypted,i,1))
      iDeCryptChar = iKeyChar Xor iStringChar
      strDecrypted =  strDecrypted & Chr(iDeCryptChar)
   next
   DeCrypt = strDecrypted
End Function

Function ReadKeyFromFile(strFileName)
   Dim keyFile, fso, f, ts
   set fso = Server.CreateObject("Scripting.FileSystemObject")
   set f = fso.GetFile(strFileName)
   set ts = f.OpenAsTextStream(1, -2)

   Do While not ts.AtEndOfStream
     keyFile = keyFile & ts.ReadLine
   Loop

   ReadKeyFromFile =  keyFile
End Function

%>