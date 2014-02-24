<%
	Dim g_KeyLocation, g_Key, g_KeyString
	g_KeyLocation = "D:\root\new.bicyclebuys.com\crypt\crypt_key.txt"
	g_KeyString = ReadKeyFromFile(g_KeyLocation)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>DeCrypt</title>
</head>

<body>
<form id="form1" name="form1" method="post" action="decrypt.asp">
  <p>Enter Encrypted String to De-Crypt</p>
  <p>
    <textarea name="Crypt" id="Crypt" cols="45" rows="3"></textarea>
    <br />
    <input type="submit" name="button" id="button" value="Submit" />
  </p>
</form>
<p>&nbsp;</p>

<%
	DIM CRYPT
	CRYPT = request.form("Crypt")

	if (CRYPT <> "") then
		g_Key = mid(g_KeyString,1,Len(CRYPT))

	%>

	<p>Decrypted Value: <%= DeCrypt(CRYPT) %></p>

	<%
	end if
%>


</body>
</html>

<%

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