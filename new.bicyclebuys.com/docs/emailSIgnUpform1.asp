
  <%
   '%%%%% FUNCTIONS DEFINED FIRST %%%%%%%
   '%% CHECKSTRING - Make sure string doesn't contain any SQL munging characters
   FUNCTION CheckString (s, endchar)
   	pos = InStr(s, "'")
   	While pos > 0
   		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
   		pos = InStr(pos + 2, s, "'")
   	Wend
    CheckString="'" & s & "'" & endchar
   END FUNCTION
   

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create the Conn Object and open it
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'Filename = "d:\database\salemail.mdb"
	'FileDSN = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Filename & ";"
	' Modified on 3/6/04 by DD
	FileDSN = Application("FileDSN")

	Set Conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Conn.Open FileDSN


   if request.form("emailField") = "" then
      vError = 1
      vRequiredString =  "<FONT ID=""body"" COLOR=""#FF0000"">Required</FONT>"
   End if

   If vError <> 1 Then
      
      vsalemail = CheckString(request.form("emailField"),"")
   
      sql = "INSERT INTO main (email) VALUES ("&vsalemail&")"
      
      Conn.Execute(sql)
      Conn.Close
      'response.redirect "http://www.bicyclebuys.com"
   End If
%>
<script language="javascript">
	window.onload= function(){
		history.back();
}
</script>