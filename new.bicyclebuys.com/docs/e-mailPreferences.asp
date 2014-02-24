<%
	dim dsn
	dsn = Application("dsn")
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open dsn
      	Set oRS1 = Server.CreateObject("ADODB.Recordset")


%>