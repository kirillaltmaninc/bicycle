<%@ language="vbscript"%>
<html><body>
<%
dim connX,rs100
Set connX = Server.CreateObject("ADODB.Connection")
connX.ConnectionString ="Provider=SQLNCLI10;Server=10.0.0.66;Database=BBC_Prod;DataTypeCompatibility=80;User ID =webuserprod;Password=bbcwebUserprod;DataTypeCompatibility=80;"
'connX.ConnectionString ="Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webuserprod;Initial Catalog=BBC_PROD;Data Source=webserver"
connX.open 
    Set rs100 = Server.CreateObject("ADODB.Recordset")
    rs100.open "select * from agegroup", connX
	while not rs100.eof
		response.write rs100.fields(1)
		rs100.movenext
	wend
    rs100.close

	connX.close
	response.write("Hello world!")
%>
</body></html>
