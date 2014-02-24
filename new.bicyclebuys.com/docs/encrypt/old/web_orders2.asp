<!--#INCLUDE virtual="/includes/template_cls.asp"-->
<!--#INCLUDE virtual="/includes/common.asp"-->
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Rs = Server.CreateObject("ADODB.Recordset")
Conn.Open "dsn=IISJK"
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "SELECT top 2500 * FROM Orders WHERE (CCNumber IS NOT NULL AND CCNumber <> '') AND (CCNumberENC IS NULL OR CCNumberENC = '') ORDER BY ID ", Conn, 1, 2, &H0001
do while not(Rs.EOF)
	g_Key = mid(g_KeyString,1,Len(Rs("CCNumber")))
	Rs.Fields("CCNumberENC") = EnCrypt(Rs("CCNumber"))
	Rs.Update
Rs.MoveNext
loop
Rs.Close
set Rs = nothing
%>
done