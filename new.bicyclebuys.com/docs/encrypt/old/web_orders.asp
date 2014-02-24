<!--#INCLUDE virtual="/includes/template_cls.asp"-->
<!--#INCLUDE virtual="/includes/common.asp"-->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "SELECT top 1000 * FROM WebOrders WHERE (CCNumber IS NOT NULL AND CCNumber <> '') AND (CCNumberENC IS NULL OR CCNumberENC = '') ORDER BY ID ", Conn, 1, 2, &H0001
'Rs.Open "SELECT top 5000 * FROM WebOrders WHERE (CCNumber IS NOT NULL AND CCNumber <> '')  ORDER BY ID ", Conn, 1, 2, &H0001
do while not(Rs.EOF)
	g_Key = mid(g_KeyString,1,Len(Rs("CCNumber")))
%>
<%= Rs("CustomerID") %> - <%= EnCrypt(Rs("CCNumber")) %><br>
<%
Rs.Fields("CCNumberENC") = EnCrypt(Rs("CCNumber"))
Rs.Update

Rs.MoveNext
loop
Rs.Close
set Rs = nothing
%>