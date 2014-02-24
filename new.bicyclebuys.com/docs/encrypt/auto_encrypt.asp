<!--#INCLUDE virtual="/includes/template_cls.asp"-->
<!--#INCLUDE virtual="/includes/common.asp"-->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "select * from Payments where isnumeric(substring(ltrim(CreditCardNumber), 0, 3)) = 1 ", Conn, 1, 2, &H0001
do while not(Rs.EOF)
	g_Key = mid(g_KeyString,1,Len(Rs("CreditCardNumber")))
	response.write(Rs("PaymentID") & "; ")
	Rs.Fields("CreditCardNumber") = EnCrypt(Rs("CreditCardNumber"))
	Rs.Update
Rs.MoveNext
loop
Rs.Close
set Rs = nothing
%>
done