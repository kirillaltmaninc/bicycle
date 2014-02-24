<!--#INCLUDE virtual="/includes/template_cls.asp"-->
<!--#INCLUDE virtual="/includes/common.asp"-->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "SELECT * FROM Payments WHERE (CreditCardNumber IS NOT NULL AND CreditCardNumber <> '') AND (CreditCardNumberOLD IS NOT NULL OR CreditCardNumberOLD IS NOT NULL)  AND PaymentID >= 124538 ORDER BY PaymentID desc ", Conn, 1, 2, &H0001
do while not(Rs.EOF)
	g_Key = mid(g_KeyString,1,Len(Rs("CreditCardNumberOLD")))
	'Rs.Fields("CreditCardNumber") = EnCrypt(Rs("CreditCardNumberOLD"))
	'Rs.Update
Rs.MoveNext
loop
Rs.Close
set Rs = nothing
%>
done