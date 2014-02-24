


 <html>
 <head>
 <title>Checkout </title>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
 <script type="text/javascript" language="javascript">
 function submitform(){
	document.payPalform.submit();
 }
 
 </script>
 <style>
	
#load {
    background: none repeat scroll 0 0 #F7F7F7;
    border: 3px double #999999;
    font-family: "Trebuchet MS",verdana,arial,tahoma;
    font-size: 14pt;
    height: 170px;
    left: 46%;
    line-height: 170px;
    margin-left: -150px;
    margin-top: -150px;
    position: absolute;
    text-align: center;
    top: 50%;
    width: 420px;
    z-index: 1;
}
</style>
 </head>
 <body onload="submitform();">
<div style="" id="load">Your Order is Being Processed...Please Wait</div>
		<form action="https://www.paypal.com/cgi-bin/webscr" method="post" name="payPalform">
			<input type="hidden" name="business" value="<%=Request.QueryString("business")%>">
			<input type="hidden" name="cmd" value="_cart">
			<input type="hidden" name="upload" value="1">
			<input type="hidden" name="item_name_1" value="<%=Request.QueryString("item_name_1")%>">
			<input type="hidden" name="item_name_2" value="<%=Request.QueryString("item_name_2")%>">
			<input type="hidden" name="item_name_3" value="<%=Request.QueryString("item_name_3")%>">
			<input type="hidden" name="item_name_4" value="<%=Request.QueryString("item_name_4")%>">
			<input type="hidden" name="item_name_5" value="<%=Request.QueryString("item_name_5")%>">
			<input type="hidden" name="item_name_6" value="<%=Request.QueryString("item_name_6")%>">
			<input type="hidden" name="item_name_7" value="<%=Request.QueryString("item_name_7")%>">
			<input type="hidden" name="item_name_8" value="<%=Request.QueryString("item_name_8")%>">
			<input type="hidden" name="item_name_9" value="<%=Request.QueryString("item_name_9")%>">
			<input type="hidden" name="item_number_1" value="<%=Request.QueryString("item_number_1")%>">
			<input type="hidden" name="item_number_2" value="<%=Request.QueryString("item_number_2")%>">
			<input type="hidden" name="item_number_3" value="<%=Request.QueryString("item_number_3")%>">
			<input type="hidden" name="item_number_4" value="<%=Request.QueryString("item_number_4")%>">
			<input type="hidden" name="item_number_5" value="<%=Request.QueryString("item_number_5")%>">
			<input type="hidden" name="item_number_6" value="<%=Request.QueryString("item_number_6")%>">
			<input type="hidden" name="item_number_7" value="<%=Request.QueryString("item_number_7")%>">
			<input type="hidden" name="item_number_8" value="<%=Request.QueryString("item_number_8")%>">
			<input type="hidden" name="item_number_9" value="<%=Request.QueryString("item_number_9")%>">
			<input type="hidden" name="amount_1" value="<%=Request.QueryString("amount_1")%>">
			<input type="hidden" name="amount_2" value="<%=Request.QueryString("amount_2")%>">
			<input type="hidden" name="amount_3" value="<%=Request.QueryString("amount_3")%>">
			<input type="hidden" name="amount_4" value="<%=Request.QueryString("amount_4")%>">
			<input type="hidden" name="amount_5" value="<%=Request.QueryString("amount_5")%>">
			<input type="hidden" name="amount_6" value="<%=Request.QueryString("amount_6")%>">
			<input type="hidden" name="amount_7" value="<%=Request.QueryString("amount_7")%>">
			<input type="hidden" name="amount_8" value="<%=Request.QueryString("amount_8")%>">
			<input type="hidden" name="amount_9" value="<%=Request.QueryString("amount_9")%>">
			<input type="hidden" name="shipping_1" value="<%=Request.QueryString("no_shipping")%>">
			<input type="hidden" name="tax_cart" value="<%=Request.QueryString("tax")%>">
			<input type="hidden" name="quantity_1" value="<%=Request.QueryString("quantity_1")%>">
			<input type="hidden" name="quantity_2" value="<%=Request.QueryString("quantity_2")%>">
			<input type="hidden" name="quantity_3" value="<%=Request.QueryString("quantity_3")%>">
			<input type="hidden" name="quantity_4" value="<%=Request.QueryString("quantity_4")%>">
			<input type="hidden" name="quantity_5" value="<%=Request.QueryString("quantity_5")%>">
			<input type="hidden" name="quantity_6" value="<%=Request.QueryString("quantity_6")%>">
			<input type="hidden" name="quantity_7" value="<%=Request.QueryString("quantity_7")%>">
			<input type="hidden" name="quantity_8" value="<%=Request.QueryString("quantity_8")%>">
			<input type="hidden" name="quantity_9" value="<%=Request.QueryString("quantity_9")%>">
			<input type="hidden" name="currency_code" value="USD">
			<input type="hidden" name="return" value="<%=Request.QueryString("return")%>">

			<input type="hidden" name="first_name" value="<%=Request.QueryString("fName")%>">
			<input type="hidden" name="last_name" value="<%=Request.QueryString("lName")%>">
			<input type="hidden" name="address1" value="<%=Request.QueryString("addr1")%>">
			<input type="hidden" name="city" value="<%=Request.QueryString("city")%>">
			<input type="hidden" name="state" value="<%=Request.QueryString("state")%>">
			<input type="hidden" name="zip" value="<%=Request.QueryString("zip")%>">
			<input type="hidden" name="country" value="<%=Request.QueryString("country")%>">
			<input type="hidden" name="email" value="<%=Request.QueryString("email")%>">
		</form>

			
			
			
 </body>
 </html>

<%

  'response.write(Request.ServerVariables("QUERY_STRING"))
  'response.end
%>