<%
response.buffer = true


Set Cart  = Server.CreateObject("iiscart2000.store")

%><!--#INCLUDE VIRTUAL="/includes/cartconfig.asp" --><%

Cart.LoadCart(Session("Cart"))

Cart.EmptyCart

Session("Cart") = Cart.SaveCart
Set Cart = Nothing
response.redirect "/includes/displaycart.asp" 
%>
