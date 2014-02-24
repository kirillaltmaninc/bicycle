<!--#INCLUDE VIRTUAL="/includes/template_cls.asp" -->
<!--#INCLUDE VIRTUAL="/includes/common.asp" -->
<!--#INCLUDE VIRTUAL="/includes/cartconfig.asp" --><%

response.buffer = true


' Process form input
vITEMID = Request.Form("ITEMID")
vITEMNUMBER = Request.Form("ITEMNUMBER")
vITEMNAME = Request.Form("ITEMNAME")
vQNum = "Q" & vITEMNUMBER
vQUANTITY = Request.Form(vQNum)
vORIGQUANTITY = Request.Form("ORIGQUANTITY")
vITEMWEIGHT = Request.Form("WEIGHT")
vCUSTOM1 = Request.Form("CUSTOM1")
vCUSTOM2 = Request.Form("CUSTOM2")
vCUSTOM3 = Request.Form("CUSTOM3")
vCUSTOM4 = Request.Form("CUSTOM4")
vCUSTOM5 = Request.Form("CUSTOM5")
vCUSTOM6 = Request.Form("CUSTOM6")
vCUSTOM7 = Request.Form("CUSTOM7")
vCUSTOM8 = Request.Form("CUSTOM8")
vProp = Request.Form("Prop")

' If the item id of the item in the cart doesn't match
' the itemid of the size/color selected then we need to
' change things.

IF not isnumeric(vQuantity) then	
	response.redirect "/displaycart/"
elseif vITEMID <> vProp or vQUANTITY <> vORIGQUANTITY then

   ' Cart.LoadCart(Session("Cart"))

   ' Response.write "<PRE>"
   ' response.write "Changing item: " & vITEMID & " to " & vProp & "<br>"

   If vQUANTITY <> vORIGQUANTITY then
      vChangeQTY = vQUANTITY - vORIGQUANTITY
      if vChangeQTY > 0 then
'         response.write vITEMID & " - " & vQUANTITY & " - " & vORIGQUANTITY & " - " & vChangeQTY
         Cart.AddtoCart vITEMID, vChangeQTY
      Else
         Cart.RemoveFromCart vITEMID, abs(vChangeQTY)
      End If
   Cart.Calculate
   Session("Cart") = Cart.SaveCart
   vORIGQUANTITY = vQUANTITY
   End If
   
  if vITEMID <> vProp then
      '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
      '% Create the Conn Object and open it
      '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
      dim conn2
      DSN = Application("DSN")
      Set Conn2 = Server.CreateObject("ADODB.Connection")
      Set rs = Server.CreateObject("ADODB.Recordset")
      Conn2.Open DSN

      vSQL = "SELECT description,price,marketingdescription FROM vwWebproducts WITH (NOLOCK) WHERE SKU='" & vProp & "' For Browse"
      ' response.write vSQL
      rs.open vSQL,Conn,3

     ' response.write "Product retrieved from DB: " & rs("SKU") & " - " & rs("description") & "<BR>"

      Cart.RemoveFromCart vITEMID, vORIGQUANTITY
      Cart.AddtoCart vProp, vQUANTITY, replace(rs("description"),"""", "''"), rs("price") , vITEMWEIGHT, rs("marketingdescription"), 0, 0, 0, 0, vCustom1, vCustom2, vCustom3, vCustom4, vCustom5, vCustom6, vCustom7, vCustom8

      rs.close
      Conn2.close
   End if

   Cart.Calculate
   Session("Cart") = Cart.SaveCart
   Set Cart = Nothing
end if   

response.redirect "/displaycart/"
%>