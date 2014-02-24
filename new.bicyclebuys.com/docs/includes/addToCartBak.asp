<!--#INCLUDE virtual="/includes/template_cls.asp" -->
<!--#INCLUDE virtual="/includes/common.asp" -->
<% response.buffer = false
dim mTracking
   ' Get form input
   vITEMID=Request.Form("ITEMID")
   
   vITEMNAME = Request.Form("ITEMNAME")
   vPRICE = Request.Form("PRICE")
   vURL = Request.Form("URL")
   vReferer = Request.Form("Referer")
   vParent = Request.Form("Parent")
   vFreeFreight = Request.Form("FreeFreight")
   vOverWeightFlags = Request.Form("OverWeightFlags")
   vProp = Request.Form("PropDATA")
   VPropID = Request.Form("PropIDDATA")
   Session("Referer") = vReferer
   mTracking = Request.Form("mTracking")
'if mTracking <> "" and right(Session("Referer"),5)<>right(mTracking,5) then
'	vReferer=vReferer & "|TK=" & mTracking
'	Session("Referer") = vReferer
'	mTracking=""
'end if



   if vProp <> "" then
      ' Get array's from the properties
      vPropA = Split(right(vProp, len(vProp)-13), ",")
      vPropIDA = Split(right(vPropID,len(vPropID)-15), ",")
   
      ' Stick the current SKU into PropID's value
      vPropID = replace(vPropID, "~", vITEMID)
   
      ' Stick the correct selection into Prop's value
      for i = 0 to ubound(vPropIDA)
   '    response.write "I=" & i & " Id=" & vITEMID & " PROP=" & vPropIDA(i) & "<BR>"
         
       if vITEMID = vPropIDA(i) then exit for
       
      next
   '   response.write "I=" & i & " Id=" & vITEMID & " PROP=" & vPropID
      vProp = replace(vProp, "~", vPropA(i))
   End If
   
   ' load in the product data
   oProd1.clearitem
   oProd1.getitemSKU(vITEMID)

   ' load the cart configuration and current cart
   %><!--#INCLUDE VIRTUAL="/includes/cartconfig.asp" --><%

   ' Here we add the product to the cart
   '
   ' Cart.AddToCart (ItemID, Quantity, ItemName, ItemPrice, ItemWeight, ItemDescription, ItemAdjustRate, ItemAdjust, ItemTaxRate, ItemTax,  
   '                 ItemCustom1,  ItemCustom2, ItemCustom3, ItemCustom4,  ItemCustom5,  ItemCustom6,  ItemCustom7,  ItemCustom8, 
   '                 Property1, Property2, Property3, Property4, Property5, Property6, Property7, Property8 )
   '
   ' ItemCustom1 = URL to view item
   ' ItemCustom2 = URL of page we were just on
   ' ItemCustom3 = Parent SKU (same if stand alone)
   ' ItemCustom4 = Size/Color Info
   ' ItemCustom5 = Size/Color Info
   ' ItemCustom6 = Overweight Type Flags
   ' ItemCustom7 = Free Freight Flag
   ' ItemCustom8 = Freight Selection/Price for overweight items

   ' clean up the descriptions
   '   get rid of single quotes
   vTMP1 = trim(oProd1.pfields.Item("description"))
   vTMP1 = replace(vTMP1 & " ","'", "")
   vTMP1 = left(replace(vTMP1 & " ","""", ""),255)
   'vTMP1 = CS(vTMP1, "")
   
   vTMP2 = oProd1.pfields.Item("marketingdescription") 
   vTMP2 = replace(vTMP2 & " ","'", "")
   vTMP2 = left(replace(vTMP2 & " ","""", ""),255)
   vTMP2 = CS(vTMP2, "")

   if oProd1.pfields.Item("mDiscountAmount") <> "0" then	
    	if oProd1.pfields.Item("mDiscountType") ="-1" then 'Dollar
		oProd1.pfields.Item("price") = oProd1.pfields.Item("price") - oProd1.pfields.Item("mDiscountAmount")
	else 'Percent
		oProd1.pfields.Item("price") = oProd1.pfields.Item("price") * (1- oProd1.pfields.Item("mDiscountAmount"))
	end if
   end if 	
   if vZeroTaxItems then 
      ' figure out if item is taxable
      ' -- critera
      '     item has webtypeid for clothes or shoes
      '     item price is under $110.01
   
   
      ' NEW 10/2006
      ' USE NEW 'DISCOUNTTAX' SCRIPTING DICTIONARY TO DETERMINE TAX
      Dim vWT
      vWT = oProd1.pfields.Item("WebTypeID") & ""
      vTax = ""
   
      ' only use it when the state matches and price for item is under the amount specified
     'jk  if vDiscountTaxSD.Exists(vWT) and Cart.Info.StateProvince = vThisState Then
	if vDiscountTaxSD.Exists(vWT) Then
         If oProd1.pfields.Item("price") < vDiscountTaxMaxSD(vWT) Then
            vTax = vDiscountTaxSD(vWT) * oProd1.pfields.Item("price")
         end if
      end if
   
      if vDebugx Then
         response.write "vWT=" & vWT & " vTax=" & vTax & " S=" & Cart.Info.StateProvince
         response.write "<br>DT=" & vDiscountTaxSD(vWT)
         response.write "<br>DTM=" & vDiscountTaxMaxSD(vWT)
         response.end
      end if
   
      Cart.AddtoCart vITEMID, 1, vTMP1, oProd1.pfields.Item("price"), 0, vTMP2, 0, 0, vTax, , vURL, vReferer, vParent, vProp, vPropID, vOverWeightFlags, vFreeFreight
    '  Cart.AddtoCart vITEMID, 1, vTMP1, oProd1.pfields.Item("price"), 0, vTMP2, 0, 0, .1, , vURL, vReferer, vParent, vProp, vPropID, vOverWeightFlags, vFreeFreight
   else
      ' if the vZeroTaxItems flag is not set then use this addtocart method and avoid the above tax figuring
      Cart.AddtoCart vITEMID, 1, vTMP1, oProd1.pfields.Item("price"), , vTMP2, , ,,, vURL, vReferer, vParent, vProp, vPropID, vOverWeightFlags, vFreeFreight
   end if

   ' --- add item to cart
   'Cart.AddtoCart vITEMID, 1, vTMP1, oProd1.pfields.Item("price"),, vTMP2,,,,, vURL, vReferer, vParent, vProp, vPropID, vOverWeightFlags, vFreeFreight

   Cart.Calculate

   ' Save cart to session
   Session("Cart") = Cart.SaveCart
   Set Cart = Nothing

   ' Done adding to cart - redirect to show cart contents

   response.redirect "/displaycart/" 


%>