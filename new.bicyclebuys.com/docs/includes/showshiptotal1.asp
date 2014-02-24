<!--#INCLUDE VIRTUAL="/includes/template_cls.asp" -->
<!--#INCLUDE VIRTUAL="/includes/common1.asp" -->
<!--#INCLUDE VIRTUAL="/includes/cartconfig1.asp" --><%
 dim adjRebate, adjRebateTotal, adjRebateTotalAll, adjRebate1, adjRebateTotal1, adjRebateTotalAll1
   call zeroRebateArray()

   ' get the template engine ready
   set objTemplate = new template_cls

   if vDebugx then vShipDebug = -1

dim vPromoFreeShipping
vPromoFreeShipping = 0 
if   1=1 or Request.ServerVariables("REMOTE_ADDR")  = "69.127.248.96" or Request.ServerVariables("REMOTE_ADDR")  = "10.0.0.78" or   "12/31/2009"  >left( (now()),10)  then
    if Cart.Info.ShipCountry ="US" and Cart.GridTotal>99 then
        vPromoFreeShipping = -1  
    end if
   ' vshipdebug = -1
       'response.write vsection & "<br>" & vsql1 & "<br>" & vsql2 & "<hr>"
end if
   vOUT1 = ""
   vOUT2 = ""
'   For each item in Cart.Items
'      response.write "<hr>IC4: " & Item.Custom4 & "<br>"
'   next

%><!--#INCLUDE VIRTUAL="/includes/cartdisplay_checkout.asp" --><%

      ' Put checkout and empty buttons on the display if there is an item in the cart
      ' (disabled with < -1)
      Dim vCheckEmpty
      vCheckEmpty = ""
      if Cart.GridTotalQuantity < -1 then
         vCheckEmpty = "<a href=""" &  vThisProto & vThisServer & "/checkout/""><img src=""/cartimages/checkout.gif"" alt=""Check out"" border=""0"" WIDTH=""100"" HEIGHT=""20""></a>" _
                       & "<a href=""" &  vThisProto & vThisServer & "/emptycart/""><img src=""/cartimages/emptycart.gif"" alt=""Remove ALL items from the cart"" border=""0"" WIDTH=""100"" HEIGHT=""20""></a>"
      'elseif  Cart.Items.Count>0 and Request.QueryString("SHIPPINGTYPE") = ""  and Cart.ShippingType = "" then
	'call Cart.SaveCartDB
	'response.write "SAVED " & Cart.ShippingType 
      end if


      ' begin handling of form chooser
      ' shippingtype will equal something if the browser submits
      if Request.QueryString("SHIPPINGTYPE") <> "" Then
         Dim vRST
         vRST =  replace(Request.QueryString("SHIPPINGTYPE"), "+", " ")
         if vShipDebug then response.write "<hr>1: Enter shippingtype: " & vRST & "<BR>"
         if vShipDebug then response.write "<hr>2: " & left(vRST, 6) & "<br>" & mid(vRST, 7) & "<br>"
         if left(vRST, 6) = "/ship/" then vRST = mid(vRST, 7)
         if left(vRST, 20) = "?SHIPPINGTYPE=/ship/" then vRST = mid(vRST, 21)
         if vShipDebug then response.write "<hr>3: " &vRST & "<br>"
      	Cart.CalculateShipping
      	Cart.Calculate
      	vShipTypeA = Split(vRST, ";")
         if vShipTypeA(0) <> "DONE" and vShipTypeA(0)<>"NONE" Then
         	Cart.ShippingType = vShipTypeA(0)
         	Cart.Info.ShipCustom8 = vShipTypeA(1)
         Else
            Cart.ShippingType = ""
         	Cart.Info.ShipCustom8 = 0
         End If
       	Session("Cart") = Cart.SaveCart
      End If

      ' we're done now when this PAY.x has a value
      '(usually the coords of the "done" picking shipping image button
      if Request("PAY.x") > 0 Then

         ' must pick a shipping type if going to the US
         if Cart.Info.ShipCountry = "US" and (Cart.ShippingType = "" or Cart.ShippingType = "NONE") Then vErrString = "Please select a shipping method.<BR>If experiencing difficulties select method then click update."

         if vShipDebug Then Response.write "CS:" & Cart.ShippingType

         ' if no errors on form then redirect to the billing page
       	if vErrString = "" Then
      		if not vShipDebug then Response.Redirect "/billing/"
      	End If
      End If
   end if

   ' we got these application vars from global.asa
   ' SCPZ = Shipping Cost Per Zone table
   ' OverWeight = Overweight cost per zone
   vSCPZ = Application("SCPZ")
response.write("ggg"&vSCPZ(x,1)&"jjj")
response.write("ggg"&vSCPZ(x,2)&"jjj")
response.write("ggg"&vSCPZ(x,3)&"jjj")
response.write("ggg"&vSCPZ(x,4)&"jjj")
response.write("ggg"&vSCPZ(x,5)&"jjj")
   vOverWeight = Application("OverWeight")

   if cart.gridtotalquantity > 0 then

      ' Figure out shipping costs based on shipping database
      Dim vOverSizedItems(5)
      Dim vOverSizedFreeItems(5)
      vNetShippingTotal = 0
      vNetShippingItems = 0
      vNetOverSizedItems = 0
      vNetOverSizedFreeItems = 0
      vNetFreeFreightItems = 0
      vNetIgnoreFreeItems = 0
      vNetIgnoreFreeTotal = 0

      ' iterate through each item to figure out shipping for each shipping type
      ' and free/over scenarios
      'vOverWeightFlags custom6, vFreeFreight=custom7,
     
      adjRebateTotal1=0
      For Each Item in Cart.Items

        adjRebateTotalAll1 = getitemrow(vCount, vBGColor, _
                        Item.ItemID, Item.Name, Item.Weight, Item.Quantity, _
                        Item.Custom1, Item.Custom2, Item.Custom3, Item.Custom4, _
                        Item.Custom5, Item.Custom6, Item.Custom7, Item.Custom8, _
                        vItemOptions,vFreightMsg,Item.Price, Item.Adjust)
               adjRebate1=Item.Adjust*Item.Quantity


 ' overweight
        'If IsNumeric(vIC6) Then vIC6 = Item.Custom6
		'If Not IsNumeric(vIC6) Then vIC6 = 0
        'If Not IsNumeric(vIC7) Then vIC7 = 0
		response.Write("adjRebate  - " & adjRebate1)
'response.end
		
		if (adjRebate1=0) then
			vIC6 = (Item.Custom6 + 0)
			If Not IsNumeric(vIC6) Then vIC6 = 0
			vIC7 = (Item.Custom7 + 0)
			If Not IsNumeric(vIC7) Then vIC7 = 0
			
		End If



        if vDebugx then RESPONSE.WRITE "IC6|" & vIC6 & "| / IC7|" & vIC7 & "|<BR>"
       
        ' ---- FREE SHIPPING ON NON-OVERWEIGHT ITEMS
        If (vIC7 = -1 or (vPromoFreeShipping ) ) AND vIC6 < 1 or (vPromoFreeShipping and cint(vIC6) < 1) Then
            ' AND THEY DIDN'T SELECT THE FREE SHIPPING METHOD
            If (Cart.Info.ShipCustom8+0) <> vFreeShippingMethodID then
               if vShipDebug Then Response.write Item.ItemID & "-" & "Free, but not FEDEX Ground - " & Cart.Info.ShipCustom8 & "/" &  vFreeShippingMethodID & "<br>"
               vNetFreeFreightItems = vNetFreeFreightItems + Item.Quantity
               vNetIgnoreFreeTotal = vNetIgnoreFreeTotal + (Item.Price * Item.Quantity)
               vNetIgnoreFreeItems = vNetIgnoreFreeItems + Item.Quantity

            ' AND THEY'VE SELECTED THE FREE SHIPPING METHOD
            Else
              if vShipDebug Then Response.write Item.ItemID & "-" & "Free freight<br>"
              vNetFreeFreightItems = vNetFreeFreightItems + Item.Quantity
              vNetIgnoreFreeTotal = vNetIgnoreFreeTotal + (Item.Price * Item.Quantity)
              vNetIgnoreFreeItems = vNetIgnoreFreeItems + Item.Quantity
            End If

            ' EITHER WAY KEEP TOTAL FOR WHEN THEY DISPLAYING THE NON-FREE METHODS
'            vNetShippingTotal = vNetShippingTotal + (Item.Price * Item.Quantity)
'            vNetShippingItems = vNetShippingItems + 1

        ' ---- OVERWEIGHT AND NOT FREE FREIGHT
        ElseIf vIC6 > 0 and vIC7 <> -1 then
           if vShipDebug Then response.write "<br>" & vic7 & " - " & (vIC7 = "-1") & " " & vartype(vic7) & "<br>"
           if vShipDebug Then Response.write Item.ItemID & "-" & "Overweight type:" & vic7 & "|" & (vIC6-1) & "<br>"
           vNetOverSizedItems = vNetOverSizedItems + Item.Quantity
           vOverSizedItems(vIC6) = vOverSizedItems(vIC6) + Item.Quantity

        ' ---- OVERWEIGHT AND FREE FREIGHT
        ElseIf vIC6 > 0 and vIC7 = -1 then
           if vShipDebug Then Response.write Item.ItemID & "-" & "Overweight/Free type:" & (vIC6-1) & "<br>"
           vNetOverSizedFreeItems = vNetOverSizedFreeItems + Item.Quantity
           vOverSizedFreeItems(vIC6) = vOverSizedFreeItems(vIC6) + Item.Quantity

        ' ---- NOT FREE AND NOT OVERWEIGHT
        ElseIf vIC7 <> -1 and vIC6 < 1 then
           if vShipDebug Then Response.write Item.ItemID & "-" & "Not free<br>"
           vNetShippingTotal = vNetShippingTotal + (Item.Price * Item.Quantity)
           vNetShippingItems = vNetShippingItems + Item.Quantity

           vNetIgnoreFreeTotal = vNetIgnoreFreeTotal + (Item.Price * Item.Quantity)
           vNetIgnoreFreeItems = vNetIgnoreFreeItems + Item.Quantity
        Else
           response.write "Shipping error: " & Item.Custom6 & "<br>"
        End If



      Next

      if vShipDebug Then
         response.write "vNetShippingTotal: " & vNetShippingTotal & "<br>"
         response.write "vNetShippingItems: " & vNetShippingItems & "<br>"
         response.write "vNetOverSizedItems: " & vNetOverSizedItems & "<br>"
         response.write "vNetOverSizedFreeItems: " & vNetOverSizedFreeItems & "<br>"
         response.write "vNetFreeFreightItems: " & vNetFreeFreightItems & "<br>"
         response.write "vNetIgnoreFreeItems: " & vNetIgnoreFreeItems & "<br>"
         response.write "vNetIgnoreFreeTotal: " & vNetIgnoreFreeTotal & "<br>"
         for x = 0 to ubound(vOverSizedItems)
            if vOverSizedItems(x) Then Response.write "OSI" & x-1 & ": " & vOverSizedItems(x) & "<BR>"
         next
      End If

      If Cart.Info.ShipCountry  = "US" Then

'            sql = "SELECT Zone FROM ShippingStateZones WHERE State='" & Cart.Info.ShipStateProvince & " 
'            rs.open sql & " For Browse",Conn,3
'            vShipZone = rs("Zone")
'            rs.close

            vShipZone = vZonesSD.Item(Cart.Info.ShipStateProvince)
            if vShipDebug then response.write cart.info.shipstateprovince & " - " & vShipZone & "<br>"
      Else
            vShipZone = 4
            Cart.Shipping = 0
            Cart.ShippingTax = 0
            Cart.Calculate
            Session("Cart") = Cart.SaveCart
      End If

      if vShipDebug then Response.write "vShipZone = " & vShipZone & "<BR>"

      ' Handle all US shipments (zones 1-3, 5)
      If vShipZone = 1 or vShipZone = 2 or vShipZone = 3 or vShipZone = 5 Then
         ' to get our shipping costs we need many different totals;
         '   Total minus freefreight item prices and overweight items to calculate the shipping cost for the freefreight shipping method.
         '   Total with freefrieght item prices, no overweight items, to calculate the shipping cost for all other shipping methods.
         '   Individual shipping costs per overweight item

         Dim vSMCAF(50)
         ' Here we handle all normal and freefreight shipping totals
         If vNetIgnoreFreeItems > 0 Then
            // find the allowable shipping methods/costs
            for x = 0 to UBound(vSCPZ)

               // gotta make sure we're using integers
               vTMP1 = vSCPZ(x,0) + 0
               vTMP2 = vShipZone + 0

               if vTMP1 = vTMP2 Then
                  ' If this record isn't the free shipping method then we'll need to
                  ' calculate a different cost based on the ignore-freefreight total
               '   if vTMP1 <> vFreeShippingMethodID Then
		if vSCPZ(x,3) <> vFreeShippingMethodID Then
                     if vNetIgnoreFreeTotal >= vSCPZ(x,1) and vNetIgnoreFreeTotal <= vSCPZ(x,2) Then vSMCAF(vSCPZ(x,3)) = vSCPZ(x, 4)
                  else
                     if vNetShippingTotal >= vSCPZ(x,1) and vNetShippingTotal <= vSCPZ(x,2) Then vSMCAF(vSCPZ(x,3)) = vSCPZ(x, 4)
                  end if
                  if vShipDebug then response.write vSCPZ(x, 0) & "---" & vSCPZ(x,4) & " - " &  vSCPZ(x,5) & "<br>"
               End if
            next


            ' If there are no non-freefreight items...
            if vNetShippingTotal = 0 Then
               vTMP1 = vFreeShippingMethodID
               vTMP2 = 0
               vSMCAF(vTMP1) = vTMP2
            End If

            vTMP3 = (Cart.Info.ShipCustom8+0)
            vShippingCost = vSMCAF(vTMP3)

         End If

         Cart.Shipping = vShippingCost
         Cart.Calculate
         Session("Cart") = Cart.SaveCart
      End If

   ' Build display array for shipping select box

   ' 0/ShipID, 1/shipname, 2/shipcost
   ' vShipSelect(1,0) = 1
   ' vShipSelect(1,1) = "FEDEX Ground"
   ' vShipSelect(1,2) = 9.95

   Dim vShipSelect(10,2)
   vShippingNames = Application("ShippingNames")
   vShipCount = 0

   ' Only handle U.S. Shipping
   If vShipZone = 1 or vShipZone = 2 or vShipZone = 3 or vShipZone = 5 Then

      ' Normal shipments of non-oversized items
      If vNetOverSizedItems = 0 and (vNetOverSizedFreeItems = 0) Then
         For x = 1 to Ubound(vSMCAF)
            if vSMCAF(x) or (x = vFreeShippingMethodID and vNetFreeFreightItems > 0 and vShipZone < 3) then
               vTMP1 = vShippingNames(x)
               vShipSelect(vShipCount,0) = x
               vShipSelect(vShipCount,1) = vTMP1
               vShipSelect(vShipCount,2) = vSMCAF(x)
               if vShipDebug then response.write "SA" & vShipCount & ": " & x & ":" & vTMP1 & ":" & vSMCAF(x) & "<BR>"
               vShipCount = vShipCount + 1
            End If
         Next

      ' Orders with overweight items are handled differently
      '    If there are any number of overweight non-freefreight items then
      '    we charge the overweight freight price for each overweight item
      '    and we don't charge freight on the non-overweight items.
      '
      '    If there is an overweight item with free freight then we charge
      '    normal freight on the rest if the items
      '
      Else
         vShipCount = 0
         ' iterate through the overweight array
         for x = 0 to ubound(vOverWeight)
            ' if this array item has a name then add it to the select display
            if vOverWeight(x,vShipZone,2) <> "" Then

               ' These are the 'fall-through' values
               vTMP1 = vShippingNames(x)
               vShipSelect(vShipCount,0) = x
               vShipSelect(vShipCount,1) = vTMP1

               ' Set the freight pricing for each overweight item
               '
               ' y = oversized type (1,2 or 3)
               for y = 0 to ubound(vOverSizedItems)
                  If vOverSizedItems(y) Then
                     vShipSelect(vShipCount,2) = vShipSelect(vShipCount,2) + (vOverWeight(x, vShipZone, y-1) * vOverSizedItems(y))
                     if vShipDebug Then response.write vShipCount & "Num:" & vOverSizedItems(y) & " Type: " & x & "  OSType:" & y & " Cost:" & vOverWeight(x, vShipZone, y-1) & "<BR>"
                  End If
               next

               ' Add the non-free freight pricing for all the free freight items
               if x <> vFreeShippingMethodID Then
                  for y = 0 to ubound(vOverSizedFreeItems)
                     If vOverSizedFreeItems(y) Then
                        vShipSelect(vShipCount,2) = vShipSelect(vShipCount,2) + (vOverWeight(x,vShipZone,y-1) * vOverSizedFreeItems(y))
                        if vShipDebug Then response.write vOverSizedFreeItems(y) & "B" & x & " - " & y & "<BR>"
                     End If
                  next
               ElseIf vNetOverSizedItems = 0 Then
                  vShipSelect(vShipCount,2) = vShipSelect(vShipCount,2) + vSMCAF(x)
               End If

               ' Increase the array index
               vShipCount = vShipCount + 1
            End If
         Next
      End If
   End If

   ' if NOT international then show a shipping selection
   If vShipZone = 1 or vShipZone = 2 or vShipZone = 3  or vShipZone = 5 Then

      ' get the dropdown list of all available shipping methods
'HACK TO FIX SHIPPING TAX..... AUTOMATICALLY SET SHIPPING TAX TO ZERO AFTER CALCULATE...
      For x = 0 to vShipCount - 1
         vSelected = ""
         if Cart.ShippingType=vShipSelect(x,1) then
            vSelected = " SELECTED"
            Cart.Shipping = vShipSelect(x,2)
	    Cart.ShippingTax =0
            Cart.Calculate
	    Cart.ShippingTax =0
            Session("Cart") = Cart.SaveCart
         End If
         vOUT3 = vOUT3 & "<option VALUE=""/ship/" & Server.URLEncode(vShipSelect(x,1)) & ";" & vShipSelect(x,0) & """ " & vSelected & ">&nbsp;" & vShipSelect(x,1) & "&nbsp;"
         if vShipSelect(x,2) > 0 then
            vOUT3 = vOUT3 & FormatCurrency(vShipSelect(x,2),2,0,0)
         else
            vOUT3 = replace(vOUT3, "Ground-US Mail", "") & "Free Shipping" & "&nbsp;"
            ' & vShipSelect(x,1)
         end if

         vOUT3 = vOUT3 & "</option>" & vbcrlf
      Next

      ' get shipping tax (disabled)
      if Cart.ShippingTax = -999999 Then
         vOUT4 = "<TR><TD ALIGN=""RIGHT"">Shipping Tax: </TD><TD COLSPAN=""2"">" & formatcurrency(Cart.ShippingTax) & "</TD></TR>" & vbcrlf
      end if
adjRebateTotal=0
         For each item in Cart.Items
        adjRebateTotalAll = getitemrow(vCount, vBGColor, _
                        Item.ItemID, Item.Name, Item.Weight, Item.Quantity, _
                        Item.Custom1, Item.Custom2, Item.Custom3, Item.Custom4, _
                        Item.Custom5, Item.Custom6, Item.Custom7, Item.Custom8, _
                        vItemOptions,vFreightMsg,Item.Price, Item.Adjust)
               adjRebate=Item.Adjust*Item.Quantity
              
               if (adjRebate) then
                  adjRebateTotal=adjRebateTotal+adjRebate
                end if 
           next
           
           if (TotalDiscount15 > 0) then
                vOUT5 = formatcurrency((Cart.GridTotal + Cart.Shipping)+ TotalDiscount15,2,0,0)           
            elseif (adjRebateTotal < 0) then
                vOUT5 = formatcurrency((Cart.GridTotal-vAdjustTotal + Cart.Shipping)+ adjRebateTotal,2,0,0)
            else
                vOUT5 = formatcurrency(Cart.GridTotal + Cart.Shipping)
            end if
      ' order total with shipping
      'vOUT5 = formatcurrency(Cart.GridTotal + Cart.Shipping)

      ' use the US shipping template
      with objTemplate
      	.TemplateFile = TMPLDIR & "shipping-us.html"
         .AddToken "shippingandhandling", 1, formatcurrency(Cart.Shipping)
         .AddToken "chooseship", 1, vOUT3
         .AddToken "shippingtax", 1, vOUT4
         .AddToken "totalwithshipping", 1, vOUT5
         .AddToken "errstring", 1, vErrString
      	vOUT10 = .getParsedTemplateFile
      end with
   else
      ' use the INTERNATIONAL shipping template
      with objTemplate
      	.TemplateFile = TMPLDIR & "shipping-int.html"
         .AddToken "shipping", 1, formatcurrency(Cart.Shipping)
         .AddToken "chooseship", 1, vOUT3
         .AddToken "shippingtax", 1, vOUT4
         .AddToken "gridtotal", 1, FormatCurrency(Cart.GridTotal, 2, 0, 0)
      	vOUT10 = .getParsedTemplateFile
      end with
   end if

      ' shipping as well as cart display built, now show them
      with objTemplate
      	.TemplateFile = TMPLDIR & "showshiptotal.html"
         .AddToken "breadcrumb", 1, vOUT1

			if (TotalDiscount15 > 0) then
				vOUT2 = vOUT2 & "<TR ><TD colspan=3 align=right class=cart style=""text-align: center; border-top: 1px solid #C4EAE6;"">&nbsp;</TD><TD align=right class=cart style=""text-align: center; border-top: 1px solid #C4EAE6;"">Discount</TD><TD colspan=5 align=right style=""text-align: center; border-top: 1px solid #C4EAE6;"" class=cart>-" & FormatCurrency(TotalDiscount15, 2, 0, 0) & "</TD></TR>"
			end if
      
       adjRebateTotal=0
         For each item in Cart.Items
        adjRebateTotalAll = getitemrow(vCount, vBGColor, _
                        Item.ItemID, Item.Name, Item.Weight, Item.Quantity, _
                        Item.Custom1, Item.Custom2, Item.Custom3, Item.Custom4, _
                        Item.Custom5, Item.Custom6, Item.Custom7, Item.Custom8, _
                        vItemOptions,vFreightMsg,Item.Price, Item.Adjust)
               adjRebate=Item.Adjust*Item.Quantity
              
               if (adjRebate) then
                  adjRebateTotal=adjRebateTotal+adjRebate
                end if 
           next
          .AddToken "adjusttotal", 1, rebateHtml(adjRebateTotal) 
          


        if (TotalDiscount15 > 0) then
            .AddToken "cgridtotal", 1, FormatCurrency(Cart.GridTotal - TotalDiscount15,2,0,0)
            'response.Write("1 adjRebateTotal - " & adjRebateTotal)
        elseif (adjRebateTotal < 0) then
            .AddToken "cgridtotal", 1, FormatCurrency((Cart.GridTotal-vAdjustTotal) + adjRebateTotal,2,0,0)
            'response.Write("2 adjRebateTotal - " & adjRebateTotal)
        else
            .AddToken "cgridtotal", 1, FormatCurrency(Cart.GridTotal,2,0,0)
            'response.Write("3 adjRebateTotal - " & adjRebateTotal)
        end if

         .AddToken "displaycart", 1, vOUT2

         .AddToken "thisserver", 1, vThisServer

         .AddToken "shippingpart", 1, vOUT10

         .AddToken "sessionreferer", 1, Session("Referer")
         '.AddToken "cgridtotal", 1, FormatCurrency(Cart.GridTotal, 2, 0, 0)
         '.AddToken "adjusttotal",1, rebateHtml(vAdjustTotal) 
         .AddToken "checkempty", 1, vCheckEmpty

      	.AddToken "header", 3, vCartHeaderSummaryShipping      	
      	.AddToken "footer", 3, TMPLDIR & "cart_footer_ship.html"
        .AddToken "PromoCode", 1, getRebates() 
      	vOUT11 = .getParsedTemplateFile
      	'.parseTemplateFile
      end with
      response.write  vOUT11

   Else
      ' empty cart! now show that
      with objTemplate
      	.TemplateFile = TMPLDIR & "displayemptycart.html"
      	.AddToken "header", 3, vCartHeaderNoSummary
      	.AddToken "footer", 3, vCartFooterNoHelp
      	.parseTemplateFile
      end with

   End If

   ' Page is done. save and cleanup
   Session("Cart") = Cart.SaveCart
   Set Cart = Nothing

   ' continue shopping referer should not be touched
   '   Session("Referer") = Trim(Request.ServerVariables("HTTP_REFERER"))

   'from old site. not sure yet what it's for
   Session("NavID") = ""

%>
