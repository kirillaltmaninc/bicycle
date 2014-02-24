<%
   '''''''''
   '  puts the cart in vOUT2
   '  this version ALLOWS FOR changing per item values (qty, itemoptions
   '''''''''
	Dim last_added_id

    vAdjustTotal = 0
	last_added_id = ""
   if cart.gridtotalquantity > 0 then
      For each item in Cart.Items
         IF 1=0 then
            response.write "<hr>"
         	for x = 1 to Item.PropertyCount
         		response.write "Name: " & Item.PropertyName(x) & " <br>"
         		response.write "Value: " & Item.PropertyValue(x) & " <br>"
         		response.write "Custom4: " & Item.Custom4  & " <br>"
            Next
         End If

         'track number of items
      	vCount = vCount + 1

      	' alternate row colors
      	vBGColor = "#DEF4F7"
      	if vCount/2 = int(vCount/2) then
      	     vBGColor = "#FFFFFF"
      	end if

         ' define the item option output
         if Item.Custom4 <> "" then
         	' Get properties from cart
         	vProp = Item.Custom4
         	vPropID = Item.Custom5

         	' Get array's from properties
         	' Sample: Prop=17 - Blue&COMBO;
         	x = instr(vProp, ";")
         	'response.write vProp & " " & x & ":" & right(vProp, Len(vProp)-x) & "<br>"
         	vPropA = Split(right(vProp, Len(vProp)-x), ",")
         	x = instr(vPropID, ";")
         	vPropIDA = Split(right(vPropID, Len(vPropID)-x), ",")
         	' response.write vPropID & " " & x & "<br>"

            vItemOptions = "<div align=""product_options""><select name=""Prop"">" & vbCRLF
            for y = 0 to ubound(vPropA)
      	      vItemOptions = vItemOptions & "<option value='" & vPropIDA(y) &"'" & vbCRLF
      	      if vPropIDA(y) = Item.ItemID then vItemOptions = vItemOptions & " SELECTED"
      	      vItemOptions = vItemOptions & ">" & vPropA(y) & "</option>" & VBCrLf
      	   Next
            vItemOptions = "<BR>" & vItemOptions & "</select></div>" & vbCRLF
         Else
            vItemOptions = "<input TYPE=""HIDDEN"" NAME=""Prop"" VALUE=""" & Item.ItemID & """>" & vbCRLF
         End If

         ' deal with null values in the overweight IC6 field
         if isnumeric(Item.Custom6) Then
            vTMP1 = (Item.Custom6 + 0)
         else
            vTMP1 = 0
         end if
         ' deal with null values in the freefreight IC7 field
         if isnumeric(Item.Custom7) Then
            vTMP2 = (Item.Custom7 + 0)
         else
            vTMP2 = 0
         end if

         ' clear freightmsg for each item
         vFreightMsg = ""

         ' if overweight and not freefreight
         if (vTMP1) = 0 then
            if (vTMP2) = -1 then
               vFreightMsg = "<hr width=""50%"" size=""1""><div class=""product_freefreight"">(Free Shipping!)</div>"
            end if
         elseif (vTMP1) > 0 then
            ' if free shipping for only a specific type of shipping
            if (vTMP2) = -1 then
               vFreightMsg = "<hr width=""50%"" size=""1""><div class=""product_freefreight"">(Free " & vFreeShippingMethod & " shipping; Freight costs will apply to other shipping methods.)</div>"
            else
               ' if item is overweight
               vFreightMsg = "<hr width=""50%"" size=""1""><div class=""product_freefreight""></div>"
            end if
         end if

         ' response.write "<hr>IC6: " & (vTMP1 > 0) & "<br>"

         vOUT2 = vOUT2 & getitemrow(vCount, vBGColor, _
                        Item.ItemID, Item.Name, Item.Weight, Item.Quantity, _
                        Item.Custom1, Item.Custom2, Item.Custom3, Item.Custom4, _
                        Item.Custom5, Item.Custom6, Item.Custom7, Item.Custom8, _
                        vItemOptions,vFreightMsg,Item.Price, Item.Adjust)
        vAdjustTotal = vAdjustTotal + Item.Adjust
		last_added_id = Item.ItemID		
      Next
 end if
 
 
%>