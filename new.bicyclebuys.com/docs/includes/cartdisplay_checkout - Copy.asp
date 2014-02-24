<%

   '''''''''
   '  puts the cart in vOUT2
   '  this version does not allow for changing per item values (qty, itemoptions
   '''''''''
   vAdjustTotal = 0
   if cart.gridtotalquantity > 0 then
      For each item in Cart.Items
         'track number of items
      	vCount = vCount + 1

      	' alternate row colors
      	vBGColor = "#DEF4F7"
      	if vCount/2 = int(vCount/2) then vBGColor = "#FFFFFF"

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

            vItemOptions = "<div class=""product_options"">" & vbCRLF
            for y = 0 to ubound(vPropA)
      	      if vPropIDA(y) = Item.ItemID then vItemOptions = vItemOptions & "[You chose: " & vPropA(y) & "]" & VBCrLf
      	   Next
            vItemOptions = "<BR>" & vItemOptions & "</div>" & vbCRLF
         Else
            vItemOptions = "<input TYPE=""HIDDEN"" NAME=""Prop"" VALUE=""" & Item.ItemID & """>" & vbCRLF
         End If

         ' clear freightmsg for each item
         vFreightMsg = ""

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

         ' response.write "<hr>IC6: " & Item.Custom6 & "<br>"

         ' add each item to the display
         ' getitemrowco will give us a bare html rep of an item line
         vOUT2 = vOUT2 & getitemrowco(vCount, vBGColor, _
                        Item.ItemID, Item.Name, Item.Weight, Item.Quantity, _
                        Item.Custom1, Item.Custom2, Item.Custom3, Item.Custom4, _
                        Item.Custom5, Item.Custom6, Item.Custom7, Item.Custom8, _
                        vItemOptions, vFreightMsg, Item.Price, Item.Adjust)
            vAdjustTotal = vAdjustTotal + Item.Adjust
      Next ' done with items loop
%>