<!-- #INCLUDE FILE="template_cls.asp" -->
<!-- #INCLUDE FILE="common.asp" -->
<!-- #INCLUDE FILE="cartconfig.asp" -->


<html>
<head>
<link rel="stylesheet" type="text/css" href="/index.css" title="index">
<title>BicycleBuys.com Online Bike Shop</title>
<!--#include virtual="/script.inc"-->
</head>

<BODY BGCOLOR="#E5E5F0" TEXT="#000000" LINK="#3770A8" VLINK="#3770A8" ALINK="#FFFFFF" TOPMARGIN="0" MARGINHEIGHT="0" LEFTMARGIN="0" MARGINWIDTH="0">

<TABLE WIDTH=100% HEIGHT=100% BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR><TD WIDTH="214" HEIGHT="45" background="/cartimages/shiptop_bkg.gif">

	<IMG SRC="/cartimages/viewshipcharges_title.gif" WIDTH="214" HEIGHT="45" BORDER=0><BR>

</TD><TD WIDTH="186" HEIGHT="45" ALIGN="right" background="/cartimages/shiptop_bkg.gif">

	<a href="javascript:window.close()"><IMG SRC="/cartimages/closewindow_top.gif" WIDTH="86" HEIGHT="45" BORDER=0></A><BR>

</TD></TR>

<TR><TD COLSPAN="2" ALIGN="center" vALIGN="top" bgcolor="#FFFFFF">
<BR>

	<TABLE WIDTH=90% BORDER=0 CELLPADDING=0 CELLSPACING=0>

<%

Set Cart  = Server.CreateObject("iiscart2000.store")%><%
Cart.LoadCart(Session("Cart"))

vShipDebug = -1
if vShipDebug and Request.ServerVariables("REMOTE_ADDR") = "204.117.211.19" Then vShipDebug = -1 Else vShipDebug = 0
'vShipDebug = -1

' First we need to backup all the shipping data
' we might have collected from the checkout process

dim vBKPShipping, vBKPShippingType, vBKPInfoShipCustom8, vBKPInfoIsStateResident, vBKPInfoShipStateProvince, vBKPInfoShipCountry, vBKPShippingTaxRate
dim vTMP

vBKPShipping = Cart.Shipping
vBKPShippingType = Cart.ShippingType
vBKPInfoShipCustom8 = Cart.Info.ShipCustom8
vBKPInfoIsStateResident = Cart.Info.IsStateResident
vBKPInfoShipStateProvince = Cart.Info.ShipStateProvince
vBKPInfoShipCountry = Cart.Info.ShipCountry

'  Get the proper info about this states shipping zone number
'  for use in cross referencing the shipping methods
'
'  This should really be an application scope lookup table
'
vTMP = Request.QueryString("SHIPSTATEPROVINCE")
if vShipDebug then response.write "State: " & vTMP & "<BR>"

if vTMP <> "" Then
	if vTMP = vThisState then
	   Cart.Info.IsStateResident = -1
	   Cart.ShippingTaxRate = "8.63%"
      	   vBKPShippingTaxRate = "8.63%"
	Else
	   Cart.Info.IsStateResident = 0
	   Cart.ShippingTaxRate = "0%"
           vBKPShippingTaxRate = "0%"
	End If
  	Cart.Info.ShipStateProvince = vTMP
	Cart.CalculateShipping
	Cart.Calculate
   if vShipDebug then Response.write "ShippingType:" & Cart.ShippingType & "<BR>"
   if (Cart.ShippingType = "" or Cart.ShippingType = "NONE") Then
  	   Cart.ShippingType = "NONE"
  	   Cart.Info.ShipCustom8 = 0
   Else
'  	   Cart.ShippingType = "UPS 3 Day Select"
'  	   Cart.Info.ShipCustom8 = 5
  	   Cart.ShippingType = "NONE"
  	   Cart.Info.ShipCustom8 = 0
  	End If
   	Cart.Info.ShipCountry = "US"
  	'Session("Cart") = Cart.SaveCart
End if

vTMP = Request.QueryString("SHIPCOUNTRY")
if vShipDebug then response.write "Country: " & vTMP & "<BR>"
if vTMP <> "" Then
  	Cart.Info.ShipCountry = vTMP
  	If vTMP = "OTHER" then Cart.Info.ShipStateProvince = "OTHER"
	Cart.CalculateShipping
	Cart.Calculate
   if Cart.ShippingType = "" Then
'  	   Cart.ShippingType = "UPS Ground"
'  	   Cart.Info.ShipCustom8 = 11
  	   Cart.ShippingType = "NONE"
  	   Cart.Info.ShipCustom8 = 0
  End If
  'Session("Cart") = Cart.SaveCart
End if

if vShipDebug then Response.write "ShippingType:" & Cart.ShippingType & "<BR>"
%>
      <TR>
            <TD ALIGN="center" COLSPAN="2">
            <FONT id="cartnormal"> 
            <B>Choose shipping location to view the cost of our delivery options</B>            </FONT>          </TD>
      </TR>
      
      <TR><TD COLSPAN="2">&nbsp;</td></tr>
      
      <TR>
         <FORM name="shipnavstate" method="get">
         <TD ALIGN="right"><FONT ID="cartnormal">U.S. State:</FONT></td>
         <TD ALIGN="left"><select name="SHIPSTATEPROVINCE" onChange="javascript:if ((document.forms.shipnavstate.SHIPSTATEPROVINCE.selectedIndex!=0) && (document.forms.shipnavstate.SHIPSTATEPROVINCE[document.forms.shipnavstate.SHIPSTATEPROVINCE.selectedIndex].value!='x')) {window.location=document.forms.shipnavstate.SHIPSTATEPROVINCE[document.forms.shipnavstate.SHIPSTATEPROVINCE.selectedIndex].value}">
           <!--           <SELECT NAME="SHIPSTATEPROVINCE" onChange="load3(this.form,parent.frames)">  -->
           <option value="smallship.asp">Select State</option>
           <% For Each vState in vStateSD.Keys
               vSelected = ""
               If vState = Cart.Info.ShipStateProvince then vSelected = " SELECTED" %>
           <option value="smallship.asp?SHIPSTATEPROVINCE=<%=vState%>"<%=vSelected%>><%=vStateSD.Item(vState)%></option>
           <% Next %>
         </select>
           <BR>         </TD>
         </FORM>
      </TR>

      <TR>
         <FORM name="shipnavcountry" method="get">
         <TD ALIGN="right">
            <FONT ID="cartnormal">Or Country:</FONT>            </td>
            <TD ALIGN="left">
            <SELECT NAME="SHIPCOUNTRY" onChange="javascript:if ((document.forms.shipnavcountry.SHIPCOUNTRY.selectedIndex!=0) && (document.forms.shipnavcountry.SHIPCOUNTRY[document.forms.shipnavcountry.SHIPCOUNTRY.selectedIndex].value!='x')) {window.location=document.forms.shipnavcountry.SHIPCOUNTRY[document.forms.shipnavcountry.SHIPCOUNTRY.selectedIndex].value}">
<!--            <select name="SHIPCOUNTRY" onChange="load4(this.form,parent.frames)">  -->
               <option value="smallship.asp">Select Country</option>
               <option value="smallship.asp?SHIPCOUNTRY=US"<% if Cart.Info.ShipCountry="US" then response.write " SELECTED"%>>US&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
               <option value="smallship.asp?SHIPCOUNTRY=OTHER"<% if Cart.Info.ShipCountry="OTHER" then response.write "SELECTED"%>>Outside the U.S.</option>
            </select>         </TD>
         </FORM>
      </TR>
<%
vOverWeight = Application("OverWeight")
if Cart.Info.ShipStateProvince <> "OTHER" and Cart.Info.ShipStateProvince <> "" Then

   vSCPZ = Application("SCPZ")
   vOverWeight = Application("OverWeight")

   if vShipDebug then response.write "STATE CHOSEN<BR>"
   ' Figure out shipping costs based on shipping database
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
      
      For Each Item in Cart.Items
      
        ' ---- FREE SHIPPING ON NON-OVERWEIGHT ITEMS
        If (Item.Custom7+0) = -1 AND (Item.Custom6+0) < 1 Then
            ' AND THEY DIDN'T SELECT THE FREE SHIPPING METHOD
            If Cart.Info.ShipCustom8 <> vFreeShippingMethodID then
               if vShipDebug Then Response.write Item.ItemID & "-" & "Free, but not UPS Ground<br>"
               vNetIgnoreFreeTotal = vNetIgnoreFreeTotal + (Item.Price * Item.Quantity)
               vNetIgnoreFreeItems = vNetIgnoreFreeItems + Item.Quantity
 
 
               ' ONLY ON THE SMALLSHIP DO WE DO THIS
               ' -- it keeps free freight item qty in order
               ' -- regardless of 'chosen' shipping method
               '--------------------------------------
               vNetFreeFreightItems = vNetFreeFreightItems + Item.Quantity
               vNetIgnoreFreeItems = vNetIgnoreFreeItems + Item.Quantity
               '--------------------------------------
      
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
        ElseIf (Item.Custom6+0) > 0 and (Item.Custom7+0) <> -1 then
           if vShipDebug Then Response.write Item.ItemID & "-" & "Overweight type:" & (Item.Custom6-1) & "<br>"
           vNetOverSizedItems = vNetOverSizedItems + Item.Quantity
           vOverSizedItems(Item.Custom6+0) = vOverSizedItems(Item.Custom6+0) + Item.Quantity

        ' ---- OVERWEIGHT AND FREE FREIGHT
        ElseIf (Item.Custom6+0) > 0 and (Item.Custom7+0) = -1 then
           if vShipDebug Then Response.write Item.ItemID & "-" & "Overweight/Free type:" & (Item.Custom6-1) & "<br>"
           vNetOverSizedFreeItems = vNetOverSizedFreeItems + Item.Quantity
           vOverSizedFreeItems(Item.Custom6+0) = vOverSizedFreeItems(Item.Custom6+0) + Item.Quantity
        
        ' ---- NOT FREE AND NOT OVERWEIGHT
        ElseIf (Item.Custom7+0) <> -1 and (Item.Custom6+0) < 1 then
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

         ' show oversized type (1,2 or 3) and # of oversizeditems
         for x = 0 to ubound(vOverSizedItems)
            if vOverSizedItems(x) Then Response.write "OverSizedItem Shipping type:" & x-1 & " - " & vOverSizedItems(x) & "<BR>"
         next

         ' show oversized type (1,2 or 3) and # of oversizedFREEitems
         for x = 0 to ubound(vOverSizedFreeItems)
            if vOverSizedFreeItems(x) Then Response.write "OverSizedFreeItem Shipping type:" & x-1 & " - " & vOverSizedFreeItems(x) & "<BR>"
         next
      End If


      If Cart.Info.ShipCountry  = "US" Then

         vShipZone = vZonesSD.Item(Cart.Info.ShipStateProvince)
         if vShipDebug then response.write cart.info.shipstateprovince & " - " & vShipZone & "<br>"

      Else
         vShipZone = 4
         'Cart.Shipping = 0
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
                'jk  if vTMP1 <> vFreeShippingMethodID Then
		if vSCPZ(x,3) <> vFreeShippingMethodID Then 
                     if vNetIgnoreFreeTotal >= vSCPZ(x,1) and vNetIgnoreFreeTotal <= vSCPZ(x,2) Then vSMCAF(vSCPZ(x,3)) = vSCPZ(x, 4)
               ' response.write vSCPZ(x, 0) & "---" & vSCPZ(x,1) & "-" & vSCPZ(x,2) & "-" & vSCPZ(x,3) & "-" & vSCPZ(x,4) & " - " &  vSCPZ(x,5) & "<br>"  
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

            if Request.ServerVariables("SCRIPT_NAME") <> "/smallship.asp" Then
               vTMP3 = (Cart.Info.ShipCustom8+0)
               vShippingCost = vSMCAF(vTMP3)
            End If
         End If
      End If
   End If

   ' Build display array for shipping select box
   '   
   ' 0/ShipID, 1/shipname, 2/shipcost
   ' Example:
   '  vShipSelect(1,0) = 5
   '  vShipSelect(1,1) = "UPS 3 Day Select"
   '  vShipSelect(1,2) = 19.95

   Dim vShipSelect(10,2)
   vShippingNames = Application("ShippingNames")
   vShipCount = 0

   ' Only handle U.S. Shipping
   If vShipZone = 1 or vShipZone = 2 or vShipZone = 3 or vShipZone = 5 Then

      ' Normal shipments of non-oversized items
      If vNetOverSizedItems = 0 and vNetOverSizedFreeItems = 0 Then
         For x = 1 to Ubound(vSMCAF)
            if vSMCAF(x) or (x = vFreeShippingMethodID and vNetFreeFreightItems > 0 and vShipZone < 3) then
               vTMP = vShippingNames(x)
               vShipSelect(vShipCount,0) = x
               vShipSelect(vShipCount,1) = vTMP
               vShipSelect(vShipCount,2) = vSMCAF(x)
               if vShipDebug then response.write "SA" & vShipCount & ": " & x & ":" & vTMP & ":" & vSMCAF(x) & "<BR>"
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
               vTMP = vShippingNames(x)
               vShipSelect(vShipCount,0) = x
               vShipSelect(vShipCount,1) = vTMP

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

   If vShipZone = 1 or vShipZone = 2 or vShipZone = 3 or vShipZone = 5 Then %>
      <TR>
         <TD ALIGN="LEFT"><FONT ID="carttitle"><B>SHIPPER</B></FONT></TD>
         <TD ALIGN="RIGHT"><FONT ID="carttitle"><B>COST</B></FONT></TD>
      </TR>
<%
      dim vCurrent, vShipInfo
	  vCurrent = 0
      For vCurrent = 0 to ubound(vShipSelect)
         if Len(vShipSelect(vCurrent,1)) > 0 Then
         
            Select Case vShipSelect(vCurrent,1)
               
               Case "UPS Ground"
                  vShipInfo = " East of the Mississippi is 1-4 days. West of the Mississippi is 4-6 business days. The first ship date is the date after UPS picks-up at our warehouse."
               
               Case "UPS 3 Day Select"
                  vShipInfo = "UPS will deliver in 3 business days. The first shipping day is the day after pickup from our warehouse. Weekends are not considered as part of the 3 day period."
      
               Case "UPS Next Day"
                  vShipInfo = "Next business day. Orders taken on Friday will be delivered on Monday. For Saturday delivery, please call for special rates."
            End Select %>

      <TR BGCOLOR="#BCBCBC"><FORM>
         <TD WIDTH="80%" ALIGN="LEFT"><FONT ID="cartnormal"><B><%=vShipSelect(vCurrent,1)%></B></FONT></TD>
         <TD WIDTH="20%" ALIGN="RIGHT"><FONT ID="cartnormal"><B><%=FormatCurrency(vShipSelect(vCurrent,2),2,0,0)%></B></FONT></TD>
      </TR>
      <TR>
          <TD COLSPAN="2" ALIGN="LEFT"> <FONT ID="cartnormal"><%=vShipInfo%></FONT></TD>
      </TR>
<%
         End if
      Next

     ElseIf vShipZone = 4 Then   %>
      <TR><TD COLSPAN="2">
         <TABLE BORDER="0">
         <TR COLSPAN="2"><TD>&nbsp;</TD></TR>
         <TR>
            <TD VALIGN="TOP"><FONT ID="carttitle"><B>International <%=vShipZone%></B></FONT></TD>
            <TD><FONT ID="cartnormal">
            <I>A member of the BicycleBuys.com staff will contact you after your order is submitted.</I></FONT></TD>
         </TR>
         </TABLE>
      </TD></TR>

<%   End If  %>

<% End If  %>
<% If Cart.Info.ShipCountry = "OTHER" Then %>
<TR><TD COLSPAN="2">
<table border="0">
<TR COLSPAN="2"><TD>&nbsp;</TD></TR>
<TR>
   <TD vALIGN="top"><FONT ID="carttitle"><B>International Shipping</B></FONT></td>
   <TD><FONT ID="cartnormal">
   <i>A member of the BicycleBuys.com staff will contact you after your order is submitted.</i></FONT></td>
</tr>
</table>

</TD></TR>
<% End If %>
<TR><TD COLSPAN="2">
<table border="0">
<TR COLSPAN="2"><TD>&nbsp;</TD></TR>
<TR>
                <TD vALIGN="top">&nbsp;</td>
				    <TD>
				      <FONT ID="cartnormal">
				      <B>At checkout the shipping method can be decided. We try our best 
                  to ship all orders within 48 hours of receiving. In the event 
                  that an item is not shippable within 4 days of date of order, 
                  you will be contacted.</B>                  </FONT>                </TD>
</tr>
</table>
 </TABLE>

</TD></TR>

<TR><TD width="214" HEIGHT="29" background="/cartimages/shipbottom_bkg.gif">

	<IMG SRC="/cartimages/bb_mini.gif" WIDTH="214" HEIGHT="29" BORDER=0><BR>

</TD><TD width="186" HEIGHT="29" ALIGN="right" background="/cartimages/shipbottom_bkg.gif">

	<a href="javascript:window.close()"><IMG SRC="/cartimages/closewindow_bottom.gif" WIDTH="86" HEIGHT="29" BORDER=0></A><BR>

</TD></TR>
</TABLE>

</body>
</html>
<%

' Now we put everything back to the way it was

Cart.Shipping = vBKPShipping
Cart.ShippingType = vBKPShippingType
Cart.Info.ShipCustom8 = vBKPInfoShipCustom8
Cart.Info.IsStateResident = vBKPInfoIsStateResident
Cart.ShippingTaxRate = vBKPShippingTaxRate
Cart.Info.ShipStateProvince = vBKPInfoShipStateProvince
Cart.Info.ShipCountry = vBKPInfoShipCountry
Cart.Calculate
'Session("Cart") = Cart.SaveCart
%>
