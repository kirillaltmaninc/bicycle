<!-- #INCLUDE FILE="template_cls.asp" -->
<!-- #INCLUDE FILE="common.asp" -->
<!-- #INCLUDE FILE="cartconfig.asp" -->
<html>
<head>
    <title>BicycleBuys.com Online Bike Shop</title>
    <link rel="stylesheet" type="text/css" href="/index.css" title="index">
    <!--#include virtual="/script.inc"-->
</head>
  <%

function getShipCountries
    dim shC, sel
    Set shC = Server.CreateObject("ADODB.Recordset")
    shC.open "exec bbc_prod.dbo.getShipCountries", conn, 3
    response.Write("<option value=""smallship.asp?ShipCountry=" & """"  & sel & "></option>")
    while not shC.eof
        
        if vTMP =  shC.fields("ShipCountry") then 
            sel= " SELECTED"
        else
            sel = ""
        end if
       response.Write("<option value=""smallship.asp?ShipCountry=" & shC.fields("ShipCountry") & """"  & sel & ">" & shC.fields("ShipCountry") & "</option>")
        shC.movenext
    wend
    shC.close
    set shC = nothing
end function

function getEstimateShipping(country)
    dim shC, sel
    Set shC = Server.CreateObject("ADODB.Recordset")
    shC.open "exec bbc_prod.dbo.getShipCountryHistory '" & country & "'", conn, 3    
     response.Write("<table class=""shipping""><tr><td colspan=""5"">The table below contains approximate shipping costs, in US dollars,based on previous shipments to your country.</td></tr>")
      '<td>Ship Country</td>
     response.Write("</table><table class=""shipping""><tr><td>Weight<BR>(lbs)</td><td>Shipping<BR>Method</td><td>Min<BR>Freight</td><td>Average<BR>Freight</td><td>Max<BR>Freight</td></tr><tr><td><br></td></tr>")    
  
    while not shC.eof         
       'response.Write("<tr><td>" & shC.fields("ShipCountry") & "</td>")
       response.Write("<tr><td>" & shC.fields("Weight") & "</td>") 
       response.Write("<td>" & shC.fields("ShippingMethod") & "</td>") 
       response.Write("<td>" & FormatCurrency(shC.fields("minFreight"),2,0,0)  & "</td>") 
       response.Write("<td>" & FormatCurrency(shC.fields("AverageFreight"),2,0,0)  & "</td>") 
       response.Write("<td>" & FormatCurrency(shC.fields("maxFreight"),2,0,0)  & "</td></tr>") 
        shC.movenext
    wend
    shC.close
    set shC = nothing
    response.Write("</table>")
end function

Set Cart  = Server.CreateObject("iiscart2000.store")
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
'  	   Cart.ShippingType = "FEDEX 3 Day Select"
'  	   Cart.Info.ShipCustom8 = 5
  	   Cart.ShippingType = "NONE"
  	   Cart.Info.ShipCustom8 = 0
  	End If
   	Cart.Info.ShipCountry = "US"
  	'Session("Cart") = Cart.SaveCart
End if

vTMP = Request.QueryString("ShipCountry")
 
if vShipDebug then response.write "Country: " & vTMP & "<BR>"
if vTMP <> "" Then
  	Cart.Info.ShipCountry = vTMP
  	If vTMP = "OTHER" then Cart.Info.ShipStateProvince = "OTHER"
	Cart.CalculateShipping
	Cart.Calculate
   if Cart.ShippingType = "" Then
'  	   Cart.ShippingType = "FEDEX Ground"
'  	   Cart.Info.ShipCustom8 = 11
  	   Cart.ShippingType = "NONE"
  	   Cart.Info.ShipCustom8 = 0
  End If
  'Session("Cart") = Cart.SaveCart
End if

if vShipDebug then Response.write "ShippingType:" & Cart.ShippingType & "<BR>"
%>

<body bgcolor="#E5E5F0" text="#000000" link="#3770A8" vlink="#3770A8" alink="#FFFFFF"
    topmargin="0" marginheight="0" leftmargin="0" marginwidth="0">

    <center><IMG SRC="images/freeship150.gif" alt="Free Shipping* on orders over $200."></center> 
    <table width="400px" height="90%" border="0" cellpadding="0" cellspacing="0" id="tb100P">
        <tr>
            <td width="214" height="45" background="/cartimages/shiptop_bkg.gif">
                <img src="/cartimages/viewshipcharges_title.gif" width="214" height="45" border="0"><br>
            </td>
            <td width="186" height="45" align="right" background="/cartimages/shiptop_bkg.gif">
                <a href="javascript:window.close()">
                    <img src="/cartimages/closewindow_top.gif" width="86" height="45" border="0"></a><br>
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF">
                <br>
                <table width="300px" border="0" cellpadding="0" cellspacing="0" id="tb90P"> 
                    <tr>
                        <td align="center" colspan="3">
                            <font id="cartnormal"><b>Choose shipping location to view the cost of our delivery options</b>
                            </font>
                        </td>
                    </tr>
                    <tr>
                        <td >
                            &nbsp;
                        </td>
                        <td width="200px">
                            &nbsp;
                        </td>
                        <td  >
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <form name="shipnavstate" method="get">
                        <td align="right" nowrap="nowrap">
                            <font id="cartnormal">U.S. State:</font>
                        </td>
                        <td align="left" colspan="2">
                            <select name="SHIPSTATEPROVINCE" onchange="javascript:if ((document.forms.shipnavstate.SHIPSTATEPROVINCE.selectedIndex!=0) && (document.forms.shipnavstate.SHIPSTATEPROVINCE[document.forms.shipnavstate.SHIPSTATEPROVINCE.selectedIndex].value!='x')) {window.location=document.forms.shipnavstate.SHIPSTATEPROVINCE[document.forms.shipnavstate.SHIPSTATEPROVINCE.selectedIndex].value}">
                                <!--           <SELECT NAME="SHIPSTATEPROVINCE" onChange="load3(this.form,parent.frames)">  -->
                                <option value="smallship.asp">Select State</option>
                                
                                <% For Each vState in vStateSD.Keys
                                    vSelected = ""
                                    If vState = Cart.Info.ShipStateProvince then vSelected = " SELECTED"
                                    response.write("<option value=""smallship.asp?SHIPSTATEPROVINCE=" & vState & """" & vSelected & ">" & vStateSD.Item(vState) & "</option>")
                                Next %>
                            </select>
                            <br>
                        </td>
                        </form>
                    </tr>
                    <tr>
                        <form name="shipnavcountry" method="get">
                        <td align="right">
                            <font id="cartnormal">Or Country:</font>
                        </td>
                        <td align="left" colspan="2">
                            <select name="ShipCountry" onchange="javascript:if ((document.forms.shipnavcountry.ShipCountry.selectedIndex!=0) && (document.forms.shipnavcountry.ShipCountry[document.forms.shipnavcountry.ShipCountry.selectedIndex].value!='x')) {window.location=document.forms.shipnavcountry.ShipCountry[document.forms.shipnavcountry.ShipCountry.selectedIndex].value}">
                                <!--            <select name="ShipCountry" onChange="load4(this.form,parent.frames)">  -->
                                <option value="smallship.asp">Select Country</option>
                                <option value="smallship.asp?ShipCountry=US" <% if Cart.Info.ShipCountry="US" then response.write " SELECTED"%>>
                                    US&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
                                <option value="smallship.asp?ShipCountry=OTHER" <% if Cart.Info.ShipCountry="OTHER" then response.write " SELECTED"%>>
                                    Outside the U.S.</option>
                                <%call getShipCountries() %>
                            </select>
                        </td>
                        </form>
                    </tr>
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
dim vPromoFreeShipping
vPromoFreeShipping = 0 
'if 1=1 or   Request.ServerVariables("REMOTE_ADDR")  = "69.127.248.96" or Request.ServerVariables("REMOTE_ADDR")  = "10.0.0.78" or   "12/31/2009"  >left( (now()),10)  then
    if Cart.Info.ShipCountry ="US" and Cart.GridTotal>150 then
        vPromoFreeShipping = -1  
    end if
   ' vshipdebug = -1
       'response.write vsection & "<br>" & vsql1 & "<br>" & vsql2 & "<hr>"
'end if      
      For Each Item in Cart.Items
      
        ' ---- FREE SHIPPING ON NON-OVERWEIGHT ITEMS
        If ((Item.Custom7+0) = -1 or (vPromoFreeShipping  ) ) AND (Item.Custom6+0) < 1  or  (vPromoFreeShipping and cint(Item.Custom6+0) < 1) Then
            ' AND THEY DIDN'T SELECT THE FREE SHIPPING METHOD
            If Cart.Info.ShipCustom8 <> vFreeShippingMethodID then
               if vShipDebug Then Response.write Item.ItemID & "-" & "Free, but not FEDEX Ground<br>"
               vNetIgnoreFreeTotal = vNetIgnoreFreeTotal + (Item.Price * Item.Quantity)
               vNetIgnoreFreeItems = vNetIgnoreFreeItems + Item.Quantity
 
 
               ' ONLY ON THE smallship DO WE DO THIS
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
        ElseIf (Item.Custom6+0) > 0 and (Item.Custom7+0) = 0 then
           if vShipDebug Then Response.write Item.ItemID & "-" & "Overweight type:" & (Item.Custom6) & "<br>"
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
   '  vShipSelect(1,1) = "FEDEX 3 Day Select"
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
                    <tr>
                        <td align="LEFT" colspan="2">
                            <font id="carttitle"><b>SHIPPER</b></font>
                        </td>
                        <td align="RIGHT">
                            <font id="carttitle"><b>COST</b></font>
                        </td>
                    </tr>
                    <%
      dim vCurrent, vShipInfo
	  vCurrent = 0
      For vCurrent = 0 to ubound(vShipSelect)
         if Len(vShipSelect(vCurrent,1)) > 0 Then
         
            Select Case vShipSelect(vCurrent,1)
               
               Case "FEDEX Ground-US Mail"
                  vShipInfo = " East of the Mississippi is 1-4 days. West of the Mississippi is 4-6 business days. The first ship date is the date after FEDEX picks-up at our warehouse."
               
               Case "FEDEX 3 Day Select"
                  vShipInfo = "FEDEX will deliver in 3 business days. The first shipping day is the day after pickup from our warehouse. Weekends are not considered as part of the 3 day period."
      
               Case "FEDEX Next Day"
                  vShipInfo = "Next business day. Orders taken on Friday will be delivered on Monday. For Saturday delivery, please call for special rates."
            End Select %>
                    <tr bgcolor="#BCBCBC">
                        <form>
                        <td colspan="2" width="80%" align="LEFT" nowrap="nowrap">
                            <font id="cartnormal"><b>
                                <%=vShipSelect(vCurrent,1)%></b></font>
                        </td>
                        <td width="20%" align="RIGHT">
                            <font id="cartnormal"><b>
                                <%=FormatCurrency(vShipSelect(vCurrent,2),2,0,0)%></b></font>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3" align="LEFT">
                            <font id="cartnormal">
                                <%=vShipInfo%></font>
                        </td>
                    </tr>
                    <%
         End if
      Next

     ElseIf vShipZone = 4 Then   %>
                    <tr>
                        <td colspan="3">
                            <table border="0">
                                <tr>
                                    <td>
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td valign="TOP">
                                        <font id="carttitle"><b>International
                                            <%=vShipZone%></b></font>
                                    </td>
                                    <td>
                                        <font id="cartnormal"><i>A member of the BicycleBuys.com staff will contact you after
                                            your order is submitted.</i></font>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <%   End If  %>
                    <% End If  %>
                    <% If Cart.Info.ShipCountry <> "US" Then %>
                    <tr>
                        <td colspan="2">
                            <table border="0">
                                <tr colspan="2">
                                    <td>
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td valign="top">
                                        <font id="carttitle"><b>International Shipping</b></font>
                                    </td>
                                    <td>
                                        <font id="cartnormal"><i>A member of the BicycleBuys.com staff will contact you after
                                            your order is submitted.</i></font>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" align="center">
                            <% End If
   if Request.QueryString("SHIPSTATEPROVINCE") ="" and vTMP<>"US" and vTMP<>"" then 
    call getEstimateShipping(vTMP)
   end if
                            %>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <table border="0">
                                <tr colspan="2">
                                    <td>
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td valign="top">
                                        &nbsp;
                                    </td>
                                    <td>
                                        <font id="cartnormal"><b>At checkout the shipping method can be decided. We try our
                                            best to ship all orders within 48 hours of receiving. In the event that an item
                                            is not shippable within 4 days of date of order, you will be contacted.</b>
                                        </font>
                                    </td>
                                </tr>
                            </table>
                </table>
            </td>
        </tr>
        <tr background="/cartimages/shipbottom_bkg.gif" style="background: /cartimages/shipbottom_bkg.gif">
            <td width="214" height="29" background="/cartimages/shipbottom_bkg.gif">
                <img src="/cartimages/bb_mini.gif" width="214" height="29" border="0">
            </td>
            <td height="29" align="right" background="/cartimages/shipbottom_bkg.gif">
                <a href="javascript:window.close()">
                    <img src="/cartimages/closewindow_bottom.gif" width="86" height="29" border="0"></a><br>
            </td>
        </tr>
    </table>

    <script>
        try {
            _uacct = 'UA-6280466-1';
            urchinTracker("/2170435499/goal");
        } catch (err) { }
    </script>

    <script type="text/javascript">
        var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
        document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
    </script>

    <script type="text/javascript">
        try {
            var pageTracker = _gat._getTracker("UA-6280466-2");
            pageTracker._trackPageview();
        } catch (err) { }</script>

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