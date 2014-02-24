<!--#INCLUDE file="includes/template_cls.asp"-->
<!--#INCLUDE file="includes/common.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Bicycle Buys | BicycleBuys.com | Online Bike Shop | Bicycles | Bike Parts | Frames | Pedals</title>
<LINK rel=stylesheet type="text/css" href="bicyclebuys.css" title="bicyclebuys">
</head>

<body>
<div align="left">
<img src="images/smallheader.jpg" /><br />


</div>

<%
vItem = request.querystring("SKU")
oProd1.getitemSKU(vItem)

' build the proper larger image height and width
Dim vHL, vWL, vHS, vWS
vItemPicture = oProd1.pfields.Item("picture")
If vItemPicture <> "" then
	if oProd1.pfields.Item("Height_Large") > 0 then
		vHL = oProd1.pfields.Item("Height_Large")
		vWL = oProd1.pfields.Item("Width_Large")
	else
		vHL = -1
	End If

	If oProd1.pfields.Item("Height_Small") > 0 then
		vHS = oProd1.pfields.Item("Height_Small")
		vWS = oProd1.pfields.Item("Width_Small")
	else
		vHS = -1
	End If

   ' if we have a defined height small
   if vHS <> -1 then
	  ' and we have a defined height large
	  if vHL <>  -1 then
		 ' then we can output the large image java popopen
		 ' used to preface the image display
		 ' vItemImageOut1 = "<A HREF=""javascript:win('/productimages/" & vItemPicture & "'," & vHL & ", " & vWL & ")"">"
	  End If
	  ' this is the flat image output
	  vItemImageOut1 = vItemImageOut1 & "<IMG class=""productimage"" SRC=""/productimages/" & vItemPicture & """ height=""" & vHS & """ width=""" & vWS & """ alt=""" & vItemDesc & """>"
	  ' if we're using the popopen, we need to end the href
	  ' if vHL <>  -1 then vItemImageOut1 = vItemImageOut1 & "</A>"
   End If
End If


vOrigPrice = ""
vSavings = ""



If (not(isNull(oProd1.pfields.Item("MSRP"))) AND (oProd1.pfields.Item("MSRP") <> "") AND IsNumeric(oProd1.pfields.Item("MSRP"))) Then
   vMSRP = oProd1.pfields.Item("MSRP")
   vPrice = oProd1.pfields.Item("price")


   ' no point in showing really low savings... over 1% and we show it
   if (vMSRP / vPrice) > 1.05 Then
	  vOrigPrice = "<div class=""minidesc"">MSRP:</div>" & FormatCurrency(vMSRP, 2, 0, 0) & "<BR>"
	  if (oProd1.pfields.Item("webnote") <> 15) then
		vSavings = "<div class=""minidesc""><span class=""product_save"">You Save:</span></div>" & FormatCurrency(vMSRP - vPrice, 2, 0, 0) & "<BR>"
	  end if
  end if
end if

vBrand = oProd1.pfields.Item("vendor")

dim RetailPrice, RebatePrice

RetailPrice = "<div class=""minidesc"">Price:</div>" & FormatCurrency(oProd1.pfields.Item("Price"), 2, 0, 0) & "<BR>"
RebatePrice = "<div class=""minidesc""><span class=price>Price After Rebate:</div><b><font size=2>" & FormatCurrency(oProd1.pfields.Item("RetailWebPrice"), 2, 0, 0) & "</font><BR>"

%>
<br>
<table width="500" border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td colspan="2"><b><font size=3><%= oProd1.pfields.Item("description") %></b><br><br></td>
  </tr>
  <tr>
    <td align="center" valign="top" width=200><%= vItemImageOut1 %></td>
    <td align="center" valign="top">
        <BR>
        <div class="minidesc">Sku#:</div>   <%= oProd1.pfields.Item("SKU") %><BR>
        <div class="minidesc">Brand:</div>   <%= vBrand %><BR>
        <%= vOrigPrice %>
        <%= RetailPrice %>
        <%= RebatePrice %>
        <%= vSavings %>
      <b>  </span>Rebate will be applied during checkout.</b>
        <BR>
    </td>
  </tr>
</table>
<BR><a href="#" onclick="window.close()"><img src="/images/close_button.gif" width="57" height="15" alt="" border="0" ></a>

	<div align="left"><br>


	    <!----------------BODY END CONTENT---------------->

    </div>
</body>
</html>
