<%
'response.write Request.Form("ITEMID")

response.buffer = true
Set Cart  = Server.CreateObject("iiscart2000.store")

%>
<!-- #INCLUDE FILE="config.asp" -->	
<%

Cart.LoadCart(Session("Cart"))

' if Len(Request.ServerVariables("HTTP_REFERER")) = 0 Then Response.Redirect("http://www.bicyclebuys.com")
' vReferer = Request.ServerVariables("HTTP_REFERER")
' if instr(vReferer, "?") > 0 Then vReferer = Left(vReferer, instr(vReferer,"?") - 1)
' if right(vReferer,27) <> ".bicyclebuys.com/Items01.asp" then Response.Redirect("http://www.bicyclebuys.com")

' Process form input
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

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% Create the Conn Object and open it
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'FileDSN = Application("FileDSN")
FileDSN = Application("dsn")

Set Conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Conn.Open FileDSN

sql = "SELECT * FROM products WHERE SKU='" & vITEMID & "';"
rs.open sql,Conn,3

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

'ON ERROR RESUME NEXT
'response.write(rs("description"))
aStr = trim(rs("description") )
aStr = replace(aStr & " ","'", "")
aStr = left(replace(aStr & " ","""", ""),255)

aStr2 = rs("marketingdescription") 
aStr2 = replace(aStr2 & " ","'", "")
aStr2 = left(replace(aStr2 & " ","""", ""),255)

Cart.AddtoCart vITEMID, 1, aStr, rs("price"),, aStr2,,,,, vURL, vReferer, vParent, vProp, vPropID, vOverWeightFlags, vFreeFreight
If Err.Number > 0 Then
   Set objNewMail = CreateObject("CDONTS.NewMail") 

   vBody =         "Client: " & Request.ServerVariables("REMOTE_ADDR") & vbCrLf
   vBody = vBody & "Error Number: " & Err.Number & vbCrLf
   vBody = vBody & "Error Description: " & Err.Description & vbCrLf
   vBody = vBody & "Error Source: " & Err.Source & vbCrLf & vbCrLf
   vBody = vBody & "ItemID: " & vITEMID & vbCrLf
   vBody = vBody & "Desc: " & replace(rs("description") & " ","""", "''") & vbCrLf
   vBody = vBody & "MDesc: " & replace(rs("marketingdescription") & " ","""", "''") & vbCrLf
   vBody = vBody & "Price: " & rs("price") & vbCrLf
   vBody = vBody & "URL: " & vURL & vbCrLf
   vBody = vBody & "Referer: " & vReferer & vbCrLf
   vBody = vBody & "Parent: " & vParent & vbCrLf
   vBody = vBody & "Prop: " & vProp & vbCrLf
   vBody = vBody & "PropID: " & vPropID & vbCrLf
   vBody = vBody & "OverWeightFlags: " & vOverWeightFlags & vbCrLf
   vBody = vBody & "FreeFreight: " & vFreeFreight & vbCrLf & vbCrLf & vbCrLf
   vBody = vBody & "CLIENT HEADERS:"  & vbCrLf & Request.ServerVariables("ALL_RAW")  & vbCrLf & vbCrLf

   objNewMail.Send "webserver.com", "webmaster@pb.net", "ERROR ON ADDTOCART.ASP", vBody, 1
   Set objNewMail = Nothing ' canNOT reuse it for another message 
End if
response.write vBody
Cart.Calculate
rs.close
Conn.close

'Cart.AddtoCart 
Session("Cart") = Cart.SaveCart
Set Cart = Nothing
response.redirect "displaycart.asp" 
%>
