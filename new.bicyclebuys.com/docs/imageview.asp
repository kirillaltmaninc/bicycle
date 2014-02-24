<% ' DISPLAY A LARGE ITEM IMAGE WITH 'CLOSE' BUTTON
 
vFN = request.querystring("fn")
vH = request.querystring("h")
vW = request.querystring("w")
vT = request.querystring("t")
vDD = ""
if vT="" then 
	vT = replace(left(vFN,len(vFN)-4),"/productimages/","")
end if
function getSkuInfo()
    dim vLoop
     
    dim oRS1, vSQL
    Dim dsn, conn
    dsn = Application("dsn")

    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open dsn

         Set oRS1 = Server.CreateObject("ADODB.Recordset")

         'response.write vSKU & "<hr>"
         'vSQL = "SELECT top 1 p.*, Vendor.Vendor " _
         '     & "FROM vwWebProducts p " _
         '     & "INNER JOIN Vendor " _
         '     & "ON Vendor.vendid = p.vendid " _
         '     & "WHERE SKU='" & vSKU & "'"  & " For Browse"
         'response.write "<hr>" & vSQL & "<hr>"
         vSQL = "exec getItemSKU '" & vT & "'"
         oRS1.open vSQL, conn, 3
         
        'response.write "<hr>" & oRS1("Caption") & "<hr>"
 

         if NOT oRS1.EOF then
            vT =  "Image: " & oRS1.fields("description")
            vDD =  replace( "Image: " & oRS1.fields("marketingdescription"),"""","'")
         end if
         oRS1.close
 end function
 getSkuInfo
%><html><title><%=vT%></title>
<meta name="description" content="<%=vDD %>" />
<META NAME="ROBOTS" CONTENT="NOINDEX, NOFOLLOW"> 
<body bgcolor="#FFFFFF">
<img src="<%=vFN%>" height="<%=vH%>" width="<%=vW%>" alt="<%=vT%>"><BR><a href="#" onclick="window.close()"><img src="/images/close_button.gif" width="57" height="15" border="0" align="right" alt="Close" ></a>
</body></html>