<%	response.buffer=false %>
<html><head><title>BICYCLEBUYS.COM SITE TABLE RELOAD UTILITY</title></head>
<body>
<B>
<font FACE="verdana, arial, helvetica" SIZE="1">
<a href="/">BicycleBuys.com</a> Site Table Reload Utility<br>
</B>
<p>
<%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%% THIS SECTION CREATES NEW TEXT FILES FOR THE LOOKUP TABLES
'%%% Lookup tables are used in an application-wide scope to replace
'%%% constant banging on the database for what essentially is, static data.
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

	Dim Filename, FileDSN
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create the Conn Object and open it
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   FileDSN = "Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webUserprod;Initial Catalog=BBC_PROD;Data Source=10.0.0.66"
   'FileDSN = "DSN=bicyclebuys;Password=bbcwebUserprod;User ID=webUserprod;"
   'Application("FileDSN")= FileDSN

	Set Conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Conn.Open FileDSN
'	response.write filedsn
	Application("dsn") = FileDSN

	Dim UNCRoot, Results

	Dim fso, vOldFile, vNewFile, vNewFile1
	Set fso = CreateObject("Scripting.FileSystemObject")

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new webnavtype txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	UNCRoot = Server.MapPath("/") & "\..\tmp\"
	UNCRootSD = Server.MapPath("/") & "\writable\asp\"

	Set vOldFile = fso.GetFile(UNCRoot + "webnavtype.txt")
	If (fso.FileExists(UNCRoot + "webnavtype.bak")) then fso.DeleteFile(UNCRoot + "webnavtype.bak")
	vOldFile.Move(UNCRoot + "webnavtype.bak")

	sql = "SELECT * FROM WebNavType;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "webnavtype.txt", 2, True)

	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "webnavtype-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vWebNavTypeSD"
   vNewFileSD.WriteLine "Set vWebNavTypeSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		tWT = rs("WebTypes")
		tSC = rs("SubCats")
		tST = rs("SortType")
		if Len(tWT) > 0 then tWT = Replace(tWT, ",", "+")
		if Len(tSC) > 0 then tSC = Replace(tSC, ",", "+")
		if Len(tST) > 0 then tST = Replace(tST, ",", "+")
		vNewfile.WriteLine(rs("NavType") & "," & tWT & "|" & tSC & "|" & tST)

		vNewFileSD.Write "vWebNavTypeSD.Item(" & chr(34) & rs("NavType") & chr(34) & ")"
		vNewFileSD.Write " = " & chr(34) & tWT & "|" & tSC & "|" & tST & chr(34)
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop
	vNewFile.Close

	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	rs.close
	Response.write "WebNavType table reloaded.<BR>"
	Response.write "WebNavTypeSD reloaded.<BR>"

	sql = "SELECT * FROM JohnWebNavType;"
	rs.open sql,Conn,3

	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "johnwebnavtype-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vJohnWebNavTypeSD"
   vNewFileSD.WriteLine "Set vJohnWebNavTypeSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		tWT = rs("WebTypes")
		tSC = rs("SubCats")
		tST = rs("SortType")
		if Len(tWT) > 0 then tWT = Replace(tWT, ",", "+")
		if Len(tSC) > 0 then tSC = Replace(tSC, ",", "+")
		if Len(tST) > 0 then tST = Replace(tST, ",", "+")
		vNewFileSD.Write "vJohnWebNavTypeSD.Item(" & chr(34) & rs("NavType") & chr(34) & ")"
		vNewFileSD.Write " = " & chr(34) & tWT & "|" & tSC & "|" & tST & chr(34)
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop

	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	rs.close
	Response.write "JohnWebNavTypeSD reloaded.<BR>"

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new webnavtypeID txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "webnavtypeid.txt")
	If (fso.FileExists(UNCRoot + "webnavtypeid.bak")) then fso.DeleteFile(UNCRoot + "webnavtypeid.bak")
	vOldFile.Move(UNCRoot + "webnavtypeid.bak")

	sql = "SELECT * FROM WebNavType;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "webnavtypeid.txt", 2, True)
	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "webnavtypeid-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vWebNavTypeIDSD"
   vNewFileSD.WriteLine "Set vWebNavTypeIDSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		tWT = rs("WebTypes")
		vNewfile.WriteLine(rs("NavType") & "," & rs("NavTypeID"))

		vNewFileSD.Write "vWebNavTypeIDSD.Item(" & chr(34) & rs("NavType") & chr(34) & ")"
		vNewFileSD.Write " = " & chr(34) & rs("NavTypeID") & chr(34)
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop
	vNewFile.Close

	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	rs.close
	Response.write "WebNavTypeID table reloaded.<BR>"
	Response.write "WebNavTypeID-SD table reloaded.<BR>"


	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new webtypes txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "webtypes.txt")
	If (fso.FileExists(UNCRoot + "webtypes.bak")) then fso.DeleteFile(UNCRoot + "webtypes.bak")
	vOldFile.Move(UNCRoot + "webtypes.bak")

	sql = "SELECT WebTypeID, WebDisplay FROM WebType;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "webtypes.txt", 2, True)
	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "webtypes-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vWebTypeListingSD"
   vNewFileSD.WriteLine "Set vWebTypeListingSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		vNewfile.WriteLine(rs("WebTypeID") & "," & rs("WebDisplay"))

		vNewFileSD.Write "vWebTypeListingSD.Item(" & chr(34) & rs("WebTypeID") & chr(34) & ")"
		vNewFileSD.Write " = " & chr(34) & rs("WebDisplay") & chr(34)
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop
	vNewFile.Close

	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	rs.close
	Response.write "WebTypeListing table reloaded.<BR>"
	Response.write "WebTypeListingSD table reloaded.<BR>"

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new webtypesaz txt file for the A-Z lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	Set vOldFile = fso.GetFile(UNCRoot + "webtypesaz.txt")
'	If (fso.FileExists(UNCRoot & "webtypesaz.bak")) then fso.DeleteFile(UNCRoot & "webtypesaz.bak")
'	vOldFile.Move(UNCRoot & "webtypesaz.bak")
'
'	sql = "SELECT WebTypeID, WebDisplay FROM WebType ORDER BY WebDisplay;"
'	rs.open sql,Conn,3
'
'	Set vNewFile = fso.OpenTextFile(UNCRoot & "webtypesaz.txt", 2, True)
'	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "webtypesaz-sd.asp", 2, True)
'	vNewFileSD.WriteLine "<" & "% Dim vWebTypeListingAZSD"
 '  vNewFileSD.WriteLine "Set vWebTypeListingAZSD = Server.CreateObject(""Scripting.Dictionary"")"
'
'	do while not rs.EOF
'	   vCount = vCount + 1
'		vNewfile.WriteLine(vCount & "," & rs("WebDisplay") & "|" & rs("WebTypeID") ) 
'
'		vNewFileSD.Write "vWebTypeListingAZSD.Item(" & chr(34) & vCount & chr(34) & ")"
'		vNewFileSD.Write " = " & chr(34) & rs("WebDisplay") & "|" & rs("WebTypeID") & chr(34)
'		vNewFileSD.Write vbcrlf
'
'		rs.movenext
'	Loop
'	vNewFile.Close

'	vNewFileSD.WriteLine "%" & ">"
'	vNewFileSD.Close

'	rs.close
'	Response.write "WebTypes A-Z table reloaded.<BR>"
'	Response.write "WebTypes A-Z SD table reloaded.<BR>"

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new vendors txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "vendors.txt")
	If (fso.FileExists(UNCRoot + "vendors.bak")) then fso.DeleteFile(UNCRoot + "vendors.bak")
	vOldFile.Move(UNCRoot + "vendors.bak")

	sql = "SELECT VendID,Vendor FROM Vendor;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "vendors.txt", 2, True)
	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "vendors-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vVendorListingSD"
   vNewFileSD.WriteLine "Set vVendorListingSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		vNewfile.WriteLine(rs("VendID") & "," & rs("Vendor"))

		vNewFileSD.Write "vVendorListingSD.Item(" & chr(34) & rs("VendID") & chr(34) & ")"
		vNewFileSD.Write " = " & chr(34) & rs("Vendor") & chr(34)
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop
	vNewfile.WriteLine("999,All Vendors")

	vNewFileSD.Write "vVendorListingSD.Item(" & chr(34) & "999" & chr(34) & ")"
	vNewFileSD.Write " = " & chr(34) & "All Vendors" & chr(34)
	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	vNewFile.Close
	rs.close
	Response.write "Vendors table reloaded.<BR>"
	Response.write "VendorsSD table reloaded.<BR>"


	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new subcatwd txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "subcatwd.txt")
	If (fso.FileExists(UNCRoot + "subcatwd.bak")) then fso.DeleteFile(UNCRoot + "subcatwd.bak")
	vOldFile.Move(UNCRoot + "subcatwd.bak")

	Set vOldFile = fso.GetFile(UNCRoot + "subcatmfg.txt")
	If (fso.FileExists(UNCRoot + "subcatmfg.bak")) then fso.DeleteFile(UNCRoot + "subcatmfg.bak")
	vOldFile.Move(UNCRoot + "subcatmfg.bak")

	sql = "SELECT SubcatID,OnlineBikeMfg,Webdisplay FROM Subcategory WHERE OnlineBikeMfg<>'';"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "subcatwd.txt", 2, True)
	Set vNewFile1 = fso.OpenTextFile(UNCRoot + "subcatmfg.txt", 2, True)

	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "subcatwd-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vSubcatWDSD"
   vNewFileSD.WriteLine "Set vSubcatWDSD = Server.CreateObject(""Scripting.Dictionary"")"

	Set vNewFileSD1 = fso.OpenTextFile(UNCRootSD & "subcatmfg-sd.asp", 2, True)
	vNewFileSD1.WriteLine "<" & "% Dim vSubcatMFGSD"
   vNewFileSD1.WriteLine "Set vSubcatMFGSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		vNewfile1.WriteLine(rs("SubcatID") & "," & rs("OnlineBikeMfg"))
		vNewfile.WriteLine(rs("SubcatID") & "," & rs("Webdisplay"))

		vNewFileSD1.Write "vSubCatMFGSD.Item(" & chr(34) & rs("SubcatID") & chr(34) & ")"
		vNewFileSD1.Write " = " & chr(34) & rs("OnlineBikeMFG") & chr(34)
		vNewFileSD1.Write vbcrlf

		vNewFileSD.Write "vSubcatWDSD.Item(" & chr(34) & rs("SubcatID") & chr(34) & ")"
		vNewFileSD.Write " = " & chr(34) & replace(rs("WebDisplay"), """", "&quot;") & chr(34)
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop
	vNewFile.Close
	vNewFile1.Close

	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	vNewFileSD1.WriteLine "%" & ">"
	vNewFileSD1.Close

	rs.close
	Response.write "SubCatWD & SubCatMFG tables reloaded.<BR>"
	Response.write "SubCatWDSD & SubCatMFGSD tables reloaded.<BR>"

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new colors txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "colors.txt")
	If (fso.FileExists(UNCRoot + "colors.bak")) then fso.DeleteFile(UNCRoot + "colors.bak")
	vOldFile.Move(UNCRoot + "colors.bak")

	sql = "SELECT * FROM Color;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "colors.txt", 2, True)
	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "colors-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vColorListingSD"
   vNewFileSD.WriteLine "Set vColorListingSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		vNewfile.WriteLine(rs("colorid") & "," & rs("Color"))

		vNewFileSD.Write "vColorListingSD.Item(" & chr(34) & rs("colorid") & chr(34) & ")"
		vNewFileSD.Write " = " & chr(34) & replace(rs("color"), """", "&quot;") & chr(34)
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop
	vNewFile.Close

	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	rs.close
	Response.write "Colors table reloaded.<BR>"
	Response.write "ColorsSD table reloaded.<BR>"

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new sizes txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "sizes.txt")
	If (fso.FileExists(UNCRoot + "sizes.bak")) then fso.DeleteFile(UNCRoot + "sizes.bak")
	vOldFile.Move(UNCRoot + "sizes.bak")

	sql = "SELECT * FROM Size;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "sizes.txt", 2, True)
	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "sizes-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vSizeListingSD"
   vNewFileSD.WriteLine "Set vSizeListingSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		vNewfile.WriteLine(rs("sizeid") & "," & rs("size"))

		vNewFileSD.Write "vSizeListingSD.Item(" & chr(34) & rs("sizeid") & chr(34) & ")"
		vNewFileSD.Write " = " & chr(34) &  replace(rs("size"), """", "&quot;") & chr(34)
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop
	vNewFile.Close

	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	rs.close
	Response.write "Sizes table reloaded.<BR>"
	Response.write "SizesSD table reloaded.<BR>"

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new specials txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	Set vOldFile = fso.GetFile(UNCRoot + "specials.txt")
'	If (fso.FileExists(UNCRoot + "specials.bak")) then fso.DeleteFile(UNCRoot + "specialss.bak")
'	vOldFile.Move(UNCRoot + "specials.bak")

'	sql = "SELECT * FROM HTML_Special_SaleItems;"
'	rs.open sql,Conn,3

'	Set vNewFile = fso.OpenTextFile(UNCRoot + "specials.txt", 2, True)
'	do while not rs.EOF
'		vNewfile.WriteLine(rs("sizeid") & "," & rs("size"))
'		rs.movenext
'	Loop
'	vNewFile.Close
'	rs.close
'	conn.close
'	Response.write "Specials table reloaded.<BR>"

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Load up all the new lookuptable files
	'% Lookup tables will reload every 4 hours
	'% automatically.
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'Results = WebNavTypeLU.LoadValuesEX(UNCRoot + "webnavtype.txt", 10, 14400)
	'Results = WebNavTypeIDLU.LoadValuesEX(UNCRoot + "webnavtypeid.txt", 10, 14400)
	'Results = ColorListingLU.LoadValuesEX(UNCRoot + "colors.txt", 12, 14400)
	'Results = SizeListingLU.LoadValuesEX(UNCRoot + "sizes.txt", 12, 14400)
	'Results = VendorListingLU.LoadValuesEX(UNCRoot + "vendors.txt", 12, 14400)
	'Results = SubcatWDLU.LoadValuesEX(UNCRoot + "subcatwd.txt", 12, 14400)
	'Results = SubcatMFGLU.LoadValuesEX(UNCRoot + "subcatmfg.txt", 12, 14400)
	'Results = WebTypeListingLU.LoadValuesEX(UNCRoot + "webtypes.txt", 12, 14400)
	
	
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Load up an application scope array with mainpage item info
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Dim vMPSpecials(8,2)
	sql = "SELECT * FROM MainPage;"
	rs.open sql,Conn,3
	for x = 1 to 8
	   vMPSpecials(x,1) = rs("ProdID" & x)
	   vMPSpecials(x,2) = rs("OrigPriceProdID" & x)
	Next
	rs.close
	Application("MPSpecials") = vMPSpecials
	Conn.Close

	response.write "Mainpage Specials loaded.<BR>"


   response.write "<P><B>Loading Merged Backends...</B><BR>"
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Make a connection to the backends database
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   'FileDSN = "DSN=bicyclebuys;Password=bbcwebUserprod;User ID=webUserprod;"
   'Application("FileDSN")= FileDSN

	Conn.Open FileDSN

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new state txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "state.txt")
	If (fso.FileExists(UNCRoot + "state.bak")) then fso.DeleteFile(UNCRoot + "state.bak")
	vOldFile.Move(UNCRoot + "state.bak")

	sql = "SELECT * FROM state ORDER BY State;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "state.txt", 2, True)
	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "state-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vStateSD"
   vNewFileSD.WriteLine "Set vStateSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		vNewfile.WriteLine(rs("abbreviation") & "," & rs("State"))

		vNewFileSD.Write "vStateSD.Item(" & chr(34) & rs("abbreviation") & chr(34) & ")"
		vNewFileSD.Write " = " & chr(34) &  replace(rs("State"), """", "&quot;") & chr(34)
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop
	vNewFile.Close

	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	rs.close
	'Results = StateLU.LoadValuesEX(UNCRoot + "state.txt", 0, 14400)
   Response.write "State table reloaded.<BR>"
   Response.write "StateSD table reloaded.<BR>"

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new zone txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "zones.txt")
	If (fso.FileExists(UNCRoot + "zones.bak")) then fso.DeleteFile(UNCRoot + "zones.bak")
	vOldFile.Move(UNCRoot + "zones.bak")

	sql = "SELECT * FROM ShippingStateZones ORDER BY State;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "zones.txt", 2, True)
	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "zones-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vZonesSD"
   vNewFileSD.WriteLine "Set vZonesSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		vNewfile.WriteLine(rs("State") & "," & rs("Zone"))

		vNewFileSD.Write "vZonesSD.Item(" & chr(34) & rs("State") & chr(34) & ")"
		vNewFileSD.Write " = " & chr(34) &  replace(rs("Zone"), """", "&quot;") & chr(34)
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop
	vNewFile.Close

	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	rs.close
	'Results = ZonesLU.LoadValuesEX(UNCRoot + "zones.txt", 0, 14400)
   Response.write "Zone table reloaded.<BR>" 
   Response.write "ZoneSD table reloaded.<BR>" 

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new webnotes txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "webnotes.txt")
	If (fso.FileExists(UNCRoot + "webnotes.bak")) then fso.DeleteFile(UNCRoot + "webnotes.bak")
	vOldFile.Move(UNCRoot + "webnotes.bak")

	sql = "SELECT * FROM Webnotes ORDER BY WebNoteID;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "webnotes.txt", 2, True)
	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "webnotes-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vWebNoteSD"
   vNewFileSD.WriteLine "Set vWebNoteSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		vNewfile.WriteLine(rs("WebNoteID") & "," & rs("Caption"))

		vNewFileSD.Write "vWebNoteSD.Item(" & chr(34) & rs("WebNoteID") & chr(34) & ")"
		vNewFileSD.Write " = " & chr(34) &  replace(rs("Caption"), """", "&quot;") & chr(34)
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop
	vNewFile.Close

	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	rs.close
	'Results = WebnoteLU.LoadValuesEX(UNCRoot + "webnotes.txt", 0, 14400)
   Response.write "Webnotes table reloaded.<BR>" 
   Response.write "WebnotesSD table reloaded.<BR>" 

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create tax lookup SD  (10/2006)
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	sql = "SELECT * FROM [Discount Tax]"
	rs.open sql,Conn,3

	Set vNewFileSD = fso.OpenTextFile(UNCRootSD & "discounttax-sd.asp", 2, True)
	vNewFileSD.WriteLine "<" & "% Dim vDiscountTaxSD,vDiscountTaxMaxSD"
   vNewFileSD.WriteLine "Set vDiscountTaxSD = Server.CreateObject(""Scripting.Dictionary"")"
   vNewFileSD.WriteLine "Set vDiscountTaxMaxSD = Server.CreateObject(""Scripting.Dictionary"")"

	do while not rs.EOF
		'response.write rs("WebTypeID") & "," & rs("Tax")
		vNewFileSD.Write "vDiscountTaxSD.Item(" & chr(34) & rs("WebTypeID") & chr(34) & ")"
		vNewFileSD.Write " = " & rs("Tax")
		vNewFileSD.Write vbcrlf

		vNewFileSD.Write "vDiscountTaxMaxSD.Item(" & chr(34) & rs("WebTypeID") & chr(34) & ")"
		vNewFileSD.Write " = " & rs("SaleCostLessThan")
		vNewFileSD.Write vbcrlf

		rs.movenext
	Loop
	vNewFileSD.WriteLine "%" & ">"
	vNewFileSD.Close

	rs.close
   Response.write "DiscountTaxSD/DiscountTaxMaxSD table reloaded.<BR>" 

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a country name array
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   sql = "SELECT * FROM Country ORDER BY Country;"
   rs.open sql,Conn,3

   Dim vCountry(300)
   vCount = 0
	do while not rs.EOF
	   vCountry(vCount) = rs("Country")
	   vCount = vCount + 1
		rs.movenext
	Loop
	rs.close
   Application("Country") = vCountry
   Application("CountryCount") = vCount
   Response.write "Country array loaded.<BR>"

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create an shipping name array
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   sql = "SELECT * FROM [Shipping Methods];"
'   response.write sql & "<BR>"
   rs.open sql,Conn,3

   Dim vShippingNames(50)
	do while not rs.EOF
	   vShippingNames((rs("ShippingMethodID")+0)) = rs("ShippingMethod")
		rs.movenext
	Loop
	rs.close
   Application("ShippingNames") = vShippingNames
   Response.write "Shipping names array loaded.<BR>"

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create an overweight shipping array
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

   sql = "SELECT OverweightCostPerZone.*, [Shipping Methods].ShippingMethod "
   sql = sql & "FROM OverweightCostPerZone INNER JOIN [Shipping Methods] ON OverweightCostPerZone.ShippingType = [Shipping Methods].ShippingMethodID;"
'   response.write sql & "<BR>"
   rs.open sql,Conn,3

'                  zone, shippingtype, overweight cost
   Dim vOverWeight(50,         10,          3)
	do while not rs.EOF
%>
<!--   Zone: <%=rs("zone")%>  Shipping Type: <%=rs("shippingtype")%>/<%=rs("shippingmethod")%> 1/Train: <%=rs("TrainersCost")%> 2/Wheels: <%=rs("WheelsCost")%> 3/Bikes: <%=rs("BikesCost")%><BR> -->
<%
	   vOverWeight((rs("ShippingType")+0),(rs("Zone")+0), 1) = rs("TrainersCost")
	   vOverWeight((rs("ShippingType")+0),(rs("Zone")+0), 2) = rs("WheelsCost")
	   vOverWeight((rs("ShippingType")+0),(rs("Zone")+0), 3) = rs("BikesCost")
		rs.movenext
	Loop
   Application("OverWeight") = vOverWeight
   Response.write "Overweight shipping cost array loaded.<BR>"
   
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create the shipping cost per zone array
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

   rs.close
   sql = "SELECT ShippingCostPerZone.*, [Shipping Methods].ShippingMethod "
   sql = sql & "FROM ShippingCostPerZone INNER JOIN [Shipping Methods] ON ShippingCostPerZone.ShippingType = [Shipping Methods].ShippingMethodID;"
   rs.open sql,Conn,3

   Dim vSCPZ(100,10)
   Index = 0
	do while not rs.EOF
	   vSCPZ(Index, 0) = rs("Zone")
	   vSCPZ(Index, 1) = rs("PurchasePriceLow")
	   vSCPZ(Index, 2) = rs("PurchasePriceHigh")
	   vSCPZ(Index, 3) = rs("ShippingType")
	   vSCPZ(Index, 4) = rs("ShippingCost")
	   vSCPZ(Index, 5) = rs("ShippingMethod")
	   Index = Index + 1
		rs.movenext
	Loop
	rs.close
   Application("SCPZ") = vSCPZ
   Response.write "Shipping cost per zone array loaded.<BR>"

   vShipZone = 5
   vNetShippingTotal = 54.99

      ' SLIDING ADVERTISEMENTS
      ' put the message into application variables for use in moving_message.inc
      ' datediff returns a 1 for same day
      sql = "SELECT *,DATEDIFF(DAY, EndDate, GetDate()) as XX " _
          & "FROM SlideAdvertiseMent " _
          & "WHERE Active LIKE 'Y' " _
          & "AND DATEDIFF(DAY, StartDate, GetDate()) >= 1 " _
          & "AND DATEDIFF(DAY, EndDate, GetDate()) <= 1 " _
          & "ORDER BY Sequence "
      rs.open sql, Conn, 3

      Dim vSATitles(100), vSATitleColors(100), vSATitleBackgrounds(100), vSATexts(100), vSATextColors(100), vSATextBackgrounds(100), vSAImages(100), cnt
      cnt = 0
      do while NOT rs.EOF
         vSATitle = rs("Title") & ""
         vSAText = rs("Text") & ""
         vSALink = rs("Link") & ""
         vSATarget = rs("Target") & ""
         vSASequence = rs("Sequence") & ""
         vSAActive = rs("Active") & ""
         vSAStartDate = rs("StartDate") & ""
         vSAEndDate = rs("EndDate") & ""
         vSADisplay = rs("Display") & ""
         vSAImage = rs("Image") & ""
         vSATextColor = rs("TextColor") & ""
         vSATitleColor = rs("TitleColor") & ""
         
         vSATitles(cnt) = vSATitle
         if vSALink <> "" Then
            vTMP = "<a href=""" & vSALink & """"
            if vSATarget <> "" Then vTMP = vTMP & " target=""" & vSATarget & """"
            vTMP = vTMP & "><font color=""" & vSATextColor & """>" & vSAText & "</font></a>"

            if vSAImage <> "" Then
               vTMP2 = "<a href=""" & vSALink & """"
               if vSATarget <> "" Then vTMP2 = vTMP2 & " target=""" & vSATarget & """"
               vTMP2 = vTMP2 & "><img src=""" & vSAImage & """ border=""0"" align=""right""></a>"
            else
               vTMP2 = ""
            end if
         else
            vTMP = vSAText
         end if
         vSATexts(cnt) = vTMP
         vSAImages(cnt) = vTMP2

         ' set up the backgrounds
         vSATitleColors(cnt) = vSATitleColor
         vSATextColors(cnt) = vSATextColor
         vSATitleBackgrounds(cnt) = rs("TitleBackground") & ""
         vSATextBackgrounds(cnt) = rs("TextBackground") & ""

'         response.write cnt & " - " & vSATitle & "/" & vTMP & "/" & vTMP2 & "<br>"
         cnt = cnt + 1
         rs.movenext
      loop
      vSATotal = (cnt - 1)
      Application("SliderTitles") = vSATitles
      Application("SliderTexts") = vSATexts
      Application("SliderImages") = vSAImages
      Application("SliderNum") = vSATotal
      Application("SliderLast") = 0

      Application("SliderTitleColors") = vSATitleColors
      Application("SliderTextColors") = vSATextColors
      Application("SliderTitleBackgrounds") = vSATitleBackgrounds
      Application("SliderTextBackgrounds") = vSATextBackgrounds

      Response.write "Sliding Advertisement array loaded.<BR>"

	rs.close

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new welcome text file for scrolling java app
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   Dim WriteRoot
   WriteRoot = Server.MapPath("/") & "\writable\"
	Set vOldFile = fso.GetFile(WriteRoot + "welcome.txt")
	If (fso.FileExists(WriteRoot + "welcome.bak")) then fso.DeleteFile(WriteRoot + "welcome.bak")
	vOldFile.Move(WriteRoot + "welcome.bak")

	sql = "SELECT * " _
	    & "FROM MainScroll " _
	    & "WHERE Active = 1 " _
	    & "AND (GETDATE() BETWEEN StartDate AND EndDate) " _
	    & "ORDER BY Sequence"
	
	rs.open sql,Conn, 3

	Set vNewFile = fso.OpenTextFile(WriteRoot + "welcome.txt", 2, True)
	do while not rs.EOF
		vNewfile.WriteLine(rs("Title") & "|" & rs("Text") & "|" & rs("Link") & "|" & rs("Target"))
		rs.movenext
	Loop
	vNewFile.Close
	rs.close
	Response.write "Mainscroll table reloaded.<BR>"

   Conn.close

%><BR>
<B>Reload complete.</B>
<br>
<br>
<br>

</body>
</html>