<%
' *********************************
' global.asa  
' Purpose: security & site integrity  
' This script executes when the first user comes to the site.
' *********************************
 

' Set the main page we use to redirect people to if 
' they jump into the middle of the site
	response.write("Starting ASA Check"  & "<BR>")

	SiteStart = "http://www.bicyclebuys.com/"
	Application("SiteStart") = SiteStart

	' FOR CYBERSOURCE
	' set up the MerchantConfig object and store it in the Application object.
	response.write("CreateObject(CyberSourceWS.MerchantConfig)" & "<BR>")
	
	dim oMerchantConfig
	set oMerchantConfig = Server.CreateObject( "CyberSourceWS.MerchantConfig")
	oMerchantConfig.MerchantID = "v2438728"
	oMerchantConfig.SendToProduction = "1"
	oMerchantConfig.KeysDirectory = "D:\Cybersource\keys\"
	oMerchantConfig.TargetAPIVersion = "1.25"
	'Visit https://ics2ws.ic3.com/commerce/1.x/transactionProcessor/ for the
	'latest version.
	oMerchantConfig.EnableLog = "0"
	oMerchantConfig.LogDirectory = "D:\Cybersource\logs\"

	set Application( "MerchantConfig" ) = oMerchantConfig
	' END CYBERSOURCE
	response.write("END CYBERSOURCE"  & "<BR>")
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'%%% THIS SECTION CREATES NEW TEXT FILES FOR THE LOOKUP TABLES
	'%%% Lookup tables are used in an application-wide scope to replace
	'%%% constant banging on the database for what essentially is, static data.
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   ' 9/2006 - removed lookuptables due to incompatibility with IIS6

	Dim Filename, FileDSN
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create the Conn Object and open it
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'Filename = "D:\database\products.mdb"
	'FileDSN = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Filename & ";"
	

	Set Conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")

   ' this dsn works on the bb server
	FileDSN = "Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webuserprod;Initial Catalog=BBC_PROD;Data Source=webserver"
	Application("FileDSN") = FileDSN
	Application("dsn") = FileDSN

   ' this dsn works on the lihq server
	'FileDSN = "DSN=bicyclebuys;Password=bbcwebUserprod;User ID=webuserprod;"
	'Application("FileDSN") = FileDSN
	response.write("opening Conn"  & "<BR>")

	Conn.Open FileDSN 
	response.write("Successfully opened: " &  FileDSN  & "<BR>")
	Dim UNCRoot, Results
	UNCRoot = Server.MapPath("/") & "\..\tmp\"
	Dim fso, vOldFile, vNewFile, vNewFile1
	Set fso = CreateObject("Scripting.FileSystemObject")

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new webnavtype txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% 
response.write("opening file: " & UNCRoot  & "<BR>")
	Set vOldFile = fso.GetFile(UNCRoot + "webnavtype.txt")

response.write("Finished Step1 "  & "<BR>")
	If (fso.FileExists(UNCRoot + "webnavtype.bak")) then fso.DeleteFile(UNCRoot + "webnavtype.bak")

response.write("Finished Step2 "  & "<BR>")
	vOldFile.Move(UNCRoot + "webnavtype.bak")

response.write("Finished Step3 "  & "<BR>")

	sql = "SELECT * FROM WebNavType;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "webnavtype.txt", 2, True)
	do while not rs.EOF
		tWT = rs("WebTypes")
		tSC = rs("SubCats")
		tST = rs("SortType")
		if Len(tWT) > 0 then tWT = Replace(tWT, ",", "+")
		if Len(tSC) > 0 then tSC = Replace(tSC, ",", "+")
		if Len(tST) > 0 then tST = Replace(tST, ",", "+")
		vNewfile.WriteLine(rs("NavType") & "," & tWT & "|" & tSC & "|" & tST)
		rs.movenext
	Loop
	vNewFile.Close
	rs.close 

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new webnavtypeID txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "webnavtypeid.txt")
	If (fso.FileExists(UNCRoot + "webnavtypeid.bak")) then fso.DeleteFile(UNCRoot + "webnavtypeid.bak")
	vOldFile.Move(UNCRoot + "webnavtypeid.bak")

	sql = "SELECT * FROM WebNavType;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "webnavtypeid.txt", 2, True)
	do while not rs.EOF
		tWT = rs("WebTypes")
		vNewfile.WriteLine(rs("NavType") & "," & rs("NavTypeID"))
		rs.movenext
	Loop
	vNewFile.Close
	rs.close

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new webtypes txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "webtypes.txt")
	If (fso.FileExists(UNCRoot + "webtypes.bak")) then fso.DeleteFile(UNCRoot + "webtypes.bak")
	vOldFile.Move(UNCRoot + "webtypes.bak")

	sql = "SELECT WebTypeID, WebDisplay FROM WebType ORDER BY WebDisplay;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "webtypes.txt", 2, True)
	do while not rs.EOF
		vNewfile.WriteLine(rs("WebTypeID") & "," & rs("WebDisplay"))
		rs.movenext
	Loop
	vNewFile.Close
	rs.close

'	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	'% Create a new webtypesaz txt file for the A-Z lookuptable
'	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	Set vOldFile = fso.GetFile(UNCRoot + "webtypesaz.txt")
'	If (fso.FileExists(UNCRoot & "webtypesaz.bak")) then fso.DeleteFile(UNCRoot & "webtypesaz.bak")
'	vOldFile.Move(UNCRoot & "webtypesaz.bak")
'
'	sql = "SELECT WebTypeID, WebDisplay FROM WebType ORDER BY WebDisplay;"
'	rs.open sql,Conn,3
'
'	Set vNewFile = fso.OpenTextFile(UNCRoot & "webtypesaz.txt", 2, True)
'	do while not rs.EOF
'	   vCount = vCount + 1
'		vNewfile.WriteLine(vCount & "," & rs("WebDisplay") & "|" & rs("WebTypeID") ) 
'		rs.movenext
'	Loop
'	vNewFile.Close
'	rs.close
'
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new vendors txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "vendors.txt")
	If (fso.FileExists(UNCRoot + "vendors.bak")) then fso.DeleteFile(UNCRoot + "vendors.bak")
	vOldFile.Move(UNCRoot + "vendors.bak")

	sql = "SELECT VendID,Vendor FROM Vendor;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "vendors.txt", 2, True)
	do while not rs.EOF
	   dim vVendName
	   vVendName = replace( lcase(rs("Vendor")), " ", "")
	   vVendorString = vVendorString & "|" & vVendName & "|" & rs("VendID") & "|" '  We use this for the search facility
		vNewfile.WriteLine(rs("VendID") & "," & rs("Vendor"))
		rs.movenext
	Loop
	vNewfile.WriteLine("999,All Vendors")
	vNewFile.Close
	rs.close
   Application("VendorString") = vVendorString

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
	do while not rs.EOF
		vNewfile1.WriteLine(rs("SubcatID") & "," & rs("OnlineBikeMfg"))
		vNewfile.WriteLine(rs("SubcatID") & "," & rs("Webdisplay"))
		rs.movenext
	Loop
	vNewFile.Close
	vNewFile1.Close
	rs.close

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new colors txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "colors.txt")
	If (fso.FileExists(UNCRoot + "colors.bak")) then fso.DeleteFile(UNCRoot + "colors.bak")
	vOldFile.Move(UNCRoot + "colors.bak")

	sql = "SELECT * FROM Color;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "colors.txt", 2, True)
	do while not rs.EOF
		vNewfile.WriteLine(rs("colorid") & "," & rs("Color"))
		rs.movenext
	Loop
	vNewFile.Close
	rs.close

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new sizes txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "sizes.txt")
	If (fso.FileExists(UNCRoot + "sizes.bak")) then fso.DeleteFile(UNCRoot + "sizes.bak")
	vOldFile.Move(UNCRoot + "sizes.bak")

	sql = "SELECT * FROM Size;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "sizes.txt", 2, True)
	do while not rs.EOF
		vNewfile.WriteLine(rs("sizeid") & "," & rs("size"))
		rs.movenext
	Loop
	vNewFile.Close
	rs.close
'	Conn.close

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
	'Results = WebTypeListingAZLU.LoadValuesEX(UNCRoot + "webtypesaz.txt", 12, 14400)

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

   '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   '% Load up an application scope array with the non-flash navigation links
   '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	UNCRoot = Server.MapPath("/") & "\flash\"

   Dim vNF_Navigation(3,30,2)     ' 3=Bike, Body & Fun; 30 and 2 are dynamic
   Dim vNavArray, vTmpArray
   Dim vFileArray(2)
   
   vFileArray(0) = "flash_bikedata.txt"
   vFileArray(1) = "flash_bodydata.txt"
   vFileArray(2) = "flash_fundata.txt"

'   For vMenuType = 0 to 2
'      Response.write "Loading Non-Flash Menu Array: " & vFileArray(vMenuType)
'      Set vNewFile = fso.OpenTextFile(UNCRoot & vFileArray(vMenuType), 1, False)
'      vLine = vNewFile.ReadLine
'      response.write vLine & "<BR><BR><BR>"
   
'      vNavArray = Split(vLine, "&")
'      response.write Ubound(vNavArray) & "<BR>"
      
'      vTmpArray = Split(vNavArray(0), "=")
'      vMenuCount = Int(vTmpArray(1))
'      vNF_Navigation(vMenuType,0,0) = vMenuCount
   
'      vCounter = 0
'      For z = 1 to vMenuCount*3 Step 3
'         response.write "z=" & z & "<br>"
'         vCounter = vCounter + 1

'         For y = 0 to 2
'           response.write "y=" & y & "/" & z + y & ":" & vNavArray(z + y) & "<BR>"
'            vTmpArray = Split(vNavArray(z + y), "=")
'            if ubound(vTmpArray) < 2 Then
'               vNF_Navigation(vMenuType, vCounter , y) = vTmpArray(1)
'            Else
'               vNF_Navigation(vMenuType, vCounter, y) = vTmpArray(1) & "=" & vTmpArray(2)
'            End If
'           Response.write "<BR>MT=" & vMenuType & "/Cnt=" & vCounter & "/y=" & y & " - " & vNF_Navigation(vMenuType, vCounter, y) & "<br>"
'         Next
'        Response.write "-----<BR><BR>"
'      Next
 '  Next
  ' vNewFile.Close

'   vNF_Navigation(0,0,0) = "2" ' chang
   Application("NFNavigation") = vNF_Navigation


	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Make a connection to the backends database
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'Filename = "D:\database\MergedDBs_BE.mdb"
	'FileDSN = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Filename & ";"

	 FileDSN = "Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webuserprod;Initial Catalog=BBC_PROD;Data Source=webserver"
	' Conn.Open FileDSN

	'FileDSN = "DSN=bicyclebuys;Password=bbcwebUserprod;User ID=webuserprod;"
	Conn.Open FileDSN

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new state txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	UNCRoot = Server.MapPath("/") & "\..\tmp\"
	Set vOldFile = fso.GetFile(UNCRoot + "state.txt")
	If (fso.FileExists(UNCRoot + "state.bak")) then fso.DeleteFile(UNCRoot + "state.bak")
	vOldFile.Move(UNCRoot + "state.bak")

	sql = "SELECT * FROM state ORDER BY State;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "state.txt", 2, True)
	do while not rs.EOF
		vNewfile.WriteLine(rs("abbreviation") & "," & rs("State"))
		rs.movenext
	Loop
	vNewFile.Close
	rs.close
	'Results = StateLU.LoadValuesEX(UNCRoot + "state.txt", 12, 14400)

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new zone txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "zones.txt")
	If (fso.FileExists(UNCRoot + "zones.bak")) then fso.DeleteFile(UNCRoot + "zones.bak")
	vOldFile.Move(UNCRoot + "zones.bak")

	sql = "SELECT * FROM ShippingStateZones ORDER BY State;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "zones.txt", 2, True)
	do while not rs.EOF
		vNewfile.WriteLine(rs("State") & "," & rs("Zone"))
		rs.movenext
	Loop
	vNewFile.Close
	rs.close
	'Results = ZonesLU.LoadValuesEX(UNCRoot + "zones.txt", 0, 14400)

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new webnotes txt file for the lookuptable
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Set vOldFile = fso.GetFile(UNCRoot + "webnotes.txt")
	If (fso.FileExists(UNCRoot + "webnotes.bak")) then fso.DeleteFile(UNCRoot + "webnotes.bak")
	vOldFile.Move(UNCRoot + "webnotes.bak")

	sql = "SELECT * FROM Webnotes ORDER BY WebNoteID;"
	rs.open sql,Conn,3

	Set vNewFile = fso.OpenTextFile(UNCRoot + "webnotes.txt", 2, True)
	do while not rs.EOF
		vNewfile.WriteLine(rs("WebNoteID") & "," & rs("Caption"))
		rs.movenext
	Loop
	vNewFile.Close
	rs.close
	'Results = WebNoteLU.LoadValuesEX(UNCRoot + "webnotes.txt", 0, 14400)

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
	   vOverWeight((rs("ShippingType")+0),(rs("Zone")+0), 1) = rs("TrainersCost")
	   vOverWeight((rs("ShippingType")+0),(rs("Zone")+0), 2) = rs("WheelsCost")
	   vOverWeight((rs("ShippingType")+0),(rs("Zone")+0), 3) = rs("BikesCost")
		rs.movenext
	Loop
	rs.close
   Application("OverWeight") = vOverWeight

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create the shipping cost per zone array
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

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


      ' put the message into application variables for use in moving_message.inc
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

      Application("SliderDay") = Day(Now)

   rs.close

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create a new welcome text file for scrolling java app
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 
   Dim WriteRoot
   WriteRoot = Server.MapPath("/") & "\writable\"
response.write("opening2 file: " & WriteRoot & "<BR>")
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


   Conn.close
 

' *********************************
' This script executes when the server shuts down or when global.asa changes.
' *********************************
SUB Application_OnEnd()
END SUB

SUB xSession_OnStart()
	' Set the Session timeout
	' 1440 = 24 hours
	Session.Timeout=1440
	Session("RunOnce") = 0
	session("lastrundate")=""

	' Make sure that new users start on the correct
	' page index.asp...Can't jump into middle of site
	' Do a case insensitive compare, and if they
	' don't match, send the user to the start page.
	currentPage = Request.ServerVariables("SCRIPT_NAME")
		
'	If strcomp(currentPage,startPage,1) Then
'		Response.Redirect(Application("SiteStart"))
'	End If

   ' track the user's referral site
   vReferral = Request.ServerVariables("HTTP_REFERER")

   if vReferral = "" Then
      vQS = Request.ServerVariables("QUERY_STRING")
      if vQS <> "" Then
         vSID = Request.QueryString("SID")
         if vSID <> "" Then
            Session("ReferredBy") = "SALE: " & vSID & " (" & vQS & ")"
         Else
            Session("ReferredBy") = vQS
         End If
      End If
   Else
      Session("ReferredBy") = vReferral
   End If

END SUB

SUB Session_OnEnd()	
END SUB
%>
