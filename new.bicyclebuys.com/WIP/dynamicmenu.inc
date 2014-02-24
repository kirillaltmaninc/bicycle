	<TABLE WIDTH=130 BORDER=0 CELLPADDING=7 CELLSPACING=0>
	<TR><TD align="center" background="/images/data_bkg.gif">

	<!-- -------BEGIN DYNAMIC MENU INFORMATION------------ -->
	

<%
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% SUBROUTINES
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   function nz(aval,nval)
      if isnull(aval) then 
         nz=nval
      else
         nz=nval
      end if
   end function



	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Add vendor info to dictionary
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Sub AddVendors (vSID,vVIDs)
		vArr = split(vVIDs, ",")
		for x = 0 to ubound(vArr)
'			response.write vVIDS & " - " & vArr(x) & "<br>"
			if Not SubCatVendors.Exists(vArr(x)) then
					SubCatVendors.Item(vArr(x)) = VendorListingLU.LookupValue(vArr(x))
			End If

			vV1 = "," & VendorSubcats(vArr(x)) & ","
			vV2 = "," & vArr(x) & ","
			vV3 = instr(vV1, vV2)
'			response.write "<pre>D1:V1:" & VendorSubcats(vArr(x)) & "-V2:" & vV2 & "-V3:" & vV3 & "-VSID:" & vSID & "</pre>"
			if instr(vV1, vV2) < 1 then
				if VendorSubcats(vArr(x)) <> "" then
					VendorSubcats(vArr(x)) = VendorSubcats(vArr(x)) & "," & vSID
'					response.write "<pre>D2a:X:" & x & "-VA:"  & VendorSubcats(vArr(x)) & "-V1:" & vV1 & "-V2:" & vV2 & "-V3:" & vV3 & "-VSID:" & vSID & "</pre>"
'					response.write "inif1"
				Else
					VendorSubcats(vArr(x)) = VendorSubcats(vArr(x)) & vSID
				End If
'				response.write "<pre>D2b:X:" & x & "-VA:"  & VendorSubcats(vArr(x)) & "-V1:" & vV1 & "-V2:" & vV2 & "-V3:" & vV3 & "-VSID:" & vSID & "</pre>"
			End If
'			response.write "<pre>D3:X:" & x & "-VA:"  & VendorSubcats(vArr(x)) & "-V1:" & vV1 & "-V2:" & vV2 & "-V3:" & vV3 & "-VSID:" & vSID & "</pre>"
		Next
	End Sub

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Get all Subcategories in an array
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Sub GetSubcats(SC)
		SubCats = NULL
		ReDim Subcats(300,4)

		vSCL = "," & SC & ","
		sc = 0
		for x = 0 to SubCatWDLU.Count - 1
			vInstr = instr(vSCL, SubCatWDLU.Key(x))
'			response.write x & ":" & SubCatWDLU.Key(x) & ":" & SubCatMFGLU.Value(x) & ":" & vSCL & ":" & vInstr & "<br>"
			if vInstr > 0 then
				Subcats(SubCatWDLU.Key(x),0) = SubCatWDLU.Value(x)
				Subcats(SubCatWDLU.Key(x),1) = " "
				Subcats(SubCatWDLU.Key(x),2) = SubCatMFGLU.Value(x)
				AddVendors SubCatWDLU.Key(x), SubCatMFGLU.Value(x)
				sc = sc + 1
			end if
		Next
		sc = sc - 1
	End Sub

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Get all the WebType names in an array
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Sub GetWebTypeList
		ReDim vWebTypeList(ubound(NavIDArray))
		For x = 0 to Ubound(NavIDArray)
			vWebTypeList(x) = WebTypeListingLU.LookupValue(NavIDArray(x))
		Next
	End Sub

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Get all info required to display the dynamic menu
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Sub GetNavInfo(NID)

		NavIDArray = NULL
		SubCatArray = NULL
		SortTypeArray = NULL

'		response.write ">>>" & WebNavTypeLU.LookupValue(NID) & " >>> " & NID
		vTmpArr = split(WebNavTypeLU.LookupValue(NID), "|", -1, 1)
		vWebTypes = Replace(vTmpArr(0), "+", ",")
		vSubCats = Replace(vTmpArr(1), "+", ",")
		vSortType = Replace(vTmpArr(2), "+", ",")

		if vWebTypes <> "" then NavIDArray = split(vWebTypes, ",")
		if vSubCats <> "" then SubCatArray = split(vSubCats, ",")
		if vSortType <> "" then SortTypeArray = split(vSortType,",")

		GetWebTypeList
		vSortLevels = uBound(SortTypeArray)
		
	End Sub

	Sub DisplayWebTypeNav
		CategoryList = ""
		for x = 0 to ubound(NavIDArray)
			vCatURL = "<a href=""/Items01.asp?NavID=" & vNavID & "&M=0&T=" & NavIDArray(x) & """>" & vWebTypeList(x) & "</a><br>" & chr(13)
			CategoryList = CategoryList + vCatURL
		next
		Response.Write CategoryList
	End Sub



	Sub DisplayVendorDropdown(WTID)
		CategoryList = "<FORM name=""mfgnav"" method=""get"">" & chr(13)
		CategoryList = CategoryList & "<DIV ALIGN=LEFT vAlign=Top><IMG SRC=""/cartimages/arrows_left.gif"" WIDTH=""30"" HEIGHT=""15"" BORDER=0>" & chr(13)
'		CategoryList = CategoryList & "<SELECT NAME=""vendorid"" onChange=""load1(this.form,parent.frames)"" size=""1"">" & chr(13)
      CategoryList = CategoryList & "<SELECT Size=4 NAME=""nav_select"" onChange=""javascript:if ((document.forms.mfgnav.nav_select.selectedIndex!=0) && (document.forms.mfgnav.nav_select[document.forms.mfgnav.nav_select.selectedIndex].value!='x')) {window.location=document.forms.mfgnav.nav_select[document.forms.mfgnav.nav_select.selectedIndex].value}"">" & chr(13)

		CategoryList = CategoryList & "<option value=""x"" selected>-- CHOOSE MANUFACTURER --</option>" & chr(13)
		vVend = GetVendorListByWebType(WTID)
		vCatURL = "<option value=""/Items01.asp?NavID=" & vNavID & "&M=999&T=" & vSC & """>( ALL MANUFACTURERS )</option>" & chr(13)
		CategoryList = CategoryList & vCatURL
		for x = 0 to 30
			if vVend(x,1) <> 0 then
'				response.write vVend(x,0) & "-" & vVend(x,1) & "<br>"
				a = vVend(x,0)
				b = vVend(x,1)
				vCatURL = "<option value=""/Items01.asp?NavID=" & vNavID & "&M=" & b & "&T=" & vSC & """>" & a & "</option>" & chr(13)
				CategoryList = CategoryList & vCatURL
			end if
		next
		Response.Write CategoryList & "</SELECT><IMG SRC=""/cartimages/arrows_right.gif"" WIDTH=""30"" HEIGHT=""15"" BORDER=0></DIV></FORM>" & chr(13)

		vSCx = WebTypeListingLU.LookupValue(vSC)

      ' Re-write some of the webtype names
      ' the carret "^" is special and will result in two non-breaking-spaces "&nbsp;&nbsp;"
      Select Case  vSCx
      Case "Shocks/Forks"
         vSCx = "shock/fork"
      Case "Shock Accessories"
         vSCx = "Shock Accessory"
      Case "clothes"
         vSCx = "clothing"
      Case "hydrofood"
         vSCx = "hydro/food"
      Case "accessories"
         vSCx = "accessory"
      Case "carcarriers"
         vSCx = "car^carriers"
      Case "babyseats"
         vSCx = "baby^seats"
      End Select

      ' remove trailing "s" so it sounds right with "specials" after it
      if right(vSCx,1) = "s" then vSCx = left(vSCx, Len(vSCx)-1)
      if Instr("aeiou", lcase(mid(vSCx,1,1))) = 0 then vSCx = "a &quot;" & vSCx & "&quot;" else vSCx = "an &quot;" & vSCx & "&quot;"
'		Response.Write "<FONT ID=""selectmfgmsg""><B>Please select " & vSCx & " manufacturer from the listing above.</B></FONT><P>"
      Response.Write "<FONT  ID=""selectmfgmsg""><B>Select " & vSCx & " Manufacturer for an expanded Product View.</B></FONT><P>"
	End Sub

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Get all vendor names in an array based on web type
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Function GetVendorListByWebType(WTID)
		'If we need the webtypes sorted, then generate an SQL statement
'		response.write "WTID:" & WTID

		if WTID <> 0 then
			vNavIDsql = "AND (products.WebTypeID=" & WTID & ") "

		if instr(vSortType, "WEBTYPE") = -2 then
			if vWebTypes <> "" then
				vNavIDsql = " AND ("
				For x = 0 to UBound(NavIDArray)
					vNavIDsql = vNavIDsql & "products.WebTypeID=" & NavIDArray(x)
					If x < UBound(NavIDArray) then vNavIDsql = vNavIDsql & " OR "
				Next
				vNavIDsql = vNavIDsql & ") "
			Else
				vNavIDsql = " AND (products.WebTypeID=" & vWebTypes & ") "
			End If
		End If

'		response.write "<pre>NavID:" & vNavIDsql & "</pre>"

		'If we need the subcats sorted, then generate a SQL statement
		if instr(vSortType, "SUBCAT") = -2 then
			if vSubCats <> "" then
				vSubCatsql = "("
				For x = 0 to UBound(SubCatArray)
					vSubCatsql = vSubCatsql & "products.SubcatID=" & SubCatArray(x)
					If x < UBound(SubCatArray) then vSubCatsql = vSubCatsql & " OR "
				Next
				vSubCatsql = vSubCatsql & ") "
			Else
				vSubcatsql = "(products.SubCatID=" & vSubcats & ") "
			End If
		End If

'		response.write "<pre>Subcat:" & vSubCatsql & "</pre>"

		if vSortLevels > 0 then
			vFinalSortSql = " AND "
			For x = 0 to vSortLevels
				if SortTypeArray(x) = "SUBCAT" then vFinalSortSql =  "X" & vFinalSortSql & vSubCatsql
				if SortTypeArray(x) = "WEBTYPE" then vFinalSortSql = vFinalSortSql & vNavIDsql
				if SortTypeArray(x) = "VENDORID" then vFinalSortSql = vFinalSortSql & ""
			Next
		Else
			if vSortType = "SUBCAT" then vFinalSortSql = " AND " & vSubCatsql & "xxx"
			if vSortType = "WEBTYPE" then vFinalSortSql = "" & vNavIDSql
			if vSortType = "VENDORID" then vFinalSortSql = "" & "xxx"
		End If
'		Response.write "<pre>Final: " & vFinalSortSql & "<br></pre>"

		'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
		'% Create a vendor listing
		'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
		sql= "SELECT DISTINCT products.VendID, vendor.Vendor FROM products "
		sql = sql & "INNER JOIN vendor ON products.VendID = vendor.VendID "
		sql = sql & "WHERE products.webposted LIKE 'Yes' "
		sql = sql & vFinalSortSql
		sql = sql & "ORDER BY vendor.Vendor;"

'		response.write "<pre>Complete:" & sql & "<br></pre>"

		rs.open sql,Conn,3
		CategoryList = "<form><select name=""vendorid"" onChange=""load1(this.form,parent.window)"" size=""1""><option selected>-- Vendor --</option>" & chr(13)

		vV = 0
		Dim vVendors(30,2)

		If rs.EOF = 0 then
			Do While not rs.EOF
				a = rs("Vendor")
				b = rs("VendID")
'				response.write "<pre>" & rs("sku") & "-" & b & "/" & a & chr(13) & "</pre>"
				vCatURL = "<option value=""/Items01.asp?NavID=" & vNavID & "&M=" & b & "&T=" & vSC & """>" & a & "</option>" & chr(13)
				CategoryList = CategoryList + vCatURL
				vVendors(vV,0) = a
				vVendors(vV,1) = b
				vV = vV + 1
				rs.movenext
			Loop
		rs.close
		End If	
	End If
	GetVendorListByWebType = vVendors
	End Function

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% START PROGRAM HERE
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	Dim sc, k
	Dim sql, rs, Conn, Filename, FileDSN
	Dim VendorSubcats(500), Subcats
	Dim vNavID, vMFG, VSC, vPID, vSKU, NavIDArray, SubCatArray, SortTypeArray
	Dim vWebTypes, vSubCats, vSortType, vSortLevels
	Dim vWebTypeList
	Dim vSpecialDisplayType

	vSpecialDisplayType = 0

	Set rs = Server.CreateObject("ADODB.Recordset")
	Set SubCatVendors = Server.CreateObject("Scripting.Dictionary")

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Define some working variables
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	sc = 0
	vMFG = 0
	vSC = 0
	vNavID = ""
	vPID = ""

	Dim vMv, vPageNo, vPage, vPageSize
	vPage = 1
	vPageSize = 10

	'these variables show the caption for the form submit buttons
	Dim vfirst, vlast, vnext, vprevious, vlistanum, vRec
	vFirst="FP"
	vLast="LP"
	vNext="NP"
	vPrevious="PP"
	vlistanum=vPageSize

	vMFG1 = Request.Querystring("M")	' Mfg
	vSC1 = Request.Querystring("T")		' Subcat Type
	vMv = Request.Querystring("D")		' Direction
	vPageNo = Request.Querystring("P")	' Pagenumber
	vPID = Request.Querystring("PID")	' Product ID (for specials)
	vSKU = Request.Querystring("SKU")	' Product ID (for moreinfo)

	if vPageNo = "" then
		vPageNo = 1
	End If

	if vMFG1 <> ""  and IsNULL(vMFG1)=False then
	   Session("M") = vMFG1
	Else
	   Session("M") = 0
	End If

	if vSC1 <> ""   and IsNULL(vSC1)=False then
	   Session("T") = vSC1
	Else
	   Session("T") = 0
	End If

	vNavID1 = Request.Querystring("NavID")
	if vNavID1 <> "" and IsNULL(vNavID1)=False then
	   Session("NavID") = vNavID1
	Else
	   Session("NavID") = ""
	End If		

	vNavID = Session("NavID")
	vMFG = Session("M")
	vSC = Session("T")
	if vNavID = "" then vNavID ="main"

	if IsNull(vMFG) = True or vMFG = "" or IsNumeric(vMFG)=False then vMFG = 0
	if IsNull(vSC) = True or vSC = "" or IsNumeric(vSC)=False then vSC = 0


	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create the Conn Object and open it
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	FileDSN = Application("dsn")

	Set Conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Conn.Open FileDSN


	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Display Choose Category and/or CHOOSE VENDOR if required
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	vDispMFG = False
	vDispCat = False
	vCVImage = ""
	vCVAlt = ""

	if (vSC <> 0 and (vTmp<>"/index.asp" and vTmp<>"/main.asp")) and vNavID<>"bikemfg" and vNavID<>"shoemfg" then
		vDispMFG = True
'		response.write "<BR><img src=""/images/vendor.gif"" alt=""Please choose a vendor"" height=""20"" width=""100""><BR>"
	End If

	if instr(vTmp, "cgi-bin") <> 0 then vTmp = "NOCAT"
	if vTmp <> "NOCAT" and vTmp <> "/main.asp" and vTmp <> "/index.asp" and vTmp <> "/contact.asp" and vTmp <> "/contact_confirm.asp" and vNavID<>"special" and vNavID<>"main" then
		vDispCat = True
'		response.write "<img src=""/images/category.gif"" alt=""Please pick a category"" height=""20"" width=""100""><BR><BR>"
	End If

	if vDispCat = True and vDispMFG = True then
		vCVImage = "/images/mancat.gif"
		vCVAlt = "Pick a Manufacturer from the right or a Category from below."
	ElseIf vDispCat = True and vDispMFG = False then
		vCVImage = "/images/justcat.gif"
		vCVAlt = "Pick a Category from below."
	End If

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% PROCESS MAIN PAGE
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	if vNavID="main"   then
		if vCVImage <> "" then response.write "<img src=""" & vCVImage & """ alt=""" & vCVAlt & """ height=""50"" width=""115""><BR><BR>"

		' Make sure the gray bar is using the right font
		Response.write "<FONT ID=""dynamic"">" & VbCrLf

	        Set rsitem = Server.CreateObject("ADODB.Recordset")
        
	        Dim vMPSpecials
	        vMPSpecials = Application("MPSpecials")
		response.write "<BR><BR><BR><table width=""100%"" border=""0"" align=""center"">"
		For Rec = 7 to 8
			sql = "SELECT * FROM products WHERE ProdID=" & vMPSpecials(Rec,1) & ";"
'			response.write "<pre>" & sql & "</pre><br>"
			rsitem.open sql, Conn, 3
			if NOT rsitem.EOF then
				vImage = rsitem("picture")
				if instr(vImage, "\") <> -1 then vImage = replace(vImage, "\", "/")
				vOP = vMPSpecials(Rec,2)
				vCP = int(rsitem("IsChildorParentorItem"))
				if isnull(vCP) or vCP = "" then vCP = 0
				response.write "<tr><td align=""center"">"
	on error resume next
				ShowProductVerySmall rsitem("ProdID"), vImage, rsitem("price"), replace(rsitem("description"),"""", "''"), rsitem("MarketingDescription"), nz(rsitem("MarketDescriptwo"),""), vCP, rsitem("SKU"), rsitem, vOP, 0
	on error goto 0
				if Rec = 7 then Response.write "<BR><hr><BR>"
				response.write "</td></tr>"
			End If
			rsitem.close
		Next
		response.write "</table>"
	End If
	
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% PROCESS BICYCLES
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	if vNavID="bicycles" or vNavID="bikemfg" then
		if vCVImage <> "" then response.write "<img src=""" & vCVImage & """ alt=""" & vCVAlt & """ height=""50"" width=""115""><BR><BR>"

	   vSpecialDisplayType = 1

		' Make sure the gray bar is using the right font
		Response.write "<FONT ID=""dynamic"">" & VbCrLf

		GetNavInfo("bicycles")
		GetSubcats(vSubCats)
		k = SubCatVendors.Keys
		for x = 0 to SubCatVendors.Count - 1
			response.write "<b>" & SubCatVendors.Item(k(x)) & "</b><br>"
			vTmpArr = split(VendorSubcats(k(x)), ",")
			for y = 0 to ubound(vTmpArr)
				response.write "<a href=""/Items01.asp?NavID=bikemfg&M=" & k(x) & "&T=" & vTmpArr(y) &""">"
				response.write SubCatWDLU.LookupValue(vTmpArr(y)) & "</a><br>"
			next
			response.write "<br>"
		next
	End if

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% PROCESS BICYCLE BY MFG
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	if vNavID="xbikemfg" then
		GetNavInfo("bicycles")
		GetSubcats(vSubCats)
		k = SubCatVendors.Keys
		for x = 0 to SubCatVendors.Count - 1
			if SubCatVendors.Item(k(x)) <> "" then
				response.write "<b>" & SubCatVendors.Item(k(x)) & "</b><br>"
				vTmpArr = split(VendorSubcats(k(x)), ",")
				for y = 0 to ubound(vTmpArr)
					response.write "<a href=""/Items01.asp?NavID=bikemfg&M=" & k(x) & "&T=" & vTmpArr(y) &""">"
					response.write SubCatWDLU.LookupValue(vTmpArr(y)) & "</a><br>"
				next
				response.write "<br>"
			end if
		next
	End if

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% PROCESS TOOLS
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	if vNavID="xtools" then
		GetNavInfo("tools")
		GetSubcats(vSubCats)
		k = SubCatVendors.Keys
'		response.write subcatvendors.count & "-" & vSubCats & "<br>"
		for x = 0 to SubCatVendors.Count - 1
'			response.write "<pre>" & k(x) & "</pre>"
			response.write "<b>" & SubCatVendors.Item(k(x)) & "</b><br>"
			vTmpArr = split(VendorSubcats(k(x)), ",")
			for y = 0 to ubound(vTmpArr)
'				response.write "<pre>VTA:" & vTmpArr(y) & "/" & y & "-VK:" & VendorSubcats(k(x)) & "</pre>"
				response.write "<a href=""/Items01.asp?NavID=toolmfg&M=" & k(x) & "&T=" & vTmpArr(y) &""">"
				response.write SubCatWDLU.LookupValue(vTmpArr(y)) & "</a><br>"
			next
			response.write "<br>"
		next
		if SubCatVendors.Count = 0 then Response.write "No products found."
	End if

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% PROCESS TOOLS BY MFG
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	if vNavID="xtoolmfg" then
		GetNavInfo("tools")
		GetSubcats(vSubCats)
		k = SubCatVendors.Keys
		for x = 0 to SubCatVendors.Count - 1
			if SubCatVendors.Item(k(x)) <> "" then
				response.write "<b>" & SubCatVendors.Item(k(x)) & "</b><br>"
				vTmpArr = split(VendorSubcats(k(x)), ",")
				for y = 0 to ubound(vTmpArr)
					response.write "<a href=""/Items01.asp?NavID=toolmfg&M=" & k(x) & "&T=" & vTmpArr(y) &""">"
					response.write SubCatWDLU.LookupValue(vTmpArr(y)) & "</a><br>"
				next
				response.write "<br>"
			end if
		next
		if SubCatVendors.Count = 0 then Response.write "No products found."
	End if

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% PROCESS SHOES
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	if vNavID="shoes" then
		if vCVImage <> "" then response.write "<img src=""" & vCVImage & """ alt=""" & vCVAlt & """ height=""50"" width=""115""><BR><BR>"
	   vSpecialDisplayType = 1

		' Make sure the gray bar is using the right font
		Response.write "<FONT ID=""dynamic"">" & VbCrLf
		GetNavInfo("shoes")
		GetSubcats(vSubCats)
		k = SubCatVendors.Keys
'		response.write subcatvendors.count & "-" & vSubCats
		for x = 0 to SubCatVendors.Count - 1
'			response.write "<pre>" & k(x) & "</pre>"
			response.write "<b>" & SubCatVendors.Item(k(x)) & "</b><br>"
			vTmpArr = split(VendorSubcats(k(x)), ",")
			for y = 0 to ubound(vTmpArr)
'				response.write "<pre>VTA:" & vTmpArr(y) & "/" & y & "-VK:" & VendorSubcats(k(x)) & "</pre>"
				response.write "<a href=""/Items01.asp?NavID=shoemfg&M=" & k(x) & "&T=" & vTmpArr(y) &""">"
				response.write SubCatWDLU.LookupValue(vTmpArr(y)) & "</a><br>"
			next
			response.write "<br>"
		next
		if SubCatVendors.Count = 0 then Response.write "No products found."
	End if

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% PROCESS SHOES BY MFG
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	if vNavID="shoemfg" then
		if vCVImage <> "" then response.write "<img src=""" & vCVImage & """ alt=""" & vCVAlt & """ height=""50"" width=""115""><BR><BR>"

		' Make sure the gray bar is using the right font
		Response.write "<FONT ID=""dynamic"">" & VbCrLf
		GetNavInfo("shoes")
		GetSubcats(vSubCats)
		k = SubCatVendors.Keys
		for x = 0 to SubCatVendors.Count - 1
			if SubCatVendors.Item(k(x)) <> "" then
				response.write "<b>" & SubCatVendors.Item(k(x)) & "</b><br>"
				vTmpArr = split(VendorSubcats(k(x)), ",")
				for y = 0 to ubound(vTmpArr)
					response.write "<a href=""/Items01.asp?NavID=shoemfg&M=" & k(x) & "&T=" & vTmpArr(y) &""">"
					response.write SubCatWDLU.LookupValue(vTmpArr(y)) & "</a><br>"
				next
				response.write "<br>"
			end if
		next
		if SubCatVendors.Count = 0 then Response.write "No products found."
	End if

	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% PROCESS EVERYTHING ELSE
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	if vNavID <> "bicycles" and vNavID <> "bikemfg" and vNavID <> "shoes" and vNavID<>"shoemfg" and vNavID<>"special" and vNavID<>"moreinfo" and vNavID<>"" and vNavID<>"closeouts" and vNavID<>"newitems" and vNavID<>"main" and vNavID<>"search" then
		GetNavInfo(vNavId)

      ' If we have something in our navidarray then
      ' we must have webtype selections to display
      '
      ' also, set the special display type so we know which ShowSpecial subroutine to use later in Items01.inc
      '
		if ubound(navidarray) > 0 then
		   if vSC= 0 then vSpecialDisplayType = 1 else vSpecialDisplayType = 2
         ' show the "choose cat" or "choose vendor" image
			if vCVImage <> "" then response.write "<img src=""" & vCVImage & """ alt=""" & vCVAlt & """ height=""50"" width=""115""><BR><BR>"
			' Make sure the gray bar is using the right font
			Response.write "<FONT ID=""dynamic"">" & VbCrLf
			DisplayWebTypeNav

      ' ... otherwise we need to display just the mfg dropdown
		Else
		   vSpecialDisplayType = 1
			vCVImage = "/images/justman.gif"
			vCVAlt = "Pick a Manufacturer from the right."
			if vCVImage <> "" then response.write "<img src=""" & vCVImage & """ alt=""" & vCVAlt & """ height=""50"" width=""115""><BR><BR>"
			' Make sure the gray bar is using the right font
			Response.write "<FONT ID=""dynamic"">" & VbCrLf
			vSC = NavIDArray(0)
			Session("T") = vSC
		End If
	End If

%>
</FONT>
	
	<!-- -------END DYNAMIC MENU INFORMATION------------ -->
	
	</TD></TR>
	</TABLE><BR>