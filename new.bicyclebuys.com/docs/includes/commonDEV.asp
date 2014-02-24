<!--#include virtual="/writable/asp/colors-sd.asp"-->
<!--#include virtual="/writable/asp/sizes-sd.asp"-->
<!--#include virtual="/writable/asp/state-sd.asp"-->
<!--#include virtual="/writable/asp/subcatwd-sd.asp"-->
<!--#include virtual="/writable/asp/subcatmfg-sd.asp"-->
<!--#include virtual="/writable/asp/vendors-sd.asp"-->
<!--#include virtual="/writable/asp/webnavtype-sd.asp"-->
<!--#include virtual="/writable/asp/webnavtypeid-sd.asp"-->
<!--#include virtual="/writable/asp/webtypes-sd.asp"-->
<!--#include virtual="/writable/asp/webtypesaz-sd.asp"-->
<!--#include virtual="/writable/asp/webnotes-sd.asp"-->
<!--#include virtual="/writable/asp/zones-sd.asp"-->
<!--#include virtual="/writable/asp/discounttax-sd.asp"-->

<%    
   ' BICYCLEBUYS.COM
   '
   ' (c)2006 - LIHQ all rights reserved
   '
   ' common.asp

   Dim vRemote_IP, vDEBUGGING, vDebug
   vRemote_IP = Request.ServerVariables("REMOTE_ADDR")
   vDEBUGGING = 0

	'---- encryption key info
	Dim g_KeyLocation, g_Key, g_KeyString
	g_KeyLocation = "D:\root\new.bicyclebuys.com\crypt\crypt_key.txt"
	g_KeyString = ReadKeyFromFile(g_KeyLocation)


   Dim i, x, y, z, rs100, rs110, a, b, vOUTnav, TEMPCC
   Dim vCountry, vCountryCount, vShipSame, vShipType, vShipTypeA, vShipTotal, vPriceWhere, vSQLDrop

   Dim Item, vSearchPage, vType, vTermCount, vSearchVendID, vSubmit, vCount
   Dim vSearchTerm, vSearchCategory, vPriceRange

   Dim numberperpage, pagenumber, showonly, showonlybrands, counter, pagenavout, loccounter

	Dim TotalDiscount15
	TotalDiscount15 = 0

   ' pagination vars
	Dim vMv, vPageNo, vPage, vPageSize
	vPage = 1
	vPageSize = 10

	'these variables show the caption for the form submit buttons
	Dim vFirst, vLast, vNext, vPrevious, vListanum, vRec
	vFirst="FP"
	vLast="LP"
	vNext="NP"
	vPrevious="PP"
	vlistanum=vPageSize


   Dim vBGColor, vSelected, vShippingNames, vShipCount, vErrorFlag, vErrString
   Dim vITEMID, vITEMNAME, vITEMNUMBER, vPRICE, vURL, vReferer, vReferer1, vParent, vFreeFreight, vItemPicture, vItemImageOut1, vItemDesc
   Dim vTax, vSCPZ, vNetShippingTotal, vNetShippingItems, vOverWeight, vOverWeightFlags, vProp, vPropID, vPropA, vPropIDA
   Dim vNetOverSizedItems, vNetOverSizedFreeItems, vNetFreeFreightItems, vNetIgnoreFreeItems, vNetIgnoreFreeTotal
   Dim vIC6, vIC7, vDebugx, vShipDebug, vShipZone, vShippingCost
   Dim vQNum, vQUANTITY, vORIGQUANTITY, vITEMWEIGHT, vChangeQTY
   Dim vCUSTOM1, vCUSTOM2, vCUSTOM3, vCUSTOM4, vCUSTOM5, vCUSTOM6, vCUSTOM7, vCUSTOM8
   Dim vSQL100, ITEMID_1, vSQL101

   dim oProd1
   set oProd1 = new bb_product

   dim tempProd
   set tempProd = new bb_product

   dim vSection, vItem, vDept, vManufacturer, vSKU, vSect, vDepts, vItemOptions, vDesc, vDesc2, vUDept, vFreightMsg, vBrand, vWebNote

   ' final checkout variables (billing.asp)
   Dim res, vCartID, vCCYear, vCC_First
   Dim vSalesTax, vGrandTotal, vSPTitle, vSPDisp, vSP2, vSSPTitle, vSSPDisp, vSSP2, vCDisp, vSCDisp, vShippingType
   Dim vAddressListed, vPaymentTypes, vCCTypes, vCCMonths, vCCYears
   Dim fso, f, vDebugLog, vDebugLogFile
   Dim eheader, eheaderI, ebillto, eshipto, epayment, epaymentI, eitems, eitemsI, ettl, efoot
   Dim sendtrue, vOrderFile, vOrderFileName
   Const ForReading = 1, ForWriting = 2, ForAppend = 8


   Dim vTMPA
   dim vTMP1, vTMP2, vTMP3, vTMP4, vTMP5, vTMP6, vTMP7, vTMP8, vTMP9, vTMP10
   dim vTMP11, vTMP12, vTMP13, vTMP14, vTMP15, vTMP16, vTMP17, vTMP18, vTMP19, vTMP20
   dim vOUT1,vOUT2,vOUT3,vOUT4,vOUT5,vOUT6,vOUT7,vOUT8,vOUT9,vOUT10,vOUT11, vOUT100, vOUT101, vOUT102, vOUT103, vOUT104, vOUT105
   Dim vSQLMFG

   Dim vMetaTitle, vMetaDescription, vMetaKeywords, vBlurb

   Dim vPriceCount
   Set vPriceCount = CreateObject("Scripting.Dictionary")

   Dim vRecentlyViewed
   Dim vRecentlyViewedArr, vRecentlyViewedArr2

   vRecentlyViewedArr = Array("","","","","","","","","","","","","","","","","","","","")
   vRecentlyViewedArr2 = Array("","","","","","","","","","","","","","","","","","","","")

   Dim vMFG, vMFGName, vMFGID, vKey, vKeys
   Set vMFG = CreateObject("Scripting.Dictionary")
   Set vMFGName = CreateObject("Scripting.Dictionary")
   Set vMFGID = CreateObject("Scripting.Dictionary")

   dim vMSRP, vOrigPrice, vSavings, vPageTitle, vImageWidth, vFilterList

   ' --- Set these to the database values found in the shipping method table
   Dim vFreeShippingMethodID, vFreeShippingMethod
   vFreeShippingMethodID = 11
   vFreeShippingMethod = "UPS Ground"

   ' --- Flag 0% NYS Tax for clothes and shoes
   Dim vZeroTaxItems
   vZeroTaxItems = TRUE





   ' --- We need this for display purposes just about everywhere during the checkout process
   Dim vStates, vState
   Set vStates = Server.CreateObject("Scripting.Dictionary")
	vStates.Item("AP") = "APO/FPO"
	vStates.Item("AL") = "Alabama"
	vStates.Item("AK") = "Alaska"
	vStates.Item("AZ") = "Arizona"
	vStates.Item("AR") = "Arkansas"
	vStates.Item("CA") = "California"
	vStates.Item("CO") = "Colorado"
	vStates.Item("CT") = "Connecticut"
	vStates.Item("DE") = "Delaware"
	vStates.Item("DC") = "District of Columbia"
	vStates.Item("FL") = "Florida"
	vStates.Item("GA") = "Georgia"
	vStates.Item("HI") = "Hawaii"
	vStates.Item("ID") = "Idaho"
	vStates.Item("IL") = "Illinois"
	vStates.Item("IN") = "Indiana"
	vStates.Item("IA") = "Iowa"
	vStates.Item("KS") = "Kansas"
	vStates.Item("KY") = "Kentucky"
	vStates.Item("LA") = "Louisiana"
	vStates.Item("ME") = "Maine"
	vStates.Item("MD") = "Maryland"
	vStates.Item("MA") = "Massachusetts"
	vStates.Item("MI") = "Michigan"
	vStates.Item("MN") = "Minnesota"
	vStates.Item("MS") = "Mississippi"
	vStates.Item("MO") = "Missouri"
	vStates.Item("MT") = "Montana"
	vStates.Item("NE") = "Nebraska"
	vStates.Item("NV") = "Nevada"
	vStates.Item("NH") = "New Hampshire"
	vStates.Item("NJ") = "New Jersey"
	vStates.Item("NM") = "New Mexico"
	vStates.Item("NY") = "New York"
	vStates.Item("NC") = "North Carolina"
	vStates.Item("ND") = "North Dakota"
	vStates.Item("OH") = "Ohio"
	vStates.Item("OK") = "Oklahoma"
	vStates.Item("OR") = "Oregon"
	vStates.Item("PA") = "Pennsylvania"
	vStates.Item("PR") = "Puerto Rico"
	vStates.Item("RI") = "Rhode Island"
	vStates.Item("SC") = "South Carolina"
	vStates.Item("SD") = "South Dakota"
	vStates.Item("TN") = "Tennessee"
	vStates.Item("TX") = "Texas"
	vStates.Item("UT") = "Utah"
	vStates.Item("VT") = "Vermont"
	vStates.Item("VA") = "Virginia"
	vStates.Item("WA") = "Washington"
	vStates.Item("WV") = "West Virginia"
	vStates.Item("WI") = "Wisconsin"
	vStates.Item("WY") = "Wyoming"
	vStates.Item("AB") = "ALberta"
	vStates.Item("BC") = "British Columbia"
	vStates.Item("MB") = "Manitoba"
	vStates.Item("NB") = "New Brunswick"
	vStates.Item("NL") = "Newfoundland and Labrador"
	vStates.Item("NT") = "Northwest Territories"
	vStates.Item("NS") = "Nova Scotia"
	vStates.Item("NU") = "Nunavut"
	vStates.Item("ON") = "Ontario"
	vStates.Item("PE") = "Prince Edward Island"
	vStates.Item("QC") = "Quebec"
	vStates.Item("SK") = "Saskatchewan"
	vStates.Item("YT") = "Yukon"

   ' defines which nav items use subcatid
   ' subcat or webtype category display
   ' put webtype categories in here
   Dim vNavTypes
   vNavTypes = "indoortrainers"

   ' get the template engine ready
   dim objTemplate, objTemplateHeader, objTemplateFooter

   const TMPLDIR = "/templates/bb/tmpl/"
   const IMGDIR = "/templates/bb/images/"

   set objTemplate = new template_cls

   Dim vHeader, vHeaderHTML, vSearchSection, vFooter, vFinalHeader, vFinalFooter
   Dim vCartHeader, vCartHeaderSummary, vCartFooter, vCartFooterNoHelp, vCartHeaderNoSummary
   Dim vCartHeaderSummaryCheckout, vCartHeaderSummaryShipping, vCartHeaderSummaryPayment
   vHeader = TMPLDIR & "home_base_header.html"
   vSearchSection = TMPLDIR & "search_section.html"
   vFooter = TMPLDIR & "home_base_footer.html"
   vCartHeader = TMPLDIR & "cart_header.html"
   vCartHeaderSummary = TMPLDIR & "cart_header_summary.html"
   vCartFooter = TMPLDIR & "cart_footer.html"
   vCartFooterNoHelp = TMPLDIR & "cart_footer_nohelp.html"
   vFinalHeader = TMPLDIR & "final_header.html"
   vFinalFooter = TMPLDIR & "final_footer.html"
   vCartHeaderNoSummary = TMPLDIR & "cart_header.html"
   vCartHeaderSummaryCheckout = TMPLDIR & "cart_header_summary_checkout.html"
   vCartHeaderSummaryShipping = TMPLDIR & "cart_header_summary_shipping.html"
   vCartHeaderSummaryPayment = TMPLDIR & "cart_header_summary_payment.html"

   ' some working variables
   dim oRS1, oRS2, vSQL, rs, rs1, rs2
   Set RS1 = Server.CreateObject("ADODB.Recordset")
   Set RS2 = Server.CreateObject("ADODB.Recordset")
   Set rs100 = Server.CreateObject("ADODB.Recordset")
   Set rs110 = Server.CreateObject("ADODB.Recordset")


   ' We need this variable so we can define a proper reference
   ' to the secure (httpS) vs. non-secure (http) URL's.
   ' Right now, only used in displaycart.asp on the checkout button
   Dim vThisServer, vThisPort, vThisProto
   vThisServer = Request.ServerVariables("SERVER_NAME")
   vThisPort = Request.ServerVariables("SERVER_PORT")

   if vThisPort = "80" then
      vThisProto = "http://"
   else
      vThisProto = "https://"
   End if

   ' Define where the order files are saved
   Dim vSaveOrderPath
   vSaveOrderPath = "D:\JohnR\"
   'vSaveOrderPath = "D:\root\new.bicyclebuys.com\JohnR\"

   ' open the primary connection to the db
   Dim dsn, conn
   dsn = Application("dsn")

   Set conn = Server.CreateObject("ADODB.Connection")
   conn.Open dsn

   class bb_product
      '---
      'Declarations
      '---
   	public pfields
      private oRS1, oRS2
      private rsFields
      private vSQL

      private sub Class_Initialize
		   set pfields = createobject("Scripting.Dictionary")
         pfields.CompareMode = 1
      end sub

      private sub Class_Terminate
      	set pfields = nothing
      end sub

      public sub clearitem
      	set pfields = nothing
		   set pfields = createobject("Scripting.Dictionary")
      end sub

    public sub getitemPID(vProdID)
         dim vLoop

         Set oRS1 = Server.CreateObject("ADODB.Recordset")

         vSQL = "SELECT top 100 p.*, vendor.vendor, webnotes.* " _
              & "FROM products p " _
              & "INNER JOIN WebNotes " _
              & "ON webnotes.webnoteid = p.webnote " _
              & "INNER JOIN Vendor " _
              & "ON vendor.vendid = p.vendid " _
              & "WHERE ProdID=" & vProdID & " For Browse"
         'response.write "<hr>" & vSQL & "<hr>"
         oRS1.open vSQL, conn, 3
         Set rsFields = oRS1.Fields

         if NOT oRS1.EOF then
	    'with rsFields
            for vLoop = 0 to (rsFields.Count - 1)
               'response.write "<hr>" & rsFields.Item(vLoop).Name & "<br>" & rsFields.Item(vLoop).Value
               pfields.Add  rsFields.Item(vLoop).Name,  rsFields.Item(vLoop).Value
            next
 	    'end with
         end if
         oRS1.close
    	 getDiscountProd(pfields)
    end sub

Function getSQL(pFields)
'on error goto serr
	dim sql	
	' response.write("<input type=hidden name=KKK tag=" & pfields.item("SKU") & ">")
        sql = "SELECT top 1   dp.State, dp.discount,discounttype"
        sql = sql  & ", convert(varchar(10),dp.DateFrom,1) DF, convert(varchar(10),dp.DateTo,1) DT,DateDiff(d,getDate(),DateTo) DaysLeft"
        sql = sql  & " FROM [Discount Price] dp"
        sql = sql  & " WHERE dp.SetupID = 2"
        sql = sql  & " And (dp.SKU = '" & pFields.item("SKU") & "' Or dp.SKU Is Null) "
        If Not (IsNull(pFields.item("WebTypeID"))  or pFields.item("WebTypeID")="")  Then sql = sql  & " AND (dp.WebTypeID=" & pFields.item("WebTypeID") & " Or dp.WebTypeID Is Null)"
       ' If Not IsNull(pFields.item("DeptID"))    Then sql = sql  & " AND (dp.DeptID=" & pFields.item("DeptID") & " Or dp.DeptID Is Null)"
        If Not( IsNull(pFields.item("VendID")) or pFields.item("VendID")="") Then sql = sql  & " AND (dp.VendID=" & pFields.item("VendID") & " Or dp.VendID Is Null)"
        If Not (IsNull(pFields.item("SubCatID")) or pFields.item("SubCatID")="") Then sql = sql  & " AND (dp.SubCatID=" & pFields.item("SubCatID") & " Or dp.SubCatID Is Null)"
        'Ignore State clause sql = sql  & " And (dp.state = '" & state & "' Or dp.state Is Null) " 
        sql = sql  & " And dp.dateFrom <= convert(varchar(10), getDate(),101)  And dp.DateTo >=convert(varchar(10), getDate(),101) "
        sql = sql  & " ORDER BY dp.SKU DESC  , dp.subCatID DESC , dp.WebTypeID DESC , dp.VendID DESC , dp.State DESC For Browse"'
        'response.Write(sql & "<BR>")
	    getSQL = sql		
End Function    

    public sub getDiscountProd(pfields)
    
        dim sql
        dim oRSDP
        dim mProductDiscount
        dim mDaysLeft
        dim sDiscount 
        dim sDiscountType 
        dim sPrice
        
        'If Product webNote = 15 don't apply discounts
        if pfields.item("webNote") = 15 then
            pfields.Add "mSpecialPricing", ""
            pfields.Add "mDiscountType", ""
            pfields.Add "mDiscountAmount", 0    
            exit sub        
        end if
	    'Check for and Get the discount amount	
    	sql = getSQL(pfields)
        Set oRSDP = Server.CreateObject("ADODB.Recordset")	
	    oRSDP.open sql, conn, 3
        If Not oRSDP.EOF Then
        	sDiscount = oRSDP.Fields("Discount")
        	sDiscountType = oRSDP.Fields("DiscountType")

        	If Not IsNull(sDiscount) Then
		        mProductDiscount ="Special Pricing Good Till " & oRSDP.Fields("DT") & "<BR>"
		        mDaysLeft = oRSDP.Fields("DaysLeft")
		        if mDaysLeft > 1 then
			        mProductDiscount = mProductDiscount  & "Only " & mDaysLeft & " Days Remaining"  
		        elseif mDaysLeft = 1 then
			        mProductDiscount = mProductDiscount  & "Only One Day Remaining"  
		        else
			        mProductDiscount = mProductDiscount  & "Last Day"  
		        end if
                if sDiscount <> "0" then	
                    if sDiscountType ="-1" then 'Dollar
                            sPrice = pfields.Item("price") - sDiscount
                    else 'Percent
                        'response.Write( pfields.Item("price") )  ' *(1.0 - cdbl(sDiscount))  
                        sPrice =  cdbl(pfields.Item("price"))  * (1.0 - cdbl(sDiscount))                            
                    end if
                else
                    sPrice = pfields.Item("price")
                end if
		        mProductDiscount =mProductDiscount & "<BR>ONLY <B><font size=2 >" & formatcurrency(sPrice,2,0,0)  & "</font>" & ""
		        pfields.Add "mSpecialPricing", mProductDiscount
		        pfields.Add "mDiscountType", sDiscountType
		        pfields.Add "mDiscountAmount", sDiscount

            else
	            pfields.Add "mSpecialPricing", ""
	            pfields.Add "mDiscountType", ""
	            pfields.Add "mDiscountAmount", ""
            End If
        else
            pfields.Add "mSpecialPricing", ""
            pfields.Add "mDiscountType", ""
            pfields.Add "mDiscountAmount", 0    
	    End If
	oRSDP.close
	set oRSDP = nothing
  end sub



    public sub getitemSKU(vSKU)
         dim vLoop
	 
         Set oRS1 = Server.CreateObject("ADODB.Recordset")

         'response.write vsku & "<hr>"
         vSQL = "SELECT top 100 p.*, vendor.vendor " _
              & "FROM products p " _
              & "INNER JOIN Vendor " _
              & "ON vendor.vendid = p.vendid " _
              & "WHERE SKU='" & vSKU & "'"  & " For Browse"
         'response.write "<hr>" & vSQL & "<hr>"
         oRS1.open vSQL, conn, 3
         Set rsFields = oRS1.Fields

        'response.write "<hr>" & oRS1("webnote") & "<hr>"

         if NOT oRS1.EOF then
	     
            for vLoop = 0 to (rsFields.Count - 1)
               ' response.write "<hr>" & rsFields.Item(vLoop).Name & "<br>" & rsFields.Item(vLoop).Value
               pfields.Add rsFields.Item(vLoop).Name, rsFields.Item(vLoop).Value
            next
	     
	    call getDiscountProd(pfields)
'	    response.write(vSKU &  pFields("mDiscountAmount") & "DDD")
         end if
         oRS1.close

		'response.write "<hr>" & pfields.Item("WebNote") & "<hr>"

		 if (pfields.Item("webnote") <> "") then
			 vSQL = "SELECT webnotes.* " _
				  & "FROM WebNotes " _
				  & "WHERE webnoteid=" & pfields.Item("webnote") & " For Browse"
			 'response.write "<hr>" & vSQL & "<hr>"
			 oRS1.open vSQL, conn, 3
			 Set rsFields = oRS1.Fields
			 if NOT oRS1.EOF then
				for vLoop = 0 to (rsFields.Count - 1)
				   ' response.write "<hr>" & rsFields.Item(vLoop).Name & "<br>" & rsFields.Item(vLoop).Value
				   pfields.Add rsFields.Item(vLoop).Name, rsFields.Item(vLoop).Value
				next
			 end if
			 oRS1.close
		end if
		

      end sub

      public function val(vPField)
         val = pfields.Item(vPField)
      end function

   end class

   ' --- Function to clean up strings for use in db
   ' it takes the string (s), puts in two single quotes when it finds one
   ' and returns a string surrounded by single quotes
   FUNCTION CS (s, endchar)
      Dim pos
   	pos = InStr(s, "'")
   	While pos > 0
   		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
   		pos = InStr(pos + 2, s, "'")
   	Wend
    CS="'" & s & "'" & endchar
   END FUNCTION

' protection from sql injection and an overly long form value
FUNCTION reqform(vFieldName)
   Dim vFTMP
   if Not isNull(request(vFieldName)) Then
      vFTMP = Left(Server.HTMLEncode(Request(vFieldName)), 20)
   else
      vFTMP = Left(Server.HTMLEncode(Request.QueryString(vFieldName)), 20)
   end if
   reqform = vFTMP
   ' response.write "<hr>" & vfieldname & "/" & vTMP
END Function

   ' turns an encoded querystring back into straight text
   ' i.e. %20 = " " (space)
   Function URLDecode(sConvert)
       Dim aSplit
       Dim sOutput
       Dim I
       If IsNull(sConvert) Then
          URLDecode = ""
          Exit Function
       End If

       ' convert all pluses to spaces
       sOutput = REPLACE(sConvert, "+", " ")

       ' next convert %hexdigits to the character
       aSplit = Split(sOutput, "%")

       If IsArray(aSplit) Then
 	 	 if (Ubound(aSplit) <> -1) then
			 sOutput = aSplit(0)
    	 end if
	     For I = 0 to UBound(aSplit) - 1
           sOutput = sOutput & _
             Chr("&H" & Left(aSplit(i + 1), 2)) &_
             Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
         Next
       End If

       URLDecode = sOutput
   End Function

   ' given a site department, get the actual departments
   ' that should be displayed.
   public function getcatlinks2(vSection)
      dim rs1, rs2, rsFields, vSQL
      dim vLoop, vSCA, vSC, vSCs, vSubCats

      Set rs1 = Server.CreateObject("ADODB.Recordset")
      Set rs2 = Server.CreateObject("ADODB.Recordset")

      vSQL = "SELECT * " _
           & "FROM NewWebNavTypes " _
           & "WHERE WebNavID = " & vWebNavID & " For Browse"
      'response.write "<hr>" & vSQL
      'response.end

      rs1.open vSQL, conn, 3
      if Not rs1.EOF Then
         vSubCats = rs1("SubCats")
         if Not IsEmpty(vSubCats) Then
            vSCA = split(vSubCats, ",")
            for each vSC in vSCA
               if vSCs <> "" Then vSCs = vSCs & ","
               vSCs = vSCs & "'" & vSC & "'"
            next
'            vSQL = "SELECT * " _
'                 & "FROM SubCategory " _
'                 & "WHERE subcatid IN (" & vSCs & ") For Browse"
'            ' response.write "<hr>" & vsql
'            rs2.open vSQL, conn, 3
'            do while not rs2.eof

'               rs2.movenext
'            loop
         end if
      end if
      rs1.Close
      'rs2.Close
   end function

   ' to transfer a dictionary object's keys to an array
   ' --- used for sorting
   Sub BuildArray(objDict, aTempArray)
     Dim nCount, strKey
     nCount = 0

     '-- Redim the array to the number of keys we need
     Redim aTempArray(objDict.Count - 1)

     '-- Load the array
     For Each strKey In objDict.Keys

       '-- Set the array element to the key
       aTempArray(nCount) = strKey

       ' response.write "<br>++ " & nCount & "=" & strKey

       '-- Increment the count
       nCount = nCount + 1

     Next
   End Sub

   ' used after BuildArray
   Sub SortArray(aTempArray)
     Dim iTemp, jTemp, strTemp

     For iTemp = 0 To UBound(aTempArray)
       For jTemp = 0 To iTemp
         ' response.write aTempArray(jTemp) & "<br>"
         If strComp(vMFGName.Item(aTempArray(jTemp)), vMFGName.Item(aTempArray(iTemp))) > 0 Then
           'Swap the array positions
           strTemp = aTempArray(jTemp)
           aTempArray(jTemp) = aTempArray(iTemp)
           aTempArray(iTemp) = strTemp
         End If

       Next
     Next
   End Sub

   ' this will display a dictionary based on keys sorted into aTempArray
   Sub PrintDictionary(objDict, aTempArray)
     Dim iTemp
     For iTemp = 0 To UBound(aTempArray)
       Response.Write(aTempArray(iTemp) & " - " & _
                      objDict.Item(aTempArray(iTemp)) & "<br>")
     Next
   End Sub

   ' some debugging, also puts all the above subs to use
   Sub PrintSortedDictionary(objDict)
     Dim aTemp
     Call BuildArray(objDict, aTemp)
     Call SortArray(aTemp)
     Call PrintDictionary(objDict, aTemp)
   End Sub

   Public Function getcatlinksc (vSubCatID)
   'select * from johnwebnavtype where webtypes like '%148%'
      vSQL = "select * from johnwebnavtype where subcats like '%" & vSubCatID & "%' For Browse"
      rs2.open vSQL, conn, 3

      if Not rs2.EOF then
         getcatlinksc = rs2("NavType")
      else
         getcatlinksc = "ERRORa: " & vSubCatID
      end if
      rs2.close

   end function

   Public Function getcatlinkwt (vWebTypeID)
      vSQL = "select * from johnwebnavtype where WebTypes like '%" & vWebTypeID & "%' For Browse"
      rs2.open vSQL, conn, 3

      if Not rs2.EOF then
         getcatlinkwt = rs2("NavType")
      else
         getcatlinkwt = "ERRORb: " & vWebTypeID
      end if
      rs2.close

   end function

   public function getpopcategories

      vSQL = "SELECT * FROM JohnWebNavType WHERE Popular = 1 ORDER BY WebDisplayForNavType For Browse"
      'response.write vsql & "<br>"
      rs2.open vSQL, conn, 3

      Dim vLoop
      vLoop = 1
      vOUT1 = "<table border=""0"" width=""100%""><tr>"
      do while not rs2.eof

         vTMP2 = rs2("NavType")
         vTMP3 = rs2("WebDisplayForNavType")
         vOUT1 = vOUT1 & "<td><span style=""white-space:nowrap""><img src=""images/orange-arrow.gif"" width=""10"" height=""9"" border=0><a href=""/" & lcase(vTMP2) & "/"">" & vTMP3 & "</a></span></td>" & vbcrlf

         ' make sure table row is terminated afer 4 cells
         if vLoop / 4 = int(vLoop / 4) and vLoop > 1 then
            vOut1 = vOut1 & "</tr>" & chr(13)
         end if

         vLoop = vLoop + 1
         rs2.movenext
      loop

      ' put in blank cells to make sure we 4 total columns
      for x = vLoop to 4
         vOut1 = vOut1 & "<td>&nbsp;</td>"
      next
      vOut1 = vOut1 & "</table>"

      getpopcategories = vOUT1
      rs2.close

   end function


   ' returns a sql statement that will provide the product section
   Public Function gettypesql (vWebType, vSection, vManufacturer)

'     response.write "<hr>gtsql: vNavTypes= |" & vNavTypes & "| vWebType=|" & vWebType & "| vSection=|" & vSection & "| vManufacturer=|" & vManufacturer & "|"

      Dim vMFGWhere

     ' if we're dealing with a subcat or a webtype...
      if instr(vNavTypes, vSection) then
         vSect = "W"
      else
         vSect = "S"
      end if

      ' if we have a mfg then we need to display differently
      If vManufacturer <> "" Then
         vMFGWhere = " AND vendor.vendor LIKE '" & replace(vManufacturer,"_", " ") & "' "
   	end if

	   Select Case vSection
	      Case "closeouts"
	         vSect = "1"
	      Case "newitems"
	         vSect = "2"
	      Case "holiday"
	         vSect = "3"
	      Case "category"
	         vSect = "4"
	   End Select

      Select Case vSect
         Case "S"
            vTMP3 = getsubcatid(vWebType)
            if vTMP3 <> "-1" then
               vTMP3 = " AND subcatid=" & vTMP3
            else
               vTMP3 = ""
            end if

            ' get the valid departments for this subcat
            vSQL = "SELECT * " _
                 & "FROM NewWebNavTypes " _
                 & "WHERE WebNavType = '" & vSection & "' For Browse"

            ' response.write "<hr>" & vSQL
            rs1.open vSQL, conn, 3
            if Not rs1.EOF Then
               vDepts = rs1("depts")
            end if
            rs1.close

            if vDepts <> "" Then
               vTMP2 = " AND DeptID IN (" & vDepts & ") "
            end if

            ' vTMP3 = getwebtypeid(vWebType)
            vSQL = "SELECT top 100 products.*,vendor.* " _
                 & "FROM products " _
                 & "INNER JOIN Vendor " _
                 & "ON vendor.vendid = products.vendid " _
                 & "WHERE 1=1 " _
                 &  vTMP3 _
                 & vTMP2 _
                 & vMFGWhere _
                 & " AND webposted LIKE 'yes' " _
                 & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
                 & " ORDER BY retailwebprice, products.VendID, description" & " For Browse"


         ' build this sql for use with the prod display filter pulldown
         vSQLMFG = "SELECT DISTINCT products.vendid, vendor.* " _
                 & "FROM products " _
                 & "INNER JOIN Vendor " _
                 & "ON vendor.vendid = products.vendid " _
                 & "WHERE subcatid = " & vTMP3 _
                 & vTMP2 _
                 & " AND webposted LIKE 'yes' " _
                 & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
                 & " ORDER BY products.vendid"  & " For Browse"

         Case "1"    ' Closeout items
      	   vSQL = "SELECT top 100 HTML_Special_SaleItems.*, Products.* " _
      	       & "FROM HTML_Special_SaleItems " _
      	       & "INNER JOIN Products ON HTML_Special_SaleItems.Col1_ProductID = Products.ProdID " _
      	       & "WHERE HTML_Special_SaleItems.Type=" & vSect _
      	       & " AND Products.WebPosted LIKE 'yes' " _
      	       & "ORDER BY HTML_Special_SaleItems.Sort" & " For Browse"

         Case "2"    ' New items
      	   vSQL = "SELECT top 100 HTML_Special_SaleItems.*, Products.* " _
      	       & "FROM HTML_Special_SaleItems " _
      	       & "INNER JOIN Products ON HTML_Special_SaleItems.Col1_ProductID = Products.ProdID " _
      	       & "WHERE HTML_Special_SaleItems.Type=" & vSect _
      	       & " AND Products.WebPosted LIKE 'yes' " _
      	       & "ORDER BY HTML_Special_SaleItems.Sort" & " For Browse"

         Case "3"    ' Holiday Specials
      	   vSQL = "SELECT top 100 HTML_Special_SaleItems.*, Products.* " _
      	       & "FROM HTML_Special_SaleItems " _
      	       & "INNER JOIN Products ON HTML_Special_SaleItems.Col1_ProductID = Products.ProdID " _
      	       & "WHERE HTML_Special_SaleItems.Type=" & vSect _
      	       & " AND Products.WebPosted LIKE 'yes' " _
      	       & "ORDER BY HTML_Special_SaleItems.Sort" & " For Browse"

         Case "4"    ' Category Specials ???
      	   vSQL = "SELECT top 100 HTML_Special_SaleItems.*, Products.* " _
      	       & "FROM HTML_Special_SaleItems " _
      	       & "INNER JOIN Products ON HTML_Special_SaleItems.Col1_ProductID = Products.ProdID " _
      	       & "WHERE HTML_Special_SaleItems.Type=" & vSect _
      	       & " AND Products.WebPosted LIKE 'yes' " _
      	       & "ORDER BY HTML_Special_SaleItems.Sort" & " For Browse"
         'Case "5"    ' prod listing by mfg and dept
            'vSQL =

         Case "W"
            vSQL = "SELECT * " _
                 & "FROM NewWebNavTypes " _
                 & "WHERE WebNavType = '" & vSection & "' For Browse"

            ' response.write "<hr>" & vSQL

            rs1.open vSQL, conn, 3
            if Not rs1.EOF Then
               vDepts = rs1("depts")
            end if
            if vDepts <> "" Then
               vTMP2 = " AND DeptID IN (" & vDepts & ") "
            end if

            vTMP3 = getwebtypeid(vWebType)
            if vTMP3 <> "-1" then
               vTMP3 = " AND WebTypeID=" & vTMP3
            else
               vTMP3 = ""
            end if

            vSQL = "SELECT top 100 products.*, vendor.* " _
                 & "FROM products " _
                 & "INNER JOIN Vendor " _
                 & "ON vendor.vendid = products.vendid " _
                 & "WHERE 1=1 " _
                 &  vTMP3 _
                 & vTMP2 _
                 & vMFGWhere _
                 & " AND webposted LIKE 'yes' " _
                 & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
                 & " ORDER BY retailwebprice, products.VendID, description"        & " For Browse"
      end select
      ' response.write "<hr>:" & vsql
      gettypesql = vSQL
   End Function

   ' returns a sql statement that will provide the product section
   Public Function gettypesql2 (vWebType, vSection, vManufacturer)

     ' response.write "<hr>gtsql2: vWebType=|" & vWebType & "| vSection=|" & vSection & "| vManufacturer=|" & vManufacturer & "|"

      vTMP2 = "": vTMP3 = ""

      ' if we have a mfg then we need to make sure the SELECT includes only that one mfg
      Dim vMFGWhere
      If vManufacturer <> "" Then
         vMFGWhere = " AND vendor.vendor LIKE '" & replace(vManufacturer,"_", " ") & "' "
   	  end if

        if (vPriceRange = "100") then
			vPriceWhere = " AND price < 100 "
		elseif (vPriceRange = "500") then
			vPriceWhere = " AND price < 500 AND price >= 100 "
		elseif (vPriceRange = "1000") then
			vPriceWhere = " AND price < 1000 AND price >= 500 "
		elseif (vPriceRange = "2000") then
			vPriceWhere = " AND price < 2000 AND price >= 1000 "
		elseif (vPriceRange = "3000") then
			vPriceWhere = " AND price < 3000 AND price >= 2000 "
		elseif (vPriceRange = "more") then
			vPriceWhere = " AND price >= 3000 "
		else
			vPriceWhere = " "
   	    end if




'      response.write "<hr>MFGW: " & vMFGWhere & "<br>"

      ' if we're dealing with a subcat or a webtype set the flag in vSect
      if instr(vNavTypes, vSection) then
         vSect = "W"
      else
         vSect = "S"
      end if

      ' based on that flag, generate some sql
      Select Case vSect
         ' subcat's
         Case "S"

            If vWebType <> "all" Then
				if vWebType <> "" then
				   vTMP3 = getsubcatid(vWebType)
				   if vTMP3 <> -1 Then
					  vTMP3 = " AND subcatid =" & vTMP3 & " "
				   else
					  vTMP3 = getsubcatids2(vWebType)
					  if vTMP3 <> "" then
						 vTMP3 = " AND subcatid IN(" & vTMP3 & ") "
					  end if
				   end If
				 else
				  vTMP3 = " "
				end if
            Else
               vTMP3 = getsubcatids2(vSection)
               if vTMP3 <> "" then
                  vTMP3 = " AND subcatid IN(" & vTMP3 & ") "
               end if
            End If

'	response.write "type" & vWebType & " section:" & vSection & ":"& vTMP3

         ' webtypes
         Case "W"
            If vWebType <> "all" Then
               vTMP3 = getwebtypeid(vWebType)
               if vTMP3 <> "-1" then
                  vTMP3 = " AND WebTypeID IN(" & vTMP3 & ") "
               else
                  vTMP3 = getwebtypeids2(vWebType)
                  if vTMP3 <> "" then
                     vTMP3 = " AND webtypeid IN(" & vTMP3 & ") "
                  end If
               End If
            Else
               vTMP3 = getwebtypeids2(vSection)
               if vTMP3 <> "" then
                  vTMP3 = " AND webtypeid IN(" & vTMP3 & ") "
               end if
            end if
      end Select

'response.write ":" & vWebType & " : " & vSection & " : " & vManufacturer
	if (vWebType = "newitems") then
	   vSQL = "SELECT top 100 HTML_Special_SaleItems.*, Products.*, vendor.* " _
	       & "FROM HTML_Special_SaleItems " _
	       & "INNER JOIN Products ON HTML_Special_SaleItems.Col1_ProductID = Products.ProdID INNER JOIN Vendor ON vendor.vendid = products.vendid " _
	       & "WHERE HTML_Special_SaleItems.Type=" & 2 _
	       & " AND Products.WebPosted LIKE 'yes' " _
		   & vMFGWhere _
		   & vPriceWhere _
	       & "ORDER BY HTML_Special_SaleItems.Sort"  & " For Browse"


	   vSQLDrop = "SELECT top 100 HTML_Special_SaleItems.*, Products.*, vendor.* " _
	       & "FROM HTML_Special_SaleItems " _
	       & "INNER JOIN Products ON HTML_Special_SaleItems.Col1_ProductID = Products.ProdID INNER JOIN Vendor ON vendor.vendid = products.vendid " _
	       & "WHERE HTML_Special_SaleItems.Type=" & 2 _
	       & " AND Products.WebPosted LIKE 'yes' " _
	       & "ORDER BY HTML_Special_SaleItems.Sort"  & " For Browse"

	elseif (vWebType = "closeouts") then

	   vSQL = "SELECT top 100 HTML_Special_SaleItems.*, Products.*, vendor.* " _
			& "FROM HTML_Special_SaleItems " _
			& "INNER JOIN Products " _
			& "ON HTML_Special_SaleItems.Col1_ProductID = Products.ProdID INNER JOIN Vendor ON vendor.vendid = products.vendid " _
			& "WHERE HTML_Special_SaleItems.Type=1 " _
			& "AND Products.WebPosted LIKE 'yes' " _
		   & vMFGWhere _
		   & vPriceWhere _
			& "ORDER BY HTML_Special_SaleItems.Sort "  & " For Browse"

	   vSQLDrop = "SELECT top 100  HTML_Special_SaleItems.*, Products.*, vendor.* " _
			& "FROM HTML_Special_SaleItems " _
			& "INNER JOIN Products " _
			& "ON HTML_Special_SaleItems.Col1_ProductID = Products.ProdID INNER JOIN Vendor ON vendor.vendid = products.vendid " _
			& "WHERE HTML_Special_SaleItems.Type=1 " _
			& "AND Products.WebPosted LIKE 'yes' " _
			& "ORDER BY HTML_Special_SaleItems.Sort "    		& " For Browse"

	elseif (vWebType = "" AND vSection = "" AND vManufacturer <> "") then

      vSQL = "SELECT top 100 products.*,vendor.* " _
           & "FROM products " _
           & "INNER JOIN Vendor " _
           & "ON vendor.vendid = products.vendid " _
           & "WHERE 1=1 " _
           & vTMP3 _
           & vTMP2 _
           & vMFGWhere _
		   & vPriceWhere _
           & " AND webposted LIKE 'yes' " _
           & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
           & " ORDER BY description" 	   & " For Browse"
		   '& " ORDER BY retailwebprice, products.VendID, description"

      vSQLDrop = "SELECT top 100 products.*,vendor.* " _
           & "FROM products " _
           & "INNER JOIN Vendor " _
           & "ON vendor.vendid = products.vendid " _
           & "WHERE 1=1 " _
           & vTMP3 _
           & vTMP2 _
		   & vMFGWhere _
           & " AND webposted LIKE 'yes' " _
           & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
           & " ORDER BY description"   & " For Browse"

	elseif (vWebType <> "" AND vSection <> "" AND vManufacturer <> "") then

      vSQL = "SELECT top 100 products.*,vendor.* " _
           & "FROM products " _
           & "INNER JOIN Vendor " _
           & "ON vendor.vendid = products.vendid " _
           & "WHERE 1=1 " _
           & vTMP3 _
           & vTMP2 _
			& vPriceWhere _
			& vMFGWhere _
			& " AND webposted LIKE 'yes' " _
           & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
           & " ORDER BY description"   & " For Browse"
		   '& " ORDER BY retailwebprice, products.VendID, description"

		 '  response.write vSQL

      vSQLDrop = "SELECT top 100 products.*,vendor.* " _
           & "FROM products " _
           & "INNER JOIN Vendor " _
           & "ON vendor.vendid = products.vendid " _
           & "WHERE 1=1 " _
           & vTMP3 _
           & vTMP2 _
           & " AND webposted LIKE 'yes' " _
           & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
           & " ORDER BY description"   & " For Browse"


	else
      vSQL = "SELECT top 100 products.*,vendor.* " _
           & "FROM products " _
           & "INNER JOIN Vendor " _
           & "ON vendor.vendid = products.vendid " _
           & "WHERE 1=1 " _
           & vTMP3 _
           & vTMP2 _
           & vMFGWhere _
		   & vPriceWhere _
           & " AND webposted LIKE 'yes' " _
           & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
           & " ORDER BY description"    & " For Browse"
		   '& " ORDER BY retailwebprice, products.VendID, description"

      vSQLDrop = "SELECT top 100 products.*,vendor.* " _
           & "FROM products " _
           & "INNER JOIN Vendor " _
           & "ON vendor.vendid = products.vendid " _
           & "WHERE 1=1 " _
           & vTMP3 _
           & vTMP2 _
           & " AND webposted LIKE 'yes' " _
           & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
           & " ORDER BY description"    & " For Browse"
	end if




      ' build this sql for use with the prod display filter pulldown
      vSQLMFG = "SELECT DISTINCT products.vendid, vendor.* " _
           & "FROM products " _
           & "INNER JOIN Vendor " _
           & "ON vendor.vendid = products.vendid " _
           & "WHERE 1=1 " _
           & vTMP3 _
           & vTMP2 _
           & vMFGWhere _
           & " AND webposted LIKE 'yes' " _
           & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
           & " ORDER BY description"   & " For Browse"
		  ' & " ORDER BY products.vendid"
      gettypesql2 = vSQL

      ' response.write "<hr>:" & vSQL
      ' response.end
   End Function

   Public Function getwebtypeid (vWebType)
      dim rs1
      Set rs1 = Server.CreateObject("ADODB.Recordset")

      vSQL = "SELECT webtypeid " _
           & "FROM webtype " _
           & "WHERE webtype = '" & vWebType & "' For Browse"
      ' response.write "<hr>" & vsql
      rs1.open vSQL, conn, 3
      If Not rs1.EOF Then
         getwebtypeid = rs1("webtypeid")
      Else
         getwebtypeid = -1
      End If
      rs1.close
   End Function

   ' returns a single subcatid found in the 'subcategory' table that matches the
   ' subcategory field (i.e. "wallstorage")
   Public Function getsubcatid (vSubCat)
      dim rs1
      Set rs1 = Server.CreateObject("ADODB.Recordset")

      vTMP1 = replace(vSubCat, "'", "''")

      vSQL = "SELECT subcatid " _
           & "FROM subcategory " _
           & "WHERE subcategory = '" & vTMP1 & "' For Browse"
      'response.write "<hr>" & vsql
      rs1.open vSQL, conn, 3
      If Not rs1.EOF Then
         getsubcatid = rs1("subcatid")
      Else
         getsubcatid = -1
      End If
      rs1.close
   End Function

   ' returns the comma separated (##,##,##) subcats value based on the subcat
   '   -- subcat is the navtype (i.e. "storagesystems") that matches in the johnwebnavtype table)
   Public Function getsubcatids2 (vSubCat)
      dim rs1
      Set rs1 = Server.CreateObject("ADODB.Recordset")

      vSQL = "SELECT subcats " _
           & "FROM JohnWebNavType " _
           & "WHERE NavType LIKE '" & vSubCat & "' For Browse"
'      response.write "<hr>getsubcatids2: " & vsql

      rs1.open vSQL, conn, 3
      If Not rs1.EOF Then
         vTMP1 = rs1("SubCats")
      Else
         vTMP1 = ""
      End If
      rs1.close

      getsubcatids2 = vTMP1
'      response.write "<hr>" &  vTMP1

   End Function

   Public Function getwebtypeid2 (vWebType)
      dim rs1
      Set rs1 = Server.CreateObject("ADODB.Recordset")

      vSQL = "SELECT webtypes " _
           & "FROM JohnWebNavType " _
           & "WHERE NavType LIKE '" & vWebType & "' For Browse"
'      response.write "<hr>" & vsql
      rs1.open vSQL, conn, 3
      If Not rs1.EOF Then
         vTMP1 = rs1("WebTypes")
      Else
         vTMP1 = ""
      End If
      rs1.close

      getwebtypeid2 = vTMP1

   End Function

   ' returns the comma separated (##,##,##) webtypes value based on the webtype
   '   -- webtype is the navtype (i.e. "storagesystems") that matches in the johnwebnavtype table)
   Public Function getwebtypeids2 (vWebType)
      dim rs1
      Set rs1 = Server.CreateObject("ADODB.Recordset")

      vSQL = "SELECT webtypes " _
           & "FROM JohnWebNavType " _
           & "WHERE NavType LIKE '" & vWebType & "' For Browse"
      'response.write "<hr>gwtids2: " & vsql
      rs1.open vSQL, conn, 3
      If Not rs1.EOF Then
         vTMP1 = rs1("WebTypes")
      Else
         vTMP1 = ""
      End If
      rs1.close

      getwebtypeids2 = vTMP1

   End Function


   ' pass this a primary nav name and it will return
   ' out1 and out2, used for the two column category list
   ' that is displayed when that nav is clicked
   ' the HREF url is SEF - re-write enabled
   public sub getcatlinks(vSection)
      dim rs1, rs2, rsFields, vSQL1, vSQL2, vSQL3, vNavSet
      dim vLoop, vDA, vD, vDs

      Set rs1 = Server.CreateObject("ADODB.Recordset")
      Set rs2 = Server.CreateObject("ADODB.Recordset")

      ' set the title
      vOut3 = vSection

      ' set the breadcrumb link
      vTMP1 = UCase(Left(vSection,1)) & Right(vSection,Len(vSection)-1)
      vOUT4 = vTMP1


      ' we break up the sql for each section with the idea that
      ' eventually we may need this as neil works out his
      ' product categorization.   "WebDisplay" is common.
      ' response.write "<hr>" & vNavTypes & "|" & vSection & "<br>"
      if instr(vNavTypes, vSection) then vSect = "W" else vSect = "S"
      Select Case vSect
         Case "S"
            vNavSet = "SubCats NS, WebDisplayForNavType, WebDisplayForCategory "
            vSQL1 = "SELECT MetaTitle, MetaDescription, MetaKeywords, " & vNavSet _
                  & "FROM JohnWebNavType " _
                  & "WHERE NavType LIKE '" & vSection & "' For Browse"

            vSQL2 = "SELECT SubCatID, SubCategory NT, WebDisplay " _
                 & "FROM SubCategory " _
                 & "WHERE SubCatID IN ( "

         Case Else
            vNavSet = "NavType, WebTypes NS, WebDisplayForNavType, WebDisplayForCategory "
            vSQL1 = "SELECT MetaTitle, MetaDescription, MetaKeywords, " & vNavSet  _
                  & "FROM JohnWebNavType " _
                  & "WHERE NavType LIKE '" & vSection & "' For Browse"

            vSQL2 = "SELECT WebTypeID, WebType NT, WebDisplay " _
                 & "FROM Webtype " _
                 & "WHERE WebTypeID IN ( "
      end select
      ' response.write vsection & "<br>" & vsql1 & "<br>" & vsql2 & "<hr>"
      ' response.end
      rs1.open vSQL1, conn, 3
      if Not rs1.EOF Then
         vOut2 = rs1("WebDisplayForNavType")
         vOut3 = rs1("WebDisplayForCategory")
         vMetaTitle = rs1("MetaTitle")
         vMetaDescription = rs1("MetaDescription")
         vMetaKeywords = rs1("MetaKeywords")

         vTMP1 = rs1("NS")
         if Not IsEmpty(vTMP1) Then
            vDA = split(vTMP1, ",")
            for each vD in vDA
               if vDs <> "" Then vDs = vDs & ","
               vDs = vDs & "'" & vD & "'"
            next

            Select Case vSect
               Case "S"
                  vSQL2 = vSQL2 & "SELECT DISTINCT SubCatID " _
                        & "FROM products " _
                        & "WHERE subcatid IN (" & vDs & ") " _
                        & " AND subcatID IS NOT NULL " _
                        & " AND subcatID != 0 " _
                        & " AND webposted LIKE 'yes' " _
                        & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
                        & ") " _
                        & " ORDER BY WebDisplay"  & " For Browse"
               Case Else
                  vSQL2 = vSQL2 & "SELECT DISTINCT WebTypeID " _
                        & "FROM products " _
                        & "WHERE webtypeid IN (" & vDs & ") " _
                        & " AND WebTypeID IS NOT NULL " _
                        & " AND WebTypeID != 0 " _
                        & " AND webposted LIKE 'yes' " _
                        & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
                        & ") " _
                        & " ORDER BY WebDisplay"       & " For Browse"
            end select

            ' response.write "<hr>" & vsql2

            rs2.open vSQL2, conn, 3
            do while not rs2.eof
               ' Remove spaces, replace with underscores
               vTMP2 = Replace(rs2("NT"), " ", "_")
               vLoop = vLoop + 1
               vOUT1 = vOUT1 & "<span style=""white-space:nowrap""><img src=""images/orange-arrow.gif"" width=""10"" height=""9"" border=0><a href=""/" & vSection & "/" & vTMP2 & "/"">" & rs2("webdisplay") & "</a></span>" & vbcrlf
               rs2.movenext
            loop
            rs2.close
         end if
      end if
      rs1.close
   end sub


   ' pass this a primary nav name and it will return
   ' out5 and out6, used for the two column category list
   ' that is displayed when that nav is clicked
   ' the HREF url is SEF - re-write enabled

'       <form name="myform" action="handle-data.php">
'      Search: <input type='text' name='query'>
'      <A href="javascript: submitform()">Search</A>
'      </form>
'      <SCRIPT language="JavaScript">
'      function submitform()
'      {
'        document.myform.submit();
'      }
'      </SCRIPT>
   public sub getmfglinks(vSection)
      dim rs1, rs2, rsFields, vSQL1, vSQL2, vSQL3, vNavSet
      dim vLoop, vDA, vD, vDs

      Set rs1 = Server.CreateObject("ADODB.Recordset")
      Set rs2 = Server.CreateObject("ADODB.Recordset")

      ' we need the subcat listing to generate the mfg listing
      ' so we do the same thing to figure out webtype or subcat...
      ' then we can x-ref with the vendor table

      ' using new table - JohnWebNavType
      if instr(vNavTypes, vSection) then vSect = "W" else vSect = "S"
      Select Case vSect
         Case "S"
            vNavSet = "SubCats"
            vSQL1 = "SELECT " & vNavSet & " NS " _
                  & "FROM JohnWebNavType " _
                  & "WHERE NavType LIKE '" & vSection & "' For Browse"

            vSQL2 = "SELECT SubCatID, SubCategory NT, WebDisplay " _
                 & "FROM SubCategory " _
                 & "WHERE SubCatID IN ( "

         Case Else
            vNavSet = "WebTypes"
            vSQL1 = "SELECT " & vNavSet & " NS " _
                  & "FROM JohnWebNavType " _
                  & "WHERE NavType LIKE '" & vSection & "' For Browse"

            vSQL2 = "SELECT WebTypeID, WebType NT, WebDisplay " _
                 & "FROM webtype " _
                 & "WHERE WebTypeID IN ( "
      end select
      ' response.write vsection & "<br>getmfglinks: " & vNavSet & "<br>sql1=" & vsql1 & "<br>sql2=" & vsql2 & "<hr>"

      rs1.open vSQL1, conn, 3
      if Not rs1.EOF Then
         vTMP1 = rs1("NS")
         vDs = "'" & replace(vTMP1,",","','") & "'"
         ' ok - vDs is all the webtypeid's for the subcat/webtype chosen

         if Not IsEmpty(vTMP1) Then
            Select Case vSect
               Case "S"
                  vSQL2 = "SELECT DISTINCT p.VendID,v.vendor " _
                        & "FROM products p " _
                        & "INNER JOIN vendor v " _
                        & "ON v.VendID = p.VendID " _
                        & "WHERE subcatid IN (" & vDs & ") " _
                        & " AND subcatID IS NOT NULL " _
                        & " AND subcatID != 0 " _
                        & " AND webposted LIKE 'yes' " _
                        & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
                        & " ORDER BY v.vendor"      & " For Browse"

               Case Else
                  vSQL2 = "SELECT DISTINCT p.VendID,v.vendor " _
                        & "FROM products p " _
                        & "INNER JOIN vendor v " _
                        & "ON v.VendID = p.VendID " _
                        & "WHERE webtypeid IN (" & vDs & ") " _
                        & " AND WebTypeID IS NOT NULL " _
                        & " AND WebTypeID != 0 " _
                        & " AND webposted LIKE 'yes' " _
                        & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
                        & " ORDER BY v.vendor"        & " For Browse"
            end select

            ' response.write "<hr>sql2=" & vsql2 & "<br><br><br>"
            rs2.open vSQL2, conn, 3
            do while not rs2.eof
               ' Remove spaces, replace with underscores
               vTMP2 = Replace(rs2("vendor"), " ", "_")
               vLoop = vLoop + 1

               vOUT5 = vOUT5 & "<span style=""white-space:nowrap""><img src=""images/orange-arrow.gif"" width=""10"" height=""9"" border=0><a href=""/manufacturer/" & vTMP2 & "/" & vSection & """>" & rs2("vendor") & "</a></span>"

               rs2.movenext
            loop

         end if
      end if
      rs1.close
   end sub


   ' puts together a mfg linked listing based on the johnnewwebnavtype subcats field
   sub getmfglinks2 (vNavType)

      Dim vCatName, vCatDisp, vSubCats
      vSQL = "SELECT MP.ID, MP.sideheading, JWNT.NavType, JWNT.subcats, JWNT.WebDisplayForCategory " _
           & "FROM MainPage MP " _
           & "INNER JOIN JohnWebNavType JWNT " _
           & "ON MP.sideheading = JWNT.NavTypeID " _
           & "WHERE MP.ID = 1 For Browse"
      'response.write "<hr>" & vSQL
      rs1.open vSQL, conn, 3

      vCatDisp = rs1("WebDisplayForCategory")
      vCatName = LCase(rs1("NavType"))
      vSubCats = rs1("subcats")

      rs1.close



   end sub

   ' returns an html excerpt of the a product listing
   Sub getprodlinks (vDept, vSection, vManufacturer)
      dim rs1, vSP, vP, vPrefURL
      Set rs1 = Server.CreateObject("ADODB.Recordset")

      'response.write "<hr>gpl|vDept=" & vDept & "|/|vSection=" & vSection & "|/|vManufacturer=" & vManufacturer & "|<br>"

      ' if section is allmfg then we expect a mfg to be sent
      ' -- product display is dependent on vendid and dept
      if vSection = "allmfg" Then

         ' clean the dept name for use in url
         vTMP1 = UCase(Left(vDept,1)) & Right(vDept,Len(vDept)-1)
         vOUT4 = "<a href=""/" & vDept & "/"">" & vTMP1 & "</a> &gt; " & vManufacturer

         ' Now create a list of products
         ' first we need the webtypeid or subcatid
         vSQL = gettypesql(replace(vDept,"_"," "), vDept, vManufacturer)

      ' regular section name here
      else
         ' clean the section name for use in url
         vTMP1 = UCase(Left(vSection,1)) & Right(vSection,Len(vSection)-1)

         if vSection <> "closeouts" and vSection <> "newitems" then
            ' set the breadcrumb link
            if instr(vNavTypes, vSection) Then
               vUDept = getwebtypedisp(vDept)
            else
               vUDept = getsubcatdisp(vDept)
            end if
            vOUT4 = "<a href=""/" & vSection & "/"">" & vTMP1 & "</a> &gt; " & vUDept
         else
            vOUT4 =  vTMP1
         end if

         ' Now create a list of products
         ' first we need the webtypeid or subcatid
         vSQL = gettypesql(replace(vDept,"_"," "), vSection, vManufacturer)
      end if

      ' We need to build our dropdown list for filtering
      ' the list by cost and mfg
      vFilterList = ""

'      response.write "<hr>" & vsqlmfg
'      rs1.open vSQLMFG, conn, 3
'      do while not rs1.eof
'         vOUT5 = vOUT5 & "<option value=""" & rs1("VendID") & """>"
'         vOUT5 = vOUT5 & rs1("vendor") & "</option>"
'         rs1.movenext
'      loop
'      rs1.close

      ' response.write "<hr>--- " & vsql
      rs1.open vSQL, conn, 3

      do while not rs1.eof

         ' keep a running count of displayed vendors
         if vSection <> "closeouts" and vSection <> "newitems" and vSection <> "allmfg" Then
            vKey = rs1("vendid")
            vMFG.Item(vKey) = cInt(vMFG.Item(vKey)) + 1
            vMFGName.Item(vKey) = rs1("vendor")
            vMFGID.Item(vKey) = rs1("vendid")

            'response.write "<hr>" & vKey & "/" & vMFG.Item(vKey) & "/" & vMFGName.Item(vKey)& "/" & vMFGID.Item(vKey)  & "/" & vMFG.Count
         end if

         vTMP1 = rs1("MSRP")
         If IsNumeric(vTMP1) Then
            vSP = formatcurrency(vTMP1,2,0,0) & ""
         Else
            vSP = "&nbsp;"
         End If
         vP = rs1("price")

         ' for filtering on price
         if vP <100 Then
            vPriceCount.Item("100") = cInt(vPriceCount.Item("100")) + 1
         elseif vP >100 and vP < 500 Then
            vPriceCount.Item("500") = cInt(vPriceCount.Item("500")) + 1
         elseif vP > 499  and vP < 1000 Then
            vPriceCount.Item("1000") = cInt(vPriceCount.Item("1000")) + 1
         elseif vP > 999  and vP < 2000 Then
            vPriceCount.Item("2000") = cInt(vPriceCount.Item("2000")) + 1
         elseif vP > 1999  and vP < 3000 Then
            vPriceCount.Item("3000") = cInt(vPriceCount.Item("3000")) + 1
         else
            vPriceCount.Item("more") = cInt(vPriceCount.Item("more")) + 1
         End if

         ' build the preface to the item detail URL
         if vSection <> "closeouts" and vSection <> "newitems" and vSection <> "allmfg" Then
            vPrefURL = "/" & vSection & "/" & vDept & "/"
         elseif vSection = "allmfg" Then
            vPrefURL = "/item/"
         else
            vPrefURL = "/item/"
         end if

         vOUT1 = vOUT1 & "   <TR>" _
               & "      <TD class=""productlist"" style=""background: transparent;""><a href=""" & vPrefURL & rs1("sku") & """><img src=""/ProductImages/" & rs1("picture") & """ border=""0"" width=""80"" alt=""" & rs1("description") & """></a></TD>" _
               & "      <TD class=""productlistfoot"" style=""background: transparent;""><a href=""" & vPrefURL & rs1("sku") & """>" & rs1("description") & "</a><BR>" _
               & "      " & vSP & "<BR>"_
               & "      <span class=""price"">YOU PAY: " & formatcurrency(vP,2,0,0) & "</span><BR>" _
               & "      <a href=""" & vPrefURL & rs1("sku") &""">MORE INFO</a></TD>" _
               & "   </TR>"
         rs1.movenext
      loop
      rs1.close

      dim vKeys
      vKeys = vPriceCount.Keys
      for x = 0 to vPriceCount.Count - 1
         ' response.write "<hr>" & vPriceCount.Item(vKeys(x)) & "/" & vKeys(x)
         vOUT5 = vOUT5 & "<option value=""" & vPriceCount.Item(vKeys(x)) & """>"
         if vKeys(x) = "100" Then vOUT5 = vOUT5 & " Less than $100"
         if vKeys(x) = "500" Then vOUT5 = vOUT5 & " $100 - $500"
         if vKeys(x) = "1000" Then vOUT5 = vOUT5 & " $500 - $1000"
         if vKeys(x) = "2000" Then vOUT5 = vOUT5 & " $1000 - $2000"
         if vKeys(x) = "3000" Then vOUT5 = vOUT5 & " $2000 - $3000"
         if vKeys(x) = "more" Then vOUT5 = vOUT5 & " > $3000"
         vOUT5 = vOUT5 & " (" & vPriceCount.Item(vKeys(x)) & ")</option>"
      next

      ' call PrintSortedDictionary(vMFGName)

      Call BuildArray(vMFGName, vKeys)
      Call SortArray(vKeys)
      for x = 0 to UBound(vKeys)
         ' response.write "<br>Key=" & vKeys(x)
         vOUT6 = vOUT6 & "<option value=""" & vMFGID.Item(vKeys(x)) & """>"
         vOUT6 = vOUT6 &  vMFGName.Item(vKeys(x)) & " (" & vMFG.Item(vKeys(x)) & ")</option>"
      next

'      vKeys = vMFG.Keys
'      for x = 0 to vMFG.Count - 1
'         ' response.write vMFG.Item(vKeys(x))
'         vOUT6 = vOUT6 & "<option value=""" & vMFGID.Item(vKeys(x)) & """>"
'         vOUT6 = vOUT6 &  vMFGName.Item(vKeys(x)) & " (" & vMFG.Item(vKeys(x)) & ")</option>"
'      next

 End Sub

   ' subroutine creats a standard product listing sql then
   ' runs it through the pagination system
   sub getprodlinks3 (vDept, vSection, vManufacturer)
      dim rs1, vSP, vP, vPrefURL
      Set rs1 = Server.CreateObject("ADODB.Recordset")

      ' response.write "<hr>gpl2|vDept=" & vDept & "|/|vSection=" & vSection & "|/|vManufacturer=" & vManufacturer & "|<br>"

      ' if section is allmfg then we expect a mfg to be sent
      ' -- product display is dependent on vendid and dept
      if vSection = "allmfg" Then

         ' clean the dept name for use in url
         vTMP1 = UCase(Left(vDept,1)) & Right(vDept,Len(vDept)-1)
         vOUT4 = "<a href=""/" & vDept & "/"">" & vTMP1 & "</a> &gt; " & vManufacturer

         ' Now create a list of products
         ' first we need the webtypeid or subcatid
         vSQL = gettypesql2(replace(vDept,"_"," "), vDept, vManufacturer)

     ' regular section name here - also handles 'all' dept listing for a section
      else
         ' clean the section name for use in url
         vTMP1 = UCase(Left(vSection,1)) & Right(vSection,Len(vSection)-1)

         if vSection <> "closeouts" and vSection <> "newitems" then
            ' set the breadcrumb link
            if instr(vNavTypes, vSection) Then
               vUDept = getwebtypedisp(vDept)
            else
               vUDept = getsubcatdisp(vDept)
            end if
            if vUDept = -1 then vUDept = "All Subcategories"
            vOUT4 = "<a href=""/" & vSection & "/"">" & vTMP1 & "</a> &gt; " & vUDept
         else
            vOUT4 =  vTMP1
         end if

         ' Now create a list of products
         ' first we need the webtypeid or subcatid
         vSQL = gettypesql2(replace(vDept,"_"," "), vSection, vManufacturer)
      end if

      ' We need to build our dropdown list for filtering
      ' the list by cost and mfg
      vFilterList = ""

      getprodlist(vSQL)

   end sub


   Sub getprodlinks2 (vDept, vSection, vManufacturer)
      dim rs100, rs110, vSP, vP, vPrefURL
      Set rs100 = Server.CreateObject("ADODB.Recordset")
	  Set rs110 = Server.CreateObject("ADODB.Recordset")

      ' response.write "<hr>gpl2|vDept=" & vDept & "|/|vSection=" & vSection & "|/|vManufacturer=" & vManufacturer & "|<br>"

      ' if section is allmfg then we expect a mfg to be sent
      ' -- product display is dependent on vendid and dept
      if vSection = "allmfg" Then

         ' clean the dept name for use in url
		 if (vDept <> "") then
			 vTMP1 = UCase(Left(vDept,1)) & Right(vDept,Len(vDept)-1)
			 vOUT4 = "<a href=""/" & vDept & "/"">" & vTMP1 & "</a> &gt; " & vManufacturer
		 end if

         ' Now create a list of products
         ' first we need the webtypeid or subcatid
         vSQL = gettypesql2(replace(vDept,"_"," "), vDept, vManufacturer)




      elseif vSection = "newitems" OR vSection = "closeouts" Then

         ' clean the dept name for use in url
         vTMP1 = vSection
         vOUT4 = "<a href=""/" & vDept & "/"">" & vTMP1 & "</a> &gt; " & vManufacturer

         ' Now create a list of products
         ' first we need the webtypeid or subcatid
         vSQL = gettypesql2(vSection, vDept, vManufacturer)


     ' regular section name here - also handles 'all' dept listing for a section
      else
         ' clean the section name for use in url
         vTMP1 = UCase(Left(vSection,1)) & Right(vSection,Len(vSection)-1)

         if vSection <> "closeouts" and vSection <> "newitems" then
            ' set the breadcrumb link
            if instr(vNavTypes, vSection) Then
               vUDept = getwebtypedisp(vDept)
            else
               vUDept = getsubcatdisp(vDept)
            end if
            if vUDept = -1 then vUDept = "All Subcategories"
            vOUT4 = "<a href=""/" & vSection & "/"">" & vTMP1 & "</a> &gt; " & vUDept
         else
            vOUT4 =  vTMP1
         end if

         ' Now create a list of products
         ' first we need the webtypeid or subcatid
         vSQL = gettypesql2(replace(vDept,"_"," "), vSection, vManufacturer)
      end if


      ' We need to build our dropdown list for filtering
      ' the list by cost and mfg
      vFilterList = ""

'      response.write "<hr>" & vsqlmfg
'      rs100.open vSQLMFG, conn, 3
'      do while not rs100.eof
'         vOUT5 = vOUT5 & "<option value=""" & rs100("VendID") & """>"
'         vOUT5 = vOUT5 & rs100("vendor") & "</option>"
'         rs100.movenext
'      loop
'      rs100.close

      ' response.write "<hr>--- " & vSQL
      rs100.open vSQL, conn, 3
	  counter = 0
      do while not rs100.eof
 		
	tempProd.ClearItem
	tempProd.GetItemPID(rs100("ProdID"))

         ' keep a running count of displayed vendors

            vKey = rs100("vendid")
            vMFG.Item(vKey) = cInt(vMFG.Item(vKey)) + 1
            vMFGName.Item(vKey) = rs100("vendor")
            vMFGID.Item(vKey) = rs100("vendid")
		
			'showonlybrands = showonlybrands & "<option value=""" & rs100("vendor") & """>" & rs100("vendor") & "</option>"
            'response.write "<hr>" & vKey & "/" & vMFG.Item(vKey) & "/" & vMFGName.Item(vKey)& "/" & vMFGID.Item(vKey)  & "/" & vMFG.Count


         vTMP1 = rs100("MSRP")
         If IsNumeric(vTMP1) AND IsNumeric(IsNumeric(vTMP1)) Then
         	if (vTMP1 / rs100("price")) > 1.05 Then
            	vSP = "MSRP: " & formatcurrency(vTMP1,2,0,0) & ""
            else
            	vSP = "&nbsp;"
            end if
         Else
            vSP = "&nbsp;"
         End If
         vP = rs100("price")

         ' for filtering on price
         if vP <100 Then
            vPriceCount.Item("100") = cInt(vPriceCount.Item("100")) + 1
         elseif vP >100 and vP < 500 Then
            vPriceCount.Item("500") = cInt(vPriceCount.Item("500")) + 1
         elseif vP > 499  and vP < 1000 Then
            vPriceCount.Item("1000") = cInt(vPriceCount.Item("1000")) + 1
         elseif vP > 999  and vP < 2000 Then
            vPriceCount.Item("2000") = cInt(vPriceCount.Item("2000")) + 1
         elseif vP > 1999  and vP < 3000 Then
            vPriceCount.Item("3000") = cInt(vPriceCount.Item("3000")) + 1
         else
            vPriceCount.Item("more") = cInt(vPriceCount.Item("more")) + 1
         End if

         ' build the preface to the item detail URL
         if vSection <> "closeouts" and vSection <> "newitems" and vSection <> "allmfg" Then
            vPrefURL = "/" & vSection & "/" & vDept & "/"
         elseif vSection = "allmfg" Then
            vPrefURL = "/item/"
         else
            vPrefURL = "/item/"
         end if

		vWebNote = ""
		if rs100("webnote") <> 1 then
				vWebNote = "<div class=""product_notes"">" & vWebNoteSD(cstr(rs100("webnote"))) & "</div>"
		end if

		if rs100("FreeFreight") = True then
			vFreeFreight = -1
		   vWebNote = vWebNote & "<div class=""product_freefreight"">(Free Shipping with " & vFreeShippingMethod & "!)</div>"
		   Else
			vFreeFreight = 0
		End If
		if rs100("OverWeight") > 0 then
			vOverWeight = rs100("OverWeight") + 1
		'   vWebNote = vWebNote & "<div class=""product_freefreight"">(Overweight shipping costs apply!)</div>"
		else
			vOverWeight = 0
		End If

		vCP = int(rs100("ischildorparentoritem"))
		if isnull(vCP) or vCP = "" then vCP = 0
		if vCP Then
		   vItemOptions = ShowOptions2(rs100("ProdID"),  rs100("description"),  rs100("SKU"),  rs100("price")) & "<BR>"
		   ITEMID_1 = "NOTINUSE"
		else
		   vItemOptions = ""
		   ITEMID_1 = "ITEMID"
		end if


		if (vWebNote <> "") then
			vWebNote = vWebNote & "<br />"
		end if
		if ((counter >= (numberperpage * pagenumber) - (numberperpage)) AND (counter < (numberperpage * pagenumber))) then
         vOUT1 = vOUT1 & "   <TR>" _
		       & "" _
			  & "<FORM METHOD=""post"" action=""/addtocart/"">" _
			  & "<INPUT TYPE=""hidden"" name=""ITEMNAME"" value=""" & replace("""", "", rs100("description")) & """>" _
			  & "<INPUT TYPE=""hidden"" name=""PRICE"" value=""" & rs100("price") & """>" _
			  & "<INPUT TYPE=""hidden"" name=""Referer"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""Referer1"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""URL"" value=""" & "/item/" & rs100("sku") & """>" _
			  & "<INPUT TYPE=""hidden"" name=""Parent"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""PID"" value=""" & rs100("ProdID") & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""FreeFreight"" VALUE=""" & vFreeFreight & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""OverWeightFlags"" VALUE=""" & vOverWeight & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""" & ITEMID_1 & """ VALUE=""" & rs100("sku") & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""mDiscountType"" VALUE=""" & tempProd.pfields.Item("mDiscountType") & """>" _
	          & "<INPUT TYPE=""hidden"" NAME=""mDiscountAmount"" VALUE=""" & tempProd.pfields.Item("mDiscountAmount") & """>" _
              & "<INPUT TYPE=""hidden"" NAME=""mSpecialPricing"" VALUE=""" & tempProd.pfields.Item("mSpecialPricing") & """>" _
               & "      <TD class=""productlist"" style=""background: transparent;"" valign=top><a href=""" & vPrefURL & rs100("sku") & """><img src=""" & resizepic("/productimages/" & rs100("picture"), rs100("Width_Small"), rs100("Height_Small")) & """ border=""0"" alt=""" & replace("""", "", rs100("description")) & """></a></TD>" _
               & "      <TD class=""productlistfoot"" style=""background: transparent;""><a href=""" & vPrefURL & rs100("sku") & """>" & rs100("description") & "</a><BR>" _
               & "      " & vSP & "<BR>"

              if (rs100("WebNote") <> 15) then
              	vOUT1 = vOUT1 & "      <span class=""price"">YOU PAY: " &  FormatCurrencyDiscount("<BR>On Special",vP,tempProd.pfields.item("mDiscountAmount") ) & "</span><BR>"
              else
                 		vOUT1 = vOUT1  & "      <span class=""price"">YOU PAY: " & formatcurrency(vP,2,0,0) & "</span><BR>"
						vOUT1 = vOUT1 & "<a href=""javascript:void(0)"" onClick=""window.open('/rebate_price.asp?SKU=" & rs100("SKU") & "', 'BikePopUp',  'width=520,height=400,toolbar=0,scrollbars=1,screenX=50,screenY=50,left=50,top=50')""><font color=blue>Click here to View the Instant Rebate<br>Price you will see in the Checkout</font></a><br>"
              end if

              vOUT1 = vOUT1 & "      <a href=""" & vPrefURL & rs100("sku") &""">MORE INFO</a>" & vWebNote & "<div align=right>" _
			  & vItemOptions & "<input name=""SUBMIT"" VALUE=""ADD"" type=image src=""images/addtocart.jpg"" alt=""View Cart"" width=""100"" height=""22"" border=0 style=""margin: 5px 0 0 0;""></div></TD>" _
               & " </FORM></TR>"
		end if
		counter = counter + 1
         rs100.movenext
      loop
      rs100.close


	  pagenavout = ""
	  loccounter = 1
	  do while (counter > 0)
	  	if (pagenumber = loccounter) then
			pagenavout = pagenavout & "<b>" & loccounter & "</b> "
		else
			pagenavout = pagenavout & "<a href=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & numberperpage & "&pagenumber=" & loccounter & """>" & loccounter & "</a> "
		end if
		counter = counter - numberperpage
	  	loccounter = loccounter + 1
	  loop
	  if (pagenavout <> "") then
	  	pagenavout = "<tr><TD colspan=""3"" class=""pages"" align=""center"">" & pagenavout
		pagenavout = pagenavout & "</TD></tr>"
	  end if

	  pagenavout = replace(pagenavout, ":", "\:")

      dim vKeys
      vKeys = vPriceCount.Keys
      for x = 0 to vPriceCount.Count - 1
         ' response.write "<hr>" & vPriceCount.Item(vKeys(x)) & "/" & vKeys(x)
         vOUT5 = vOUT5 & "<option value=""" & vPriceCount.Item(vKeys(x)) & """>"
         if vKeys(x) = "100" Then vOUT5 = vOUT5 & " Less than $100"
         if vKeys(x) = "500" Then vOUT5 = vOUT5 & " $100 - $500"
         if vKeys(x) = "1000" Then vOUT5 = vOUT5 & " $500 - $1000"
         if vKeys(x) = "2000" Then vOUT5 = vOUT5 & " $1000 - $2000"
         if vKeys(x) = "3000" Then vOUT5 = vOUT5 & " $2000 - $3000"
         if vKeys(x) = "more" Then vOUT5 = vOUT5 & " > $3000"
         vOUT5 = vOUT5 & " (" & vPriceCount.Item(vKeys(x)) & ")</option>"
      next

      ' call PrintSortedDictionary(vMFGName)

      Call BuildArray(vMFGName, vKeys)
      Call SortArray(vKeys)
      for x = 0 to UBound(vKeys)
         ' response.write "<br>Key=" & vKeys(x)
         vOUT6 = vOUT6 & "<option value=""" & vMFGID.Item(vKeys(x)) & """>"
         vOUT6 = vOUT6 &  vMFGName.Item(vKeys(x)) & " (" & vMFG.Item(vKeys(x)) & ")</option>"
      next

'      vKeys = vMFG.Keys
'      for x = 0 to vMFG.Count - 1
'         ' response.write vMFG.Item(vKeys(x))
'         vOUT6 = vOUT6 & "<option value=""" & vMFGID.Item(vKeys(x)) & """>"
'         vOUT6 = vOUT6 &  vMFGName.Item(vKeys(x)) & " (" & vMFG.Item(vKeys(x)) & ")</option>"
'      next


'response.write vSQLDrop
		if (vSQLDrop <> "") then

   Set vMFG = CreateObject("Scripting.Dictionary")
   Set vMFGName = CreateObject("Scripting.Dictionary")
   Set vMFGID = CreateObject("Scripting.Dictionary")
   Set vPriceCount = CreateObject("Scripting.Dictionary")
			  rs100.open vSQLDrop, conn, 3
			  counter = 0
			  do while not rs100.eof


				 ' keep a running count of displayed vendors

					vKey = rs100("vendid")
					vMFG.Item(vKey) = cInt(vMFG.Item(vKey)) + 1
					vMFGName.Item(vKey) = rs100("vendor")
					vMFGID.Item(vKey) = rs100("vendid")

					'showonlybrands = showonlybrands & "<option value=""" & rs100("vendor") & """>" & rs100("vendor") & "</option>"
					'response.write "<hr>" & vKey & "/" & vMFG.Item(vKey) & "/" & vMFGName.Item(vKey)& "/" & vMFGID.Item(vKey)  & "/" & vMFG.Count


				 vTMP1 = rs100("MSRP")
				 If IsNumeric(vTMP1) Then
					vSP = "MSRP: " & formatcurrency(vTMP1,2,0,0) & ""
				 Else
					vSP = "&nbsp;"
				 End If
				 vP = rs100("price")

				 ' for filtering on price
				 if vP <100 Then
					vPriceCount.Item("100") = cInt(vPriceCount.Item("100")) + 1
				 elseif vP >100 and vP < 500 Then
					vPriceCount.Item("500") = cInt(vPriceCount.Item("500")) + 1
				 elseif vP > 499  and vP < 1000 Then
					vPriceCount.Item("1000") = cInt(vPriceCount.Item("1000")) + 1
				 elseif vP > 999  and vP < 2000 Then
					vPriceCount.Item("2000") = cInt(vPriceCount.Item("2000")) + 1
				 elseif vP > 1999  and vP < 3000 Then
					vPriceCount.Item("3000") = cInt(vPriceCount.Item("3000")) + 1
				 else
					vPriceCount.Item("more") = cInt(vPriceCount.Item("more")) + 1
				 End if


				counter = counter + 1
				 rs100.movenext
			  loop
			  rs100.close
	end if





 End Sub






' this should generate a generic product listing based on passed sql
' written for search... will try to implement elsewhere


'   Sub getprodlist (vSQL, vSection)
   Sub getprodlist (vSQL)
  
      dim rs100, vP, vSP, vPrefURL, vLoop
      Set rs100 = Server.CreateObject("ADODB.Recordset")

      ' response.write "<hr>--- " & vSQL
      ' response.write "<hr>--- " & vListanum
      rs100.open vSQL, conn, 3

      If rs100.EOF then

   		'response.write "<FONT ID=""body"">No items found!</FONT><BR>"
   		vOUT6 = "No items found!"
   		vOUT2 = ""
   		rs100.Close

      else

         ' let's do some pagination
   		rs100.PageSize = vListanum

         ' if vMv equals something then we're moving within a result set
   		If vMv = vPrevious or vMv = vNext or vMv=vFirst or vMv= vLast Then
   			Select Case vMv

   				Case vFirst
   					vPageNo = 1

   				Case vLast
   					vPageNo = rs100.PageCount

   				Case vPrevious
   					If vPageNo > 1 Then
   						vPageNo = vPageNo - 1
   					Else
   						vPageNo = 1
   					End If

   				Case vNext
   					If rs100.AbsolutePage < rs100.PageCount Then
   						vPageNo = vPageNo + 1
   					Else
   						vPageNo = rs100.PageCount
   					End If

               ' if moving within result set then we start at beginning
   				Case Else
   					vPageNo = 1
   			End Select
   		End If
   		rs100.AbsolutePage = vPageNo
   		' response.write "<hr>" & vpageno

         ' begin the product page here, only vPageSize items per page
			For xx = 1 to rs100.PageSize
            ' done with this page if we've show vpagesize items
				If rs100.EOF Then
					Exit For
				End If
            
            ' define the msrp and price displays
            vTMP1 = rs100("MSRP")
            If IsNumeric(vTMP1) Then
               vSP = "MSRP: " & formatcurrency(vTMP1,2,0,0) & ""
            Else
               vSP = "&nbsp;"
            End If
            vP = rs100("price")

            ' keep a running count of displayed vendors for product displays only
            ' this count is for the mfg filter pulldown
            if vSection <> "closeouts" and vSection <> "newitems" and vSection <> "allmfg" Then
               vKey = rs100("vendid")
               vMFG.Item(vKey) = cInt(vMFG.Item(vKey)) + 1
               vMFGName.Item(vKey) = rs100("vendor")
               vMFGID.Item(vKey) = rs100("vendid")
               'response.write "<hr>" & vKey & "/" & vMFG.Item(vKey) & "/" & vMFGName.Item(vKey)& "/" & vMFGID.Item(vKey)  & "/" & vMFG.Count
            end if

            ' keep running total of price ranges for filtering on price
            if vP <100 Then
               vPriceCount.Item("100") = cInt(vPriceCount.Item("100")) + 1
            elseif vP >100 and vP < 500 Then
               vPriceCount.Item("500") = cInt(vPriceCount.Item("500")) + 1
            elseif vP > 499  and vP < 1000 Then
               vPriceCount.Item("1000") = cInt(vPriceCount.Item("1000")) + 1
            elseif vP > 999  and vP < 2000 Then
               vPriceCount.Item("2000") = cInt(vPriceCount.Item("2000")) + 1
            elseif vP > 1999  and vP < 3000 Then
               vPriceCount.Item("3000") = cInt(vPriceCount.Item("3000")) + 1
            else
               vPriceCount.Item("more") = cInt(vPriceCount.Item("more")) + 1
            End if

            ' build the preface to the item detail URL
            if vSection <> "closeouts" and vSection <> "newitems" and vSection <> "allmfg" Then
               vPrefURL = "/" & vSection & "/" & vDept & "/"
            elseif vSection = "allmfg" Then
               vPrefURL = "/item/"
            else
               vPrefURL = "/item/"
            end if

            vPrefURL = "/item/"  

		vWebNote = ""
		if rs100("webnote") <> 1 then
				vWebNote = "<div class=""product_notes"">" & vWebNoteSD(cstr(rs100("webnote"))) & "</div>"
		end if

		if rs100("FreeFreight") = True then
			vFreeFreight = -1
		   vWebNote = vWebNote & "<div class=""product_freefreight"">(Free Shipping with " & vFreeShippingMethod & "!)</div>"
		   Else
			vFreeFreight = 0
		End If
		if rs100("OverWeight") > 0 then
			vOverWeight = rs100("OverWeight") + 1
		'   vWebNote = vWebNote & "<div class=""product_freefreight"">(Overweight shipping costs apply!)</div>"
		else
			vOverWeight = 0
		End If

			if (vWebNote <> "") then
				vWebNote = vWebNote & "<br />"
			end if

			vCP = int(rs100("IsChildorParentorItem"))
			if isnull(vCP) or vCP = "" then vCP = 0
			if vCP Then
			   vItemOptions = ShowOptions2(rs100("ProdID"),  rs100("description") ,  rs100("SKU"),  rs100("price")) & "<BR>"
			   ITEMID_1 = "NOTINUSE"
			else
			   vItemOptions = ""
			   ITEMID_1 = "ITEMID"
			end if

            tempProd.clearitem
            tempProd.getitemPID(rs100("ProdID") )

            ' vOUT1 is the final output var
            vOUT1 = vOUT1 & "   <TR>" _
                  & "      <TD class=""productlist"" style=""background: transparent;""><a href=""" & vPrefURL & rs100("sku") & """><img src=""" & resizepic("/productimages/" & rs100("picture"), rs100("Width_Small"), rs100("Height_Small")) & """ border=""0""  alt=""""></a></TD>" _
                  & "      <TD class=""productlistfoot"" style=""background: transparent;""><a href=""" & vPrefURL & rs100("sku") & """>" & rs100("description") & "</a><BR>" _
                  & "      " & vSP & "<BR>"

                  if (rs100("WebNote") <> 15) then
                 		vOUT1 = vOUT1  & "      <span class=""price"">YOU PAY: " & formatcurrencyDiscount("<BR>On Special",vP,tempProd.pfields.item("mDiscountAmount")) & "</span><BR>"
                  else
                 		vOUT1 = vOUT1  & "      <span class=""price"">YOU PAY: " & formatcurrency(vP,2,0,0) & "</span><BR>"
						vOUT1 = vOUT1 & "<a href=""javascript:void(0)"" onClick=""window.open('/rebate_price.asp?SKU=" & rs100("SKU") & "', 'BikePopUp',  'width=520,height=400,toolbar=0,scrollbars=1,screenX=50,screenY=50,left=50,top=50')""><font color=blue>Click here to View the Instant Rebate<br>Price you will see in the Checkout</font></a><br>"
                  end if
                 vOUT1 = vOUT1 & "      <a href=""" & vPrefURL & rs100("sku") &""">MORE INFO</a>" & vWebNote _
				  & "<FORM METHOD=""post"" action=""/addtocart/"">" _
				  & "<INPUT TYPE=""hidden"" name=""ITEMNAME"" value=""" & replace("""", "", rs100("description")) & """>" _
				  & "<INPUT TYPE=""hidden"" name=""PRICE"" value=""" & rs100("price") & """>" _
				  & "<INPUT TYPE=""hidden"" name=""Referer"" value="""">" _
				  & "<INPUT TYPE=""hidden"" name=""Referer1"" value="""">" _
				  & "<INPUT TYPE=""hidden"" name=""URL"" value=""" & "/item/" & rs100("SKU") & """>" _
				  & "<INPUT TYPE=""hidden"" name=""Parent"" value="""">" _
				  & "<INPUT TYPE=""hidden"" name=""PID"" value=""" & rs100("ProdID") & """>" _
				  & "<INPUT TYPE=""hidden"" NAME=""FreeFreight"" VALUE=""" & vFreeFreight & """>" _
				  & "<INPUT TYPE=""hidden"" NAME=""OverWeightFlags"" VALUE=""" & vOverWeight & """>" _
				  & "<INPUT TYPE=""hidden"" NAME=""" & ITEMID_1 & """ VALUE=""" & rs100("SKU") & """>" _
			      & "<INPUT TYPE=""hidden"" NAME=""mDiscountType"" VALUE=""" & tempProd.pfields.Item("mDiscountType") & """>" _
	              & "<INPUT TYPE=""hidden"" NAME=""mDiscountAmount"" VALUE=""" & tempProd.pfields.Item("mDiscountAmount") & """>" _
                  & "<INPUT TYPE=""hidden"" NAME=""mSpecialPricing"" VALUE=""" & tempProd.pfields.Item("mSpecialPricing") & """>" _
				  & "<right>" & vItemOptions & "<input name=""SUBMIT"" VALUE=""ADD"" type=image src=""images/addtocart.jpg"" alt=""View Cart"" width=""100"" height=""22"" border=0 style=""margin: 5px 0 0 0;""></div></right>" _
				  & "</FORM>"	_
				  & "</TD>" _
                  & "   </TR>"

            ' protection from runaways
            vLoop = vLoop + 1
            if vLoop > 50 then
               response.write "Runaway detected..."
               response.end
            end if

				rs100.movenext
				If rs100.EOF Then
					Exit For
				End If
			Next
         	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
         	'% Build the navigation bar
         	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

            Dim vHRefH, vHRefT, vImgH, vImgT
            Dim vFP, vPP, vNP, vLP
            Dim vNavID, vNav1, vNav2

         	vHRefH = "<a href=""/Items01.asp?NavID=" & vNavID
         	vHRefH = vHRefH & "&M=" & vManufacturer
         	vHRefH = vHRefH & "&T=" & vSearchCategory
         	vHRefH = vHRefH & "&P=" & vPageNo
         	vHRefH = vHRefH & "&D="""

         	vHRefT = "</a>"
         	vImgH = "<img src=""/images/"
         	vImgT = "page.gif"" border=""0"" align=""absmiddle"">"

         	vFP = vImgH & "first" & vImgT
         	vPP = vImgH & "previous" & vImgT
         	vNP = vImgH & "next" & vImgT
         	vLP = vImgH & "last" & vImgT

         	if vPageNo > 1 then
               vFP = vHRefH & "FP"">" & vFP & vHRefT
               vPP = vHRefH & "PP"">" & vPP & vHRefT
         	End If

         	if vPageNo < rs100.PageCount then
         		vNP = vHRefH & "NP"">" & vNP & vHRefT
         		vLP = vHRefH & "LP"">" & vLP & vHRefT
         	End If
         	vNav1 = vFP & vPP
         	vNav2 = vNP & vLP

            ' more than 1 page needs a page nav
         	if rs100.PageCount > 1 then

         		'Response.write "<CENTER>" & vNav1 & " " & vNav2 & "<FONT id=""fineprint""><BR>Pages: "
         		vOUT2 = "<TD colspan=""3"" class=""pages"" align=""center"">" & chr(13)

       		   ' response.write "<hr>" & rs100.pagecount
         		For x = 1 to rs100.PageCount
         		   ' response.write "<hr>" & x & "<br>" & vPageNo & "/" & (x = (vPageNo -0))

                  ' show this page we're on, but do not link to it
         			If x = vPageNo + 0 then
         				'response.write "<b>" & x & " </b>"
         				vOUT2 = vOUT2 & "<b>" & x  & "&nbsp;</b>"
         				' response.write "THIS PAGE!!<HR>"

                  ' links for each number, first & last pages
         			Else
         				'response.write "<a href=""/Items01.asp?NavID=""" & vNavID & "&M=" & vManufacturer & "&T=" & vSearchCategory & "&"
         				if (vSection = "search") then
							vOUT2 = vOUT2 & "<a href=""/search"
						else
         					vOUT2 = vOUT2 & "<a href=""/manufacturer/" & vManufacturer & "/" & vDept
						end if

         				' if a section name then put that up
         				vOUT2 = vOUT2 & "/s/" & vSearchCategory

         				' if there is a mfg then put it up
         				vOUT2 = vOUT2 & "/v/" & vSearchVendID

                     ' if we're making the link to "first" - page 1
         				if x = 1 then
         					'response.write "D=" & vFirst & """>"
         					vOUT2 = vOUT2 & "/p/1/DIR/" & vFirst & """>"

                     ' if we're on the last page
         				elseif x = rs100.PageCount then
         					'response.write "D=" & vLast & """>"
         					vOUT2 = vOUT2 & "/p/" & rs100.PageCount & "/DIR/" & vLast & """>"

                     ' ok we're on an actual page number but not this one
         				else
         					'response.write "P=" & x-1 & "&D=" & vNext & """>"
         					vOUT2 = vOUT2 & "/p/" & x & "/DIR/" & vNext & """>"
         				end if
         			'response.write x & " </a>"
         			vOUT2 = vOUT2 & x & "</a>&nbsp;"
         			End if
         		next
         		'response.write "</FONT></CENTER><BR><BR>"
               vOUT2 = vOUT2 & "</TD></TR>"
         	End If
         rs100.close
      End If

		if (vLoop = 0) then
				vOUT1 = vOUT1 & "   <TR>" _
					  & "      <TD class=""productlist"" style=""background: transparent;"" colspan=2>No search results.</TD>" _
					  & "   </TR>"
		end if

      ' Save the search criteria in a session var
		Session("searchterm") = vSearchTerm
		Session("searchcategory") = vSearchCategory

      'response.write "<hr><h1>SC: "  & vsearchcategory & "</h1>"

		' Since we just displayed items we should clear the session variables
		Session("M") = 0
		Session("T") = 0
		Session("NavID") = ""

      ' this is for price filter pulldown
      dim vKeys
      vKeys = vPriceCount.Keys
      for x = 0 to vPriceCount.Count - 1
         ' response.write "<hr>" & vPriceCount.Item(vKeys(x)) & "/" & vKeys(x)
         vOUT5 = vOUT5 & "<option value=""" & vPriceCount.Item(vKeys(x)) & """>"
         if vKeys(x) = "100" Then vOUT5 = vOUT5 & " Less than $100"
         if vKeys(x) = "500" Then vOUT5 = vOUT5 & " $100 - $500"
         if vKeys(x) = "1000" Then vOUT5 = vOUT5 & " $500 - $1000"
         if vKeys(x) = "2000" Then vOUT5 = vOUT5 & " $1000 - $2000"
         if vKeys(x) = "3000" Then vOUT5 = vOUT5 & " $2000 - $3000"
         if vKeys(x) = "more" Then vOUT5 = vOUT5 & " > $3000"
         vOUT5 = vOUT5 & " (" & vPriceCount.Item(vKeys(x)) & ")</option>"
      next

      ' call PrintSortedDictionary(vMFGName)

      ' this is for mfg filter pulldown
      BuildArray vMFGName, vKeys
      SortArray vKeys
      for x = 0 to UBound(vKeys)
         ' response.write "<br>Key=" & vKeys(x)
         vOUT6 = vOUT6 & "<option value=""" & vMFGID.Item(vKeys(x)) & """>"
         vOUT6 = vOUT6 &  vMFGName.Item(vKeys(x)) & " (" & vMFG.Item(vKeys(x)) & ")</option>"
      next

'      vKeys = vMFG.Keys
'      for x = 0 to vMFG.Count - 1
'         ' response.write vMFG.Item(vKeys(x))
'         vOUT6 = vOUT6 & "<option value=""" & vMFGID.Item(vKeys(x)) & """>"
'         vOUT6 = vOUT6 &  vMFGName.Item(vKeys(x)) & " (" & vMFG.Item(vKeys(x)) & ")</option>"
'      next

 End Sub



Public  Sub getproductdetail (vSection, vDept, vItem )

   oProd1.getitemSKU(vItem)

End Sub

	function ShowOptions(ProdID, Desc, SKU, Price)
	 'price = cdbl(Price)
	 'response.Write(price)
	 'exit function
      Dim vPDesc, rschild, vValue, vValue1, vODesc, vUseDesc, vDiff, vPosNeg

      ' clean slate
      vOUT1 = ""

		' Need to get rid of double spacing or the cart options wont work right
		vPDesc = replace(Desc, "  ", " ")
		vPDesc = replace(Desc, "  ", " ")
		vPDesc = replace(Desc, "  ", " ")
		vPDesc = replace(Desc, "  ", " ")

		Set rschild = Server.CreateObject("ADODB.Recordset")
		vSQL = "SELECT top 100 [Products Children].*,Products.*,Size.* " _
		     & "FROM  [Products Children] " _
		     & "INNER JOIN (Products INNER JOIN Size ON Products.SizeID = Size.SizeID) " _
		     & "ON [Products Children].ChildProdID = Products.ProdID " _
		     & "WHERE [Products Children].ProdID=" & ProdID _
		     & " AND ((Products.Discontinued LIKE 'yes' AND Products.QtyAvailable > 0) or (Products.Discontinued LIKE 'no')) " _
		     & "ORDER BY Size.Sort, Products.Description"    & " For Browse"

'		response.write "<pre>" & vSQL & "</pre><br>"
		rschild.open vSQL, Conn, 3

		If rschild.EOF then
			vSQL = "SELECT top 100 [Products Children].*,Products.* " _
			     & "FROM [Products Children] " _
			     & "INNER JOIN Products ON [Products Children].ChildProdID = Products.ProdID " _
			     & "WHERE [Products Children].ProdID=" & ProdID & " AND ((Products.Discontinued LIKE 'yes' AND Products.QtyAvailable > 0) or (Products.Discontinued LIKE 'no'))" _
			     '& " For Browse"
			rschild.close
'			response.write "<pre>" & vsql & "</pre><br>"
			rschild.open vSQL & " For Browse",Conn,3
		End If

		vValue = "Prop=~&COMBO;"
		vValue1 = "PropID=~&COMBO;"

			If NOT rschild.EOF then
			vOUT1 = vbcrlf & "<SELECT NAME=""ITEMID"" SIZE=1 style=""font-size: 12px;"">" & chr(13)
				Do While not rschild.EOF
				   ' response.write "<hr>" & rschild("sku")
					' Need to strip out any quote characters
					vODesc = replace(rschild("description"), """", "''")
					' Need to get rid of double spacing or the cart options wont work right
					vODesc = replace(vODesc, "  ", " ")
					vODesc = replace(vODesc, "  ", " ")
					vODesc = replace(vODesc, "  ", " ")
					vODesc = replace(vODesc, "  ", " ")
					' four times should be enough... we hope

					vUseDesc = ""
					if rschild("SizeID") <> 0 and IsNULL(rschild("SizeID"))=False then
						vUseDesc = vSizeListingSD.Item(rschild("SizeID"))
					End If
					if rschild("ColorID") <> 0 and IsNULL(rschild("ColorID"))=False then
						if vColorListingSD.Item(rschild("ColorID")) <> "" then
							if vUseDesc <> "" then vUseDesc = vUseDesc & " - "
							vUseDesc = vUseDesc & vColorListingSD.Item(rschild("ColorID"))
						End if
					End If
					if vUseDesc = "" then
						vUseDesc = shortdesc(vPDesc, vODesc) & " "
					end if

					' If the parent price is different than the child price we
					' display the difference in the dropdown box
					if Price <> rschild("price")  then
					   vDiff = rschild("price") - Price
					   if vDiff <> 0 then
					   if vDiff > 0 Then vPosNeg = " [+" else vPosNeg = " ["
					   vUseDesc = vUseDesc & vPosNeg & formatcurrency(vDiff,2,0,0,0) & "]"
					   end if
					end if
'					vUseDesc = vUseDesc & " Q:" & rschild("QtyAvailable") & " D:" & rschild("Discontinued")
					vValue = vValue & vUseDesc & ","
					vValue1 = vValue1 & rschild("SKU") & ","
					vOUT1 = vOUT1 & chr(9) & chr(9) & "<option value=""" & rschild("SKU") & """>" & vUseDesc  & chr(13)
					rschild.movenext
				Loop
				vValue = left(vValue, Len(vValue) - 1)
				vValue1 = left(vValue1, Len(vValue1) - 1)
				vOUT1 = vOUT1 &  "	</SELECT>" & VBCrLf
				vOUT1 = vOUT1 & "	<input type=hidden name=""PropDATA"" value=""" & vValue & """>" & VBCrLf
				vOUT1 = vOUT1 & "	<input type=hidden name=""PropIDDATA"" value=""" & vValue1 & """>" & VBCrLf
			End If
		rschild.Close

      ShowOptions = vOUT1
	End Function



	function ShowOptions2(ProdID, Desc, SKU, Price)
      Dim vPDesc, rschild, vValue, vValue1, vODesc, vUseDesc, vDiff, vPosNeg

      ' clean slate
      vOUT105 = ""

		' Need to get rid of double spacing or the cart options wont work right
		vPDesc = replace(Desc, "  ", " ")
		vPDesc = replace(Desc, "  ", " ")
		vPDesc = replace(Desc, "  ", " ")
		vPDesc = replace(Desc, "  ", " ")

		Set rschild = Server.CreateObject("ADODB.Recordset")
		vSQL101 = "SELECT top 100 [Products Children].*,Products.*,Size.* " _
		     & "FROM  [Products Children] " _
		     & "INNER JOIN (Products INNER JOIN Size ON Products.SizeID = Size.SizeID) " _
		     & "ON [Products Children].ChildProdID = Products.ProdID " _
		     & "WHERE [Products Children].ProdID=" & ProdID _
		     & " AND ((Products.Discontinued LIKE 'yes' AND Products.QtyAvailable > 0) or (Products.Discontinued LIKE 'no')) " _
		     & "ORDER BY Size.Sort, Products.Description"    & " For Browse"

'		response.write "<pre>" & vSQL & "</pre><br>"
		rschild.open vSQL101, Conn, 3

		If rschild.EOF then
			vSQL101 = "SELECT top 100 [Products Children].*,Products.* " _
			     & "FROM [Products Children] " _
			     & "INNER JOIN Products ON [Products Children].ChildProdID = Products.ProdID " _
			     & "WHERE [Products Children].ProdID=" & ProdID & " AND ((Products.Discontinued LIKE 'yes' AND Products.QtyAvailable > 0) or (Products.Discontinued LIKE 'no')) " 			  
'   & " For Browse"
			rschild.close
'			response.write "<pre>" & vsql & "</pre><br>"
			rschild.open vSQL101 & " For Browse",Conn,3
		End If

		vValue = "Prop=~&COMBO;"
		vValue1 = "PropID=~&COMBO;"

			If NOT rschild.EOF then
			vOUT105 = vbcrlf & "	<SELECT NAME=""ITEMID"" style=""font-size: 9px;"">" & chr(13)
				Do While not rschild.EOF
				   ' response.write "<hr>" & rschild("sku")
					' Need to strip out any quote characters
					vODesc = replace(rschild("description"), """", "''")
					' Need to get rid of double spacing or the cart options wont work right
					vODesc = replace(vODesc, "  ", " ")
					vODesc = replace(vODesc, "  ", " ")
					vODesc = replace(vODesc, "  ", " ")
					vODesc = replace(vODesc, "  ", " ")
					' four times should be enough... we hope

					vUseDesc = ""
					if rschild("SizeID") <> 0 and IsNULL(rschild("SizeID"))=False then
						vUseDesc = vSizeListingSD.Item(rschild("SizeID"))
					End If
					if rschild("ColorID") <> 0 and IsNULL(rschild("ColorID"))=False then
						if vColorListingSD.Item(rschild("ColorID")) <> "" then
							if vUseDesc <> "" then vUseDesc = vUseDesc & " - "
							vUseDesc = vUseDesc & vColorListingSD.Item(rschild("ColorID"))
						End if
					End If
					if vUseDesc = "" then
						vUseDesc = shortdesc(vPDesc, vODesc) & " "
					end if

					' If the parent price is different than the child price we
					' display the difference in the dropdown box
					if Price <> rschild("price") then
					   vDiff = rschild("price") - Price
					   if vDiff <> 0 then 
					   if vDiff > 0 Then vPosNeg = " [+" else vPosNeg = " ["
					   vUseDesc = vUseDesc & vPosNeg & formatcurrency(vDiff,2,0,0,0) & "]"
					   end if
					end if
'					vUseDesc = vUseDesc & " Q:" & rschild("QtyAvailable") & " D:" & rschild("Discontinued")
					vValue = vValue & vUseDesc & ","
					vValue1 = vValue1 & rschild("SKU") & ","
					vOUT105 = vOUT105 & chr(9) & chr(9) & "<option value=""" & rschild("SKU") & """>" & vUseDesc & chr(13)
					rschild.movenext
				Loop
				vValue = left(vValue, Len(vValue) - 1)
				vValue1 = left(vValue1, Len(vValue1) - 1)
				vOUT105 = vOUT105 &  "	</SELECT>" & VBCrLf
				vOUT105 = vOUT105 & "	<input type=hidden name=""PropDATA"" value=""" & vValue & """>" & VBCrLf
				vOUT105 = vOUT105 & "	<input type=hidden name=""PropIDDATA"" value=""" & vValue1 & """>" & VBCrLf
			End If
		rschild.Close

      ShowOptions2 = vOUT105
	End Function




	Function shortdesc(PDesc, CDesc)
	   Dim FDesc, PDescArray, CDescArray, shortcounter, breakloop

		FDesc = ""

		PDescArray = Split(PDesc, " ")
		CDescArray = Split(CDesc, " ")

		shortcounter = 0
		breakloop = 0

		do while ((shortcounter <= ubound(PDescArray)) AND (shortcounter <= ubound(CDescArray)))
			if (PDescArray(shortcounter) <> CDescArray(shortcounter)) then
				FDesc = FDesc & CDescArray(shortcounter) & " "
				breakloop = 1
			elseif (breakloop = 1) then
				FDesc = FDesc & CDescArray(shortcounter) & " "
			end if
			shortcounter = shortcounter + 1
		loop


		do while (shortcounter <= ubound(CDescArray))
			FDesc = FDesc & CDescArray(shortcounter) & " "
			shortcounter = shortcounter + 1
		loop

'FDesc = ""
'		For x = 1 to Len(CDesc)
'			if mid(CDesc, x, 1) <> mid(PDesc, x, 1) then
'				FDesc = FDesc & mid(CDesc, x)
'				Exit For
'			end If
'		Next
		shortdesc = Trim(FDesc)
'		shortdesc = CDesc
'		response.write "<pre>PDesc=:" & PDesc & ":  CDesc=:" & CDesc & ":  FDesc=:" & Trim(FDesc) & ":</PRE>" & VBCrLf
	End Function


   ' grab the dept display name
   function getdeptdisp(vDept)
      dim rs1
      Set rs1 = Server.CreateObject("ADODB.Recordset")

      vSQL = "SELECT webdisplay " _
           & "FROM departments " _
           & "WHERE dept = '" & vDept & "'"
      ' response.write "<hr>" & vsql
      rs1.open vSQL & " For Browse", conn, 3
      If Not rs1.EOF Then
         getdeptdisp = rs1("Dept")
      Else
         getdeptdisp = -1
      End If
      rs1.close
    end function

   ' grab the subcat display name
   function getsubcatdisp(vSubCat)
      dim rs1
      Set rs1 = Server.CreateObject("ADODB.Recordset")

      vSubCat = URLDecode(vSubCat)

      vSQL = "SELECT webdisplay " _
           & "FROM subcategory " _
           & "WHERE subcategory = '" & replace(replace(vSubCat, "'", "''"), "_", " ") & "' For Browse"
      ' response.write "<hr>" & vsql
      rs1.open vSQL, conn, 3
      If Not rs1.EOF Then
         getsubcatdisp = rs1("webdisplay")
      Else
         getsubcatdisp = -1
      End If
      rs1.close
   end function

   ' grab the webtype display name
   function getwebtypedisp(vWebType)
      dim rs1
      Set rs1 = Server.CreateObject("ADODB.Recordset")

      vSQL = "SELECT webdisplay " _
           & "FROM webtype " _
           & "WHERE webtype = '" & replace(replace(vWebType, "'", "''"), "_", " ") & "' For Browse"
      ' response.write "<hr>" & vsql
      rs1.open vSQL, conn, 3
      If Not rs1.EOF Then
         getwebtypedisp = rs1("webdisplay")
      Else
         getwebtypedisp = -1
      End If
      rs1.close
 end function

function getfeatured(vNavTypeID)

   ' featuring -- right column
   '    -- display the first 4 items out of "new products" listing
   vSQL = "SELECT TOP 5 HTML_Special_SaleItems.*, Products.* " _
        & "FROM HTML_Special_SaleItems " _
        & "INNER JOIN Products " _
        & "ON HTML_Special_SaleItems.Col1_ProductID = Products.ProdID " _
        & "WHERE HTML_Special_SaleItems.Type=4 "
   if vNavTypeID <> "" Then
      vSQL = vSQL & "AND HTML_Special_SaleItems.NavTypeID=" & vNavTypeID & " "
   end if
   vSQL = vSQL & "AND Products.WebPosted LIKE 'yes' " _
        & "ORDER BY NEWID()"    & " For Browse"
'        & "ORDER BY NEWID(), HTML_Special_SaleItems.NavTypeID, HTML_Special_SaleItems.Sort " 

'   response.write "<hr>GetFeatured:" & vSQL
'   response.end
   rs1.open vSQL, Conn
   do while not rs1.EOF
      
      tempProd.clearitem
      tempProd.getitemPID(rs1("ProdID"))
      vTMP4 = rs1("description")
      vTMP4 = Server.HTMLEncode(vTMP4)
      vOUT9 = vOUT9 & vbcrlf & vbcrlf & vbcrlf & "<a href=""/item/" & rs1("sku") & """><img src=""" & resizepic("/productimages/" & rs1("picture"), rs1("Width_Small"), rs1("Height_Small")) & """ border=""0"" alt=""" & vTMP4 & """ vspace=""10"" ></a><BR>" & vbcrlf _
                    & "<div class=""featuringtext""><a href=""/item/" & rs1("sku") & """>" & vTMP4 & "</a></div>" & vbcrlf _
                    & "<span class=""price"">You Pay: " & FormatCurrencyDiscount("<BR>On Special", rs1("price"), tempProd.pfields.item("mDiscountAmount")) & "</span><br>" & vbcrlf _
                    & "<a href=""/item/" & rs1("sku") & """>MORE INFO</a><BR>" & vbcrlf _
                    & "<img name=""feature_divide"" src=""/images/feature_divide.gif"" width=""159"" height=""12"" border=""0"" alt=""""><BR>" & vbcrlf & vbcrlf
      rs1.movenext
   loop
   rs1.close

   dim ot1
   set ot1 = new template_cls

   with ot1
   	.TemplateFile = TMPLDIR & "featured_leftcol.html"
      .AddToken "category_type", 1, vOUT3
      .AddToken "category_name", 1, "Products" ' vOUT2
      .AddToken "featured", 1, vOUT9
      .AddToken "breadcrumb", 1, vOUT4
      .AddToken "categories_col1", 1, vOUT1
      vOut8 = .getParsedTemplateFile
   end with
   'vOUT8 = ot1.getParsedTemplateFile

   getfeatured = vOUT8

end function

function getheader(vPage)

   dim ot1, vHeadOut
   set ot1 = new template_cls

   Select Case vSection
      Case ""     '  home page
         	ot1.TemplateFile = TMPLDIR & "home_base_header.html"

      Case "displaycart"
         	ot1.TemplateFile = TMPLDIR & "cart_header.html"
            ot1.AddToken "headertitle", 1, "Verify the contents of your cart"

      Case "checkout"
         	ot1.TemplateFile = TMPLDIR & "cart_header.html"
            ot1.AddToken "headertitle", 1, "Summary of the contents of your cart"

      Case "billing"
         	ot1.TemplateFile = TMPLDIR & "cart_header.html"
            ot1.AddToken "headertitle", 1, "Summary of the contents of your cart"

   End Select

   vHeadOut = ot1.getParsedTemplateFile
   getheader = vHeadOut

end function

function getcatheader(vMetaTitle, vMetaDescription, vMetaKeywords)

   dim ot1, vHeadOut
   set ot1 = new template_cls

   Dim vMTitle, vMDesc, vMKeywords

   vMTitle = vMetaTitle & ""
   if vMTitle = "" Then vMTitle = "BicycleBuys.com | Online Bike Shop | Bicycles | Bike Parts | Frames | Pedals"

   vMDesc = vMetaDescription & ""
   if vMDesc = "" Then vMDesc = "Bicycle Buys - BicycleBuys.com - Your Online Bike Shop - We Cycle the World"

   vMKeywords = vMetaKeywords & ""
   if vMKeywords = "" Then vMKeywords = "Bicycles, Bikes, Bicycle, Bike, Bike Parts, Clothes, Helmets, Shoes, Trainers, Crank, Handle Bars, Frames, Bike Frames, Crankset, Forks, Seat, Pedals, Bike Kits, Bicycle Kits, Wheels, Tires, Tubes, Heartrate Monitor, Cycle Computer, BikeHard, Bike Hard"

   with ot1
  	   .TemplateFile = TMPLDIR & "categoryheader.html"
      .AddToken "title", 1, vMTitle
      .AddToken "description", 1, vMDesc
      .AddToken "keywords", 1, vMKeywords
   end with
   vHeadOut = ot1.getParsedTemplateFile
   getcatheader = vHeadOut

end function

function getprodheader(vMetaTitle, vMetaDescription, vMetaKeywords)

   dim ot1, vHeadOut
   set ot1 = new template_cls

   Dim vMTitle, vMDesc, vMKeywords

   vMTitle = vMetaTitle & ""
   if vMTitle = "" Then vMTitle = "BicycleBuys.com | Online Bike Shop | Bicycles | Bike Parts | Frames | Pedals"

   vMDesc = vMetaDescription & ""
   if vMDesc = "" Then vMDesc = "Bicycle Buys - BicycleBuys.com - Your Online Bike Shop - We Cycle the World"

   vMKeywords = vMetaKeywords & ""
   if vMKeywords = "" Then vMKeywords = "Bicycles, Bikes, Bicycle, Bike, Bike Parts, Clothes, Helmets, Shoes, Trainers, Crank, Handle Bars, Frames, Bike Frames, Crankset, Forks, Seat, Pedals, Bike Kits, Bicycle Kits, Wheels, Tires, Tubes, Heartrate Monitor, Cycle Computer, BikeHard, Bike Hard"

   with ot1
  	   .TemplateFile = TMPLDIR & "productheader.html"
      .AddToken "title", 1, vMTitle
      .AddToken "description", 1, vMDesc
      .AddToken "keywords", 1, vMKeywords
   end with
   vHeadOut = ot1.getParsedTemplateFile
   getcatheader = vHeadOut

end function

function getcloseouts(vMoreLink)
'Does not display Discounts

   ' featuring -- right column
   '    -- display the first 4 items out of "new products" listing
   vSQL = "SELECT TOP 8 HTML_Special_SaleItems.*, Products.* " _
        & "FROM HTML_Special_SaleItems " _
        & "INNER JOIN Products " _
        & "ON HTML_Special_SaleItems.Col1_ProductID = Products.ProdID " _
        & "WHERE HTML_Special_SaleItems.Type=1 " _
        & "AND Products.WebPosted LIKE 'yes' " _
        & "ORDER BY NEWID(), HTML_Special_SaleItems.Sort " _
        & " For Browse"

   'response.write vSQL
   rs2.open vSQL, Conn
   do while not rs2.EOF
        tempprod.clearitem
        tempprod.getitemPID(rs2("ProdID"))
      vTMP4 = rs2("description")
      vTMP4 = Server.HTMLEncode(vTMP4)
      vOUT9 = vOUT9 & vbcrlf & vbcrlf & vbcrlf & "<a href=""/item/" & rs2("sku") & """><img src=""" & resizepic("/productimages/" & rs2("picture"), rs2("Width_Small"), rs2("Height_Small")) & """ border=""0"" alt=""" & vTMP4 & """ vspace=""10"" ></a><BR>" & vbcrlf _
                    & "<div class=""featuringtext""><a href=""/item/" & rs2("sku") & """>" & vTMP4 & "</a></div>" & vbcrlf _
                    & "<span class=""price"">You Pay: " & FormatCurrencyDiscount("",rs2("price"),tempprod.pfields("mDiscountAmount")) & "</span><br>" & vbcrlf _
                    & "<a href=""/item/" & rs2("sku") & """>MORE INFO</a><BR>" & vbcrlf _
                    & "<img name=""feature_divide"" src=""/images/feature_divide.gif"" width=""159"" height=""12"" border=""0"" alt=""""><BR>" & vbcrlf & vbcrlf
      rs2.movenext
   loop
   rs2.close

   dim ot1
   set ot1 = new template_cls

   with ot1
   	.TemplateFile = TMPLDIR & "featured_right.html"
      .AddToken "category_type", 1, vOUT3
      .AddToken "category_name", 1, vOUT2
      .AddToken "featured", 1, vOUT9
      .AddToken "breadcrumb", 1, vOUT4
      .AddToken "categories_col1", 1, vOUT1
      .AddToken "moreitemslink", 1, vMoreLink

       vOut8 = .getParsedTemplateFile
   end with

   getcloseouts = vOUT8

end function


sub cycleviewed(vName, vURL)

   ' store the array in the session variable
   Dim vRVName, vRVURL, vRVNum

   vRVName = Session("RVName")
   vRVURL = Session("RVURL")
   vRVNum = Session("RVNum")

end sub

Function getitemrow(vCount, vBGColor, ItemID, ItemName, ItemWeight, ItemQuantity, _
                   ItemCustom1,ItemCustom2,ItemCustom3,ItemCustom4,ItemCustom5, _
                   ItemCustom6,ItemCustom7,ItemCustom8,ItemOptions,FreightMsg, _
                   ItemPrice)

   oProd1.ClearItem
   oProd1.getitemSKU(ItemID)
   vItemPicture = oProd1.pfields.Item("picture")
   dim JKWebNote
   JKWebNote = oProd1.pfields.Item("WebNote")

	dim psku
	psku = ""


   vSQL100 = "SELECT  top 100 P.SKU, P.picture " _
        & "FROM   [Products Children] C " _
        & "INNER JOIN Products P " _
        & "ON C.ProdID = P.ProdID " _
        & "WHERE C.ChildProdID = " & oProd1.pfields.Item("ProdID") & "" _
        & " For Browse"

   rs100.open vSQL100, Conn
   if not rs100.EOF then
 		vItemPicture = rs100("picture")
		psku = rs100("SKU")
   end if
   rs100.close


   Dim objItemRow
   Set objItemRow = new template_cls

	'response.write vCount & ", " &  vBGColor & ", " &  ItemID & ", " &  ItemName & ", " &  ItemWeight & ", " &  ItemQuantity & ", " &  ItemCustom1 & ", " & ItemCustom2 & ", " & ItemCustom3 & ", " & ItemCustom4 & ", " & ItemCustom5 & ", " &  ItemCustom6 & ", " & ItemCustom7 & ", " & ItemCustom8 & ", " & ItemOptions & ", " & FreightMsg & ", " &  ItemPrice

   With objItemRow
      .AddToken "itemid", 1, ItemID
      .AddToken "itemparentid", 1, ItemCustom3
      .AddToken "count", 1, vCount
      .AddToken "itemname", 1, ItemName
      .AddToken "itempicture", 1, vItemPicture
      .AddToken "bgcolor", 1, vBGColor
      .AddToken "itemname", 1, ItemName
      .AddToken "itemweight", 1, ItemWeight
      .AddToken "itemquantity", 1, ItemQuantity
      .AddToken "itemcustom1", 1, ItemCustom1
      .AddToken "itemcustom2", 1, ItemCustom2
if (psku <> "") then
	.AddToken "itemcustom3", 1, psku
else
	.AddToken "itemcustom3", 1, ItemID
end if

      .AddToken "itemcustom4", 1, ItemCustom4
      .AddToken "itemcustom5", 1, ItemCustom5
      .AddToken "itemcustom6", 1, ItemCustom6
      .AddToken "itemcustom7", 1, ItemCustom7
      .AddToken "itemcustom8", 1, ItemCustom8
      .AddToken "itemoptions", 1, ItemOptions
      .AddToken "freightmsg", 1, FreightMsg
      if (JKWebNote <> 15) then
      		.AddToken "itemprice", 1, FormatCurrency(ItemPrice, 2, 0, 0)
      else
      		.AddToken "itemprice", 1, FormatCurrency(ItemPrice, 2, 0, 0)

      end if
      .AddToken "itemquantity", 1,ItemQuantity
      if (JKWebNote <> 15) then
      		.AddToken "itempriceextended", 1, FormatCurrency(ItemPrice * ItemQuantity, 2, 0, 0)
      else
      		.AddToken "itempriceextended", 1, FormatCurrency(ItemPrice * ItemQuantity, 2, 0, 0)
      		TotalDiscount15 = TotalDiscount15 + ((ItemPrice - oProd1.pfields.Item("RetailWebPrice")) * ItemQuantity)
      end if
      .TemplateFile = TMPLDIR & "displaycart-itemrow.html"
      getitemrow = .getParsedTemplateFile
   end with
   set objItemRow = nothing
End Function


Function getitemrowco(vCount, vBGColor, ItemID, ItemName, ItemWeight, ItemQuantity, _
                   ItemCustom1,ItemCustom2,ItemCustom3,ItemCustom4,ItemCustom5, _
                   ItemCustom6,ItemCustom7,ItemCustom8,ItemOptions,FreightMsg, _
                   ItemPrice)

   oProd1.ClearItem
   oProd1.getitemSKU(ItemID)
   vItemPicture = oProd1.pfields.Item("picture")
	dim psku
	psku = ""

   dim JKWebNote
   JKWebNote = oProd1.pfields.Item("WebNote")


   vSQL100 = "SELECT top 100 P.SKU, P.picture " _
        & "FROM   [Products Children] C " _
        & "INNER JOIN Products P " _
        & "ON C.ProdID = P.ProdID " _
        & "WHERE C.ChildProdID = " & oProd1.pfields.Item("ProdID") & "" _
        & " For Browse"

   rs100.open vSQL100, Conn
   if not rs100.EOF then
 		vItemPicture = rs100("picture")
		psku = rs100("SKU")
   end if
   rs100.close

   Dim objItemRow
   Set objItemRow = new template_cls

   With objItemRow
      .AddToken "itemid", 1, ItemID
      .AddToken "itemparentid", 1, ItemCustom3
      .AddToken "count", 1, vCount
      .AddToken "itemname", 1, ItemName
      .AddToken "itempicture", 1, vItemPicture
      .AddToken "bgcolor", 1, vBGColor
      .AddToken "itemname", 1, ItemName
      .AddToken "itemweight", 1, ItemWeight
      .AddToken "itemquantity", 1, ItemQuantity
      .AddToken "itemcustom1", 1, ItemCustom1
      .AddToken "itemcustom2", 1, ItemCustom2
if (psku <> "") then
	.AddToken "itemcustom3", 1, psku
else
	.AddToken "itemcustom3", 1, ItemID
end if
      .AddToken "itemcustom4", 1, ItemCustom4
      .AddToken "itemcustom5", 1, ItemCustom5
      .AddToken "itemcustom6", 1, ItemCustom6
      .AddToken "itemcustom7", 1, ItemCustom7
      .AddToken "itemcustom8", 1, ItemCustom8
      .AddToken "itemoptions", 1, ItemOptions
      .AddToken "freightmsg", 1, FreightMsg
      if (JKWebNote <> 15) then
      		.AddToken "itemprice", 1, FormatCurrency(ItemPrice, 2, 0, 0)
      else
      		.AddToken "itemprice", 1, FormatCurrency(ItemPrice, 2, 0, 0)

      end if
      .AddToken "itemquantity", 1,ItemQuantity
      if (JKWebNote <> 15) then
      		.AddToken "itempriceextended", 1, FormatCurrency(ItemPrice * ItemQuantity, 2, 0, 0)
      else
      		.AddToken "itempriceextended", 1, FormatCurrency(ItemPrice * ItemQuantity, 2, 0, 0)
      		TotalDiscount15 = TotalDiscount15 + ((ItemPrice - oProd1.pfields.Item("RetailWebPrice")) * ItemQuantity)
      end if
      .TemplateFile = TMPLDIR & "displaycart-itemrowco.html"
      getitemrowco = .getParsedTemplateFile
   end with
   set objItemRow = nothing
End Function

' put's an item into the recently viewed list
' will not put the same item in twice
sub putinrecentlyviewed(vItem)
   dim vRecArr, counter
   vRecentlyViewed = Session("RecentlyViewed")
   'response.write "<hr>IN:" & vRecentlyViewed
   ' if we've already viewed it, then remove it, move all the other items
   '    down in the showit order, and out this one back on the top
   if Instr(vRecentlyViewed, vItem) Then
      vRecentlyViewed = replace(vRecentlyViewed, vItem & "|", "")
	  vRecentlyViewed = replace(vRecentlyViewed, "||", "")
   end if
   vRecentlyViewed = vItem & "|" & vRecentlyViewed
   vRecArr = split(vRecentlyViewed, "|")
   vRecentlyViewed = ""
   counter = 0
   do while vRecArr(counter) <> ""
   		if (counter < 5) then
			vRecentlyViewed = vRecentlyViewed & vRecArr(counter) & "|"
		end if
		counter = counter + 1
   loop
   Session("RecentlyViewed") = vRecentlyViewed

'	response.write "<hr>OUT:" & vRecentlyViewed
end sub

' get the home page "most popular" category and subcategories with images
function hpmostpop

   Dim vCatName, vCatDisp, vSubCats
 	   vTMP1 = ""
	   vTMP2 = ""
 	Dim poparray, textarray, linkarray
	poparray = Array(161, 125, 126, 162, 127)
	textarray = Array("Chains", "Cassettes", "Cranks", "Rear Derailleur", "Front Derailleur")
	linkarray = Array("/drivetrain/Chains/", "/drivetrain/Cassettes/", "/drivetrain/Cranks/", "/drivetrain/DeraillRear/", "/drivetrain/DeraillFront/")

	For I = LBound(poparray) To UBound(poparray)
	   vSQL = "SELECT top 1 p.picture, p.subcatid, s.SubCategory, s.WebDisplay, p.Width_Small, p.Height_Small " _
			& "FROM products p  " _
			& "INNER JOIN subcategory S  " _
			& "ON s.SubCatID = p.subcatid " _
			& "WHERE p.WebTypeID = " & poparray(I) & " AND p.WebPosted LIKE 'yes' ORDER BY newid() " _
		        & " For Browse"
	 ' response.write "<hR>" & vSQL
	   rs1.open vSQL, conn, 3


	   if NOT rs1.EOF then
		  vTMP1 = vTMP1 & "<TD class=""driveitems"" align=center><a href=""" & lcase(linkarray(I)) & """><img src=""" & resizepic("/productimages/" & rs1("picture"), rs1("Width_Small"), rs1("Height_Small")) & """  border=0></a></TD>" & chr(13)
		  vTMP2 = vTMP2 & "<TD class=""driveitems"" align=center><a href=""" & lcase(linkarray(I)) & """>" & textarray(I) & "</a></TD>" & chr(13)
		  rs1.movenext
	   end if
	   rs1.close
	Next




   Dim ot1
   set ot1 = new template_cls
   with ot1
   	.TemplateFile = TMPLDIR & "home_base_mostpopcat.html"
      .AddToken "mpcategorydisp", 1, "Drive Train"
      .AddToken "mpcategorylink", 1, vCatName
      .AddToken "subcatline1", 1, vTMP1
      .AddToken "subcatline2", 1, vTMP2
      vOut8 = .getParsedTemplateFile
   end with

   hpmostpop = vOUT8
end function

function getsearch
   Dim ot1
   Dim locSearchTerm
   locSearchTerm = vSearchTerm
   if (locSearchTerm = "") then
   		locSearchTerm = "Keyword Search"
	end if

	locSearchTerm = replace(locSearchTerm, "%20", " ")

   set ot1 = new template_cls
   with ot1
   	.TemplateFile = vSearchSection
      .AddToken "search_options", 1, getsearchopts
      .AddToken "search_term", 1, locSearchTerm
	  .AddToken "brand_options", 1, getbrandopts
	  .AddToken "cat_display", 1, getcatdisplay
      vTMP1 = .getParsedTemplateFile
   end with
   getsearch = vTMP1
end function

' returns a list of options for search
' based on scripting dictionary vWebTypeListAZSD
function getsearchopts
   Dim vSDC, vSDCA, vWTA
   vSDC = vWebTypeListingAZSD.Count-1
   vSDCA = vWebTypeListingAZSD.Items
   vTMP2 = ""
   ' response.write "<hr>sc:" & vSearchCategory
   For vType = 0 to vSDC
      vWTA = split(vSDCA(vType), "|")
     ' vTMP1 = Left(vWTA(0), 12)
	   vTMP1 = vWTA(0)
      vSelected = ""
     ' if vWTA(1) = vSearchCategory then vSelected =" SELECTED"
      vTMP2 = vTMP2 & "<option value=""" & vWTA(1) & """" & vSelected & ">" & vTMP1 & "</option>" & chr(13)
   Next
   getsearchopts = vTMP2
end function

' returns a list of brands
function getbrandopts
	  vTMP2 = ""
      vSQL100 = "SELECT DISTINCT vendor.* " _
			   & "FROM products " _
			   & "INNER JOIN Vendor " _
			   & "ON vendor.vendid = products.vendid " _
			   & "WHERE 1=1 " _
			   & " AND webposted LIKE 'yes' " _
			   & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') " _
			   & " ORDER BY Vendor" _
		           & " For Browse"

		rs1.open vSQL100, conn, 3

		 do while not rs1.EOF
		 	vTMP2 = vTMP2 & "<option value=""" & "/manufacturer/" & replace(replace(rs1("Vendor"), " ", "_"), "'", "\'") & """" & vSelected & ">" & rs1("Vendor") & "</option>" & chr(13)
		 	rs1.movenext
		 loop
		 rs1.close


		 getbrandopts = vTMP2

end function


function getcatdisplay
	dim cathtml, cat_name, vNavSet, vSQL1, deptname, deptnameweb

	  cathtml = "<div class=""breadcrumb"">"

	  if (vSection = "search") then
	  	   cathtml =  cathtml & "<a href=""/"">Home</a>"
		   cathtml =  cathtml & " &gt;&gt;  Search"
	  elseif (vSection = "closeouts") then
	  	   cathtml =  cathtml & "<a href=""/"">Home</a>"
		   cathtml =  cathtml & " &gt;&gt; <a href=""/closeouts"">Specials</a>"
	  elseif (vSection = "newitems") then
	  	   cathtml =  cathtml & "<a href=""/"">Home</a>"
		   cathtml =  cathtml & " &gt;&gt; <a href=""/newitems"">New Products</a>"
	  elseif (vSection = "item") then
	  	   cathtml =  cathtml & "<a href=""/"">Home</a>"
			vSQL100 = "SELECT J.NavType, J.WebDisplayForNavType FROM products P, JohnWebNavType J WHERE (P.SKU LIKE '" & vItem & "') AND ((J.WebTypes LIKE '%' + CAST(P.WebTypeID AS nvarchar(20)) + '%') OR (J.SubCats LIKE '' + CAST(P.SubCatID AS nvarchar(20)) + ',%') OR (J.SubCats LIKE '%,' + CAST(P.SubCatID AS nvarchar(20)) + ',%') OR (J.SubCats LIKE '%' + CAST(P.SubCatID AS nvarchar(20)) + ''))" _
		        & " For Browse"

			rs1.open vSQL100, Conn
			if not rs1.EOF	then
				vSection = LCase(rs1("NavType"))
				cat_name = rs1("WebDisplayForNavType")
			end if
			rs1.close
			cathtml =  cathtml & " &gt;&gt; <a href=""/" & vSection & """>" & cat_name & "</a>"

			  vSQL1 = "SELECT top 100 S.subcatid, S.WebDisplay, S.SubCategory " _
				   & "FROM subcategory S INNER JOIN Products P ON S.SubCatID = P.SubCatID " _
				   & "WHERE P.SKU LIKE '" & vItem & "'" _
			           & " For Browse"
			  rs1.open vSQL1, conn, 3
			  If Not rs1.EOF Then
				deptname = rs1("WebDisplay")
				deptnameweb =  rs1("SubCategory")
			  End If
			  rs1.close
			  cathtml =  cathtml & " &gt;&gt; <a href=""/" & vSection & "/" & deptnameweb & """>" & deptname & "</a>"


	  elseif (vSection = "allmfg") then
		  cathtml =  cathtml & "<a href=""/"">Home</a>"
	      if (vDept <> "") then
			  if instr(vNavTypes, vDept) then vSect = "W" else vSect = "S"
			  Select Case vSect
				 Case "S"
					vNavSet = "SubCats NS, WebDisplayForNavType, WebDisplayForCategory "
					vSQL1 = "SELECT MetaTitle, MetaDescription, MetaKeywords, " & vNavSet _
						  & "FROM JohnWebNavType " _
						  & "WHERE NavType LIKE '" & vDept & "' For Browse"
				 Case Else
					vNavSet = "NavType, WebTypes NS, WebDisplayForNavType, WebDisplayForCategory "
					vSQL1 = "SELECT MetaTitle, MetaDescription, MetaKeywords, " & vNavSet  _
						  & "FROM JohnWebNavType " _
						  & "WHERE NavType LIKE '" & vDept & "' For Browse"
			  end select
			  rs1.open vSQL1, conn, 3
			  if Not rs1.EOF Then
				 cat_name = rs1("WebDisplayForNavType")
			  end if
			  rs1.close
			  cathtml =  cathtml & " &gt;&gt; <a href=""/" & vDept & """>" & cat_name & "</a>"
		  end if

		  if (vManufacturer <> "") then
		  	if (vDept <> "") then
				cathtml =  cathtml & " &gt;&gt; <a href=""/manufacturer/" & vManufacturer & "/" & vDept & """>" & replace(vManufacturer, "_", " ") & "</a>"
		 	else
				cathtml =  cathtml & " &gt;&gt; <a href=""/manufacturer/" & vManufacturer & """>" & replace(vManufacturer, "_", " ") & "</a>"
			end if
		  end if

	  elseif (vSection <> "") then
		  cathtml =  cathtml & "<a href=""/"">Home</a>"
	      if instr(vNavTypes, vSection) then vSect = "W" else vSect = "S"
		  Select Case vSect
			 Case "S"
				vNavSet = "SubCats NS, WebDisplayForNavType, WebDisplayForCategory "
				vSQL1 = "SELECT MetaTitle, MetaDescription, MetaKeywords, " & vNavSet _
					  & "FROM JohnWebNavType " _
					  & "WHERE NavType LIKE '" & vSection & "' For Browse"
			 Case Else
				vNavSet = "NavType, WebTypes NS, WebDisplayForNavType, WebDisplayForCategory "
				vSQL1 = "SELECT MetaTitle, MetaDescription, MetaKeywords, " & vNavSet  _
					  & "FROM JohnWebNavType " _
					  & "WHERE NavType LIKE '" & vSection & "' For Browse"
		  end select
		  rs1.open vSQL1, conn, 3
		  if Not rs1.EOF Then
			 cat_name = rs1("WebDisplayForNavType")
		  end if
		  rs1.close
		  cathtml =  cathtml & " &gt;&gt; <a href=""/" & vSection & """>" & cat_name & "</a>"

		  if (vDept = "all") then
				cathtml =  cathtml & " &gt;&gt; <a href=""/" & vSection & "/" & vDept & "/"">" & "All Products" & "</a>"
		  elseif (vDept <> "") then

			  vSQL1 = "SELECT subcatid, WebDisplay " _
				   & "FROM subcategory " _
				   & "WHERE subcategory = '" & vDept & "' For Browse"
			  rs1.open vSQL1, conn, 3
			  If Not rs1.EOF Then
				deptname = rs1("WebDisplay")
			  End If
			  rs1.close
		  	cathtml =  cathtml & " &gt;&gt; <a href=""/" & vSection & "/" & vDept & "/"">" & deptname & "</a>"
		  end if
	  end if
	  cathtml = cathtml & "</div>"

	  getcatdisplay = cathtml

end function


Function CapitalizeValue(sFullName)

   Dim  fCapitalizeNextLetter
   Dim sNewName
   Dim sChar
   Dim i

   sFullName = Trim(sFullName)
   if (isnull(sFullName)) Then
   	sFullName = ""
	end if

   if sFullName = "" then CapitalizeValue = sFullName

   fCapitalizeNextLetter = true

   For i=1 to Len(sFullName)

         sChar = Mid(sFullName,i,1)

         if (fCapitalizeNextLetter = true) then
             if IsLetter(sChar) = true then
                sChar = UCase(sChar)
                fCapitalizeNextLetter = false
             else
                fCapitalizeNextLetter = true
                sChar = LCase(sChar)
             end if
         else
            if IsLetter(sChar) = false then
               fCapitalizeNextLetter = true
            else
            	sChar = LCase(sChar)
            end if


         end if


        sNewName = sNewName & sChar

   Next

    CapitalizeValue = sNewName

end Function

function IsLetter(sChar)
	dim fRet, nASCII
     fRet = false
     nASCII = ASC(sChar)

    if ((nASCII >=65) and (nASCII <=90)) then fRet = true   ' Upper case letters
    if ((nASCII >=97) and (nASCII <=122)) then fRet = true ' Lower case letters

   IsLetter = fRet

end function

' for pagination we build the  prev...1.2.3.4...next display here
Sub ShowPageNav()
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Build the navigation bar
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

   Dim vHRefH, vHRefT, vImgH, vImgT
   Dim vFP, vPP, vNP, vLP
   Dim vNavID, vNav1, vNav2

	vHRefH = "<a href=""/Items01.asp?NavID=" & vNavID
	vHRefH = vHRefH & "&M=" & vManufacturer
	vHRefH = vHRefH & "&T=" & vSection
	vHRefH = vHRefH & "&P=" & vPageNo
	vHRefH = vHRefH & "&D="""

	vHRefT = "</a>"
	vImgH = "<img src=""/images/"
	vImgT = "page.gif"" border=""0"" align=""absmiddle"">"

	vFP = vImgH & "first" & vImgT
	vPP = vImgH & "previous" & vImgT
	vNP = vImgH & "next" & vImgT
	vLP = vImgH & "last" & vImgT

	if vPageNo > 1 then
		vFP = vHRefH & "FP"">" & vFP & vHRefT
		vPP = vHRefH & "PP"">" & vPP & vHRefT
	End If

	if vPageNo < RS1.PageCount then
		vNP = vHRefH & "NP"">" & vNP & vHRefT
		vLP = vHRefH & "LP"">" & vLP & vHRefT
	End If
	vNav1 = vFP & vPP
	vNav2 = vNP & vLP

	if rs1.pagecount > 1 then
		Response.write "<CENTER>" & vNav1 & " " & vNav2 & "<FONT id=""fineprint""><BR>Pages: "
		For x = 1 to rs1.pagecount
			If x = vPageNo then
				response.write "<b>" & x & " </b>"
			Else
				response.write "<a href=""/Items01.asp?NavID=" & vNavID & "&M=" & vManufacturer & "&T=" & vSection & "&"
				if x = 1 then
					response.write "D=" & vFirst & """>"
				elseif x = RS1.pagecount then
					response.write "D=" & vLast & """>"
				else
					response.write "P=" & x-1 & "&D=" & vNext & """>"
				end if
			response.write x & " </a>"
			End if
		next
		response.write "</FONT></CENTER><BR><BR>"
	End If

End Sub

Function resizepic(imageurl, w, h)
	if (w = h) then
		imageurl = imageurl & """ width=""80" & """ height=""80"
	elseif (w > h) then
		imageurl = imageurl & """ width=""80"
	else
		imageurl = imageurl & """ height=""80"
	end if
	resizepic = imageurl
End Function

Function EnCrypt(strCryptThis)
   Dim strChar, iKeyChar, iStringChar, i, iCryptChar, strEncrypted
   for i = 1 to Len(strCryptThis)
      iKeyChar = Asc(mid(g_Key,i,1))
      iStringChar = Asc(mid(strCryptThis,i,1))
      iCryptChar = iKeyChar Xor iStringChar
      strEncrypted =  strEncrypted & Chr(iCryptChar)
   next
   EnCrypt = strEncrypted
End Function

Function DeCrypt(strEncrypted)
Dim strChar, iKeyChar, iStringChar, i, iDeCryptChar, strDecrypted
   for i = 1 to Len(strEncrypted)
      iKeyChar = (Asc(mid(g_Key,i,1)))
      iStringChar = Asc(mid(strEncrypted,i,1))
      iDeCryptChar = iKeyChar Xor iStringChar
      strDecrypted =  strDecrypted & Chr(iDeCryptChar)
   next
   DeCrypt = strDecrypted
End Function

Function ReadKeyFromFile(strFileName)
   Dim keyFile, fso, f, ts
   set fso = Server.CreateObject("Scripting.FileSystemObject")
   set f = fso.GetFile(strFileName)
   set ts = f.OpenAsTextStream(1, -2)

   Do While not ts.AtEndOfStream
     keyFile = keyFile & ts.ReadLine
   Loop

   ReadKeyFromFile =  keyFile
End Function

Function FormatCurrencyDiscount( AdditionalText,fPrice, fDiscountAmount)

	if fDiscountAmount = 0 then
		FormatCurrencyDiscount= formatcurrency(fPrice, 2, 0, 0)
	else
		FormatCurrencyDiscount=  "<strike>" & formatcurrency(fPrice, 2, 0, 0) & "</strike>" & AdditionalText
	end if

End function


' debugging
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'response.write Request.ServerVariables("REMOTE_ADDR")

vDebugx = False
'vDebugx = True
if vDebugx and Request.ServerVariables("REMOTE_ADDR") = "68.194.179.95" Then
   vDebugx = True
Else
   vDebugx = False
end if
'response.write(Now())
%>
 