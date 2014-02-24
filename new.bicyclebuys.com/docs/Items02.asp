<!--#INCLUDE file="includes/template_cls.asp"-->
<!--#INCLUDE file="includes/common.asp"-->
<!--#INCLUDE file="includes/IndexFunctionsP.asp"-->
<% 
'        vMv 
'        vPageNo 

    dim vRecArr
    Dim vHL, vWL, vHS, vWS
    Dim vCP
    dim vVST, vSDX

    Dim vVendlist
    Dim vInSearch, vVendorString, vTerm, vTerms, vIsVend
    Dim vSearchTermEndInVS, vVendIDEndInVS
    Dim vVendSQL, vSQLVL, vVendid

    dim xx, temp_1, temp_2
    Dim vURI, vProdMeta

    'response.Write (now())
    'response.write "Recently Viewed: " & session("RecentlyViewed")
    'response.end 
    dim sPrice, sDiscount
    
   ' BICYCLEBUYS.COM
   '
   ' index.asp 
   ' protection from evil / sql injections / overflows / etc
   ' we can narrow this down further once development is complete
  
    call getRequestStrings()   
    call getPageNumbers()   
    call getSearchTerm
     
'   response.write "sect = |" & vSection & "|<br>"
'   response.write "dept = |" & vDept & "|<br>"
'   response.write "item = |" & vItem & "|<br>"
'   response.write "manufacturer = |" & vManufacturer & "|<br>"
'
'   response.write "mv = |" & vMv & "|<br>"
'   response.write "pageno = |" & vPageNo & "|<br>"
'   response.write "searchvendid = |" & vSearchVendID & "|<br>"


    vURI = Server.HTMLEncode(Request.QueryString)
    '  response.write Request.QueryString
    '  Response.write "<HR>URI=" & vURI & "<br>"
    '   dim oProd1
    '   set oProd1 = new bb_product
    '   oProd1.getitem(15)
    '
    '   dim oProd2
    '   set oProd2 = new bb_product
    '   tempProd.getitem(17)
    '
    '   ' response.write "<hr>" &  oProd1.val("RetailMarkupPerc") & "<hr>"
    '   ' response.write "<hr>" &  tempProd.val("RetailMarkupPerc") & "<hr>"
'if    Request.ServerVariables("REMOTE_ADDR")  = "207.237.26.122" then
'   response.write "<hr>|vSection=" & vSection & "|/|vDept=" & vDept & "|/|vItem=" & vItem & "|/|vManufacturer=" & vManufacturer & "|<hr>"
'end if

    ' all pages use the search, so let's get that done right away
    vSearchPage = getsearch()

if    Request.ServerVariables("REMOTE_ADDR")  = "207.237.26.122" then
   response.write "<hr>xxx|vSection=" & vSection & "|/|vDept=" & vDept & "|/|vItem=" & vItem & "|/|vManufacturer=" & vManufacturer & "|<hr>"
end if
    ' now we begin page processing
    Select Case vSection
      Case ""     '  home page

         getMostPopularProducts()
         getNewProducts()        
         ' this should be closeouts...
        vOUT9 = getcloseouts("closeouts")  
         ' popular category listing
         vOUT10 = getpopcategories
         ' most popular category w/ subcats
         
         vOUT11 = hpmostpop
%>  
		 <!--#INCLUDE file="includes/moving.asp"-->
<%
         with objTemplate
         	.TemplateFile = TMPLDIR & "home_base.html"

            .AddToken "closeouts", 1, vOUT9
            .AddToken "popcat", 1, vOUT10
            .AddToken "mostpopcat", 1, vOUT11

         	.AddToken "header", 3, vHeader
         	.AddToken "search_section", 1, vSearchPage
			.AddToken "moving", 1, moving
         	.AddToken "footer", 3, vFooter

         	.parseTemplateFile
         end with
         set objTemplate = nothing

      Case Else    ' interior pages

         ' Interior page used for:
         '     - item display  - item passed - i=sku
         '     - closeouts / newitems - c="closeouts" or "newitems"
         '     - show all product categories of a department d=xxxxx
         '          also show list of vendors
         '     - show all products in a department of one vendor c="allmfg"
         '     - show all products in a single category d=all
         '     - product search

         ' if the item field is blank then we know we are not showing a product detail page
         ' if no department then we're at the first category level
         ' and if we're not showing closeouts,newitems, search results or allmfg then...
         if vItem = "" AND vDept="" AND (vSection <> "closeouts" AND vSection <> "newitems" AND vSection <> "allmfg" AND vSection <> "search") Then

               'vMetaTitle = ucase(vSection)
               ' get the category links
               getcatlinks vSection

               vBlurb = vMetaDescription & ""
               if vBlurb = "" Then vBlurb = "We are your greatest source featuring a complete line of bicycle and bicycle parts and accessories."

               ' get the manufacturer links
               vOUT5 = ""  : vOUT6 = "" : vOUT7 = "" : vOUT8 = "" : vOUT9 = ""
               getmfglinks vSection   ' -- I don't think this mfg list is going to work here.. too early and we know too little
                               '     to show products on the next refresh...
               vOUT6 = ""  ' sb alllink
               vOUT9 = getfeatured("")

               vOUT10 = getcatheader(vMetaTitle, vMetaDescription, vMetaKeywords)

			    vOUT103 = ""
			    vOUT103 = vOUT103 & " <TR>"
               
                getMostPopularFour

         ' we have no item sku
         ' this isnt a product searc
         ' or we do have a department
         ' or the section is closeouts, newitems, or allmfg
         ElseIf vItem="" AND ((vDept <> "" AND vSection <> "search") OR (vSection = "closeouts" OR vSection = "newitems" OR vSection = "allmfg" ))Then
            ' we have an item sku --  show the actual product detail page
 
            getNonSearchPages("")

         ElseIf vSection = "search" Then
            ''''' begin search
            ' for product search
            'getSearchHTML()
            getNonSearchPages("Search")
         ElseIf vItem <> "" Then
 
            
    	 	if (vSection = "item") then
			vSQL100 = "SELECT top 1 J.NavType FROM vwWebproducts P, JohnWebNavType J WHERE (P.SKU LIKE '" & vItem & "') AND ((J.WebTypes LIKE '%' + CAST(P.WebTypeID AS nvarchar(20)) + '%') OR (J.SubCats LIKE '' + CAST(P.SubCatID AS nvarchar(20)) + ',%') OR (J.SubCats LIKE '%,' + CAST(P.SubCatID AS nvarchar(20)) + ',%') OR (J.SubCats LIKE '%' + CAST(P.SubCatID AS nvarchar(20)) + ''))  for browse"
	         	rs2.open vSQL100, Conn
         		if not rs2.EOF	then
				vSection = LCase(rs2("NavType"))
			end if
			rs2.close
		end if
 
             getcatlinks vSection
             
			 vOUT101 = vOUT1
			 vOUT1 = ""
			 vOUT102 = vOUT3
			 vOUT3 = ""
			 vOUT103 = vOUT2
			 vOUT2 = ""
            putinrecentlyviewed(vItem)


            if instr(vNavTypes, vSection) Then
               vUDept = getsubcatdisp(vDept)
            else
               vUDept = getwebtypedisp(vDept)
            end if

            getItemInformation()

            buildImageDimensions("")
            getRecentlyViewed()
'zzzzz
            with objTemplate
            	.TemplateFile = TMPLDIR & "product.html"
               .AddToken "title", 1,  oProd1.pfields.Item("description")
               .AddToken "itemdesc", 1, vDesc
               .AddToken "itemdesc2", 1, vDesc2
               .AddToken "itemname", 1, vItemDesc
               .AddToken "itemimage", 1, vItemImageOut1
               .AddToken "itembrand", 1, vBrand
               .AddToken "imagewidth", 1, vImageWidth
               .AddToken "origprice", 1, vOrigPrice
               .AddToken "savings", 1, vSavings
               .AddToken "mDiscountType", 1, oProd1.pfields.Item("aDiscountType")
               .AddToken "mDiscountAmount", 1, oProd1.pfields.Item("aDiscount")
               .AddToken "mSpecialPricing", 1, oProd1.pfields.Item("mSpecialPricing")

               if (oProd1.pfields.Item("WebNoteID") <> 15) then
               		.AddToken "price", 1,  formatcurrency(oProd1.pfields.Item("price"), 2, 0, 0)
               else
               		.AddToken "price", 1, formatcurrency(oProd1.pfields.Item("price"), 2, 0, 0)
               end if
	            .AddToken "StrikePrice",1, FormatCurrencyDiscount("", oProd1.pfields.Item("price"), oProd1.pFields("aDiscount"))
               .AddToken "referer", 1, vReferer
               .AddToken "referer1", 1, vReferer1
               .AddToken "itemurl", 1, "/item/" & oProd1.pfields.Item("sku")
               .AddToken "itemsku", 1, oProd1.pfields.Item("SKU")
               .AddToken "itemparent", 1, oProd1.pfields.Item("SKU")
               .AddToken "itemid", 1, oProd1.pfields.Item("ProdID")
               .AddToken "freefreight", 1, vFreeFreight
               .AddToken "overweightflag", 1, vOverWeight
               .AddToken "itemoptions", 1, vItemOptions
               if (oProd1.pfields.Item("WebNoteID") <> 15) then
               		.AddToken "WebNote", 1,  vwebnote & " "
               else
               '		.AddToken "WebNote", 1,  " <div class=product_notes><div class=button><a href=""javascript:void(0)"" onClick=""window.open('/rebate_price.asp?SKU=" & oProd1.pfields.Item("SKU") & "', 'BikePopUp',  'width=520,height=400,toolbar=0,scrollbars=1,screenX=50,screenY=50,left=50,top=50')"">Click here to View the Instant Rebate Price you will see in the Checkout</a></div></div>"  & " " & vwebnote & " "
               end if
                vHeader = TMPLDIR & "item_base_header.html"
                vOUT9 = getcloseouts("closeouts")
                getTop3Related(oProd1)

	       vProdMeta = newHeader(vItemDesc, oProd1.pfields.Item("marketingdescription"), -1)
               .AddToken "prodmeta",1, vProdMeta
               .AddToken "featured", 1, vOUT11
               .AddToken "closeouts", 1, vOUT9

               .AddToken "relatedproducts", 1, vOUT105
               .AddToken "previouslyviewed", 1, vOUT100

               .AddToken "linktitle", 1, vOUT3
               .AddToken "breadcrumb", 1, vOUT4
               .AddToken "categories_col1", 1, vOUT1
               .AddToken "categories_col2", 1, vOUT2

            	.AddToken "header", 3, vHeader
            	.AddToken "search_section", 1, vSearchPage
            	.AddToken "footer", 3, vFooter

                .AddToken "category_type", 1, vOUT102
                .AddToken "category_name", 1, vOUT103
                .AddToken "categories_col1", 1, vOUT101
		.AddToken "alllink", 1, "/" & vSection
		if oProd1.pfields.Item("ManuLink") <> "" then
			.AddToken "ManuLink", 1, "<DIV align=""right""><A href=""" & Server.HTMLEncode(oProd1.pfields.Item("ManuLink")) & """target=""_blank"">Manufacturer info</a></div>"
		end if
'if    Request.ServerVariables("REMOTE_ADDR")  = "207.237.26.122" then
'response.write(oProd1.pfields.Item("ManuLink")  )
'end if
            	.parseTemplateFile
            end with
         end if
         set objTemplate = nothing
      End Select

    if (conn.state = 1) then
	    conn.close
	    set conn = nothing
    end if
    if (RS1.state = 1) then
	    RS1.close
	    set RS1 = nothing
    end if
    if (RS2.state = 1) then
	    RS2.close
	    set RS2 = nothing
    end if
    if (rs100.state = 1) then
	    rs100.close
	    set rs100 = nothing
    end if


%>
