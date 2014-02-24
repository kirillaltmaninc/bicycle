
<!--#INCLUDE file="includes/template_cls.asp"-->
<!--#INCLUDE file="includes/getKeyWords.asp"-->
<!--#INCLUDE file="includes/common.asp"-->
<!--#INCLUDE file="includes/IndexFunctionsPrich.asp"-->
<!--#INCLUDE file="includes/redirects.asp"-->
<%

Dim TMPLDIR2
Dim vHeader2
Dim vFooter2
Dim leftColumn2

TMPLDIR2="/templates/bb/tmplRich/"
vHeader2="/templates/bb/tmplRich/home_base_header.html"
vFooter2="/templates/bb/tmplRich/home_base_footer.html"
leftColumn2="/templates/bb/tmplRich/home_base_leftColumn.html"


'response.write("<b>Due to maintenance the system will be offline momentarily at 10:55 AM today. <BR>Please finalize your order before 10:55 AM Eastern Time.</b><BR>Current Eastern Time: " & now())
Dim  DonsCount
    DonsCount = DonsCount + 1 
Dim FB
Dim vOriginalPageScript
Dim vOPSBottom
FB=""
vOriginalPageScript = "<script>" & Chr(13)
vOriginalPageScript = vOriginalPageScript & "function utmx_section(){}function utmx(){}" & Chr(13)
vOriginalPageScript = vOriginalPageScript & "(function(){var k='2170435499',d=document,l=d.location,c=d.cookie;function f(n){" & Chr(13)
vOriginalPageScript = vOriginalPageScript & "if(c){var i=c.indexOf(n+'=');if(i>-1){var j=c.indexOf(';',i);return c.substring(i+n." & Chr(13)
vOriginalPageScript = vOriginalPageScript & "length+1,j<0?c.length:j)}}}var x=f('__utmx'),xx=f('__utmxx'),h=l.hash;" & Chr(13)
vOriginalPageScript = vOriginalPageScript & "d.write('<sc'+'ript src=""'+" & Chr(13)
vOriginalPageScript = vOriginalPageScript & "'http'+(l.protocol=='https:'?'s://ssl':'://www')+'.google-analytics.com'" & Chr(13)
vOriginalPageScript = vOriginalPageScript & "+'/siteopt.js?v=1&utmxkey='+k+'&utmx='+(x?x:'')+'&utmxx='+(xx?xx:'')+'&utmxtime='" & Chr(13)
vOriginalPageScript = vOriginalPageScript & "+new Date().valueOf()+(h?'&utmxhash='+escape(h.substr(1)):'')+" & Chr(13)
vOriginalPageScript = vOriginalPageScript & "'"" type=""text/javascript"" charset=""utf-8""></sc'+'ript>')})();" & Chr(13)
vOriginalPageScript = vOriginalPageScript & "</script><script>utmx(""url"",'A/B');</script>" & Chr(13)


vOPSBottom = "<script>" & Chr(13)
vOPSBottom = vOPSBottom & "if(typeof(urchinTracker)!='function')document.write('<sc'+'ript src=""'+" & Chr(13)
vOPSBottom = vOPSBottom & "'http'+(document.location.protocol=='https:'?'s://ssl':'://www')+" & Chr(13)
vOPSBottom = vOPSBottom & "'.google-analytics.com/urchin.js'+'""></sc'+'ript>')" & Chr(13)
vOPSBottom = vOPSBottom & "</script>" & Chr(13)
vOPSBottom = vOPSBottom & "<script>" & Chr(13)
vOPSBottom = vOPSBottom & "try {" & Chr(13)
vOPSBottom = vOPSBottom & "_uacct = 'UA-6280466-1';" & Chr(13)
vOPSBottom = vOPSBottom & "urchinTracker(""/2170435499/test"");" & Chr(13)
vOPSBottom = vOPSBottom & "} catch (err) { }" & Chr(13)
vOPSBottom = vOPSBottom & "</script>" & Chr(13)



'        vMv
'        vPageNo
'response.write(server.urlencode("x""s"))
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
    if Session("RunOnce") = 2 Then checkSearch
sub checkSearch

   dim rs100
   if Session("RCount")>3 then exit sub
   if  vSection="search" and Session("RunOnce") = 2   then
       Set rs100 = Server.CreateObject("ADODB.Recordset")

       rs100.open getsearchsqlCnt, conn
       xx=0
       if not rs100.eof then
           xx = rs100.fields("cnt")
       end if
       rs100.close

       if xx = 0 then
            rs100.open "exec getitemsku '" & replace(vSearchTerm,"'","''") & "'", conn
            if rs100.eof then
                vSection=""
            end if
            rs100.close
       end if

    end if
    Session("RunOnce") = 1
    Session("RCount")= Session("RCount") + 1
   set rs100 = nothing
end sub
'if    Request.ServerVariables("REMOTE_ADDR")  = "69.127.248.96" then
'   response.Write    vSearchTerm  & "|<br>"
'   response.Write vSearchCategory & "|<br>"
'   response.Write session("searchterm") & "|<br>"
'end if

'if    Request.ServerVariables("REMOTE_ADDR")  = "69.127.249.205" then
 '  response.write "<hr>xxx|vSection=" & vSection & "|/|vDept=" & vDept & "|/|vItem=" & vItem & "|/|vManufacturer=" & vManufacturer & "|<hr>"
'   response.Write("R: " & Session("ReferredBy"))
'end if

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
'if    Request.ServerVariables("REMOTE_ADDR")  = "24.186.147.208" then
'   response.write "<hr>|vSection=" & vSection & "|/|vDept=" & vDept & "|/|vItem=" & vItem & "|/|vManufacturer=" & vManufacturer & "|<hr>"
'	'http://www.bicyclebuys.com/manufacturer/3TTT/barsstems
'	response.write(request.form("vSection"))

'end if
'Broken somewhere....can't find where adhock fix
'if vDept = "Trainer" then vDept = "indoortrainers"

    ' all pages use the search, so let's get that done right away
    vSearchPage = getsearch()


'if    Request.ServerVariables("REMOTE_ADDR")  = "173.52.75.153" then
'   response.write "<hr>xxx|vSection=" & vSection & "|/|vDept=" & vDept & "|/|vItem=" & vItem & "|/|vManufacturer=" & vManufacturer & "|<hr>"
'end if

    ' now we begin page processing
    Select Case vSection
      Case ""     '  home page

         getMostPopularProducts()
         getNewProducts()
         'this should be closeouts...
         vOUT9 = getcloseouts("closeouts")
         'popular category listing
         vOUT10 = getpopcategories
         'most popular category w/ subcats
         vOUT11 = hpmostpop
%>
<!--#INCLUDE file="includes/moving.asp"-->
<%

'if    Request.ServerVariables("REMOTE_ADDR")  = "10.0.1.85" then
'	response.write(TESTIT)
'end if
         with objTemplate
            .TemplateFile = TMPLDIR2 & "home_base.html"
            .AddToken "closeouts", 1, vOUT9
            .AddToken "popcat", 1, vOUT10
            .AddToken "mostpopcat", 1, vOUT11
            .AddToken "header", 3, vHeader2
            .AddToken "search_section", 1, vSearchPage
	        .AddToken "moving", 1, (moving & getSlider(1))
            .AddToken "footer", 3, vFooter2
            .AddToken "leftColumn", 3, leftColumn2
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

		    'get the manufacturer links
		    vOUT5 = ""  : vOUT6 = "" : vOUT7 = "" : vOUT8 = "" : vOUT9 = ""
		    getmfglinks vSection   ' -- I don't think this mfg list is going to work here.. too early and we know too little
		    'to show products on the next refresh...

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
             'xxxxxx

            'if vSection="indoortrainers" and vDept="Trainer" and vItem="" and vManufacturer="" then
            '   response.write (vOriginalPageScript)
            'end if
            getNonSearchPages("")

            vItem=""
         ElseIf vSection = "search" Then
            ''''' begin search
            ' for product search
            'getSearchHTML()
            getNonSearchPages("Search")
        ElseIf vItem <> "" Then
    	 	if (vSection = "item") then
			    vSQL100 = "SELECT top 1 J.NavType FROM vwWebproducts P   WITH (NOLOCK), JohnWebNavType J  WITH (NOLOCK) WHERE (P.SKU LIKE '" & vItem & "') AND ((J.WebTypes LIKE '%' + CAST(P.WebTypeID AS nvarchar(20)) + '%') OR (J.SubCats LIKE '' + CAST(P.SubCatID AS nvarchar(20)) + ',%') OR (J.SubCats LIKE '%,' + CAST(P.SubCatID AS nvarchar(20)) + ',%') OR (J.SubCats LIKE '%' + CAST(P.SubCatID AS nvarchar(20)) + ''))  for browse"
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
            if vSection<>"xxx" and mTracking="" then checkItemUrls()
            buildImageDimensions("I")
            getRecentlyViewed()
'zzzzz
            with objTemplate

'if oProd1.pfields.Item("qtyonhand") > 0 then
	'vSavings = vSavings &  "<br><b>in stock</b>" 
'else
	'vSavings = vSavings &  "<br><b>available for order</b>" 
'end if


if oProd1.pfields.Item("qtyonhand") > 0 and oProd1.pfields.Item("WebNote")<>"8" then
	vStock = vStock &  "<img src=""/images/b_instock.gif"" alt=""In Stock"">" 
else
	vStock = vStock &  "<br/><b>available for order</b>" 
end if
                if mTracking ="goog" or  mTracking ="googUK" then
             	    .TemplateFile = TMPLDIR2 & "productGoog.html"
             	    .AddToken "title", 1, oProd1.pfields.Item("description") & " | Search Results"
             	    vProdMeta = vProdMeta & " | Search Results"
               else
            	    .TemplateFile = TMPLDIR2 & "productLeft.html"
            	    .AddToken "title", 1,  oProd1.pfields.Item("description") & FB
            	end if 

               .AddToken "itemdesc", 1, vDesc
               .AddToken "itemdesc2", 1, vDesc2
               .AddToken "itemname", 1, vItemDesc
               .AddToken "itemimage", 1, vItemImageOut1
               .AddToken "itembrand", 1, vBrand
               .AddToken "imagewidth", 1, vImageWidth
               .AddToken "origprice", 1, vOrigPrice
               .AddToken "savings", 1, vSavings
               .AddToken "stock", 1, vStock
               
               .AddToken "mDiscountType", 1, oProd1.pfields.Item("aDiscountType")
               .AddToken "mDiscountAmount", 1, oProd1.pfields.Item("dollarDiscountAmount")
               .AddToken "mSpecialPricing", 1, oProd1.pfields.Item("mSpecialPricing")
               if  oProd1.pfields.Item("upc") <> "" then
               	.AddToken "upc", 1, "UPC: " & oProd1.pfields.Item("upc")
               	end if
               	if oProd1.pfields.Item("mpn") <> "" then
              	 .AddToken "mpn", 1, "Manufacturer Part Number: " &  oProd1.pfields.Item("mpn")
              	 end if
    	       .AddToken "mTracking", 1, mTracking
    	       if vSavings<>"" then
                .AddToken "YourPrice",1,"SALE "
               end if

               if (oProd1.pfields.Item("WebNoteID") <> 15) then
               		.AddToken "price", 1,  formatcurrency(oProd1.pfields.Item("price"), 2, 0, 0)
               else
               		.AddToken "price", 1, formatcurrency(oProd1.pfields.Item("RetailWebPrice"), 2, 0, 0)
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
               if (oProd1.pfields.Item("WebNoteID") <> 15) and Not(oProd1.pfields.item("RebateCode") = "N" or oProd1.pfields.item("RebateCode")="")  then
               		.AddToken "WebNote", 1,  vwebnote & " "
               elseif oProd1.pfields.Item("WebNote") = 15 then
               .AddToken "WebNote", 1,  " <div class=product_notesReb><div class=button><a href=""javascript:void(0)"" onClick=""window.open('/rebate_price.asp?SKU=" & oProd1.pfields.Item("SKU") & "', 'BikePopUp',  'width=520,height=400,toolbar=0,scrollbars=1,screenX=50,screenY=50,left=50,top=50')"">Instant Rebate will Apply During Checkout</a></div></div>" 
                elseif oProd1.pfields.Item("WebNote") = 7 then
               .AddToken "WebNote", 1,  " <div class=""product_notesDisc"">" & oProd1.pfields.Item("Caption") & "</div>"
                elseif oProd1.pfields.Item("WebNote") = 2 then
               .AddToken "WebNote", 1,  " <div class=""product_notesDisc"">" & oProd1.pfields.Item("Caption") & "</div>"
               else
                end if
           vWebNote = ""     
        if oProd1.pfields.Item("FreeFreight") = True then
			vFreeFreight = -1
		   vWebNote = vWebNote & "<div class=""product_freefreight"">(Free Shipping with FEDEX Ground-US Mail!)</div>"
		   .AddToken "vWebNote", 1,  vWebNote & " "
		   Else
			vFreeFreight = 0
		End If

if Request.ServerVariables("Https")="off"  then
    FB ="<iframe src=""http://www.facebook.com/plugins/like.php?href=www.bicyclebuys.com/item/" & vItem & "/FB" & "&amp;layout=standard&amp;show_faces=false&amp;width=150&amp;action=like&amp;colorscheme=light&amp;height=80"" scrolling=""no"" frameborder=""0"" style=""border:none; overflow:hidden; width:150px; height:80px;"" allowTransparency=""true""></iframe>"
else
    FB ="<iframe src=""https://www.facebook.com/plugins/like.php?href=www.bicyclebuys.com/item/" & vItem & "/FB" & "&amp;layout=standard&amp;show_faces=false&amp;width=150&amp;action=like&amp;colorscheme=light&amp;height=80"" scrolling=""no"" frameborder=""0"" style=""border:none; overflow:hidden; width:150px; height:80px;"" allowTransparency=""true""></iframe>"
end if
               'else
               '     FB ="<iframe src=""http://www.facebook.com/plugins/like.php?href=www.bicyclebuys.com/" & vSection & "/" & vDept & "/" & vItem & "/FB" & "&amp;layout=standard&amp;show_faces=false&amp;width=250&amp;action=like&amp;colorscheme=light&amp;height=80"" scrolling=""no"" frameborder=""0"" style=""border:none; overflow:hidden; width:450px; height:80px;"" allowTransparency=""true""></iframe>"
               'end if
               if    Request.ServerVariables("REMOTE_ADDR")  <> "173.52.75.153" then
                'response.Write FB
               end if
               .AddToken "FB",1,FB
               vHeader = TMPLDIR2 & "item_base_header.html"
               vOUT9 = getcloseouts("closeouts")
               getTop3Related(oProd1)

	           vProdMeta = newHeader(vItemDesc, oProd1.pfields.Item("marketingdescription"), -1)
               .AddToken "prodmeta",1,  vProdMeta
               .AddToken "featured", 1, vOUT11
               .AddToken "closeouts", 1, vOUT9
               .AddToken "relatedproducts", 1, vOUT105
		        if not (instr(1,Request.ServerVariables("HTTP_USER_AGENT") ,"bot/")>0  )then
        	               .AddToken "previouslyviewed", 1, vOUT100
		        else
        	               .AddToken "previouslyviewed", 1, ""
		        end if
dim sizingChart
sizingChart = ""
'response.Write(oProd1.pfields.Item("ManuLink"))
'if    Request.ServerVariables("REMOTE_ADDR")  = "74.108.49.46" then
if InStr(oProd1.pfields.Item("description"),"Shoe") > 0  then
	if vBrand = "Shimano"  then
		sizingChart = "&nbsp;&nbsp;<img src=""/images/shimanoshoesize.jpg"" alt=""Sizing Chart"">"
	elseif vBrand = "Time"  then
		sizingChart = "&nbsp;&nbsp;<img src=""/images/TimeSizingChart.jpg"" alt=""Sizing Chart"">"
	elseif vBrand = "Diadora" then
		sizingChart = "&nbsp;&nbsp;<img src=""/sizing/images/DiadoraSizing.gif"" alt=""Sizing Chart"">"
	elseif vBrand = "Cannondale" then
		sizingChart = "&nbsp;&nbsp;<img src=""/images/CannondaleShoeSize.jpg"" alt=""Sizing Chart"">"
	elseif vBrand = "Louis Garneau" then
		if oProd1.pfields.Item("discontinued")="No" or instr(oProd1.pfields.Item("SKU"),"7Part")>0 or instr(oProd1.pfields.Item("SKU"),"6Part")>0 or instr(oProd1.pfields.Item("SKU"),"5Part")>0  then
			sizingChart = "&nbsp;&nbsp;<img src=""/images/LouisGarneauShoeSize2010.jpg"" alt=""Sizing Chart"">"
		else
			sizingChart = "&nbsp;&nbsp;<img src=""/images/LouisGarneauShoeSize2010.jpg"" alt=""Sizing Chart"">"
		end if
	elseif vBrand = "Pearl Izumi" then
		if InStr(oProd1.pfields.Item("description"),"Vip") + InStr(oProd1.pfields.Item("description"),"Vag") + InStr(oProd1.pfields.Item("description"),"vap") + InStr(oProd1.pfields.Item("description"),"vortex") > 0  then
			sizingChart = "&nbsp;&nbsp;<img src=""/sizing/images/PIMenshoe.gif"" alt=""Sizing Chart"">"
		else
			sizingChart = "&nbsp;&nbsp;<img src=""/images/PearlIzumi2.jpg"" alt=""Sizing Chart"">"
		end if
	elseif vBrand = "SIDI" then
		sizingChart = "&nbsp;&nbsp;<img src=""/images/SidiSizingSmall.jpg"" alt=""Sizing Chart"">"
	elseif vBrand = "Mavic" then
		sizingChart = "&nbsp;&nbsp;<img src=""/images/MavicSizingChart.jpg"" alt=""Sizing Chart"">"
	elseif vBrand = "Giro" then
		sizingChart = "&nbsp;&nbsp;<img src=""/images/GiroSizingChart.jpg"" alt=""Sizing Chart"">"
	elseif  oProd1.pfields.Item("ManuLink")<>"" then
			sizingChart = "&nbsp;&nbsp;<img src=""/sizing/images/DiadoraSizing.gif"" alt=""Sizing Chart"">"
	else
		sizingChart = "&nbsp;&nbsp;<img src=""/sizing/images/DiadoraSizing.gif"" alt=""Sizing Chart"">"
	end if
end if
'end if

               .AddToken "linktitle", 1, vOUT3

               .AddToken "breadcrumb", 1, vOUT4
               .AddToken "categories_col1", 1, vOUT1
               .AddToken "categories_col2", 1, vOUT2

            	.AddToken "header", 3, vHeader
            	.AddToken "search_section", 1, vSearchPage
            	.AddToken "footer", 3, vFooter
        		.AddToken "categories_col1", 1, vOUT101
                .AddToken "category_type", 1, vOUT102
                .AddToken "category_name", 1, vOUT103
		        .AddToken "alllink", 1, "/" & vSection
		        'if oProd1.pfields.Item("ManuLink") <> "" then
			    '    .AddToken "ManuLink", 1, "<DIV align=""right""><A href=""" & Server.HTMLEncode(oProd1.pfields.Item("ManuLink")) & """target=""_blank"">Manufacturer info</a></div>"
		        'end if
 		.AddToken "sizingChart",1,sizingChart
            	.parseTemplateFile
            end with
         end if
         set objTemplate = nothing
      End Select

      response.write(TESTIT &  chr(13) & "</script>")
      getbrandoptsJS()

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


'if vSection="indoortrainers" and (vDept="Trainer" or  vDept = "indoortrainers") and vItem="" and vManufacturer="" then
'	response.write (vOPSBottom )
'end if

'if    Request.ServerVariables("REMOTE_ADDR")  = "69.127.248.96" then
'         response.Write("<BR>HTTP: " )
'         response.Write( Request.ServerVariables("HTTP_REFERER"))
'         response.Write("<BR>vReferer: " )
'         response.Write( vReferer )
'         response.Write("<BR>sess:ReferredBy:" )
'         response.Write( Session("ReferredBy") )
'         response.Write("<BR>Ses:Referer" )
'         response.Write( Session("Referer") )
'
'
'end if
%>
