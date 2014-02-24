<% 


   Sub getRequestStrings()
       vSection = Escape(Left(Request("c"), 40))
       vItem = Escape(Left(Request("i"), 40))
       vDept = Escape(Left(Request("d"), 40))
       vManufacturer = Escape(Left(Request("m"), 40))
       vPriceRange = Escape(Left(Request("price"), 40))
       numberperpage = Escape(Left(Request("numberperpage"), 40))
       pagenumber = Escape(Left(Request("pagenumber"), 40))
       vSubmit = Trim(Escape(Left(request("submit"), 4)))
       vSearchTerm = Trim(Escape(Left(Request("searchterm"), 100)))
       vSearchVendID = Trim(Escape(Left(Request("v"), 4)))
       vSearchCategory =  Trim(Escape(Left(Request("searchcategory"), 4)))
'if    Request.ServerVariables("REMOTE_ADDR")  = "207.237.100.71" then 
'   response.write "section=" & vSection & " item=" & vItem & " dept=" & vDept & " manu=" & vManufacturer
'end if

       'response.end 
       ' sometimes d has a slash at the end... has to do with manufacturer product lists
       if vDept <> "" then vDept = replace(vDept, "/", "")
   end sub
   
   sub getPageNumbers()
        if (numberperpage = "") then
          numberperpage = 10
        end if
        numberperpage = cint(numberperpage)
        if (pagenumber = "") then
	        pagenumber = 1
        end if
        pagenumber = cint(pagenumber)
        ' pagination...
        vMv = Escape(Left(Request("DIR"), 1))		' Direction
        vPageNo = Escape(Left(Request("p"), 2))	' Pagenumber
 
        ' if we have no pageno then start at 1
	    if vPageNo = "" then
		    vPageNo = 1
	    End If 	    
   end sub

    sub getSearchTerm()
         
        if cint(pagenumber)  >= 1 and request.servervariables("REQUEST_METHOD") = "GET" then 
            vSearchTerm = Session("searchterm")
            exit sub
        end if
        ' this sets the search term
        ' if they clicked submit and left the term blank, then it's really blank - dont pull it out of session
          If request.servervariables("REQUEST_METHOD") = "POST" then
	          If vSearchTerm = "" Then
		        Session("searchterm") = ""
	          End If
          end if

        if vSubmit = "SUBM"  And vSearchTerm = "" Then
          vSearchTerm = ""
        Else
          ' if its generally, no blank submission then we need to pull it out of a session var
          If vSearchTerm = "" Then
             vSearchTerm = Session("searchterm")
          End If
        End If
        '   response.write "searchterm|" & vSearchTerm & "|"
        '   response.write "submit|" & vSubmit & "|"
        '   response.write "searchcategory|" & vSearchCategory & "|"
        '   response.end
       ' this keeps track of the last subcat shown and pre-selects it
       vSearchCategory = Session("searchcategory")
       If vSearchCategory = "" Or vSearchCategory = "all" Then vSearchCategory = 0

    end sub  
    
    sub getNewProducts
        dim rsFields, vLoop, rs2
        Set rs2 = Server.CreateObject("ADODB.Recordset")
          '    -- display  "new products" listing
   '      vSQL = "SELECT TOP 8 p.*, H.ID, H.Sort, H.Col1_ProductID, H.Col2_ProductID, H.Col1_Original_Price, H.Col2_Original_Price," _
   '           & "H.Col1_SizeHTML, H.Col2_SizeHTML,H.Col1_PageLink, H.Col2_PageLink, H.Type, H.NavTypeID  " _
   '           & "FROM HTML_Special_SaleItems H " _
   '           & "INNER JOIN vwWebProducts p " _
   '           & "ON H.Col1_ProductID = p.ProdID " _
   '           & "WHERE H.Type=2 " _
    '          & "AND p.WebPosted LIKE 'yes' " _
    '          & "ORDER BY NEWID(), H.Sort  for browse"
    
         '  response.write vSQL
         vSQL = "exec getNewProducts"
         xx = 0
         rs2.open vSQL, Conn
         while not rs2.EOF
            vTMP4 = rs2("description")
            vTMP4 = Server.HTMLEncode(vTMP4)
            vOUT11 = vOUT11 & vbcrlf & vbcrlf & vbcrlf & "<a href=""/item/" & rs2("sku") & """><img src=""/productimages/" & rs2("picture") & """ border=""0"" alt=""" & vTMP4 & """ vspace=""10"" width=""80""></a><BR>" & vbcrlf _
                          & "<div class=""featuringtext""><a href=""/item/" & rs2("sku") & """>" & vTMP4 & "</a></div>" & vbcrlf _
                          & "<img name=""feature_divide"" src=""/images/feature_divide.gif"" width=""159"" height=""12"" border=""0"" alt=""""><BR>" & vbcrlf & vbcrlf
 
            xx = xx + 1
            oProd1.ClearItem
            Set rsFields = rs2.Fields
            'response.write "<hr>" & oRS1("WebNoteID") & "<hr>"
             if NOT rs2.EOF then             
                for vLoop = 0 to (rsFields.Count - 1)
                   ' response.write "<hr>" & rsFields.Item(vLoop).Name & "<br>" & rsFields.Item(vLoop).Value
                   oProd1.pfields.Add rsFields.Item(vLoop).Name, rsFields.Item(vLoop).Value
                next
            end if            
            vTMP1 = "newprod" & xx & "sku"
            objTemplate.AddToken vTMP1, 1, oProd1.pfields.Item("SKU")
            vTMP1 = "newprod" & xx & "name"
            objTemplate.AddToken vTMP1, 1, oProd1.pfields.Item("description")
            vTMP1 = "newprod" & xx & "image"
            objTemplate.AddToken vTMP1, 1, resizepic("/productimages/" & oProd1.pfields.Item("picture"), oProd1.pfields.Item("Width_Small"), oProd1.pfields.Item("Height_Small"))
            vTMP1 = "newprod" & xx & "price"
            vTMP2 = FormatCurrencyDiscount("", oProd1.pfields.Item("price"), oProd1.pfields.Item("aDiscount"))

            objTemplate.AddToken vTMP1, 1, vTMP2

            vTMP1 = "newcartinfo" & xx

			if oProd1.pfields.Item("FreeFreight") = True then
				vFreeFreight = -1
			   vWebNote = vWebNote & "<div class=""product_freefreight"">(Free Shipping with " & vFreeShippingMethod & "!)</div>"
			   Else
				vFreeFreight = 0
			End If
			if oProd1.pfields.Item("OverWeight") > 0 then
				vOverWeight = oProd1.pfields.Item("OverWeight") + 1
			'   vWebNote = vWebNote & "<div class=""product_freefreight"">(Overweight shipping costs apply!)</div>"
			else
				vOverWeight = 0
			End If

			vCP = int(oProd1.pfields.Item("IsChildorParentorItem"))
			if isnull(vCP) or vCP = "" then vCP = 0
			if vCP=1 Then			   
			   'vItemOptions = ShowOptions2(oProd1.pfields.Item("ProdID"),  oProd1.pfields.Item("description"),  oProd1.pfields.Item("SKU"),   oProd1.pfields.Item("price")) & "<BR>"
			   vItemOptions = ShowOptionsRS(rs2, (rs2.fields("pProdID").value),    oProd1.pfields.Item("description"),  oProd1.pfields.Item("SKU"),   oProd1.pfields.Item("price")) & "<BR>"
			   ITEMID_1 = "NOTINUSE"
			else
			   vItemOptions = ""
			   ITEMID_1 = "ITEMID"
			end if
			
         vOUT104 = "" _
			  & "<FORM METHOD=""post"" action=""https://www.bicyclebuys.com/addtocart/"">" _
			  & "<INPUT TYPE=""hidden"" name=""ITEMNAME"" value=""" & oProd1.pfields.Item("description") & """>" _
			  & "<INPUT TYPE=""hidden"" name=""PRICE"" value=""" & oProd1.pfields.Item("price") & """>" _
			  & "<INPUT TYPE=""hidden"" name=""Referer"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""Referer1"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""URL"" value=""" & "/item/" & oProd1.pfields.Item("SKU") & """>" _
			  & "<INPUT TYPE=""hidden"" name=""Parent"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""PID"" value=""" & oProd1.pfields.Item("ProdID") & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""FreeFreight"" VALUE=""" & vFreeFreight & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""OverWeightFlags"" VALUE=""" & vOverWeight & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""" & ITEMID_1 & """ VALUE=""" & oProd1.pfields.Item("SKU") & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""mDiscountType"" VALUE=""" & oProd1.pfields.Item("aDiscountType") & """>" _
      		  & "<INPUT TYPE=""hidden"" NAME=""mDiscountAmount"" VALUE=""" & oProd1.pfields.Item("aDiscount") & """>" _
      	      & "<INPUT TYPE=""hidden"" NAME=""mSpecialPricing"" VALUE=""" & oProd1.pfields.Item("mSpecialPricing") & """>" _
			  & vItemOptions & "<input name=""SUBMIT"" VALUE=""ADD"" type=image src=""images/addtocart.jpg"" alt=""View Cart"" width=""100"" height=""22"" border=0 style=""margin: 5px 0 0 0;""></div></TD>" _
              & "</FORM>"
            objTemplate.AddToken vTMP1, 1,  vOUT104


            vTMP1 = "newprod" & xx & "msrp"
            if oProd1.pfields.Item("msrp") > 0 then
              	  vTMP2 = "MSRP: " & formatcurrency(oProd1.pfields.Item("msrp"), 2, 0, 0) & "<BR>"
            else
               vTMP2 = ""
            end if
            objTemplate.AddToken vTMP1, 1, vTMP2
            if not rs2.eof then    
                if oProd1.pfields.Item("pProdID")=rs2.fields("pProdID") then            
                    rs2.movenext                
                end if
            end if
         wend
         rs2.close
	     '    -- display  "new hardware" listing
         'vSQL = "SELECT TOP 8 p.*" _
          '    & "FROM vwWebProducts p " _
           '   & "WHERE p.WebPosted LIKE 'yes' " _
			'  & "AND (p.WebTypeID = 161 " _
			 ' & "OR p.WebTypeID = 125 " _
'			  & "OR p.WebTypeID = 126 " _
'			  & "OR p.WebTypeID = 162 " _
'			  & "OR p.WebTypeID = 127) " _
 '             & "ORDER BY NEWID()  for browse "
  
        vSQL = "exec getNewHardware"
         ' response.write vSQL
         xx = 0
         rs2.open vSQL, Conn
         while not rs2.EOF
         
            vTMP4 = rs2("description")
            vTMP4 = Server.HTMLEncode(vTMP4)
            vOUT11 = vOUT11 & vbcrlf & vbcrlf & vbcrlf & "<a href=""/item/" & rs2("sku") & """><img src=""/productimages/" & rs2("picture") & """ border=""0"" alt=""" & vTMP4 & """ vspace=""10"" width=""80""></a><BR>" & vbcrlf _
                          & "<div class=""featuringtext""><a href=""/item/" & rs2("sku") & """>" & vTMP4 & "</a></div>" & vbcrlf _
                          & "<img name=""feature_divide"" src=""/images/feature_divide.gif"" width=""159"" height=""12"" border=""0"" alt=""""><BR>" & vbcrlf & vbcrlf
 

            xx = xx + 1
            oProd1.ClearItem
            Set rsFields = rs2.Fields
            'response.write "<hr>" & oRS1("WebNoteID") & "<hr>"
             if NOT rs2.EOF then             
                for vLoop = 0 to (rsFields.Count - 1)
                   ' response.write "<hr>" & rsFields.Item(vLoop).Name & "<br>" & rsFields.Item(vLoop).Value
                   oProd1.pfields.Add rsFields.Item(vLoop).Name, rsFields.Item(vLoop).Value
                next
            end if            
            
            
            vTMP1 = "newdrive" & xx & "sku"
            objTemplate.AddToken vTMP1, 1, oProd1.pfields.Item("SKU")
            vTMP1 = "newdrive" & xx & "name"
            objTemplate.AddToken vTMP1, 1, oProd1.pfields.Item("description")
            vTMP1 = "newdrive" & xx & "image"
            objTemplate.AddToken vTMP1, 1, oProd1.pfields.Item("picture")
            vTMP1 = "newdrive" & xx & "price"

            vTMP2 = FormatCurrencyDiscount("", oProd1.pfields.Item("price"), oProd1.pfields("aDiscount"))
            objTemplate.AddToken vTMP1, 1, vTMP2

            vTMP1 = "newdriveinfo" & xx

			if oProd1.pfields.Item("FreeFreight") = True then
				vFreeFreight = -1
			   vWebNote = vWebNote & "<div class=""product_freefreight"">(Free Shipping with " & vFreeShippingMethod & "!)</div>"
			   Else
				vFreeFreight = 0
			End If
			if oProd1.pfields.Item("OverWeight") > 0 then
				vOverWeight = oProd1.pfields.Item("OverWeight") + 1
			 '  vWebNote = vWebNote & "<div class=""product_freefreight"">(Overweight shipping costs apply!)</div>"
			else
				vOverWeight = 0
			End If

			vCP = int(oProd1.pfields.Item("IsChildorParentorItem"))
			if isnull(vCP) or vCP = "" then vCP = 0
			if vCP=1 Then
			   vItemOptions = ShowOptionsRS(rs2, (rs2.fields("pProdID").value),  oProd1.pfields.Item("description"),  oProd1.pfields.Item("SKU"),   oProd1.pfields.Item("price")) & "<BR>"
			   ITEMID_1 = "NOTINUSE"
			else
			   vItemOptions = ""
			   ITEMID_1 = "ITEMID"
			end if

         vOUT104 = "" _
			  & "<FORM METHOD=""post"" action=""https://www.bicyclebuys.com/addtocart/"">" _
			  & "<INPUT TYPE=""hidden"" name=""ITEMNAME"" value=""" & oProd1.pfields.Item("description") & """>" _
			  & "<INPUT TYPE=""hidden"" name=""PRICE"" value=""" & oProd1.pfields.Item("price") & """>" _
			  & "<INPUT TYPE=""hidden"" name=""Referer"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""Referer1"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""URL"" value=""" & "/item/" & oProd1.pfields.Item("SKU") & """>" _
			  & "<INPUT TYPE=""hidden"" name=""Parent"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""PID"" value=""" & oProd1.pfields.Item("ProdID") & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""FreeFreight"" VALUE=""" & vFreeFreight & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""OverWeightFlags"" VALUE=""" & vOverWeight & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""" & ITEMID_1 & """ VALUE=""" & oProd1.pfields.Item("SKU") & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""mDiscountType"" VALUE=""" & oProd1.pfields.Item("aDiscountType") & """>" _
      		          & "<INPUT TYPE=""hidden"" NAME=""mDiscountAmount"" VALUE=""" & oProd1.pfields.Item("aDiscount") & """>" _
      	                  & "<INPUT TYPE=""hidden"" NAME=""mSpecialPricing"" VALUE=""" & oProd1.pfields.Item("mSpecialPricing") & """>" _
			  & vItemOptions & "<input name=""SUBMIT"" VALUE=""ADD"" type=image src=""images/addtocart.jpg"" alt=""View Cart"" width=""100"" height=""22"" border=0 style=""margin: 5px 0 0 0;""></div></TD>" _
              & "</FORM>"
            objTemplate.AddToken vTMP1, 1,  vOUT104


            vTMP1 = "newdrive" & xx & "msrp"
            if oProd1.pfields.Item("msrp") > 0 then
               vTMP2 = "MSRP: " & formatcurrency(oProd1.pfields.Item("msrp"), 2, 0, 0) & "<BR>"
            else
               vTMP2 = ""
            end if
            objTemplate.AddToken vTMP1, 1, vTMP2

            if not rs2.eof then                
                if oProd1.pfields.Item("pProdID")=rs2.fields("pProdID") then
                    rs2.movenext                
                end if
            end if
         WEND
         rs2.close
         set rs2 = nothing       

    end sub
    
    sub getMostPopularProducts()
        dim rs1
        dim vSQL100
        Set rs1 = Server.CreateObject("ADODB.Recordset")
             ' get the home page items out of the mainpage table (most popular)
         vSQL ="exec getMainPage"
'          "SELECT prodID1 + ',' " _
'              & "FROM mainpage " _
 '             & "WHERE ID = 1 for Browse"
 '        ' get each product and put the product info into the template
 '        '  (currently only have 8 -- need 9)
 '           "SELECT top 1 p.*, vendor.vendor, webnotes.* " _
 '             & "FROM vwWebProducts p " _
 '             & "INNER JOIN webnotes " _
 '             & "ON webnotes.WebNoteID = p.webnote " _
 '             & "INNER JOIN Vendor " _
 '             & "ON vendor.vendid = p.vendid " _
 '             & "WHERE ProdID=" & vProdID & " For Browse"      
                 
         rs1.open vSQL, conn 
         xx=0
         while not rs1.eof
            xx = xx+1
            oProd1.ClearItem
           ' response.write "<hr>" & vTMP1 & ":" & rs1(vTMP1) & "|"
            oProd1.getFields(rs1)  
            vTMP1 = "prod" & xx & "sku"
            objTemplate.AddToken vTMP1, 1, oProd1.pfields.Item("SKU")
            vTMP1 = "prod" & xx & "name"
            objTemplate.AddToken vTMP1, 1, oProd1.pfields.Item("description")
            vTMP1 = "prod" & xx & "image"
            objTemplate.AddToken vTMP1, 1, resizepic("/productimages/" & oProd1.pfields.Item("picture"), oProd1.pfields.Item("Width_Small"), oProd1.pfields.Item("Height_Small"))
            vTMP1 = "prod" & xx & "price"

'            vTMP2 = formatcurrency(oProd1.pfields.Item("price"), 2, 0, 0)
            vTMP2 = FormatCurrencyDiscount("", oProd1.pfields.Item("price"), oProd1.pFields("aDiscount"))

            objTemplate.AddToken vTMP1, 1, vTMP2
            vTMP1 = "cartinfo" & xx


			if oProd1.pfields.Item("FreeFreight") = True then
				vFreeFreight = -1
			   vWebNote = vWebNote & "<div class=""product_freefreight"">(Free Shipping with " & vFreeShippingMethod & "!)</div>"
			   Else
				vFreeFreight = 0
			End If
			if oProd1.pfields.Item("OverWeight") > 0 then
				vOverWeight = oProd1.pfields.Item("OverWeight") + 1
			 '  vWebNote = vWebNote & "<div class=""product_freefreight"">(Overweight shipping costs apply!)</div>"
			else
				vOverWeight = 0
			End If

			vCP = int(oProd1.pfields.Item("IsChildorParentorItem"))
			if isnull(vCP) or vCP = "" then vCP = 0
			if vCP=1 Then
			   'vItemOptions = ShowOptions2(oProd1.pfields.Item("ProdID"),  oProd1.pfields.Item("description"),  oProd1.pfields.Item("SKU"),   oProd1.pfields.Item("price")) & "<BR>"
			   vItemOptions = ShowOptionsRS(rs1, (rs1.fields("pProdID").value),   oProd1.pfields.Item("description"),  oProd1.pfields.Item("SKU"),   oProd1.pfields.Item("price")) & "<BR>"
			   ITEMID_1 = "NOTINUSE"
			else
			   vItemOptions = ""
			   ITEMID_1 = "ITEMID"
			end if
			
         vOUT104 = "" _
   			  & "<FORM METHOD=""post"" action=""https://www.bicyclebuys.com/addtocart/"">" _
			  & "<INPUT TYPE=""hidden"" name=""ITEMNAME"" value=""" & oProd1.pfields.Item("description") & """>" _
			  & "<INPUT TYPE=""hidden"" name=""PRICE"" value=""" & oProd1.pfields.Item("price") & """>" _
			  & "<INPUT TYPE=""hidden"" name=""Referer"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""Referer1"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""URL"" value=""" & "/item/" & oProd1.pfields.Item("SKU") & """>" _
			  & "<INPUT TYPE=""hidden"" name=""Parent"" value="""">" _
			  & "<INPUT TYPE=""hidden"" name=""PID"" value=""" & oProd1.pfields.Item("ProdID") & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""FreeFreight"" VALUE=""" & vFreeFreight & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""OverWeightFlags"" VALUE=""" & vOverWeight & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""" & ITEMID_1 & """ VALUE=""" & oProd1.pfields.Item("SKU") & """>" _
			  & "<INPUT TYPE=""hidden"" NAME=""mDiscountType"" VALUE=""" & oProd1.pfields.Item("aDiscountType") & """>" _
      		  & "<INPUT TYPE=""hidden"" NAME=""mDiscountAmount"" VALUE=""" & oProd1.pfields.Item("aDiscount") & """>" _
      	      & "<INPUT TYPE=""hidden"" NAME=""mSpecialPricing"" VALUE=""" & oProd1.pfields.Item("mSpecialPricing") & """>" _
              & vItemOptions & "<input name=""SUBMIT"" VALUE=""ADD"" type=image src=""images/addtocart.jpg"" alt=""View Cart"" width=""100"" height=""22"" border=0 style=""margin: 5px 0 0 0;""></div></TD>" _
              & "</FORM>"

            objTemplate.AddToken vTMP1, 1,  vOUT104

            vTMP1 = "prod" & xx & "msrp"
            if oProd1.pfields.Item("MSRP") > 0 then
               if (oProd1.pfields.Item("MSRP") / oProd1.pfields.Item("price")) > 1.05 Then
              	 vTMP2 = "MSRP: " & formatcurrency(oProd1.pfields.Item("MSRP"), 2, 0, 0) & "<BR>"
              else
              	vTMP2 = ""
              end if
            else
               vTMP2 = ""
            end if
            objTemplate.AddToken vTMP1, 1, vTMP2
		'	vSQL100 = "SELECT top 1 J.NavType, J.WebDisplayForNavType FROM vwWebProducts P, JohnWebNavType J WHERE (P.SKU LIKE '" & oProd1.pfields.Item("SKU") & "') AND ((J.WebTypes LIKE '%' + CAST(P.WebTypeID AS nvarchar(20)) + '%') OR (J.SubCats LIKE '' + CAST(P.SubCatID AS nvarchar(20)) + ',%') OR (J.SubCats LIKE '%,' + CAST(P.SubCatID AS nvarchar(20)) + ',%') OR (J.SubCats LIKE '%' + CAST(P.SubCatID AS nvarchar(20)) + ''))  for browse"
		'	response.Write( vSQL100)
		'	rs100.open vSQL100, Conn
		'	if not rs100.EOF	then
				temp_1 = LCase(oProd1.pfields.Item("NavType"))
				temp_2 = oProd1.pfields.Item("WebDisplayForNavType")
		'	end if
		'	rs100.close
        

          '  if oProd1.pfields.Item("SortType") = "SUBCAT" Then
          '    vTMP2 = getcatlinksc(oProd1.pfields.Item("SubCatID"))
          '  else
          '     vTMP2 = getcatlinkwt(oProd1.pfields.Item("WebTypeID"))
          '  end if
          
            vTMP1 = "prod" & xx & "more"
            objTemplate.AddToken vTMP1, 1, lcase(temp_1)
            vTMP1 = "prod" & xx & "moretext"
            objTemplate.AddToken vTMP1, 1, temp_2
            if not rs1.eof then                
                if oProd1.pfields.Item("pProdID")=rs1.fields("pProdID") then
                    rs1.movenext                
                end if
            end if
         wend
         rs1.close
         set rs1 = nothing
    end sub 
       
    sub getMostPopularFour()
			dim pfields2
			dim rsFields2, vLoop
			dim oBB
			
            if instr(vNavTypes, vSection) then
               vOUT100 = getwebtypeids2(vSection)
            '    vSQL100 = "SELECT TOP 4 p.*,vendor.* FROM vwWebProducts p INNER JOIN Vendor ON vendor.vendid = p.vendid WHERE 1=1 AND webtypeid IN(" & vOUT100 & ") AND webposted LIKE 'yes' AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') ORDER BY NEWID() for browse"
                vSQL100 = "exec getMostPopularFour '" & vOUT100 & "', 0"
            else
                vOUT100 = getsubcatids2(vSection)
            '    vSQL100 = "SELECT TOP 4 p.*,vendor.* FROM vwWebProducts p INNER JOIN Vendor ON vendor.vendid = p.vendid WHERE 1=1 AND subcatid IN(" & vOUT100 & ") AND webposted LIKE 'yes' AND (IsChildorParentorItem='1' or IsChildorParentorItem='0' or IsChildorParentorItem='' or IsNull(IsChildorParentorItem,'')='') ORDER BY NEWID()  for browse"
                vSQL100 = "exec getMostPopularFour '" & vOUT100 & "', 1"
            end if
            
            'response.Write(vSQL100)
			set oBB = new bb_product
			counter = 0
			rs2.open vSQL100, Conn
		    while not rs2.EOF
				'response.write( "ZZZ" & rs2.fields("SKU"))
			    
			    'set pfields2 = createobject("Scripting.Dictionary") 
		        '    pfields2.CompareMode = 1
			    'Set rsFields2 = rs2.Fields
			    oBB.clearitem
                oBB.getFields(rs2)
                set pfields2 = oBB.pfields 
                'pfields2.Add  rsFields2.Item("ProdID").Name, rsFields2.Item("ProdID").Value       		    
			    'pfields2.Add  rsFields2.Item("SKU").Name, rsFields2.Item("SKU").Value
			    'pfields2.Add  rsFields2.Item("WebTypeID").Name, rsFields2.Item("WebTypeID").Value
			    'pfields2.Add  rsFields2.Item("SubCatID").Name, rsFields2.Item("SubCatID").Value
			    'pfields2.Add  rsFields2.Item("VendID").Name, rsFields2.Item("VendID").Value
                 
				
				if (counter = 2) then
					vOUT103 = vOUT103 & " </TR><TR>"
					vOUT103 = vOUT103 & vOUT102
					vOUT103 = vOUT103 & " </TR><TR>"
					vOUT102 = ""
				end if
				vOUT103 = vOUT103 & "<TD class=""tiny"" align=center valign=top><a href=""/item/" & rs2("SKU") & """><img src=""" & resizepic("/productimages/" & rs2("picture"), rs2("Width_Small"), rs2("Height_Small")) & """ border=""0""></a></td>"
				vOUT102 = vOUT102 & "<TD class=""popularfoot"" align=center valign=top>"
				vOUT102 = vOUT102 & "<a href=""/item/" & rs2("SKU") & """>" & rs2("description") & "</a>" & "<BR><span class=""price"">YOU PAY: "
				vOUT102 = vOUT102 & FormatCurrencyDiscount("<BR>On Special", rs2("price"), pfields2.item("aDiscount"))    
				vOUT102 = vOUT102 & "</span><br><a href=""/item/" & rs2("SKU") & """>MORE INFO</a><br />"
				vOUT102 = vOUT102 & ""

				if rs2("FreeFreight") = True then
					vFreeFreight = -1
				   Else
					vFreeFreight = 0
				End If
				if rs2("OverWeight") > 0 then
					vOverWeight = rs2("OverWeight") + 1
				else
					vOverWeight = 0
				End If

				vCP = int(rs2("IsChildorParentorItem"))
				if isnull(vCP) or vCP = "" then vCP = 0
				if vCP=1 Then
				   'vItemOptions = ShowOptionsRS(rs2("ProdID"),  rs2("description"),  rs2("SKU"),  rs2("price")) & "<BR>"
				   vItemOptions =ShowOptionsRS(rs2, (rs2.fields("pProdID").value),    oBB.pfields.Item("description"),  oBB.pfields.Item("SKU"),   oBB.pfields.Item("price")) & "<BR>"
				   ITEMID_1 = "NOTINUSE"
				else
				   vItemOptions = ""
				   ITEMID_1 = "ITEMID"
				end if
			 vOUT104 = "" _
				  & "<FORM METHOD=""post"" action=""https://www.bicyclebuys.com/addtocart/"">" _
				  & "<INPUT TYPE=""hidden"" name=""ITEMNAME"" value=""" & pfields2.Item("description") & """>" _
				  & "<INPUT TYPE=""hidden"" name=""PRICE"" value=""" & pfields2.Item("price") & """>" _
				  & "<INPUT TYPE=""hidden"" name=""Referer"" value="""">" _
				  & "<INPUT TYPE=""hidden"" name=""Referer1"" value="""">" _
				  & "<INPUT TYPE=""hidden"" name=""URL"" value=""" & "/item/" & pfields2.Item("SKU") & """>" _
				  & "<INPUT TYPE=""hidden"" name=""Parent"" value="""">" _
				  & "<INPUT TYPE=""hidden"" name=""PID"" value=""" & pfields2.Item("ProdID") & """>" _
				  & "<INPUT TYPE=""hidden"" NAME=""FreeFreight"" VALUE=""" & vFreeFreight & """>" _
				  & "<INPUT TYPE=""hidden"" NAME=""OverWeightFlags"" VALUE=""" & vOverWeight & """>" _
				  & "<INPUT TYPE=""hidden"" NAME=""" & ITEMID_1 & """ VALUE=""" & pfields2.Item("SKU") & """>" _
				  & vItemOptions & "<input name=""SUBMIT"" VALUE=""ADD"" type=image src=""/images/addtocart.jpg"" alt=""View Cart"" width=""100"" height=""22"" border=0 style=""margin: 5px 0 0 0;""></div>" _
				  & "<INPUT TYPE=""hidden"" NAME=""mDiscountType"" VALUE=""" &   pfields2.Item("aDiscountType") & """>" _
      			  & "<INPUT TYPE=""hidden"" NAME=""mDiscountAmount"" VALUE=""" & pfields2.Item("aDiscount") & """>" _
	      	      & "<INPUT TYPE=""hidden"" NAME=""mSpecialPricing"" VALUE=""" & pfields2.Item("mSpecialPricing") & """>" _
				  & "</FORM>"
 
				vOUT102 = vOUT102 & vOUT104 & "</td>"
				'vOUT102 = vOUT102 &  "</td>"

				counter = counter + 1
                if not rs2.eof then    
                    if oBB.pfields.Item("pProdID")=rs2.fields("pProdID") then            
                        rs2.movenext                
                    end if
                end if
            wend
			
			rs2.close
        
			vOUT103 = vOUT103 & " </TR><TR>"
			vOUT103 = vOUT103 & vOUT102
			vOUT103 = vOUT103 & " </TR>"



               with objTemplate
               	.TemplateFile = TMPLDIR & "interior_catpop.html"
                  .AddToken "category_type", 1, vOUT3
                  .AddToken "category_name", 1, vOUT2
                  .AddToken "breadcrumb", 1, vOUT4
                  .AddToken "mostpopular", 1, vOUT103
                  .AddToken "categories_col1", 1, vOUT1
                  .AddToken "mfg_col1", 1, vOUT5
                  .AddToken "alllink", 1, vSection
                  .AddToken "featured", 1, vOUT9
                  .AddToken "blurb", 1, vBlurb
               	.AddToken "header", 1, vOUT10
               	.AddToken "search_section", 1, vSearchPage
               	.AddToken "footer", 3, vFooter
               	.parseTemplateFile
               end with
               
    end sub        
 
    sub getNonSearchPages(mSearch)
       
			 if (vSection = "closeouts" OR vSection = "newitems") then
				' getcatlinks vSection
				 vOUT100 = vOUT1
				 vOUT1 = ""
				' getmfglinks vSection
				 vOUT101 = vOUT5
				 vOUT5 = ""
				 getprodlinks2 vDept, vSection, vManufacturer, mSearch
			 elseif (vSection = "allmfg") then
				if vDept = "" then
					'vDept = "%"
					vOUT1 = ""
					vOUT5 = ""
					vOUT101 = ""
					vOUT100 = ""
				else
					 getcatlinks vDept
					 vOUT100 = vOUT1
					 vOUT1 = ""
					 getmfglinks vDept
					 vOUT101 = vOUT5
					 vOUT5 = ""
				end if
				 getprodlinks2 vDept, vSection, vManufacturer, mSearch
			else
				 getcatlinks vSection
				 vOUT100 = vOUT1
				 vOUT1 = ""	
				 getmfglinks vSection
				 vOUT101 = vOUT5
				 vOUT5 = ""
				 getprodlinks2 vDept, vSection, vManufacturer, mSearch
			end if


            vOUT9 = getfeatured("")
			if (vSection = "closeouts") then
				vUDept = "Closeouts"
			elseif (vSection = "newitems") then
				vUDept = "New Products"
			end if


			showonlybrands = showonlybrands & "Products per page: <select name=""shownum"" onchange=""MM_jumpMenu('parent',this,0)"">"
			if (numberperpage = 5) then
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 5 & "&pagenumber=" & 1 & """ selected>5</option>"
			else
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 5 & "&pagenumber=" & 1 & """>5</option>"
			end if
			if (numberperpage = 10) then
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 10 & "&pagenumber=" & 1 & """ selected>10</option>"
			else
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 10 & "&pagenumber=" & 1 & """>10</option>"
			end if
			if (numberperpage = 20) then
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 20 & "&pagenumber=" & 1 & """ selected>20</option>"
			else
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 20 & "&pagenumber=" & 1 & """>20</option>"
			end if
			if (numberperpage = 30) then
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 30 & "&pagenumber=" & 1 & """ selected>30</option>"
			else
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 30 & "&pagenumber=" & 1 & """>30</option>"
			end if
			if (numberperpage = 50) then
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 50 & "&pagenumber=" & 1 & """ selected>50</option>"
			else
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 50 & "&pagenumber=" & 1 & """>50</option>"
			end if
			if (numberperpage = 1000) then
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 1000 & "&pagenumber=" & 1 & """ selected>ALL</option>"
			else
				showonlybrands = showonlybrands & "<option value=""/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" & vPriceRange & "&numberperpage=" & 1000 & "&pagenumber=" & 1 & """>ALL</option>"
			end if

			showonlybrands = showonlybrands & "</select> "
			showonlybrands = showonlybrands & " <select name=""showonly"" onchange=""MM_jumpMenu('parent',this,0)"">"
			showonlybrands = showonlybrands & "<option value="""" selected=""selected"">Show only...</option><option value=""""></option>"

			if ((vDept <> "" AND vSection <> "" AND vManufacturer <> "") OR (vSection = "newitems") OR (vSection = "closeouts")) then
				 showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=&price=" &  "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>All prices</option>"
				 if vPriceCount.Item("100") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=&price=100" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$0 - $99.99 (" &  vPriceCount.Item("100") & ")" & "</option>"
				 End if
				 if vPriceCount.Item("500") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=&price=500" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$100 - $499.99 (" &  vPriceCount.Item("500") & ")" & "</option>"
				 End if
				 if vPriceCount.Item("1000") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=&price=1000" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$500 - $999.99 (" &  vPriceCount.Item("1000") & ")" & "</option>"
				 End if
				 if vPriceCount.Item("2000") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=&price=2000" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$1000 - $1999.99 (" &  vPriceCount.Item("2000") & ")" & "</option>"
				 End if
				 if vPriceCount.Item("3000") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=&price=3000" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$2000 - $2999.99 (" &  vPriceCount.Item("3000") & ")" & "</option>"
				 End if
				 if vPriceCount.Item("more") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=&price=more" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$3000 or more (" &  vPriceCount.Item("more") & ")" & "</option>"
				 End if
			else
				 showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=" &  "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>All prices</option>"
				 if vPriceCount.Item("100") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=100" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$0 - $99.99 (" &  vPriceCount.Item("100") & ")" & "</option>"
				 End if
				 if vPriceCount.Item("500") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=500" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$100 - $499.99 (" &  vPriceCount.Item("500") & ")" & "</option>"
				 End if
				 if vPriceCount.Item("1000") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=1000" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$500 - $999.99 (" &  vPriceCount.Item("1000") & ")" & "</option>"
				 End if
				 if vPriceCount.Item("2000") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=2000" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$1000 - $1999.99 (" &  vPriceCount.Item("2000") & ")" & "</option>"
				 End if
				 if vPriceCount.Item("3000") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=3000" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$2000 - $2999.99 (" &  vPriceCount.Item("3000") & ")" & "</option>"
				 End if
				 if vPriceCount.Item("more") > 0 Then
					showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=more" & "&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>$3000 or more (" &  vPriceCount.Item("more") & ")" & "</option>"
				 End if
			end if

			showonlybrands = showonlybrands & "<option value=""""></option>"

			if (vDept = "" AND vSection = "allmfg") then
				showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & vManufacturer & "&price=&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>All brands</option>"
			else
				showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=&price=&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>All brands</option>"
				a=vMFGName.Items
				b=vMFG.Items
				for i = 0 To vMFGName.Count -1
						showonlybrands = showonlybrands & "<option value=""" & "/?c=" & vSection & "&d=" & replace(vDept, "'", "\'") & "&m=" & replace(a(i), " ", "_") & "&price=&numberperpage=" & numberperpage & "&pagenumber=" & 1 & """>" & a(i) & " (" & b(i) & ")" & "</option>"
				next

			end if
			showonlybrands = showonlybrands & "</select>"

            With objTemplate
             if vSection = "search" then
                .TemplateFile = TMPLDIR &    "productlist.html"
				'.AddToken "showonly", 1, showonlybrands
				.AddToken "pagenav", 1, pagenavout
			 elseif (vSection = "closeouts" OR vSection = "newitems" OR (vSection = "allmfg" AND vDept = "" AND vManufacturer <> "")) then
			 	.TemplateFile = TMPLDIR & "productlistfeat_new.html"
				.AddToken "subcategory_name", 1, replace(vManufacturer, "_", " ")
				.AddToken "showonly", 1, showonlybrands
				.AddToken "pagenav", 1, pagenavout
			 else
			 	.TemplateFile = TMPLDIR & "productlistfeat.html"
				.AddToken "subcategory_name", 1, vUDept
				.AddToken "showonly", 1, showonlybrands
				.AddToken "pagenav", 1, pagenavout
			 end if

               .AddToken "breadcrumb", 1, vOUT4
               .AddToken "productlist", 1, vOUT1
			   .AddToken "categories_col1", 1, vOUT100

			   if (vSection = "allmfg") then
			   		.AddToken "alllink", 1, "/" & vDept
			   else
			   		.AddToken "alllink", 1, "/" & vSection
			   end if
               .AddToken "mfg_col1", 1, vOUT101
				.AddToken "category_type", 1, vOUT3
				.AddToken "category_name", 1, vOUT2
               .AddToken "featured", 1, vOUT9
               .AddToken "pricefilteroptions", 1, vOUT5
               .AddToken "mfgfilteroptions", 1, vOUT6
            	.AddToken "header", 3, vHeader
            	.AddToken "search_section", 1, vSearchPage
            	.AddToken "footer", 3, vFooter

            	.parseTemplateFile
            End With

    end sub    
    
    function getSearchSQL
           'response.write vSection
            dim firstWord
            firstWord = ""
            ' this gets rid of the "oops" messages for out of stock/discontinued items
            vInSearch = True

            ' if we're not moving page to page then vMv will be blank
            ' and we've already begun a search...
            ' otherwise we'll need to get the searchterm and searchcategory from the form
            
            if cint(pagenumber) = 1 and request.servervariables("REQUEST_METHOD") = "POST"  Then

               vSearchTerm = trim(replace(reqform("searchterm"), "'", ""))
               vSearchVendID = replace(Trim(request("v")), 4, "")
               vSearchCategory = reqform("searchcategory")

				if vSearchTerm = "Keyword Search" then
					vSearchTerm = ""
				end if

               if vSearchCategory = "" then vSearchCategory = 0
               if vSearchCategory = "all" then vSearchCategory = 0
               Session("searchterm")= vSearchTerm
            ' since we're in a search already, get the searchterm and searchcategory out of the session
            Else
               vSearchTerm = Session("searchterm")
               vSearchVendID = Session("searchvendid")
               vSearchCategory = Session("searchcategory")
            End If

'response.write "<hr> msc:" & vsearchcategory & "<hr>"

            ' begin search processing

            ' the vendorstring repository that we created in global.asa
            vVendorString = Application("VendorString")

            ' set the search vendor id to "0" as a default
            If vSearchVendID = "" Then vSearchVendID = "0"

            ' turn all the user entered terms into an array
            vTerms = split(vSearchTerm, " ")

            ' clear sql vars
            vSQL = ""
            vSQLVL = ""

            ' build the elements of the sql search, one term at a time
            For Each vTerm in vTerms
                
               ' short search terms are skipped right over, bigger than 1 char
                if len(vTerm) > 1 Then
                    ' is this term a vendor name?
                    vIsVend = instr(vVendorString, "|" & lcase(vTerm) & "|")

                    ' why yes, yes it is a vendor name.  We need to handle a vendor name differently.
                    if vIsVend Then
                        '  the search term endpoint within the large vendorstring repository is...
                        vSearchTermEndInVS = (vIsVend + Len(vTerm) + 1 + 1)  ' +1 +1??

                        '  the vendor id endpoint within the vendorstring repository
                        vVendIDEndInVS = instr(vSearchTermEndInVS, vVendorString, "|")

                        '  the actual cut out vendor id from the vendorstring repository
                        vSearchVendID = mid(vVendorString, vSearchTermEndInVS, vVendIDEndInVS - vSearchTermEndInVS)

                        '                  response.write vSearchTermEndInVS & "<br>"
                        '                  response.write vVendIDEndInVS & "<br>"
                        '                  response.write Len(vVendorString) & "<br>"
                        '                  response.write mid(vVendorString, vSearchTermEndInVS, vVendIDEndInVS - vSearchTermEndInVS)

                        ' One last check to make sure we actually got
                        ' a vendor id then set the SQL statement addition
                        if Len(vSearchVendID) > 0 Then vVendSQL = " AND (p.Vendid = " & vSearchVendID & ")"

                    ' not a vendor id, just it's a regular search...
                    Else
                    ' add the searchterm words to the sql
		     
                        If vTermCount > 0 then
                        vSQL = vSQL & " AND "
                        Else
                        vSQL = vSQL & " ("
                        end if
                        vTermCount = vTermCount + 1

                        '              vSQL = vSQL & "(description LIKE '%" & vTerm & "%' AND "
                        '              vSQL = vSQL & "marketingdescription LIKE '%" & vTerm & "%') "
                        '              vSQL = vSQL & "(description LIKE '%" & vTerm & "%' or "
                        vSQL = vSQL & "(description LIKE '%" & vTerm & "%') "
                    End If
                End If
                if vTermCount = 1 then
                    firstWord = vTerm
                end if
            Next

            ' if we added searchterm words to the sql already, close the sql condition and end it to add more
            if vTermCount > 0 then vSQL = vSQL & ") AND "

            ' add the parentchild and webposted check
   			vSQL = vSQL & " (IsChildorParentorItem='1' or IsChildorParentorItem='0')"
   			vSQL = vSQL & " AND (webposted LIKE 'Yes') "
	    
            ' if we're showing a specific webtype then add it to sql here
            if vSearchCategory <> 0 Then
               vSQL = vSQL & " AND (WebTypeID = " & vSearchCategory & ") "
            End If

            ' if there's a vendor search, add it on here
   			if vVendSQL <> "" Then vSQL = vSQL & vVendSQL

            ' Make sure we close the recordset if it's open.
	        if rs1.state = 1 then rs1.close

            ' start the vendor listing sql string
            if vSearchTerm <>"" then 
		        vSQL = "(" & vSQL & ") or (SKU LIKE '%" & vSearchTerm & "%' "	
 		        vSQL = vSQL & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0')"
   		        vSQL = vSQL & " AND webposted LIKE 'Yes' )"
	        end if
            'response.write vSQL


            vSQLVL = "SELECT   vendid FROM vwWebProducts p WHERE " & vSQL & " GROUP BY vendid for browse"
            'response.write(vSQLVL)
            ' if we're not already doing a vendor search, then we should
            ' dig out a vendor listing from the result set
            if vSearchVendID = "0" then
               
                Set vVendlist = Server.CreateObject("ADODB.Recordset")
                ' if vDebug then response.write vSQLVL & "<BR>"
                vVendlist.open vSQLVL, Conn,3
                if Not vVendlist.EOF Then                  
                    vVST = ""
                    ' Response.write "<font id=""bodylg"">You can choose a manufacturer to help narrow your search:<br>" & vIsVend
                        Do While not vVendlist.EOF
                            ' if this is blank then this is the first time through the loop
                            ' if vVST <> "" then Response.write "|&nbsp;"

                            ' force it blank every time through
                            vVST = ""

                            ' if the user entered a search term make sure it's put back into the
                            ' search if they choose a vendor.
                            vSDX = vVendlist("vendid") & ""
                            if vSearchTerm <> "" then
                                vVST = "&searchterm=" & Server.URLEncode(Replace(vVendorListingSD.Item(vSDX), " ", "") & " " & vSearchTerm)
                                if vSearchCategory <> 0 then vVST = vVST & "&searchcategory=" & vSearchCategory

                            ' otherwise just search on the category (if one was selected)
                            Else
                                vVST = "&searchterm=" & Server.URLEncode(Replace(vVendorListingSD.Item(vSDX), " ", ""))
                                if vSearchCategory <> 0 then vVST = vVST & "&searchcategory=" & vSearchCategory
                            End if
                            'response.write "<a href=""Items01.asp?NavID=search" & vVST & """>"
                            'response.write vVendorListingSD.Item(vSDX) & "</a>&nbsp;"
                            vVendlist.movenext
                        Loop
                        vVendlist.Close
                  'response.write "<br><br><br>"
               End If
            End If

            ' now we do the main product sql query
         '   vSQL = "SELECT top 100 p.*, vendor.* " _
            
            vSQL = " set nocount on  if object_id('tempdb..#tmp3') is not null drop table #tmp3 SELECT top 100 " _ 
                & "p.ProdID, p.description, p.SizeID , p.ColorID, p.price, p.SKU, " _ 
                & "p.picture, Width_Small, Height_Small,FreeFreight,OverWeight,IsChildorParentorItem, " _ 
                & "MSRP,  -1 msort, p.ProdID  pProdID,  aDaysLeft, aDiscount, aDiscountType, DT, WebNoteID, NoteType, Caption " _ 
                & ", marketingdescription, marketdescriptwo, webTypeID, vendor.Vendor, vendor.vendID, -1 sort,RetailWebPrice " _ 
                 & ",SubCategory, left(description,10) aSort  into #tmp3 FROM vwWebProducts p " _
                 & "INNER JOIN Vendor " _
                 & "ON vendor.vendid = p.vendid " _
                 & "INNER JOIN Subcategory ON p.SubCatID = Subcategory.SubCatID " _   
                 & "left outer  join  webnotes w on w.WebNoteID =p.webnote " _                           
                 & "WHERE " & vSQL
           vSQL = vSQL & " Insert into #tmp3 SELECT p.ProdID, p.description, p.SizeID , p.ColorID, p.price, p.SKU, " _   
                & " p.picture, p.Width_Small, p.Height_Small,p.FreeFreight,p.OverWeight,p.IsChildorParentorItem, " _   
                & " p.MSRP,  1 sort,   pProdID,  p.aDaysLeft, p.aDiscount, p.aDiscountType, p.DT, WebNoteID, NoteType, Caption " _   
                & " , p.marketingdescription, p.marketdescriptwo, p.webTypeID, v.Vendor, v.vendID, isnull(s.sort,99),p.RetailWebPrice,t.SubCategory, aSort  " _   
                & " FROM  [Products Children] pc " _   
                & " INNER JOIN  vwWebProducts p LEFT OUTER JOIN [Size] s ON p.SizeID = s.SizeID  ON pc.ChildProdID = p.ProdID  " _   
                & " INNER JOIN #tmp3 t on pc.prodID = t.ProdID left outer  JOIN Vendor V ON V.vendid = p.vendid   " _   
                & " WHERE   ((p.Discontinued LIKE 'yes' AND p.QtyAvailable > 0) or (p.Discontinued LIKE 'no'))  "     
           vSQL = vSQL & " Select * from #tmp3 "
            'response.Write(firstWord)
           if firstWord <> "" then
                vSQL = vSQL & " order by case when SubCategory LIKE '%" & trim(firstWord) & "%' then 0 else 1 end ,aSort, pProdID,msort, Sort drop table #tmp3 " 
           else
                vSQL = vSQL & " order by  aSort,pProdID,msort, Sort drop table #tmp3 " 
           end if
           ' vSQL = vSQL & " order by case when Subcategory.SubCategory LIKE '%" & vSearchTerm & "%' then 0 else 1 end"
          ' response.write vSQL & "<HR>v: " & vsearchvendid & "<hr>"
          getSearchSQL = vSQL   
          

    end function
    
    sub getSearchHTML
'NO LONGER USED
            'response.write vSection
            dim firstWord
            firstWord = ""
            ' this gets rid of the "oops" messages for out of stock/discontinued items
            vInSearch = True

            ' if we're not moving page to page then vMv will be blank
            ' and we've already begun a search...
            ' otherwise we'll need to get the searchterm and searchcategory from the form
            if vMv = "" Then

               vSearchTerm = trim(replace(reqform("searchterm"), "'", ""))
               vSearchVendID = replace(Trim(request("v")), 4, "")
               vSearchCategory = reqform("searchcategory")

				if vSearchTerm = "Keyword Search" then
					vSearchTerm = ""
				end if

               if vSearchCategory = "" then vSearchCategory = 0
               if vSearchCategory = "all" then vSearchCategory = 0

            ' since we're in a search already, get the searchterm and searchcategory out of the session
            Else
               vSearchTerm = Session("searchterm")
               vSearchVendID = Session("searchvendid")
               vSearchCategory = Session("searchcategory")
            End If

'response.write "<hr> msc:" & vsearchcategory & "<hr>"

            ' begin search processing

            ' the vendorstring repository that we created in global.asa
            vVendorString = Application("VendorString")

            ' set the search vendor id to "0" as a default
            If vSearchVendID = "" Then vSearchVendID = "0"

            ' turn all the user entered terms into an array
            vTerms = split(vSearchTerm, " ")

            ' clear sql vars
            vSQL = ""
            vSQLVL = ""

            ' build the elements of the sql search, one term at a time
            For Each vTerm in vTerms
                
               ' short search terms are skipped right over, bigger than 1 char
                if len(vTerm) > 1 Then
                    ' is this term a vendor name?
                    vIsVend = instr(vVendorString, "|" & lcase(vTerm) & "|")

                    ' why yes, yes it is a vendor name.  We need to handle a vendor name differently.
                    if vIsVend Then
                        '  the search term endpoint within the large vendorstring repository is...
                        vSearchTermEndInVS = (vIsVend + Len(vTerm) + 1 + 1)  ' +1 +1??

                        '  the vendor id endpoint within the vendorstring repository
                        vVendIDEndInVS = instr(vSearchTermEndInVS, vVendorString, "|")

                        '  the actual cut out vendor id from the vendorstring repository
                        vSearchVendID = mid(vVendorString, vSearchTermEndInVS, vVendIDEndInVS - vSearchTermEndInVS)

                        '                  response.write vSearchTermEndInVS & "<br>"
                        '                  response.write vVendIDEndInVS & "<br>"
                        '                  response.write Len(vVendorString) & "<br>"
                        '                  response.write mid(vVendorString, vSearchTermEndInVS, vVendIDEndInVS - vSearchTermEndInVS)

                        ' One last check to make sure we actually got
                        ' a vendor id then set the SQL statement addition
                        if Len(vSearchVendID) > 0 Then vVendSQL = " AND (p.Vendid = " & vSearchVendID & ")"

                    ' not a vendor id, just it's a regular search...
                    Else
                    ' add the searchterm words to the sql
		     
                        If vTermCount > 0 then
                        vSQL = vSQL & " AND "
                        Else
                        vSQL = vSQL & " ("
                        end if
                        vTermCount = vTermCount + 1

                        '              vSQL = vSQL & "(description LIKE '%" & vTerm & "%' AND "
                        '              vSQL = vSQL & "marketingdescription LIKE '%" & vTerm & "%') "
                        '              vSQL = vSQL & "(description LIKE '%" & vTerm & "%' or "
                        vSQL = vSQL & "(description LIKE '%" & vTerm & "%') "
                    End If
                End If
                if vTermCount = 1 then
                    firstWord = vTerm
                end if
            Next

            ' if we added searchterm words to the sql already, close the sql condition and end it to add more
            if vTermCount > 0 then vSQL = vSQL & ") AND "

            ' add the parentchild and webposted check
   			vSQL = vSQL & " (IsChildorParentorItem='1' or IsChildorParentorItem='0')"
   			vSQL = vSQL & " AND (webposted LIKE 'Yes') "
	    
            ' if we're showing a specific webtype then add it to sql here
            if vSearchCategory <> 0 Then
               vSQL = vSQL & " AND (WebTypeID = " & vSearchCategory & ") "
            End If

            ' if there's a vendor search, add it on here
   			if vVendSQL <> "" Then vSQL = vSQL & vVendSQL

            ' Make sure we close the recordset if it's open.
	        if rs1.state = 1 then rs1.close

            ' start the vendor listing sql string
            if vSearchTerm <>"" then 
		        vSQL = "(" & vSQL & ") or (SKU LIKE '%" & vSearchTerm & "%' "	
 		        vSQL = vSQL & " AND (IsChildorParentorItem='1' or IsChildorParentorItem='0')"
   		        vSQL = vSQL & " AND webposted LIKE 'Yes' )"
	        end if
            'response.write vSQL


            vSQLVL = "SELECT   vendid FROM vwWebProducts p WHERE " & vSQL & " GROUP BY vendid for browse"
            'response.write(vSQLVL)
            ' if we're not already doing a vendor search, then we should
            ' dig out a vendor listing from the result set
            if vSearchVendID = "0" then
               
                Set vVendlist = Server.CreateObject("ADODB.Recordset")
                ' if vDebug then response.write vSQLVL & "<BR>"
                vVendlist.open vSQLVL, Conn,3
                if Not vVendlist.EOF Then                  
                    vVST = ""
                    ' Response.write "<font id=""bodylg"">You can choose a manufacturer to help narrow your search:<br>" & vIsVend
                        Do While not vVendlist.EOF
                            ' if this is blank then this is the first time through the loop
                            ' if vVST <> "" then Response.write "|&nbsp;"

                            ' force it blank every time through
                            vVST = ""

                            ' if the user entered a search term make sure it's put back into the
                            ' search if they choose a vendor.
                            vSDX = vVendlist("vendid") & ""
                            if vSearchTerm <> "" then
                                vVST = "&searchterm=" & Server.URLEncode(Replace(vVendorListingSD.Item(vSDX), " ", "") & " " & vSearchTerm)
                                if vSearchCategory <> 0 then vVST = vVST & "&searchcategory=" & vSearchCategory

                            ' otherwise just search on the category (if one was selected)
                            Else
                                vVST = "&searchterm=" & Server.URLEncode(Replace(vVendorListingSD.Item(vSDX), " ", ""))
                                if vSearchCategory <> 0 then vVST = vVST & "&searchcategory=" & vSearchCategory
                            End if
                            'response.write "<a href=""Items01.asp?NavID=search" & vVST & """>"
                            'response.write vVendorListingSD.Item(vSDX) & "</a>&nbsp;"
                            vVendlist.movenext
                        Loop
                        vVendlist.Close
                  'response.write "<br><br><br>"
               End If
            End If

            ' now we do the main product sql query
         '   vSQL = "SELECT top 100 p.*, vendor.* " _
            
            vSQL = " set nocount on  if object_id('tempdb..#tmp3') is not null drop table #tmp3 SELECT top 100 " _ 
                & "p.ProdID, p.description, p.SizeID , p.ColorID, p.price, p.SKU, " _ 
                & "p.picture, Width_Small, Height_Small,FreeFreight,OverWeight,IsChildorParentorItem, " _ 
                & "MSRP,  -1 msort, p.ProdID  pProdID,  aDaysLeft, aDiscount, aDiscountType, DT, WebNoteID, NoteType, Caption " _ 
                & ", marketingdescription, marketdescriptwo, webTypeID, vendor.Vendor, vendor.vendID, -1 sort,RetailWebPrice " _ 
                 & ",SubCategory into #tmp3 FROM vwWebProducts p " _
                 & "INNER JOIN Vendor " _
                 & "ON vendor.vendid = p.vendid " _
                 & "INNER JOIN Subcategory ON p.SubCatID = Subcategory.SubCatID " _   
                 & "left outer  join  webnotes w on w.WebNoteID =p.webnote " _                           
                 & "WHERE " & vSQL
           vSQL = vSQL & " Insert into #tmp3 SELECT p.ProdID, p.description, p.SizeID , p.ColorID, p.price, p.SKU, " _   
                & " p.picture, p.Width_Small, p.Height_Small,p.FreeFreight,p.OverWeight,p.IsChildorParentorItem, " _   
                & " p.MSRP,  1 sort,   pProdID,  p.aDaysLeft, p.aDiscount, p.aDiscountType, p.DT, WebNoteID, NoteType, Caption " _   
                & " , p.marketingdescription, p.marketdescriptwo, p.webTypeID, v.Vendor, v.vendID, isnull(s.sort,99),p.RetailWebPrice,t.SubCategory " _   
                & " FROM  [Products Children] pc " _   
                & " INNER JOIN  vwWebProducts p LEFT OUTER JOIN [Size] s ON p.SizeID = s.SizeID  ON pc.ChildProdID = p.ProdID  " _   
                & " INNER JOIN #tmp3 t on pc.prodID = t.ProdID left outer  JOIN Vendor V ON V.vendid = p.vendid   " _   
                & " WHERE   ((p.Discontinued LIKE 'yes' AND p.QtyAvailable > 0) or (p.Discontinued LIKE 'no'))  "     
           vSQL = vSQL & " Select * from #tmp3 "
            'response.Write(firstWord)
           if firstWord <> "" then
                vSQL = vSQL & " order by case when SubCategory LIKE '%" & trim(firstWord) & "%' then 0 else 1 end, pProdID, Sort drop table #tmp3 " 
           else
                vSQL = vSQL & " order by pProdID, Sort drop table #tmp3 " 
           end if
           ' vSQL = vSQL & " order by case when Subcategory.SubCategory LIKE '%" & vSearchTerm & "%' then 0 else 1 end"
          ' response.write vSQL & "<HR>v: " & vsearchvendid & "<hr>"
          	
            getprodlist(vSQL)
            	
			vOUT9 = getfeatured("")

            With objTemplate
                .TemplateFile = TMPLDIR & "productlist.html"
                .AddToken "breadcrumb", 1, vOUT4
                '               .AddToken "subcategory_name", 1, vUDept
                .AddToken "productlist", 1, vOUT1
                .AddToken "pagenav", 1, vOUT2
                .AddToken "featured", 1, vOUT9
                '               .AddToken "pricefilteroptions", 1, vOUT5
                '               .AddToken "mfgfilteroptions", 1, vOUT6
		'.AddToken "pagenav", 1, pagenavout
                .AddToken "header", 3, vHeader
                .AddToken "search_section", 1, vSearchPage
                .AddToken "footer", 3, vFooter
                .parseTemplateFile
            End With
            response.end

'What is all of this below here for....

exit sub
            ' Get the item list
   			rs1.open vSQL & " For Browse" ,Conn,3

   			If Not rs1.EOF then

               ' let's do some pagination
   				rs1.PageSize = vListanum 

               ' if vMv equals something then we're moving within a result set
   				If vMv = vPrevious or vMv = vNext or vMv=vFirst or vMv= vLast Then
   					Select Case vMv
      					Case vFirst
      						vPageNo = 1
      					Case vLast
      						vPageNo = RS1.PageCount
      					Case vPrevious
      						If vPageNo > 1 Then
      							vPageNo = vPageNo - 1
      						Else
      							vPageNo = 1
      						End If
      					Case vNext
      						If RS1.AbsolutePage < RS1.PageCount Then
      							vPageNo = vPageNo + 1
      						Else
      							vPageNo = RS1.PageCount
      						End If

                     ' if moving within result set then we start at beginning
      					Case Else
      						vPageNo = 1
   					End Select
   				End If
   				RS1.AbsolutePage = vPageNo

               ' start showing product results
   				response.write "<FONT ID=""bodylger"">"
               response.write "Results: <BR></FONT><FONT ID=""body"">"
'zzz Maybe...
   				if vSearchVendID <> "" Then response.write "Manufacturers limited to: &quot;" & vVendorListingSD.Item(vSearchVendID) & "&quot;<BR>"
   				if vTermCount = 1 Then response.write "Search term: &quot;" & vSearchTerm & "&quot;<BR>"
   				if vTermCount > 1 Then response.write "Search terms: &quot;" & vSearchTerm & "&quot;<BR>"
   				if vSearchCategory > 0 Then response.write "Category: &quot;" & vWebTypeListingSD.Item(vSearchCategory) &  "&quot;<BR>"
   				response.write "</FONT><HR>"
   				ShowPageNav

   				For x = 1 to rs1.PageSize

   					If rs1.EOF Then
   						Exit For
   					End If
					tempProd.clearItem
					tempProd.getfields(rs1)
   					vItemPicture = rs1("picture")
   					if instr(vItemPicture, "\") <> -1 then vItemPicture = replace(vItemPicture, "\", "/")
   					vCP = int(rs1("IsChildorParentorItem"))
   					if isnull(vCP) or vCP = "" then vCP = 0
   					response.write "<hr>" & rs1("ProdID") & ", " & vItemPicture & ", " &  FormatCurrencyDiscount("<BR>On Special", tempProd.pfields.Item("price"), tempProd.pFields("aDiscount")) & ", " & rs1("description") & ", " & rs1("MarketingDescription") & ", " & rs1("MarketDescriptwo") & ", " & vCP & ", " & rs1("SKU")
   					', rs1, 0
   					'ShowProduct rs1("ProdID"), vItemPicture, rs1("price"), replace(rs1("description") & " ","""", "''"), rs1("MarketingDescription"), rs1("MarketDescriptwo"), vCP, rs1("SKU"), rs1, 0
   					rs1.movenext
   					If rs1.EOF Then
   						Exit For
   					End If
   				Next
   				ShowPageNav
   				rs1.Close

               ' Save the search criteria in session
   				Session("searchterm") = vSearchTerm
   				Session("searchmfg") = vSearchMFG
   				Session("searchcategory") = vSearchCategory

   				' Since we just displayed items we should clear the session variables
   				Session("M") = 0
   				Session("T") = 0
   				Session("NavID") = ""

   		   ' didn't find anything...
   			Else
         		response.write "<FONT ID=""body"">No items found!</FONT><BR>"
   				rs.Close
   			End If

    ''''''''''''''''''' end search    
    end sub
    
    sub getTop3Related(oProd1)
        dim rs2 
        Set rs2 = Server.CreateObject("ADODB.Recordset")
			vOUT105 = ""
			'''vSQL100 = "SELECT TOP 3 P.* FROM vwWebProducts P WHERE P.SKU <> '" & oProd1.pfields.Item("SKU") & "' AND P.WebTypeID = " & oProd1.pfields.Item("WebTypeID") & " AND webposted like 'yes' ORDER BY NEWID()   for browse"
			vSQL100="exec getTop3Related '" & oProd1.pfields.Item("SKU")   & "'," & oProd1.pfields.Item("WebTypeID") 
			'response.Write(vSQL100)
			rs2.open vSQL100, Conn
			do while not rs2.EOF
				tempProd.ClearItem
				tempProd.getfields(rs2)
				vOUT105 = vOUT105 & " <TR><TD class=""tiny""><img src=""/productimages/" & rs2("picture") & """ width=""35"" border=""0""></a></TD> <TD class=""tiny""><a href=""/item/" & rs2("SKU") & """>" & rs2("description") & "</a></TD><TD class=""tiny2"" style=""color: #FB6600"">" &  FormatCurrencyDiscount("<BR>On Special", tempProd.pfields.Item("price"), tempProd.pFields("aDiscount"))  & "</TD></TR>"
			rs2.movenext
   			loop
			rs2.close
       set rs2 = nothing
    end sub
    
    sub getRecentlyViewed()
        dim rs2 
        dim skus
			' get recently viewed items
			vOUT100 = Session("RecentlyViewed")
			 
			vRecArr = split(vOUT100, "|")
			vOUT100 = ""
			counter = 1
			
			skus=""
			do while vRecArr(counter) <> ""
			    skus = vRecArr(counter) & "," & skus 
			    counter = counter + 1
			loop
			if counter = 1 then 
			    vOUT100=""
			    exit sub
			 end if
			skus = left(skus,len(skus)-1)
				'vSQL100 = "SELECT top 1 P.SKU, p.Description FROM vwWebProducts P WHERE P.SKU in '" & vRecArr(counter) & "' AND webposted like 'yes' ORDER BY NEWID()  for browse"
				vSQL100 = "exec getRecentlyViewed '" & skus & "'"
				'response.Write(vSQL100)
				Set rs2 = Server.CreateObject("ADODB.Recordset")
				rs2.open vSQL100, Conn
				while not rs2.EOF	 
					vOUT100 = vOUT100 & "<img src=""images/orange-arrow.gif"" width=""10"" height=""9"" border=0> <a href=""/item/" & rs2("SKU") & """>" & rs2("description") & "</a><br>"
					rs2.movenext
				wend 
				counter = counter + 1
				rs2.close
				set rs2 = nothing
			
    end sub
    
    sub buildImageDimensions(where)
            ' build the proper larger image height and width
            dim tmp
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
                     vItemImageOut1 = "<A HREF=""javascript:win('/productimages/" & vItemPicture & "'," & vHL & ", " & vWL & ")"">"
         	      End If
         	      ' this is the flat image output
         	      if where = "I" then
                    tmp = resizepic2("", vWS, vHS, 225)
                  else
                    tmp =  " height=""" & vHS & """ width=""" & vWS & """"
                  end if
                  vItemImageOut1 = vItemImageOut1 & "<IMG class=""productimage"" SRC=""/productimages/" & vItemPicture & """" & tmp & " alt=""" & vItemDesc & """><br><img src=""/images/zoom.gif"" border=0 tag=zzz>"
                  ' if we're using the popopen, we need to end the href
                  if vHL <>  -1 then vItemImageOut1 = vItemImageOut1 & "</A>"
               End If
            End If
    end sub

    Function resizepic2(imageurl, w, h, max)
	    if (w = h) then
		    imageurl = imageurl & """ width=""" & max & """ height=""" & max & """"
	    elseif (w > h) then
		    imageurl = imageurl & """ width=""" & max & """"
	    else
		    imageurl = imageurl & """ height=""" & max & """"
	    end if
	    resizepic2 = imageurl
    End Function                

 
    sub getItemInformation()
            ' load in the product data	    
	         oProd1.getitemSKU(vItem)
             
            ' set the breadcrumb link
            if vSection <> "item" Then
               vTMP1 = UCase(Left(vSection,1)) & Right(vSection,Len(vSection)-1)
               vOUT4 = "<a href=""/" & vSection & "/"">" & vTMP1 & "</a> " _
                     & "&gt; <a href=""/" & vSection & "/" & vDept & "/"">" & vUDept & "</a> " _
                     & "&gt; " & oProd1.pfields.Item("description")
            end if

            vOrigPrice = ""
            vSavings = ""

            If (not(isNull(oProd1.pfields.Item("MSRP"))) AND (oProd1.pfields.Item("MSRP") <> "") AND IsNumeric(oProd1.pfields.Item("MSRP"))) Then
                vMSRP = oProd1.pfields.Item("MSRP")
                vPrice = oProd1.pfields.Item("price")
                ' no point in showing really low savings... over 1% and we show it
                if (vMSRP / vPrice) > 1.05 Then
                    vOrigPrice = "<div class=""minidesc"">MSRP:</div>" & FormatCurrency(vMSRP, 2, 0, 0) & "<BR>"

                    sPrice= 0
                   if   (oProd1.pfields.item("RebateCode") = "N" or oProd1.pfields.item("RebateCode")="")  then

                       sDiscount =oProd1.pfields.Item("aDiscount")
			            'response.write(oProd1.pfields.Item("SKU") & oProd1.pfields.Item("mDiscountAmount"))
	                    if sDiscount <> 0 then 'Dollar
	                        if oProd1.pfields.Item("aDiscountType")=-1 then
	                            sPrice =  sDiscount
                           else 'Percent
	                            sPrice = oProd1.pfields.Item("price") * sDiscount
                           end if
                        end if
						vSavings = "<div class=""minidesc""><span class=""product_save"">You Save:</span></div>" & FormatCurrency(vMSRP - vPrice + sPrice , 2, 0, 0) & "<BR>"
                   end if  
                end if
            end if
            

            vImageWidth = cInt(oProd1.pfields.Item("width_large"))

            if  vImageWidth > 250 then
               vImageWidth = 250
            end if

           
'           response.write "<hr>"  & oProd1.pfields.Item("ischildorparentoritem")
			vCP = int(oProd1.pfields.Item("ischildorparentoritem"))
			if isnull(vCP) or vCP = "" then vCP = 0
            if vCP=1 Then
               vItemOptions = ShowOptions(oProd1.pfields.Item("ProdID"),  oProd1.pfields.Item("description"),  oProd1.pfields.Item("SKU"),  oProd1.pfields.Item("price")) & "<BR>"
               objTemplate.AddToken "ITEMID_1", 1, "NOTINUSE"
            else
               vItemOptions = ""
               objTemplate.AddToken "ITEMID_1", 1, "ITEMID"
            end if

            vDesc = replace(oProd1.pfields.Item("marketingdescription"), "^", "<li>")
            if oProd1.pfields.Item("marketdescriptwo") <> "" Then
               vDesc = vDesc & "<hr width=""25%"">" & replace(oProd1.pfields.Item("marketdescriptwo"), "^", "<li>")
               ' vDesc2 = vDesc2 & " " & oProd1.pfields.Item("marketdescriptwo")
            end if


            call cycleviewed(oProd1.pfields.Item("description"), "/item/" & oProd1.pfields.Item("sku"))

            vBrand = oProd1.pfields.Item("vendor")
            vWebNote = ""
            if oProd1.pfields.Item("WebNoteID") <> 1 then
               vWebNote = "<div class=""product_notes"">" & oProd1.pfields.Item("Caption") & "</div>"
            end if

         	if oProd1.pfields.Item("FreeFreight") = True then
         		vFreeFreight = -1
               vWebNote = vWebNote & "<div class=""product_freefreight"">(Free Shipping with " & vFreeShippingMethod & "!)</div>"
        	   Else
         		vFreeFreight = 0
         	End If

         	if oProd1.pfields.Item("OverWeight") > 0 then
         		vOverWeight = oProd1.pfields.Item("OverWeight") + 1
              ' vWebNote = vWebNote & "<div class=""product_freefreight"">(Overweight shipping costs apply!)</div>"
         	else
         		vOverWeight = 0
         	End If

            vReferer = ""
            vReferer1 = ""

            vItemDesc = oProd1.pfields.Item("description")
			vItemDesc = replace(vItemDesc, """", "&quot;")

			if (oProd1.pfields.Item("WebTypeID") = "") then
				oProd1.pfields.Item("WebTypeID") = 000
			end if
			
			vItemDesc = replace(vItemDesc, """", "&quot;")
    
    end sub
   %>
   