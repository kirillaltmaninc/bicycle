<!--#INCLUDE VIRTUAL="/includes/template_cls.asp" -->
<!--#INCLUDE VIRTUAL="/includes/common.asp" -->
<!--#INCLUDE VIRTUAL="/includes/cartconfig.asp" -->
 
<%



    call zeroRebateArray()
    call getRebate()
    
    
 	Cart.Adjust = 0
 	Cart.SaveCart
    Session("Cart") = Cart.SaveCart
   ' get the template engine ready
   set objTemplate = new template_cls

   vOUT1 = ""
   vOUT2 = ""
'   For each item in Cart.Items
'      response.write "<hr>IC4: " & Item.Custom4 & "<br>"
'   next

   Cart.SetPropertyFormat "nameeditvalue"
   Cart.HandleCommands
   Cart.Calculate

%>
<!--#INCLUDE VIRTUAL="/includes/cartdisplay_displaycart.asp" -->
<% 

      ' Put checkout and empty buttons on the display
      Dim vCheckEmpty, pFields, RopCnt, vTMP4CH, descCH
      vCheckEmpty = ""
'response.write( vThisProto & vThisServer )
      if cart.gridtotalquantity > 0 then
         vCheckEmpty = "<a href=""" &  vThisProto & vThisServer & "/checkout/""><img src=""/cartimages/checkout.gif"" alt=""Check out"" border=""0"" WIDTH=""100"" HEIGHT=""20""></a>" _
                       & "<a href=""" &  vThisProto & vThisServer & "/emptycart/""><img src=""/cartimages/emptycart.gif"" alt=""Remove ALL items from the cart"" border=""0"" WIDTH=""100"" HEIGHT=""20""></a>"
      end if

	  vOUT105 = ""
      if  last_added_id <> ""  then 
	  		dim vCP 
			RopCnt = 0
			vOUT101 = ""
			'vSQL100 = "SELECT P.* FROM vwWebproducts P WHERE P.SKU like '" & last_added_id & "' For Browse"
			vSQL100 = "exec getTop3RelatedOp '" & last_added_id & "'"
			vOUT102 = ""
			vOUT103 = ""
			vOUT103 = vOUT103 & " <TR>"
			if (vOUT101 = "") then
				vOUT101 = 000
			end if

			rs2.open vSQL100, Conn
			while not rs2.EOF
			    if RopCnt = 5 then
				vOUT103 = vOUT103 & "</TR><TR>"
				vOUT103 = vOUT103 & vOUT102
				vOUT103 = vOUT103 & "</tr><tr><td colspan=5><hr width=""90%""></td></tr><tr>"
				vOUT102 = ""
				RopCnt = 0
			    end if 
			    tempProd.clearitem
			    tempProd.getfields(rs2)
			    set pfields = tempProd.pfields

				vTMP4CH = rs2("description")
				descCH = replace(vTMP4CH,"""","")
				descCH = replace(descCH,"/"," ")
				descCH = replace(descCH,"'","")
				descCH = replace(descCH,"%20","-")
				descCH =  lcase(replace(descCH," ","-"))

				vOUT103 = vOUT103 & "<TD class=""tiny"" align=center valign=top><a href=""/item/" & rs2("SKU") &"/" &descCH & """><img src=""" & resizepic("/productimages/" & rs2("picture"), rs2("Width_Small"), rs2("Height_Small")) & """ border=""0""></a></td>"
				vOUT102 = vOUT102 & "<TD class=""popularfoot"" align=center valign=top>"
				vOUT102 = vOUT102 & "<a href=""/item/" & rs2("SKU") &"/" &descCH & """>" & rs2("description") & "</a>" & "<BR><span class=""price"">YOU PAY: " & formatcurrency(rs2("price"), 2, 0, 0) & "</span><br><a href=""/item/" & rs2("SKU") &"/" &descCH & """>MORE INFO</a><br />"
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
				if isnull(vCP) or vCP = "" then
				     vCP = 0
				end if
				if vCP=1 Then
				   'vItemOptions = ShowOptions2(rs2("ProdID"),  rs2("description"),  rs2("SKU"),  rs2("price")) & "<BR>"
'if  Request.ServerVariables("REMOTE_ADDR")  <> "10.0.0.66" then
'				   vItemOptions = ShowOptionsRS(rs2, (pFields.item("pProdID")),   pFields.Item("description"),  pFields.Item("SKU"),   pFields.Item("price")) & "<BR>"
'else
				   vItemOptions = ShowOptionsShort(rs2, (pFields.item("pProdID")),   pFields.Item("description"),  pFields.Item("SKU"),   pFields.Item("price")) & "<BR>"
'end if		
		   		   ITEMID_1 = "NOTINUSE"
				else
				   vItemOptions = ""
				   ITEMID_1 = "ITEMID"
				end if

			    vOUT104 = "" _
				  & "<FORM METHOD=""post"" action=""/addtocart/"">" _
				  & "<INPUT TYPE=""hidden"" name=""ITEMNAME"" value=""" &  pFields.Item("description") & """>" _
				  & "<INPUT TYPE=""hidden"" name=""PRICE"" value=""" &  pFields.Item("price") - pFields.Item("dollarDiscountAmount") & """>" _
				  & "<INPUT TYPE=""hidden"" name=""Referer"" value="""">" _
				  & "<INPUT TYPE=""hidden"" name=""Referer1"" value="""">" _
				  & "<INPUT TYPE=""hidden"" name=""URL"" value=""" & "/item/" &  pFields.Item("SKU") & """>" _
				  & "<INPUT TYPE=""hidden"" name=""Parent"" value="""">" _
				  & "<INPUT TYPE=""hidden"" name=""PID"" value=""" &  pFields.Item("ProdID") & """>" _
				  & "<INPUT TYPE=""hidden"" NAME=""FreeFreight"" VALUE=""" & vFreeFreight & """>" _
				  & "<INPUT TYPE=""hidden"" NAME=""OverWeightFlags"" VALUE=""" & vOverWeight & """>" _
				  & "<INPUT TYPE=""hidden"" NAME=""" & ITEMID_1 & """ VALUE=""" &  pFields.Item("SKU") & """>" _
				  & vItemOptions & "<input name=""SUBMIT"" VALUE=""ADD"" type=image src=""/images/addtocart.jpg"" alt=""" & replace("Buy " & pFields.Item("description"),"""","'") & """ width=""100"" height=""22"" border=0 style=""margin: 5px 0 0 0;""></div></TD></FORM>"

				vOUT102 = vOUT102 & vOUT104 & "</td>"
				'vOUT102 = vOUT102 &  "</td>"
                if not rs2.eof then
                    if pfields.Item("pProdID")=rs2.fields("pProdID") then
                        rs2.movenext
                    end if
                end if
				RopCnt= RopCnt + 1
			wend
			rs2.close
			vOUT103 = vOUT103 & " </TR><TR>"
			vOUT103 = vOUT103 & vOUT102
			vOUT103 = vOUT103 & " </TR>"
      ' cart display built, now show it
      with objTemplate
      	.TemplateFile = TMPLDIR & "displaycart.html"
         .AddToken "breadcrumb", 1, vOUT1

		if (TotalDiscount15 > 0) then
			vOUT2 = vOUT2 & "<TR ><TD colspan=4 align=right class=cart style=""text-align: center; border-top: 1px solid #C4EAE6;"">&nbsp;</TD><TD align=right class=cart style=""text-align: center; border-top: 1px solid #C4EAE6;"">Discount</TD><TD colspan=5 align=right style=""text-align: center; border-top: 1px solid #C4EAE6;"" class=cart>-" & FormatCurrency(TotalDiscount15, 2, 0, 0) & "</TD></TR>"

		 	Cart.Adjust = "-" & TotalDiscount15
		 	Cart.SaveCart
		else
		 	Cart.Adjust = 0
		 	Cart.SaveCart
		end if

         .AddToken "displaycart", 1, vOUT2

         .AddToken "thisserver", 1, vThisServer
         .AddToken "checkempty", 1, vCheckEmpty

'         .AddToken "continueshopping", 1, Session("Referer")
		 .AddToken "continueshopping", 1, "/"

         '.AddToken "cgridtotal", 1, FormatCurrency(Cart.GridTotal, 2, 0, 0)
         .AddToken "cgridtotal", 1, (FormatCurrency((Cart.GridTotal-vAdjustTotal), 2, 0, 0))

      	.AddToken "date", 1, formatdatetime(now(), 3)
      	.AddToken "title", 1, "Bicyclebuys.com - Show My Shopping Cart"
      	.AddToken "images", 1, IMGDIR

      	.AddToken "header", 3, vCartHeader
      	vCartFooter =  TMPLDIR & "cart_footer_add.html"
      	.AddToken "footer", 3, vCartFooter
        vCartFooter =  TMPLDIR & "cart_footer.html"
		.AddToken "similarproducts", 1, vOUT103
	    .AddToken "HTMLLink",1,  vThisProto & vThisServer 
	    '.AddToken "adjusttotal", 1, rebateHtml(vAdjustTotal)      
	    .AddToken "PromoCode", 1, getRebates() 
      	.parseTemplateFile
      end with

   Else

      ' cart display built, now show it
      with objTemplate
      	.TemplateFile = TMPLDIR & "displayemptycart.html"
      	.AddToken "header", 3, vCartHeader      	
      	.AddToken "footer", 3, vCartFooter
      	.parseTemplateFile
      	
      end with

   End If

   ' Page is done. save and cleanup
   Session("Cart") = Cart.SaveCart
   Set Cart = Nothing

   Session("Referer") = Trim(Request.ServerVariables("HTTP_REFERER"))

   'from old site. not sure yet what it's for
   Session("NavID") = "" 

%>
