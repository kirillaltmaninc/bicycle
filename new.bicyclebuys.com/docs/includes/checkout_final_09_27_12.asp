<!--#INCLUDE VIRTUAL="/includes/template_cls.asp" -->
<!--#INCLUDE VIRTUAL="/includes/common.asp" -->
<!--#INCLUDE VIRTUAL="/includes/cartconfig.asp" --><%
   ' get the template engine ready
   call zeroRebateArray()

   set objTemplate = new template_cls

   vOUT1 = ""
   vOUT2 = ""

   ' Session.Abandon
   vCartID = (Request.QueryString("CID")+0)
   if vCartID = 0 Then response.redirect("/")
   Cart.EmptyCart
   Cart.LoadCartDB vCartID

%><!--#INCLUDE VIRTUAL="/includes/cartdisplay_checkout.asp" --><%


      If Cart.Info.ShipCountry = "US" Then
          vShipTotal = FormatCurrency(Cart.Shipping,2,0,0)
      Else
          vShipTotal = "International<br>We'll call you!"
      End If

      vSalesTax = FormatCurrency(Cart.TotalTax,2,0,0)
      vGrandTotal = FormatCurrency(Cart.Total,2,0,0)      
      

      dim ot1
      set ot1 = new template_cls

      with ot1
      	.TemplateFile = TMPLDIR & "final_header.html"
         '.AddToken "CartID", 1, vCartID
        .AddToken "CartID", 1, Cart.OrderID
         vOut8 = .getParsedTemplateFile
      end with

      with objTemplate
      	.TemplateFile = TMPLDIR & "checkout_final.html"

			if (TotalDiscount15 > 0) then
				vOUT2 = vOUT2 & "<TR ><TD colspan=3 align=right class=cart style=""text-align: center; border-top: 1px solid #C4EAE6;"">&nbsp;</TD><TD align=right class=cart style=""text-align: center; border-top: 1px solid #C4EAE6;"">Discount</TD><TD colspan=5 align=right style=""text-align: center; border-top: 1px solid #C4EAE6;"" class=cart>-" & FormatCurrency(TotalDiscount15, 2, 0, 0) & "</TD></TR>"
			end if

         .AddToken "displaycart", 1, vOUT2

         .AddToken "thisserver", 1, vThisServer
         .AddToken "sessionreferer", 1, Session("Referer")

        ' .AddToken "cartid", 1, Cart.OrderID
	    .AddToken "cartid", 1, vCartID

         .AddToken "shiptotal", 1, vShipTotal
         .AddToken "salestax", 1, vSalesTax
         .AddToken "grandtotal", 1, vGrandTotal         
         .AddToken "adjusttotal", 1, rebateHtml(vAdjustTotal)                
         .AddToken "cgridtotal", 1, FormatCurrency(Cart.GridTotal,2,0,0)

         .AddToken "CartTotal", 1, formatcurrency(Cart.Total,2,0,0)

      	.AddToken "header", 1, vOut8
      	.AddToken "footer", 3, vFinalFooter
        .AddToken "PromoCode", 1, getRebates() 

'         .AddToken "", 1, Cart.Info.

      	.parseTemplateFile
      end with

   Else

      ' empty cart!
      with objTemplate
      	.TemplateFile = TMPLDIR & "displayemptycart.html"
      	.AddToken "header", 3, vCartHeader
      	.AddToken "footer", 3, vCartFooter

      	.parseTemplateFile
      end with

   End If

   ' we're done
   Session.Abandon

%>



<script type="text/javascript">
    var _gaq = _gaq || [];
    _gaq.push(['_setAccount', 'UA-6280466-2']);
    _gaq.push(['_trackPageview']);
    _gaq.push(['_addTrans',
    '<%=vCartID%>',           // order ID - required     
	'',  // affiliation or store name     
	'<%=vGrandTotal%>',          // total - required     
	'<%=vSalesTax%>',           // tax     
	'<%=vShipTotal%>',              // shipping     
	'',       // city     
	'',     // state or province     
	''             // country   
	]);


    <%For each Item in Cart.Items%>  
        _gaq.push(['_addItem',
        '<%=vCartID%>',           // order ID - required     
        '<%=Item.ItemID%>',           // SKU/code - required     
        '<%=Item.Name%>',        // product name     
        '',   // category or variation     
        '<%=Item.Price%>',          // unit price - required     
        '<%=Item.Quantity%>'               // quantity - required   
        ]);
	<%next%>

   

    _gaq.push(['_trackTrans']); //submits transaction to the Analytics servers    
    (function() {
        var ga = document.createElement('script');
        ga.type = 'text/javascript'; ga.async = true;
        ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
        var s = document.getElementsByTagName('script')[0];
        s.parentNode.insertBefore(ga, s);
    })();  
</script>
 <% dim gOrd %>   
    <%For each Item in Cart.Items %> 
       <% gOrd = gOrd + Item.Quantity %>
	<%next%>
<script language="javascript">
   
         <!--
                 /* Performance Tracking Data */
                 var mid            = '199034';
                 var cust_type      = '1';
                 var order_value = '<%=vGrandTotal%>';
                 var order_id = '<%=vCartID%>';
                 var units_ordered  = '<%=gOrd%>';
         //-->
</script>
<script language="javascript" src="https://www.shopzilla.com/css/roi_tracker.js"></script>
