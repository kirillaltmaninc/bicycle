<!--#INCLUDE VIRTUAL="/includes/template_cls.asp" -->
<!--#INCLUDE VIRTUAL="/includes/common.asp" -->
<!--#INCLUDE VIRTUAL="/includes/cartconfig.asp" --><%
   ' get the template engine ready
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
        ' .AddToken "CartID", 1, vCartID
.AddToken "CartID", 1, Cart.OrderID
vOut8 = .getParsedTemplateFile
      end with

      with objTemplate
      	.TemplateFile = TMPLDIR & "printfriendly.html"
			if (TotalDiscount15 > 0) then
				vOUT2 = vOUT2 & "<TR ><TD colspan=3 align=right class=cart style=""text-align: center; border-top: 1px solid #C4EAE6;"">&nbsp;</TD><TD align=right class=cart style=""text-align: center; border-top: 1px solid #C4EAE6;"">Discount</TD><TD colspan=5 align=right style=""text-align: center; border-top: 1px solid #C4EAE6;"" class=cart>-" & FormatCurrency(TotalDiscount15, 2, 0, 0) & "</TD></TR>"
			end if

         .AddToken "displaycart", 1, vOUT2

         .AddToken "thisserver", 1, vThisServer
         .AddToken "sessionreferer", 1, Session("Referer")

         .AddToken "cartid", 1, vCartID

         .AddToken "shiptotal", 1, vShipTotal
         .AddToken "salestax", 1, vSalesTax
         .AddToken "grandtotal", 1, vGrandTotal
         .AddToken "cgridtotal", 1, FormatCurrency(Cart.GridTotal,2,0,0)

         .AddToken "CartTotal", 1, formatcurrency(Cart.Total,2,0,0)

      	.AddToken "header", 1, vOUT8
      	.AddToken "footer", 3, vFinalFooter

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
