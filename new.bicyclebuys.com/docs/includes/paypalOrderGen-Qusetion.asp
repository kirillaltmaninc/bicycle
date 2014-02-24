



<%

  
				 g_Key = mid(g_KeyString,1,Len(Cart.Info.CCNumber))
				 Cart.Info.CCNumber = EnCrypt(Cart.Info.CCNumber)

				 
				 if Cart.CustomerID > 0 Then
					Cart.SaveCustomerDB(Cart.CustomerID)
				 Else
					Cart.SaveCustomerDB
				End If
				 'Cart.SaveOrderDB
				 'Cart.SaveItemsDB
				 'vCartID = Cart.SaveCartDB
		        
				'Call SendEmail()
		         

				 'if vDEBUGGING = 1 Then
					
				 'End if
				
dim my_paypal_email, return_page, notify_page, item_price, item_name, item_quantity, item_cod, currency_type, url, itemPPcount
				
my_paypal_email = "neil@bicyclebuys.com"                                                           
return_page = "https://www.bicyclebuys.com/includes/checkout_final.asp"
currency_type = "USD"  
'url = "https://api-3t.paypal.com/nvp"
url = "paypalBill.asp"
url = url & "?business=" & my_paypal_email
url = url & "&return=" & return_page
url = url & "&lName=" & Cart.Info.Custom3
url = url & "&fName=" & Cart.Info.Custom5
url = url & "&addr1=" & Cart.Info.Address1
url = url & "&city=" & Cart.Info.City
url = url & "&state=" & Cart.Info.StateProvince
url = url & "&zip=" & Cart.Info.ZipPostal
url = url & "&phone=" & Cart.Info.Phone
url = url & "&email=" & Cart.Info.Email
url = url & "&tax=" & Cart.TotalTax
url = url & "&country=" & Cart.Info.Country




itemPPcount=0

   For Each Item in Cart.Items
		itemPPcount = itemPPcount +1
		url = url & "&item_name_"&itemPPcount&"=" & Item.Name
		url = url & "&quantity_"&itemPPcount&"=" & Item.Quantity
		url = url & "&item_number_"&itemPPcount&"=" & Item.ItemID
		url = url & "&amount_"&itemPPcount&"=" & Item.Price
   Next
'url = url & "&amount=" & Cart.Total
url = url & "&no_shipping=" & Cart.Shipping
'url = url & "&currency_code=" & currency_type
Response.Redirect(url)

 
'Response.Redirect "https://www.paypal.com/cgi-bin/webscr"
	
			   'Response.Redirect "checkout_final.asp?CID=" & vCartID
			   'Response.Redirect "faxorder.asp"
			End If
%>