<!--#include file="ccfuncs.asp"--><!--#include file="configvars_cs.inc"--><%

Sub SendEmail( )

	FileDSN = "Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webuserprod;Initial Catalog=BBC_PROD;Data Source=webserver"


	Set Conn = Server.CreateObject("ADODB.Connection")

	Conn.Open FileDSN


   sql = "UPDATE Payments SET ApprovalCode = '" & Cart.CC.AuthCode & "' WHERE Orderid = " & Cart.OrderID
	Conn.Execute(sql)

conn.close


   eheader = "BICYCLEBUYS.COM POS CUSTOMER ORDER ACKNOWLEDGEMENT" & vbcrlf & vbcrlf

   eheader = eheader & "POS Order Number: " & Left("0000000000", 10 - Len(Cart.OrderID)) & Cart.OrderID & vbcrlf
   eheader = eheader & "Date: " & Date & vbcrlf
   eheader = eheader & "Time: " & Time & vbcrlf & vbcrlf
   eheader = eheader & "Customer Name: " & Cart.Info.Name & vbcrlf & vbcrlf

   eheaderI = eheaderI & "POS Order Number: " & Left("0000000000", 10 - Len(Cart.OrderID)) & Cart.OrderID & vbcrlf
   eheaderI = eheaderI & "Date: " & Date & vbcrlf
   eheaderI = eheaderI & "Time: " & Time & vbcrlf

   ebillto = ebillto & "Bill-To:" & vbcrlf
   ebillto = ebillto & "  " & Cart.Info.Name &  vbcrlf
   ebillto = ebillto & "  " & Cart.Info.Address1 &  vbcrlf
   ebillto = ebillto & "  " & Cart.Info.City & ", " & Cart.Info.StateProvince & "   " & Cart.Info.ZipPostal & vbcrlf
   ebillto = ebillto & "  " & Cart.Info.Country & vbcrlf
   ebillto = ebillto & "  Ph: " & Cart.Info.Phone & vbcrlf
   ebillto = ebillto & "  " & Cart.Info.Email & vbcrlf & vbcrlf

	
   ebillto = ebillto & "Ship-To Address: " & vbcrlf
   ebillto = ebillto & "  " & Cart.Info.shipName &  vbcrlf
   ebillto = ebillto & "  " & Cart.Info.shipCompany &  vbcrlf
   ebillto = ebillto & "  " & Cart.Info.shipAddress1 &  vbcrlf
   ebillto = ebillto & "  " & Cart.Info.shipCity & ", " & Cart.Info.shipStateProvince & "   " & Cart.Info.shipZipPostal & vbcrlf
   ebillto = ebillto & "  " & Cart.Info.shipCountry & vbcrlf
   ebillto = ebillto & "  Ph: " & Cart.Info.shipPhone & vbcrlf
   ebillto = ebillto & "  " & Cart.Info.Email & vbcrlf


   epayment = ""
   epayment = epayment & "Payment: Credit Card" & vbcrlf
   epayment = epayment & "  Card Type: " & Cart.Info.CCType & vbcrlf & vbcrlf
   epayment = epayment & "  Card Holder: " & Cart.Info.CCName & vbcrlf
   epayment = epayment & "  Authorization Code: " & Cart.CC.AuthCode & vbcrlf

   ' This is the internal version of the epayment paragraph
   epaymentI = ""
   epaymentI = epaymentI & "Payment: " & Cart.Info.CCType & vbcrlf
   epaymentI = epaymentI & "  Card Number: " & Cart.Info.CCNumber & vbcrlf
   epaymentI = epaymentI & "  Card Expiration: " & Cart.Info.CCMonth & "/" & Cart.Info.CCYear & vbcrlf
   epaymentI = epaymentI & "  Card Holder: " & Cart.Info.CCName & vbcrlf
   epaymentI = epaymentI & "  Authorization Code: " & Cart.CC.AuthCode & vbcrlf
   epaymentI = epaymentI & "  AVS Code: " & vAVSResponseCode & vbcrlf
   epaymentI = epaymentI & "  AVS Message: " & vAVSResponseMessage & vbcrlf

   ettl = ""
   ettl = ettl & vbcrlf & "Total:                       " & formatcurrency(Cart.Total,2,0,0) & vbcrlf & vbcrlf

   efoot = ""
   efoot = efoot & "Please refer to this whenever contacting BicycleBuys.com customer" & vbcrlf
   efoot = efoot & "service. If you have any questions please just reply to this e-mail" & vbcrlf
   efoot = efoot & "or call 1-888-4-BIKE-BUY." & vbcrlf & vbcrlf
   efoot = efoot & "Thanks again for shopping with us." & vbcrlf
   efoot = efoot & "---------------------------------------------------------------------" & vbcrlf
   efoot = efoot & "BicycleBuys.com" & vbcrlf
   efoot = efoot & """We Cycle The World""" & vbcrlf
   efoot = efoot & "http://www.bicyclebuys.com" & vbcrlf

   ' We're going to save the order into a file on the web server for
   ' order fulfillment by the BB team
   vOrderFileName = vSaveOrderPath & "POS-Order-" & Left("0000000000", 10 - Len(Cart.OrderID)) & Cart.OrderID & ".txt"
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   Set vOrderFile = fso.OpenTextFile(vOrderFileName, 2, True)
   vOrderFile.WriteLine(vbcrlf & vbcrlf & eheaderI & ebillto & vbcrlf & epaymentI & ettl)
   vOrderFile.Close

   ' ---- Send Email to Customer
   Cart.Mail.Host = "localhost"
   Cart.Mail.From = "Sales@BicycleBuys.com"
   Cart.Mail.FromName = "BicycleBuys.com Sales"
   Cart.Mail.Subject = "BicycleBuys.com Store Receipt - POS Order Number: " & Left("0000000000", 10 - Len(Cart.OrderID)) & Cart.OrderID

   Cart.Mail.Body = eheader & ebillto & vbcrlf & epayment & eshipto & eitems & ettl & efoot
   
   ' Send it to the shipping email address too
   If len(Cart.Info.Email) <> 0 Then
    Cart.Mail.AddAddress Cart.Info.Email
    sendtrue = 1
   End If

   ' Send it to the shipping email address too if it's different
   If (len(Cart.Info.ShipEmail) <> 0) and (Cart.Info.Email <> Cart.Info.ShipEmail) Then
       Cart.Mail.AddAddress Cart.Info.ShipEmail
       sendtrue = 1
   End if

   If sendtrue = 1 Then
      'Cart.mail.addattachment "/orders/" & Cart.orderid & ".enc"
      Cart.Mail.Send ' send to buyer
   End if
   Cart.Mail.Reset

   response.write "<pre><font size=""3"">" & (vbcrlf & vbcrlf & eheaderI & ebillto & vbcrlf & epaymentI & ettl) & "</pre>"

End Sub


   Set Cart  = Server.CreateObject("iiscart2000.store")

   Cart.key = "lii"
   Cart.CurrencyFormat = "$,2"

   ' live or test skipjack
   If Request("ordernumber") <> "999999" Then
      ' live system
      Cart.cc.Processor = "skipjack"
      Cart.cc.Login = "000293293270"
      Cart.cc.host = "www.skipjackic.com"
   Else
      ' test system
      Cart.cc.Processor = "skipjack"
      Cart.cc.Login = "000785615111"
      Cart.cc.host = "developer.skipjackic.com"
      response.write "Developer ID used"
   End if

	Cart.cc.timeOut = 30
	Cart.cc.TransType = "capture"
   Cart.cc.logging = "../logs"
   Cart.StateTaxRate = "8.625%"
   Cart.CountryTaxRate = "0%"

'   Cart.cc.code = request.form("cvv2")
	   Cart.Info.Custom3 = request("custFN")
   Cart.Info.Custom5 = request("custLN")
   
   Cart.Info.Name = request("sjname")
   Cart.Info.Address1 = request("streetaddress")
   Cart.Info.Address2 = ""
   Cart.Info.City = request("city")

   'if request("province") = "" Then
   '   Cart.Info.StateProvince = request("state")
   'else
      Cart.Info.StateProvince = request("province")
   'end if   

   Cart.Info.ZipPostal = request("zipcode")
   Cart.Info.Country = request("country")
   Cart.Info.Phone = request("shiptophone")
   Cart.Info.Email = request("email")

   Cart.Info.ShipName = request("ShipName")
   Cart.Info.ShipCompany = request("ShipCompany")
   Cart.Info.ShipAddress1 = request("shipAddress1")
   Cart.Info.ShipAddress2 = request("shipAddress2")
   Cart.Info.ShipCity = request("shipCity")
   Cart.Info.ShipStateProvince = request("shipStateProvince")
   Cart.Info.ShipZipPostal = request("shipZipPostal")
   Cart.Info.ShipCountry = request("shipCountry")
   Cart.Info.OrderCustom8 = "OFF"
   if not trim(request("ShipPhoneNumber")) ="" then
	   Cart.Info.ShipPhone = request("ShipPhoneNumber")
   else
	   Cart.Info.ShipPhone = request("shiptophone")
   end if
   Cart.Info.ORDERCUSTOM8 = "OFF"  
   Cart.Info.ShipFax = Cart.Info.Fax 
   Cart.Info.ShipEmail = request("shipEmail")

   Cart.Info.CCMonth = request("month")
   Cart.Info.CCYear = request("year")
   Cart.Info.CCNumber = request("accountnumber")

   ' figure out the card type based on first digit of the card number
   vCC_First = left(Cart.Info.CCNumber, 1)
   if vCC_First = "3" then Cart.Info.CCType = "American Express"
   if vCC_First = "4" then Cart.Info.CCType = "Visa"
   if vCC_First = "5" then Cart.Info.CCType = "Mastercard"
   if vCC_First = "6" then Cart.Info.CCType = "Discover"

   Cart.Info.CCName = request("sjname")
   Cart.OrderID = request("ordernumber")

	Cart.Calculate
   Cart.Total = request("transactionamount")

   Cart.SaveCart

   ' figure out the card type based on first digit of the card number
   vCC_First = left(Cart.Info.CCNumber, 1)
   if vCC_First = "3" then
      Cart.Info.CCType = "American Express"
      vCCType = "003"
   end if
   if vCC_First = "4" then
      Cart.Info.CCType = "Visa"
      vCCType = "001"
   end if
   if vCC_First = "5" then
      Cart.Info.CCType = "Mastercard"
      vCCType = "002"
   end if
   if vCC_First = "6" then
      Cart.Info.CCType = "Discover"
      vCCType = "004"
   end if

      	' set up the request by creating a Hashtable and adding fields to it
      	dim oRequest	
      	set oRequest = Server.CreateObject( "CyberSourceWS.Hashtable" )
      
      	oRequest( "ccAuthService_run" ) = "true"

      	' we will let the Client get the merchantID from the MerchantConfig object
      	' and insert it into the Hashtable.
      
      
      	' this is your own tracking number.  This sample uses a hardcoded value.
      	' CyberSource recommends that you use a unique one for each order.
         ' the length of this order number, from processspike.asp, needs to be 10 or less chars
         '    or a 500 server error will be thrown
         '
         '  this whole thing just adds leading zero's to the order number... therefore optional
      	oRequest( "merchantReferenceCode" ) = Left("0000000000", 10 - Len(Cart.OrderID)) & Cart.OrderID

         oRequest( "taxService_run" ) = "false"

         ' need to break name into firstname/lastname
         ' last space defines where the lastname begins
         if instr(Cart.Info.Name, " ") Then
           ' vBTFN = Left(Cart.Info.Name, instrrev(Cart.Info.Name, " "))
           ' vBTLN = Right(Cart.Info.Name, len(Cart.Info.Name) - instrrev(Cart.Info.Name, " "))
            vBTFN = Cart.Info.Custom3
            vBTLN = Cart.Info.Custom5
         else
            vBTFN = "NONAME"
            vBTLN = Cart.Info.Name
         end if

         if instr(Cart.Info.ShipName, " ") Then
            vSTFN = Left(Cart.Info.ShipName, instrrev(Cart.Info.ShipName, " "))
            vSTLN = Right(Cart.Info.ShipName, len(Cart.Info.ShipName) - instrrev(Cart.Info.ShipName, " "))
         else
            vSTFN = "[NO NAME]"
            vSTLN = Cart.Info.ShipName
         end if
         
         ' use the country code and province
         Dim vBTCountryCode, vSTCountryCode
         
         ' check if custom2 is blank... if so then we didn't get an "other than us" country
         if IsEmpty(Cart.Info.Custom2) OR IsNULL(Cart.Info.Custom2) OR Cart.Info.Custom2="" Then
            vBTStateProvince = Cart.Info.StateProvince
            vBTCountryCode = "US"
         else
            ' since it is an "other than us" country, use custom2 as country, and province (custom1) instead of state 
            vBTStateProvince = Cart.Info.Custom1
            vBTCountryCode = vCountryCodes.Item(Cart.Info.Custom2)
         end if            

         ' check if shipcustom2 is blank... if so then we didn't get an "other than us" country
         if IsEmpty(Cart.Info.ShipCustom2) OR IsNULL(Cart.Info.ShipCustom2) OR Cart.Info.ShipCustom2="" Then
            vSTStateProvince = Cart.Info.ShipStateProvince
            vSTCountryCode = "US"
         else
            ' since it is an "other than us" country, use shipcustom2 as country, and province (shipcustom1) instead of state 
            vSTStateProvince = Cart.Info.ShipCustom1
            vSTCountryCode = vCountryCodes.Item(Cart.Info.ShipCustom2)
         end if

		'''''''''New Country & State code --- Jerry
        vBTCountryCode = Cart.Info.Country
        vSTCountryCode = Cart.Info.ShipCountry
        'if Cart.Info.Country = "USA" then
            vBTStateProvince = Cart.Info.StateProvince
            vSTStateProvince = Cart.Info.ShipStateProvince
        'else
        '    vBTStateProvince = Cart.Info.Custom1
        '    vSTStateProvince = Cart.Info.ShipCustom1
        'end if

         
      	oRequest( "billTo_firstName" ) = vBTFN
      	oRequest( "billTo_lastName" ) = vBTLN
      	oRequest( "billTo_company" ) = Cart.Info.Company
      	oRequest( "billTo_street1" ) = Cart.Info.Address1
      	oRequest( "billTo_street2" ) = Cart.Info.Address2
      	oRequest( "billTo_city" ) = Cart.Info.City
      	oRequest( "billTo_state" ) = vBTStateProvince
      	oRequest( "billTo_postalCode" ) = Cart.Info.ZipPostal
      	oRequest( "billTo_country" ) = vBTCountryCode
      	oRequest( "billTo_email" ) = Cart.Info.Email
      	oRequest( "billTo_phoneNumber" ) = Cart.Info.Phone

      	oRequest( "shipTo_firstName" ) = vSTFN
      	oRequest( "shipTo_lastName" ) = vSTLN
      	oRequest( "shipTo_company" ) = Cart.Info.ShipCompany
      	oRequest( "shipTo_street1" ) = Cart.Info.ShipAddress1
      	oRequest( "shipTo_street2" ) = Cart.Info.ShipAddress2
      	oRequest( "shipTo_city" ) = Cart.Info.ShipCity
      	oRequest( "shipTo_state" ) = vSTStateProvince
      	oRequest( "shipTo_postalCode" ) = Cart.Info.ShipZipPostal
      	oRequest( "shipTo_country" ) = vSTCountryCode
      	oRequest( "shipTo_email" ) = Cart.Info.ShipEmail
      	oRequest( "shipTo_phoneNumber" ) = Cart.Info.ShipPhone

      	oRequest( "card_accountNumber" ) = Cart.Info.CCNumber
      	oRequest( "card_cardType" ) = vCCTYpe
      	oRequest( "card_expirationMonth" ) = Cart.Info.CCMonth
      	oRequest( "card_expirationYear" ) = Cart.Info.CCYear
      	oRequest( "card_fullName" ) = Cart.Info.CCName
      	oRequest( "purchaseTotals_currency" ) = "USD"

      	' obtain visitor's IP Address.  This is the method used by many ASP
      	' developers to obtain the visitor's IP address.  It is not guaranteed to
      	' work all the time.  For instance, some proxies do not set the
      	' HTTP_X_FORWARDED_FOR header, in which case, you'll end up using
      	' REMOTE_ADDR, which is the IP address of the proxy, not of the visitor.
      	dim strIPAddress
      	strIPAddress = Request.ServerVariables( "HTTP_X_FORWARDED_FOR" )
      	if strIPAddress = "" then
      		strIPAddress = Request.ServerVariables( "REMOTE_ADDR" )
      	end if
      
      	if strIPAddress <> "" then
      		oRequest( "billTo_ipAddress" ) = strIPAddress
      	end if

      	' set the line item fields using information in the shopping basket
      	Dim ii, nNumItems, oItem
         For each Item in Cart.Items
            ii = ii + 1
      		oRequest( "item_" & ii & "_productName" ) = Item.Name
      		oRequest( "item_" & ii & "_productSKU" ) = Item.ItemID
      		oRequest( "item_" & ii & "_quantity" ) = Item.Quantity
      		oRequest( "item_" & ii & "_unitPrice" ) = Item.Price
      	next

         ' BY USING THE GRANDTOTAL WE DONT NEED TO WORRY ABOUT TAXES OR FREIGHT
         ' if debugging then 1$ grand total
         if vDEBUGGING = 1 then
            oRequest( "purchaseTotals_grandTotalAmount" ) = 1
         else
            oRequest( "purchaseTotals_grandTotalAmount" ) = Cart.Total
         end if            
         
      	' create Client object
      	dim oClient
      	set oClient = Server.CreateObject( "CyberSourceWS.Client" )
      	
      	' send request now
      	dim varReply, nStatus, strErrorInfo
      	nStatus = oClient.RunTransaction( _
      				Application( "MerchantConfig" ), Nothing, Nothing, _
      				oRequest, varReply, strErrorInfo )

         GetCCStatus nStatus, varReply, res, vCCErrorMessage, vAVSResponseCode, vAVSResponseMessage

         'response.write "<hr>" & Cart.CC.AuthCode & "<br>" & vAVSResponseCode & "<br>" & vAVSResponseMessage
         'response.write "<pre>res=" & res & vbcrlf
         'res = 3               

			if (res = 3) Then
  				   ErrString = "The following message was returned from our Credit Card processor:<BR>" & vCCErrorMessage & vbcrlf
  				   vErrorFlag = True
			End If 

			if (res = 2) Then 
  				   ErrString =  "There was an error processing your card:<BR>" & vCCErrorMessage & vbcrlf
  				   vErrorFlag = True
			End If 

			if (res = 0) Then 
      	   ErrString = "Unrecognized error:<BR>" & vCCErrorMessage & vbcrlf
      	   vErrorFlag = True
			End If 

'Response.write "Response: " & res & " / " & ErrString & ":" & vErrorFlag & "<br><br>"
'response.end


' ---------- we're good
%>
<html>
<body>
<div align="center">
  <center>
      <table border="0" cellpadding="2" cellspacing="0" width="600" bgcolor="#3399FF" style="border-collapse: collapse" bordercolor="#111111">
        <tr> 
          <td colspan="2"> 
            <p align="center"><font face="Verdana" size="2"><b><br>
              </b></font><b><font face="Verdana" size="4"><i>BicycleBuys.com</i></font></b> 
            <p align="center"> <font face="Verdana"><b>Order Entry Form Results</b></font> 
          </td>
        </tr>
        <tr> 
          <td colspan="2"> 
            <p align="center">
            <br><br>
            <font face="Verdana" size="2"><b>
         
<%

      if vErrorFlag = True Then
         response.write "<font color=""white"">" & ErrString & "</font><br><br><br><br><br>"
      Else

        ' res should equal 1 here
         Cart.SaveCustomerDB
         Cart.SaveOrderDB
         vCartID = Cart.SaveCartDB
         Session("Cart") = Cart.SaveCart

         dim vAVSResponseMessage, vAVSResponseCode
         Call SendEmail()
         response.write "<pre>Customer's receipt has been mailed.</pre>"

         Set Cart = Nothing
      End If
%>
          </b></font></td>
        </tr>
      </table>
  </center>
</div>
</body>
</html>
