<!--#INCLUDE VIRTUAL="/includes/template_cls.asp" -->
<!--#INCLUDE VIRTUAL="/includes/common.asp" -->
<!--#INCLUDE VIRTUAL="/includes/cartconfig.asp" -->
<%
 dim adjRebate, adjRebateTotal, adjRebateTotalAll
     call zeroRebateArray()

   ' get the template engine ready
   set objTemplate = new template_cls

   vOUT1 = ""
   vOUT2 = ""
'   For each item in Cart.Items
'      response.write "<hr>IC4: " & Item.Custom4 & "<br>"
'   next

'response.buffer = true

dim ErrString, sql, vCCType, vBTFN, vBTLN, vSTFN, vSTLN, vBTStateProvince, vSTStateProvince, vCCErrorMessage, vAVSResponseCode
dim vAVSResponseMessage


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' problems with the cart.saveorderdb
' so we do it here instead
'
'   SUBROUTINE ON HOLD, GOT THE CART.SAVEORDERDB METHOD WORKING
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SaveOrderToDB(vOrderID)
 

   Set Conn = Server.CreateObject("ADODB.Connection")
   rs = Server.CreateObject("ADODB.Recordset")
   Conn.Open "dsn=liidsn;uid=iiscart;pwd=iiscart"

   ' first make sure it didn't save there already
   sql = "SELECT * FROM Orders   WITH (NOLOCK) WHERE id=" & vOrderID & "  For Browse"
   Set rs = Conn.Execute(sql)

   if rs.EOF Then
      response.write "Order " & vOrderID & " not found in Orders table.  Inserting..."

'      sql = "INSERT INTO Orders " _
'            " (id,CustomerID,SessionID,Comments,ShipComments,Tax,Shipping,OrderTotal,CCType,CCNumber,CCMonth,CCYear,CCName, " _
'            " OrderDate,ShippingType,PaymentType,Processed,CCAuthCode, " _
'            " Custom1,Custom2,Custom3,Custom4,Custom5,Custom6,Custom7,Custom8) " _
'            " VALUES( " _
'            " & vOrderID
   else
      response.write "Order " & vOrderID & " found in Orders table."
   end if

' to set the order ID we need the following done first
'SET IDENTITY_INSERT Orders ON

End Sub


'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------

' debug subroutine to dump the contents of the cart object
' to a text file for examination.
Sub SaveCartInfoDEBUG

   vDebugLogFile = "D:\root\new.bicyclebuys.com\logs\NewCheckoutDebug-" & Year(Date) & Month(Date) & Day(Date) & ".log"
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set vDebugLog = fso.OpenTextFile(vDebugLogFile, ForAppend, True)

   vDebugLog.WriteLine(chr(13) & "------------------------------------------------------------" & chr(13))
   vDebugLog.WriteLine("Name: " & Len(Cart.Info.Name) & "|" & Cart.Info.Name & chr(13))
   vDebugLog.WriteLine("Address1: " & Len(Cart.Info.Address1) & "|" & Cart.Info.Address1 & chr(13))
   vDebugLog.WriteLine("Address2: " & Len(Cart.Info.Address2) & "|" & Cart.Info.Address2 & chr(13))
   vDebugLog.WriteLine("Company: " & Len(Cart.Info.Company) & "|" & Cart.Info.Company & chr(13))
   vDebugLog.WriteLine("City: " & Len(Cart.Info.City) & "|" & Cart.Info.City & chr(13))
   vDebugLog.WriteLine("StateProvince: " & Len(Cart.Info.StateProvince) & "|" & Cart.Info.StateProvince & chr(13))
   vDebugLog.WriteLine("ZipPostal: " & Len(Cart.Info.ZipPostal) & "|" & Cart.Info.ZipPostal & chr(13))
   vDebugLog.WriteLine("Country: " & Len(Cart.Info.Country) & "|" & Cart.Info.Country & chr(13))
   vDebugLog.WriteLine("Phone: " & Len(Cart.Info.Phone) & "|" & Cart.Info.Phone & chr(13))
   vDebugLog.WriteLine("Fax: " & Len(Cart.Info.Fax) & "|" & Cart.Info.Fax & chr(13))
   vDebugLog.WriteLine("Email: " & Len(Cart.Info.Email) & "|" & Cart.Info.Email & chr(13))
   vDebugLog.WriteLine("Custom1: " & Len(Cart.Info.Custom1) & "|" & Cart.Info.Custom1 & chr(13))
   vDebugLog.WriteLine("Custom2: " & Len(Cart.Info.Custom2) & "|" & Cart.Info.Custom2 & chr(13))
   vDebugLog.WriteLine("Custom3: " & Len(Cart.Info.Custom3) & "|" & Cart.Info.Custom3 & chr(13))
   vDebugLog.WriteLine("Custom4: " & Len(Cart.Info.Custom4) & "|" & Cart.Info.Custom4 & chr(13))
   vDebugLog.WriteLine("Custom5: " & Len(Cart.Info.Custom5) & "|" & Cart.Info.Custom5 & chr(13))
   vDebugLog.WriteLine("Custom6: " & Len(Cart.Info.Custom6) & "|" & Cart.Info.Custom6 & chr(13))
   vDebugLog.WriteLine("Custom7: " & Len(Cart.Info.Custom7) & "|" & Cart.Info.Custom7 & chr(13))
   vDebugLog.WriteLine("Custom8: " & Len(Cart.Info.Custom8) & "|" & Cart.Info.Custom8 & chr(13))

   vDebugLog.WriteLine("ShipName: " & Len(Cart.Info.ShipName) & "|" & Cart.Info.ShipName & chr(13))
   vDebugLog.WriteLine("ShipAddress1: " & Len(Cart.Info.ShipAddress1) & "|" & Cart.Info.ShipAddress1 & chr(13))
   vDebugLog.WriteLine("ShipAddress2: " & Len(Cart.Info.ShipAddress2) & "|" & Cart.Info.ShipAddress2 & chr(13))
   vDebugLog.WriteLine("ShipCompany: " & Len(Cart.Info.ShipCompany) & "|" & Cart.Info.ShipCompany & chr(13))
   vDebugLog.WriteLine("ShipCity: " & Len(Cart.Info.ShipCity) & "|" & Cart.Info.ShipCity & chr(13))
   vDebugLog.WriteLine("ShipStateProvince: " & Len(Cart.Info.ShipStateProvince) & "|" & Cart.Info.ShipStateProvince & chr(13))
   vDebugLog.WriteLine("ShipZipPostal: " & Len(Cart.Info.ShipZipPostal) & "|" & Cart.Info.ShipZipPostal & chr(13))
   vDebugLog.WriteLine("ShipCountry: " & Len(Cart.Info.ShipCountry) & "|" & Cart.Info.ShipCountry & chr(13))
   vDebugLog.WriteLine("ShipPhone: " & Len(Cart.Info.ShipPhone) & "|" & Cart.Info.ShipPhone & chr(13))
   vDebugLog.WriteLine("ShipFax: " & Len(Cart.Info.ShipFax) & "|" & Cart.Info.ShipFax & chr(13))
   vDebugLog.WriteLine("ShipEmail: " & Len(Cart.Info.ShipEmail) & "|" & Cart.Info.ShipEmail & chr(13))
   vDebugLog.WriteLine("ShipCustom1: " & Len(Cart.Info.ShipCustom1) & "|" & Cart.Info.ShipCustom1 & chr(13))
   vDebugLog.WriteLine("ShipCustom2: " & Len(Cart.Info.ShipCustom2) & "|" & Cart.Info.ShipCustom2 & chr(13))
   vDebugLog.WriteLine("ShipCustom3: " & Len(Cart.Info.ShipCustom3) & "|" & Cart.Info.ShipCustom3 & chr(13))
   vDebugLog.WriteLine("ShipCustom4: " & Len(Cart.Info.ShipCustom4) & "|" & Cart.Info.ShipCustom4 & chr(13))
   vDebugLog.WriteLine("ShipCustom5: " & Len(Cart.Info.ShipCustom5) & "|" & Cart.Info.ShipCustom5 & chr(13))
   vDebugLog.WriteLine("ShipCustom6: " & Len(Cart.Info.ShipCustom6) & "|" & Cart.Info.ShipCustom6 & chr(13))
   vDebugLog.WriteLine("ShipCustom7: " & Len(Cart.Info.ShipCustom7) & "|" & Cart.Info.ShipCustom7 & chr(13))
   vDebugLog.WriteLine("ShipCustom8: " & Len(Cart.Info.ShipCustom8) & "|" & Cart.Info.ShipCustom8 & chr(13))

   vDebugLog.WriteLine("Comments: " & Len(Cart.Info.Comments) & "|" & Cart.Info.Comments & chr(13))
   vDebugLog.WriteLine("BillingSameAsShipping: " & Len(Cart.Info.BillingSameAsShipping) & "|" & Cart.Info.BillingSameAsShipping & chr(13))
   vDebugLog.WriteLine("Password: " & Len(Cart.Info.Password) & "|" & Cart.Info.Password & chr(13))
   vDebugLog.WriteLine("OrderCustom1: " & Len(Cart.Info.OrderCustom1) & "|" & Cart.Info.OrderCustom1 & chr(13))
   vDebugLog.WriteLine("OrderCustom2: " & Len(Cart.Info.OrderCustom2) & "|" & Cart.Info.OrderCustom2 & chr(13))
   vDebugLog.WriteLine("OrderCustom3: " & Len(Cart.Info.OrderCustom3) & "|" & Cart.Info.OrderCustom3 & chr(13))
   vDebugLog.WriteLine("OrderCustom4: " & Len(Cart.Info.OrderCustom4) & "|" & Cart.Info.OrderCustom4 & chr(13))
   vDebugLog.WriteLine("OrderCustom5: " & Len(Cart.Info.OrderCustom5) & "|" & Cart.Info.OrderCustom5 & chr(13))
   vDebugLog.WriteLine("OrderCustom6: " & Len(Cart.Info.OrderCustom6) & "|" & Cart.Info.OrderCustom6 & chr(13))
   vDebugLog.WriteLine("OrderCustom7: " & Len(Cart.Info.OrderCustom7) & "|" & Cart.Info.OrderCustom7 & chr(13))
   vDebugLog.WriteLine("OrderCustom8: " & Len(Cart.Info.OrderCustom8) & "|" & Cart.Info.OrderCustom8 & chr(13))
   vDebugLog.WriteLine("IsStateResident: " & Len(Cart.Info.IsStateResident) & "|" & Cart.Info.IsStateResident & chr(13))
   vDebugLog.WriteLine("IsCountryResident: " & Len(Cart.Info.IsCountryResident) & "|" & Cart.Info.IsCountryResident & chr(13))
   vDebugLog.WriteLine("CCTYpe: " & Len(Cart.Info.CCTYpe) & "|" & Cart.Info.CCTYpe& chr(13))
   vDebugLog.WriteLine("CCNumber: " & Len(Cart.Info.CCNumber) & "|" & Cart.Info.CCNumber & chr(13))
   vDebugLog.WriteLine("CCName: " & Len(Cart.Info.CCName) & "|" & Cart.Info.CCName& chr(13))
   vDebugLog.WriteLine("CCMonth: " & Len(Cart.Info.CCMonth) & "|" & Cart.Info.CCMonth & chr(13))
   vDebugLog.WriteLine("CCYear: " & Len(Cart.Info.CCYear) & "|" & Cart.Info.CCYear& chr(13))
   vDebugLog.WriteLine("PaymentType: " & Len(Cart.Info.PaymentType) & "|" & Cart.Info.PaymentType & chr(13))

'   vDebugLog.WriteLine(": " & Len(Cart.Info.) & "|" & Cart.Info. & chr(13))

   vDebugLog.WriteLine(chr(13) & "ITEMS    -----------" & chr(13))
   For each Item in Cart.Items
      vDebugLog.WriteLine("ItemID: " & Len(Item.ItemID) & "|" & Item.ItemID & chr(13))
      vDebugLog.WriteLine("Name: " & Len(Item.Name) & "|" & Item.Name & chr(13))
      vDebugLog.WriteLine("Description: " & Len(Item.Description) & "|" & Item.Description & chr(13))
      vDebugLog.WriteLine("Quantity: " & Len(Item.Quantity) & "|" & Item.Quantity & chr(13))
      vDebugLog.WriteLine("Price: " & Len(Item.Price) & "|" & Item.Price & chr(13))
      vDebugLog.WriteLine("Weight: " & Len(Item.Weight) & "|" & Item.Weight & chr(13))
      vDebugLog.WriteLine("Custom1: " & Len(Item.Custom1) & "|" & Item.Custom1 & chr(13))
      vDebugLog.WriteLine("Custom2: " & Len(Item.Custom2) & "|" & Item.Custom2 & chr(13))
      vDebugLog.WriteLine("Custom3: " & Len(Item.Custom3) & "|" & Item.Custom3 & chr(13))
      vDebugLog.WriteLine("Custom4: " & Len(Item.Custom4) & "|" & Item.Custom4 & chr(13))
      vDebugLog.WriteLine("Custom5: " & Len(Item.Custom5) & "|" & Item.Custom5 & chr(13))
      vDebugLog.WriteLine("Custom6: " & Len(Item.Custom6) & "|" & Item.Custom6 & chr(13))
      vDebugLog.WriteLine("Custom7: " & Len(Item.Custom7) & "|" & Item.Custom7 & chr(13))
      vDebugLog.WriteLine("Custom8: " & Len(Item.Custom8) & "|" & Item.Custom8 & chr(13))
      vDebugLog.WriteLine("Tax: " & Item.Tax & chr(13))
      vDebugLog.WriteLine("TaxRate: " & Item.TaxRate & chr(13))
      vDebugLog.WriteLine("Adjust: " & Item.Adjust & chr(13))
'      vDebugLog.WriteLine("AdjustRate: " & Item.AdjustRate & chr(13))
      For i = 1 to Item.PropertyCount
         vDebugLog.WriteLine("Property Name: " & Len(Item.PropertyName(i)) & "|" & Item.PropertyName(i) & chr(13))
         vDebugLog.WriteLine("Property Value: " & Len(Item.PropertyValue(i)) & "|" & Item.PropertyValue(i) & chr(13))
         vDebugLog.WriteLine("Property Data: " & Len(Item.PropertyData(i)) & "|" & Item.PropertyData(i) & chr(13))
      Next
   Next

   vDebugLog.Close

End Sub

Sub DigOutAVS ( )
   ' the use of this subroutine has been depreciated
   ' use of the Cart.CC.avscode property has taken it's place
   '  -- reason, cart.cc.logging is NOT working    (9/30/2006)

   Const ForReading = 1, ForWriting = 2
   Dim fso, f

   vCCLogFile = "D:\root\logs\CC" & Year(Date) & Month(Date) & Day(Date) & ".log"
'   response.write vCCLogFile  & "<br>"

   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile(vCCLogFile, ForReading)
   FileContents =   f.ReadAll
'   response.write FileContents

   FileLines = split(FileContents, chr(13))
   FileCount = UBound(FileLines)
'   response.write FileCount & "<br>"

   LastTrans = FileLines(FileCount - 10)
'   response.write "<HR>" & LastTrans

   Values = split(LastTrans, "--><!--")
   for x = 0 to ubound(Values)
      if instr(Values(x), "AVS") > 0 Then
'        response.write x & ":" & values(x) & "<br>"
         vWork = replace(Values(x), "<", "")
         vWork = replace(vWork, ">", "")
         vWork = replace(vWork, "sz", "")

         SmallArr = split(vWork, "=")
         if Ubound(SmallArr) > 0 Then
            vKey = SmallArr(0)
            vVal = SmallArr(1)
'            response.write "<pre>" & vKey & ": <font color=""white"">" & vVal & "</font></pre>"
            if vKey = "AVSResponseCode" Then
               vAVSResponseCode = vVal
            end if
            if vKey = "AVSResponseMessage" Then
               vAVSResponseMessage = vVal
            end if
         End if
      End If
   next
End Sub

Sub SendEmail( )
   Set Conn = Server.CreateObject("ADODB.Connection")
   rs = Server.CreateObject("ADODB.Recordset")
   Conn.Open "dsn=liidsn;uid=iiscart;pwd=iiscart"

 '  sql = "UPDATE Orders SET CCAuthCode = '" & Cart.CC.AuthCode & "' WHERE id = " & Cart.OrderID
   
   
'   if  Request.ServerVariables("REMOTE_ADDR")  = "10.0.1.85"  then
    sql =  "exec spSaveCCInfo " & Cart.OrderID & ",'" & replace( vAVSResponseCode,"'","") & "','" & replace(vAVSResponseMessage,"'","") & "','"  & replace(Cart.CC.AuthCode,"'","") & "'"
     
 '  end if
   
	Conn.Execute(sql)

conn.close

   eheader = "BICYCLEBUYS.COM ONLINE CUSTOMER WEB ORDER ACKNOWLEDGEMENT" & vbcrlf & vbcrlf

   eheader = eheader & "Web Order Number: " & Left("000000", 6 - Len(Cart.OrderID)) & Cart.OrderID & vbcrlf
   eheader = eheader & "Date: " & Date & vbcrlf
   eheader = eheader & "Time: " & Time & vbcrlf & vbcrlf
   eheader = eheader & "Customer Name: " & Cart.Info.Name & vbcrlf & vbcrlf

   eheaderI = eheaderI & "Web Order Number: " & Left("000000", 6 - Len(Cart.OrderID)) & Cart.OrderID & vbcrlf
   eheaderI = eheaderI & "Date: " & Date & vbcrlf
   eheaderI = eheaderI & "Time: " & Time & vbcrlf

   if Len(Session("ReferredBy")) > 0 Then
      eheaderI = eheaderI & "Referred By: " & Session("ReferredBy") & vbcrlf & vbcrlf
   End If

   If Cart.Info.OrderCustom8 <> "ON" Then
      ebillto = ebillto & "Bill-To Address: " & vbcrlf
   else
      ebillto = ebillto & "Bill-To/Ship-To Address: " & vbcrlf
   end If

   ebillto = ebillto & "  " & Cart.Info.Name &  vbcrlf

   If len(Cart.Info.Company) > 0 Then ebillto = ebillto & " " & Cart.Info.Company & vbcrlf

   ebillto = ebillto & "  " & Cart.Info.Address1 &  vbcrlf

   If len(Cart.Info.Address2) > 0 Then ebillto = ebillto & "  " & Cart.Info.Address2 & vbcrlf

   If Cart.Info.StateProvince <> "ZZ" Then
      ebillto = ebillto & "  " & Cart.Info.City & ", " & Cart.Info.StateProvince & "   " & Cart.Info.ZipPostal & vbcrlf
	  ebillto = ebillto & "  " & Cart.Info.Country & vbcrlf
   Else
      ebillto = ebillto & "  " & Cart.Info.City & ", " & Cart.Info.Custom1
		if (Cart.Info.ZipPostal = "None") then
			ebillto = ebillto & "" & vbcrlf & "  " & Cart.Info.Country & vbcrlf
		else
			ebillto = ebillto & " " & Cart.Info.ZipPostal & vbcrlf & "  " & Cart.Info.Country & vbcrlf
		end if
   End If

   ebillto = ebillto & "  Ph: " & Cart.Info.Phone & vbcrlf

   If Len(Cart.Info.Fax) > 0 then ebillto = ebillto & "  Fx: " & Cart.Info.Fax & vbcrlf
   If Len(Cart.Info.Email) > 0 Then ebillto = ebillto & "  " & Cart.Info.Email & vbcrlf
   If Len(Cart.Info.Comments) > 0 Then ebillto = ebillto & "  Comments: " & Cart.Info.Comments & vbcrlf

   If Cart.Info.OrderCustom8 <> "ON" Then
      eshipto = eshipto & vbcrlf & "Ship-To Address:" & vbcrlf
      eshipto = eshipto & "  " & Cart.Info.ShipName &  vbcrlf
      If Cart.Info.ShipCompany <> "" Then eshipto = eshipto & "  " & Cart.Info.ShipCompany & vbcrlf
      eshipto = eshipto & "  " & Cart.Info.ShipAddress1 & vbcrlf
      If Cart.Info.ShipAddress2 <> "" Then eshipto = eshipto & "  " & Cart.Info.ShipAddress2 & vbcrlf

      If Cart.Info.ShipStateProvince <> "ZZ" Then
         eshipto = eshipto & "  " & Cart.Info.ShipCity & ", "
         eshipto = eshipto & Cart.Info.ShipStateProvince & "   " & Cart.Info.ShipZipPostal & vbcrlf
		 eshipto = eshipto & "  " & Cart.Info.ShipCountry & vbcrlf
      Else
         eshipto = eshipto & "  " & Cart.Info.ShipCustom1 '& "   " & Cart.Info.ShipCustom2 & "  " & Cart.Info.ShipZipPostal & vbcrlf
			if (Cart.Info.ShipZipPostal = "None") then
				eshipto = eshipto & vbcrlf & Cart.Info.ShipCountry & vbcrlf
			else
				eshipto = eshipto & " " & Cart.Info.ShipZipPostal & vbcrlf & "  " & Cart.Info.ShipCountry & vbcrlf
			end if

      End If

      eshipto = eshipto & "  Ph: " & Cart.Info.ShipPhone & vbcrlf
      If Len(Cart.Info.ShipFax) > 0 then eshipto = eshipto & "  Fx: " & Cart.Info.ShipFax & vbcrlf
      If Len(Cart.Info.ShipEmail) > 0 Then eshipto = eshipto & "  " & Cart.Info.ShipEmail & vbcrlf
      if Len(Cart.Info.ShipComments) > 0 Then eshipto = eshipto & "  Comments: " & Cart.Info.ShipComments & vbcrlf
   End If

   epayment = ""
   If Cart.Info.PaymentType = "Credit Card" Then
      epayment = epayment & "Payment: Credit Card" & vbcrlf
      epayment = epayment & "  Card Type: " & Cart.Info.CCType & vbcrlf & vbcrlf
      epayment = epayment & "  Card Holder: " & Cart.Info.CCName & vbcrlf
      epayment = epayment & "  Authorization Code: " & Cart.CC.AuthCode & vbcrlf
   Else
      epayment = epayment & "Payment: Fax/Call-In" & vbcrlf
      epayment = epayment & "  Fax: 1(516) 673-2220" & vbcrlf
      epayment = epayment & "  Phone: 1(888) 4BIKE BUY or 1(888) 424-5328" & vbcrlf
   End If

   ' This is the internal version of the epayment paragraph
   epaymentI = ""
   If Cart.Info.PaymentType = "Credit Card" Then
      epaymentI = epaymentI & "Payment: " & Cart.Info.CCType & vbcrlf
      epaymentI = epaymentI & "  Card Number: " & Cart.Info.CCNumber & vbcrlf
      epaymentI = epaymentI & "  Card Expiration: " & Cart.Info.CCMonth & "/" & Cart.Info.CCYear & vbcrlf
      epaymentI = epaymentI & "  Card Holder: " & Cart.Info.CCName & vbcrlf
      epaymentI = epaymentI & "  Authorization Code: " & Cart.CC.AuthCode & vbcrlf
      epaymentI = epaymentI & "  AVS Code: " & vAVSResponseCode & vbcrlf
      epaymentI = epaymentI & "  AVS Message: " & vAVSResponseMessage & vbcrlf
   Else
      epaymentI = epaymentI & "Payment: Fax/Call-In" & vbcrlf
   End If

   ' Build the item detail listing
   eitems = vbcrlf
   eitems = eitems & "Item Details" & vbcrlf
   eitems = eitems & "==========================================================" & vbcrlf
   For Each Item in Cart.Items
      eitems = eitems & "  Name:                      " & Item.Name & vbcrlf
      eitems = eitems & "  Sku#:                      " & Item.ItemID & vbcrlf
      If Item.Custom4 <> "" Then
      	' Get properties from cart
      	vProp = Item.Custom4
      	vPropID = Item.Custom5
   	   x = instr(vProp, ";")
   	   vPropA = Split(right(vProp, Len(vProp)-x), ",")
   	   x = instr(vPropID, ";")
   	   vPropIDA = Split(right(vPropID, Len(vPropID)-x), ",")
   	   for i = 0 to ubound(vPropA)
      	   If vPropIDA(i) = Item.ItemID Then eitems = eitems & "  Size/Color:                " & vPropA(i) & vbcrlf
   	   Next
      End If
      eitems = eitems & "  Quantity:                  " & Item.Quantity & vbcrlf
      eitems = eitems & "  Unit Price:                " & FormatCurrency(Item.Price,2,0,0) & vbcrlf
      eitems = eitems & "  Extended Price:            " & FormatCurrency(Item.Quantity * Item.Price,2,0) & vbcrlf & vbcrlf
   Next
   eitems = eitems & "==========================================================" & vbcrlf

   ' Build the item detail listing
   eitemsI = vbcrlf
   eitemsI = eitemsI & "Item Details" & vbcrlf

   eitemsI = eitemsI & "==========================================================" & vbcrlf
   For Each Item in Cart.Items
      eitemsI = eitemsI & Item.ItemID
      eitemsI = eitemsI & Item.Name
      If Item.Custom4 <> "" Then
      	' Get properties from cart
      	vProp = Item.Custom4
      	vPropID = Item.Custom5
   	   x = instr(vProp, ";")
   	   vPropA = Split(right(vProp, Len(vProp)-x), ",")
   	   x = instr(vPropID, ";")
   	   vPropIDA = Split(right(vPropID, Len(vPropID)-x), ",")
   	   for i = 0 to ubound(vPropA)
      	   If vPropIDA(i) = Item.ItemID Then eitemsI = eitemsI & vPropA(i) & vbcrlf
   	   Next
      End If
      eitemsI = eitemsI & Item.Quantity
      eitemsI = eitemsI & FormatCurrency(Item.Price,2,0,0)
      eitemsI = eitemsI & FormatCurrency(Item.Quantity * Item.Price,2,0) & vbcrlf
   Next
   eitemsI = eitemsI & "==========================================================" & vbcrlf
   ettl = ettl & getTextRebates()

   ettl = ettl & "Subtotal:                    " & formatcurrency(Cart.Gridtotal,2,0,0) & vbcrlf
   ettl = ettl & "Tax:                         " & formatcurrency(Cart.TotalTax,2,0,0) & vbcrlf
   if Cart.Info.ShipCountry = "US" Then
      ettl = ettl & "Shipping:                    " & formatcurrency(Cart.Shipping,2,0,0) & " (" & Cart.ShippingType & ")" & vbcrlf
   Else
      ettl = ettl & "Shipping:                    International -- We will contact you." & vbcrlf
   End If
   ettl = ettl & "Total:                       " & formatcurrency(Cart.Total,2,0,0) & vbcrlf & vbcrlf

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
   vOrderFileName = vSaveOrderPath & "Web-Order-" & Left("000000", 6 - Len(Cart.OrderID)) & Cart.OrderID & ".txt"
   Set fso = CreateObject("Scripting.FileSystemObject")

   Set vOrderFile = fso.OpenTextFile(vOrderFileName, 2, True)
   vOrderFile.WriteLine(vbcrlf & vbcrlf & eheaderI & ebillto & vbcrlf & epaymentI & eshipto & eitems & ettl)
   vOrderFile.Close

   ' ---- Send Email to Customer
       Cart.Mail.Host = "webserver"
   Cart.Mail.From = "Sales@BicycleBuys.com"
   Cart.Mail.FromName = "BicycleBuys.com Sales"
   Cart.Mail.Subject = "BicycleBuys.com Store Receipt - Web Order Number: " & Left("000000", 6 - Len(Cart.OrderID)) & Cart.OrderID

   Cart.Mail.Body = eheader & ebillto & vbcrlf & epayment & eshipto & eitems & ettl & efoot

   ' Send it to the shipping email address too
   If len(Cart.Info.Email) <> 0 Then
    Cart.Mail.AddAddress Cart.Info.Email
    sendtrue = 1
   End If

   ' Send it to the shipping email address too
   If (len(Cart.Info.ShipEmail) <> 0) and (Cart.Info.Email <> Cart.Info.ShipEmail) Then
       Cart.Mail.AddAddress Cart.Info.ShipEmail
       sendtrue = 1
   End if

   If sendtrue = 1 Then
      'Cart.mail.addattachment "/orders/" & Cart.orderid & ".enc"
      Cart.Mail.Send ' send to buyer
   End if
   Cart.Mail.Reset

End Sub
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
%><!--#INCLUDE VIRTUAL="/includes/cartdisplay_checkout.asp" --><%
      ' Put checkout and empty buttons on the display
      ' (disabled with < -1)
      Dim vCheckEmpty
      vCheckEmpty = ""
      if cart.gridtotalquantity > 0 then
         vCheckEmpty = "<a href=""" &  vThisProto & vThisServer & "/checkout/""><img src=""/cartimages/checkout.gif"" alt=""Check out"" border=""0"" WIDTH=""100"" HEIGHT=""20""></a>" _
                       & "<a href=""" &  vThisProto & vThisServer & "/emptycart/""><img src=""/cartimages/emptycart.gif"" alt=""Remove ALL items from the cart"" border=""0"" WIDTH=""100"" HEIGHT=""20""></a>"
      end if


      If Request("SAVE") <> "" Then
		   Cart.Info.CCType = "American Express"
			ErrString = Cart.AcceptBillingInfo
			vErrString = ErrString
			vCCYEAR =  Request.Form("CCYEAR")
'			if (vCCYEAR = "" or vCCYEAR="0" or IsEmpty(vCCYEAR) or IsNULL(vCCYEAR)) AND (Cart.Info.PaymentType = "Credit Card") Then ErrString = vErrString & "CREDIT CARD EXPIRATION YEAR;"
			Session("Cart") = Cart.SaveCart


			' response.write ErrString
'DON CODE FOR CYBER SCREWUP XXXXX
			If ErrString = ""   Then
      		TEMPCC = Cart.Info.CCNumber
				g_Key = mid(g_KeyString,1,Len(Cart.Info.CCNumber))
				Cart.Info.CCNumber = EnCrypt(Cart.Info.CCNumber)
				Cart.SaveOrderItemsDB
 	    
 				Cart.SaveOrderDB
				Cart.Info.CCNumber = TEMPCC

	
		'----------------- CREDIT CARD PROCESSING -----------------------------------------------------
				If Cart.Info.PaymentType = "Credit Card" Then

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



			  if vCyberSource = 1 Then


			 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			 ''' S/B DONE IN GLOBAL.ASA -- DONE HERE FOR TESTING
			 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				dim oMerchantConfig
				set oMerchantConfig = Server.CreateObject( "CyberSourceWS.MerchantConfig" )
				oMerchantConfig.MerchantID = "v2438728"
				oMerchantConfig.SendToProduction = "1"
				oMerchantConfig.KeysDirectory = "D:\Cybersource\keys\"
				oMerchantConfig.TargetAPIVersion = "1.25"
				'Visit https://ics2ws.ic3.com/commerce/1.x/transactionProcessor/ for the
				'latest version.
				oMerchantConfig.EnableLog = "1"
				oMerchantConfig.LogDirectory = "D:\Cybersource\logs\"

				set Application( "MerchantConfig" ) = oMerchantConfig
				' END CYBERSOURCE
			 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

				' set up the request by creating a Hashtable and adding fields to it
				dim oRequest
				set oRequest = Server.CreateObject( "CyberSourceWS.Hashtable" )

				oRequest( "ccAuthService_run" ) = "true"

				' we will let the Client get the merchantID from the MerchantConfig object
				' and insert it into the Hashtable.

				' this is your own tracking number.  This sample uses a hardcoded value.
				' CyberSource recommends that you use a unique one for each order.
				oRequest( "merchantReferenceCode" ) = Left("000000", 6 - Len(Cart.OrderID)) & Cart.OrderID
				 oRequest( "taxService_run" ) = "false"

				 ' need to break name into firstname/lastname
				 ' last space defines where the lastname begins
				 if instr(Cart.Info.Name, " ") Then
					vBTFN = Left(Cart.Info.Name, instrrev(Cart.Info.Name, " "))
					vBTLN = Right(Cart.Info.Name, len(Cart.Info.Name) - instrrev(Cart.Info.Name, " "))
				 else
					vBTFN = "[NO NAME]"
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


			   ''''''Original Country & State Code

				' check if custom2 is blank... if so then we didn't get an "other than us" country
				' if IsEmpty(Cart.Info.Custom2) OR IsNULL(Cart.Info.Custom2) OR Cart.Info.Custom2="" Then
				'    vBTStateProvince = Cart.Info.StateProvince
				'    vBTCountryCode = "US"
				' else
				 '   ' since it is an "other than us" country, use custom2 as country, and province (custom1) instead of state
				 '   vBTStateProvince = Cart.Info.Custom1
				'    vBTCountryCode = vCountryCodes.Item(Cart.Info.Custom2)
				' end if
				 ' check if shipcustom2 is blank... if so then we didn't get an "other than us" country
				 'if IsEmpty(Cart.Info.ShipCustom2) OR IsNULL(Cart.Info.ShipCustom2) OR Cart.Info.ShipCustom2="" Then
				 '   vSTStateProvince = Cart.Info.ShipStateProvince
				 '   vSTCountryCode = "US"
				 'else
				 '   ' since it is an "other than us" country, use shipcustom2 as country, and province (shipcustom1) instead of state
				 '   vSTStateProvince = Cart.Info.ShipCustom1
				 '   vSTCountryCode = vCountryCodes.Item(Cart.Info.ShipCustom2)
				 'end if


				'''''''''New Country & State code --- Jerry
				vBTCountryCode = Cart.Info.Country
				vSTCountryCode = Cart.Info.ShipCountry

				'get the actual country code...
			   sql = "SELECT * FROM Country   WITH (NOLOCK) WHERE Country LIKE '" & vBTCountryCode & "'  For Browse"
			   rs100.open sql,Conn,3
				if not rs100.EOF then
				   vBTCountryCode = rs100("CountryAbbreviation")
				end if
				rs100.close

			   sql = "SELECT * FROM Country   WITH (NOLOCK) WHERE Country LIKE '" & vSTCountryCode & "' For Browse"
			   rs100.open sql,Conn,3
				if not rs100.EOF then
				   vSTCountryCode = rs100("CountryAbbreviation")
				end if
				rs100.close



				if ((vBTCountryCode  = "US") OR (vBTCountryCode  = "CA")) then
				   vBTStateProvince = Cart.Info.StateProvince
				else
					vBTStateProvince = Cart.Info.Custom1

				end if
				if ((vSTCountryCode  = "US") OR (vSTCountryCode  = "CA")) then
					vSTStateProvince = Cart.Info.ShipStateProvince
				else
					vSTStateProvince = Cart.Info.ShipCustom1
				end if



				'oRequest( "billTo_firstName" ) = vBTFN
				'oRequest( "billTo_lastName" ) = vBTLN

				oRequest( "billTo_firstName" ) = Cart.Info.Custom3
				oRequest( "billTo_lastName" ) = Cart.Info.Custom5

				oRequest( "billTo_company" ) = Cart.Info.Company
				oRequest( "billTo_street1" ) = Cart.Info.Address1
				oRequest( "billTo_street2" ) = Cart.Info.Address2
				oRequest( "billTo_city" ) = Cart.Info.City
				oRequest( "billTo_state" ) = vBTStateProvince
				oRequest( "billTo_postalCode" ) = Cart.Info.ZipPostal
				oRequest( "billTo_country" ) = vBTCountryCode
				oRequest( "billTo_email" ) = Cart.Info.Email
				oRequest( "billTo_phoneNumber" ) = Cart.Info.Phone

				oRequest( "shipTo_firstName" ) = Cart.Info.ShipCustom3
				oRequest( "shipTo_lastName" ) = Cart.Info.ShipCustom5
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
				oRequest( "card_cvNumber") = request.form("CVVN")


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
					'oRequest( "purchaseTotals_grandTotalAmount" ) = 1
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

				 ' response.write "<hr>" & Cart.CC.AuthCode & "<br>" & vAVSResponseCode & "<br>" & vAVSResponseMessage
				 ' res = 3
			  else

				 ' using the Cart32 CC module
					Cart.cc.timeOut = 30		'seconds
					Cart.cc.TransType = "capture"
				   Cart.cc.logging = "D:/root/new.bicyclebuys.com/logs/"

					res = Cart.cc.charge()
					vCCErrorMessage = Cart.cc.LastCCErrorMsg
				 vAVSResponseCode = Cart.CC.AVSCode

			  end if
    
		'         response.write "<pre>res=" & res & vbcrlf
					if (res = 3) Then
						   vErrString = vErrString & "The following message was returned from our Credit Card processor:<BR>" & vCCErrorMessage & vbcrlf
						   vErrorFlag = True
					End If

					if (res = 2) Then
						   vErrString =  vErrString & "There was an error processing your card:<BR>" & vCCErrorMessage & vAVSResponseCode & vAVSResponseMessage & vbcrlf
						   vErrorFlag = True
					End If

					if (res = 0) Then
				   vErrString = vErrString & "Unrecognized error:<BR>" & vCCErrorMessage & vbcrlf
				   vErrorFlag = True
					End If

				 ' res should equal 1 here
				 If vErrorFlag <> True Then

					' debugging problems
					'call SaveCartInfoDEBUG

					if Cart.CustomerID > 0 Then
					   Cart.SaveCustomerDB(Cart.CustomerID)
					Else
					   Cart.SaveCustomerDB
					End If

					Cart.SaveOrderDB
					Cart.SaveItemsDB
					vCartID = Cart.SaveCartDB
					Session("Cart") = Cart.SaveCart

					'dim vAVSResponseMessage, vAVSResponseCode
					'Call DigOutAVS()
					Call SendEmail()

					if vDEBUGGING = 1 Then
					   'SaveOrderToDB Cart.OrderID
					   'response.end
					End if
                    Session("Cart") = Cart.SaveCart
					Set Cart = Nothing
					    Response.Redirect "checkout_final.asp?CID=" & vCartID
					   'Response.Redirect "thankyou.asp"
				 End If
				End if

		'----------------- FAX ORDER FORM --------------------------------------------------------------

				If Cart.Info.PaymentType = "Fax/Call" Then
				 ' debugging problems

				 g_Key = mid(g_KeyString,1,Len(Cart.Info.CCNumber))
				 Cart.Info.CCNumber = EnCrypt(Cart.Info.CCNumber)

				 'call SaveCartInfoDEBUG
				 if Cart.CustomerID > 0 Then
					Cart.SaveCustomerDB(Cart.CustomerID)
				 Else
					Cart.SaveCustomerDB
				 End If
				 Cart.SaveOrderDB
				 Cart.SaveItemsDB
				 vCartID = Cart.SaveCartDB
		'         Session("Cart") = Cart.SaveCart
				Call SendEmail()
		'         Set Cart = Nothing

				 if vDEBUGGING = 1 Then
					'SaveOrderToDB Cart.OrderID
					'response.end
				 End if

				   Response.Redirect "checkout_final.asp?CID=" & vCartID
			   'Response.Redirect "faxorder.asp"
			End If

		'----------------- ENTRY ERROR -----------------------------------------------------------------
		   Else
			  vErrString = "The following field(s) are required:<BR>" & ErrString
			  ErrString = vErrString
		   End If
	end if

      '''''''''''''''
      If Cart.Info.ShipCountry = "US" Then
          vShipTotal = FormatCurrency(Cart.Shipping,2,0,0)
      Else
          vShipTotal = "International<br>We will e-mail you!"
      End If


	if (Cart.Info.ShipStateProvince <> "NY") then
		Cart.Total = Cart.Total - Cart.TotalTax
		Cart.TotalTax = 0
	end if

      vSalesTax = FormatCurrency(Cart.TotalTax,2,0,0)
      vGrandTotal = FormatCurrency(Cart.Total,2,0,0)


      if Cart.Info.StateProvince = "OTHER" Then
         vSPTitle = "Province"
         vSPDisp = Cart.Info.Custom1
      Else
         vSPTitle = "State"
         vSPDisp = vStates.Item(Cart.Info.StateProvince)
         vSP2 = Cart.Info.StateProvince
      End If
      if Cart.Info.Country = "OTHER" Then
         vCDisp = Cart.Info.Custom2
      Else
        vCDisp = Cart.Info.Country
      End If

      If Cart.Info.ShipStateProvince = "OTHER" Then
         vSSPTitle = "Province"
         vSSPDisp = Cart.Info.Custom1
      Else
         vSSPTitle = "State"
         vSSPDisp = vStates.Item(Cart.Info.ShipStateProvince)
         vSSP2 = Cart.Info.ShipStateProvince
      End If

      If Cart.Info.ShipCountry = "OTHER" Then
         vSCDisp = Cart.Info.ShipCustom2
      Else
         vSCDisp = Cart.Info.ShipCountry
      End If

      If Cart.ShippingType <> "" Then
         vShippingType = Cart.ShippingType
      Else
         vShippingType = "None"
      End If

      If Cart.Info.OrderCustom8<>"ON" Then
         vAddressListed = "<TR ALIGN=""CENTER""><TD ALIGN=""CENTER"" COLSPAN=""2"">This address must be listed with your credit card company. For more info, please call us.</TD></TR>"
      End If

      ' had to use instr comparisons because regular comparisons were always coming up false
      '  tried reassignments and checked vartype -- no go

      vPaymentTypes = "<option VALUE=""Credit Card"" "
      If Cart.Info.PaymentType = "Credit Card" Then vPaymentTypes = vPaymentTypes & "SELECTED"
      vPaymentTypes = vPaymentTypes & ">Credit Card</option>"
      vPaymentTypes = vPaymentTypes & "<option VALUE=""Fax/Call"""
      If Cart.Info.PaymentType = "Fax/Call" Then vPaymentTypes = vPaymentTypes & "SELECTED"
      vPaymentTypes = vPaymentTypes & ">Fax/Call Order In</option>"

      vTMPA = Array("Visa", "Mastercard", "American Express", "Discover")
      For x = 0 to 3
         vCCTypes = vCCTypes & "<option VALUE=""" & vTMPA(x) & """ "
         If InStr(Cart.Info.CCType, vTMPA(x)) > 0 Then vCCTypes = vCCTypes & " SELECTED "
         vCCTypes = vCCTypes & ">" & vTMPA(x) & "</option>" & Chr(13)
      Next

      vTMPA = Array("--Month--", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
      For x = 0 to 12
         vCCMonths = vCCMonths & "<option VALUE=""" & x & """ "
         If Cart.Info.CCMonth = x Then vCCMonths = vCCMonths & " SELECTED "
         vCCMonths = vCCMonths & ">"
         If x > 0 Then vCCMonths = vCCMonths & x & "-"
         vCCMonths = vCCMonths & vTMPA(x) & "</option>" & Chr(13)
      Next

      vCCYears = "<option value=""0"">--Year--</option>" & Chr(13)
      For x = Year(Now) to Year(Now) + 12
         vCCYears = vCCYears & "<option value=""" & x & """"
         If Cart.Info.CCYear = x Then vCCYears = vCCYears & " SELECTED"
         vCCYears = vCCYears & ">" & x & "</option>" & Chr(13)
      Next

      ' cart display built, now show it
      with objTemplate
      	.TemplateFile = TMPLDIR & "billing.html"

			if (TotalDiscount15 > 0) then
				vOUT2 = vOUT2 & "<TR ><TD colspan=3 align=right class=cart style=""text-align: center; border-top: 1px solid #C4EAE6;"">&nbsp;</TD><TD align=right class=cart style=""text-align: center; border-top: 1px solid #C4EAE6;"">Discount</TD><TD colspan=5 align=right style=""text-align: center; border-top: 1px solid #C4EAE6;"" class=cart>-" & FormatCurrency(TotalDiscount15, 2, 0, 0) & "</TD></TR>"
			end if
     
       adjRebateTotal=0
         For each item in Cart.Items
        adjRebateTotalAll = getitemrow(vCount, vBGColor, _
                        Item.ItemID, Item.Name, Item.Weight, Item.Quantity, _
                        Item.Custom1, Item.Custom2, Item.Custom3, Item.Custom4, _
                        Item.Custom5, Item.Custom6, Item.Custom7, Item.Custom8, _
                        vItemOptions,vFreightMsg,Item.Price, Item.Adjust)
               adjRebate=Item.Adjust*Item.Quantity
              
               if (adjRebate) then
                  adjRebateTotal=adjRebate
                end if 
           next
          .AddToken "adjusttotal", 1, rebateHtml(adjRebateTotal)  
         .AddToken "displaycart", 1, vOUT2

         .AddToken "thisserver", 1, vThisServer
         .AddToken "checkempty", 1, vCheckEmpty

         .AddToken "sessionreferer", 1, Session("Referer")

         .AddToken "OtherStateSelected", 1, vTMP4
         .AddToken "StateSelect", 1, vTMP5

         .AddToken "shiptotal", 1, vShipTotal
         .AddToken "salestax", 1, vSalesTax
         
         '.AddToken "grandtotal", 1, vGrandTotal
         if (TotalDiscount15 > 0) then
                .AddToken "grandtotal", 1, FormatCurrency(Cart.Total - TotalDiscount15,2,0,0)
                'response.Write("TotalDiscount15  - " & TotalDiscount15)
            elseif (adjRebateTotal < 0) then
                'response.Write("adjRebate  - " & adjRebate)
                .AddToken "grandtotal", 1, FormatCurrency((Cart.Total-vAdjustTotal) + adjRebateTotal,2,0,0)
            else
                .AddToken "grandtotal", 1, vGrandTotal 
            end if
         
         '.AddToken "cgridtotal", 1, FormatCurrency(Cart.GridTotal,2,0,0)
        if (TotalDiscount15 > 0) then
            .AddToken "cgridtotal", 1, FormatCurrency(Cart.GridTotal - TotalDiscount15,2,0,0)
            'response.Write("1 adjRebateTotal - " & adjRebateTotal)
        elseif (adjRebateTotal < 0) then
            .AddToken "cgridtotal", 1, FormatCurrency((Cart.GridTotal-vAdjustTotal) + adjRebateTotal,2,0,0)
            'response.Write("2 adjRebateTotal - " & adjRebateTotal)
        else
            .AddToken "cgridtotal", 1, FormatCurrency(Cart.GridTotal,2,0,0)
            'response.Write("3 adjRebateTotal - " & adjRebateTotal)
        end if
         
         '.AddToken "adjusttotal", 1, rebateHtml(vAdjustTotal) 
         .AddToken "InfoName", 1, Cart.Info.Name
         .AddToken "InfoCompany", 1, Cart.Info.Company
         .AddToken "Address1", 1, Cart.Info.Address1
         .AddToken "Address2", 1, Cart.Info.Address2
         .AddToken "City", 1, Cart.Info.City
         .AddToken "SPTitle", 1, vSPTitle
         .AddToken "SPDisp", 1, vSPDisp
         .AddToken "SP2", 1, vSP2
         .AddToken "ZipPostal", 1, Cart.Info.ZipPostal
         .AddToken "CDisp", 1, vCDisp
         .AddToken "Phone", 1, Cart.Info.Phone
         .AddToken "Email", 1, Cart.Info.Email
         .AddToken "Fax", 1, Cart.Info.Fax
         .AddToken "Comments", 1, Cart.Info.Comments
         .AddToken "Custom1", 1, Cart.Info.Custom1
         .AddToken "Custom2", 1, Cart.Info.Custom2
         .AddToken "Custom3", 1, Cart.Info.Custom3
         .AddToken "Custom4", 1, Cart.Info.Custom4
         .AddToken "Custom5", 1, Cart.Info.Custom5
         .AddToken "Custom6", 1, Cart.Info.Custom6
         .AddToken "Custom7", 1, Cart.Info.Custom7
         .AddToken "Custom8", 1, Cart.Info.Custom8

         .AddToken "ShipName", 1, Cart.Info.ShipName
         .AddToken "ShipCompany", 1, Cart.Info.ShipCompany
         .AddToken "ShipAddress1", 1, Cart.Info.ShipAddress1
         .AddToken "ShipAddress2", 1, Cart.Info.ShipAddress2
         .AddToken "ShipZipPostal", 1, Cart.Info.ShipZipPostal
         .AddToken "ShipCity", 1, Cart.Info.ShipCity
         .AddToken "SSPTitle", 1, vSSPTitle
         .AddToken "SSPDisp", 1, vSSPDisp
         .AddToken "SSP2", 1, vSSP2
         .AddToken "SCDisp", 1, vSCDisp
         .AddToken "ShipEmail", 1, Cart.Info.ShipEmail
         .AddToken "ShipPhone", 1, Cart.Info.ShipPhone
         .AddToken "ShipComments", 1, Cart.Info.ShipComments
         .AddToken "ShipFax", 1, Cart.Info.ShipFax
         .AddToken "ShipCustom1", 1, Cart.Info.ShipCustom1
         .AddToken "ShipCustom2", 1, Cart.Info.ShipCustom2
         .AddToken "ShipCustom3", 1, Cart.Info.ShipCustom3
         .AddToken "ShipCustom4", 1, Cart.Info.ShipCustom4
         .AddToken "ShipCustom5", 1, Cart.Info.ShipCustom5
         .AddToken "ShipCustom6", 1, Cart.Info.ShipCustom6
         .AddToken "ShipCustom7", 1, Cart.Info.ShipCustom7
         .AddToken "ShipCustom8", 1, Cart.Info.ShipCustom8

         .AddToken "ShippingType",1, vShippingType
         .AddToken "AddressListed", 1, vAddressListed

            if (TotalDiscount15 > 0) then
                .AddToken "CartTotal", 1, FormatCurrency(Cart.Total - TotalDiscount15,2,0,0)
                'response.Write("TotalDiscount15  - " & TotalDiscount15)
            elseif (adjRebateTotal < 0) then
                'response.Write("adjRebate  - " & adjRebate)
                .AddToken "CartTotal", 1, FormatCurrency((Cart.Total-vAdjustTotal) + adjRebateTotal,2,0,0)
            else
                .AddToken "CartTotal", 1, vGrandTotal 
            end if



         '.AddToken "CartTotal", 1, formatcurrency(Cart.Total,2,0,0)
         
         
         
         .AddToken "PaymentTypes",1, vPaymentTypes
         .AddToken "CCTypes",1, vCCTypes
         .AddToken "CCMonths",1, vCCMonths
         .AddToken "CCYears",1, vCCYears

         .AddToken "CCNumber",1, Cart.Info.CCNumber
         .AddToken "CCName", 1, Cart.Info.CCName

         .AddToken "continueshopping", 1, Session("Referer")


      	.AddToken "ErrString", 1, "<b><font color=""#CC0000"">" & vErrString &  "</font></b>"
         .AddToken "CartString", 1, Server.HTMLEncode(Cart.SaveCart)

      	.AddToken "header", 3, vCartHeaderSummaryPayment
      	.AddToken "footer", 3,   TMPLDIR & "cart_footerPay.html"
        .AddToken "PromoCode", 1, getRebates() 
'         .AddToken "", 1, Cart.Info.

      	.parseTemplateFile
      end with
   Else

      ' empty cart!
      with objTemplate
      	.TemplateFile =  TMPLDIR & "displayemptycart.html"
      	.AddToken "header", 3, vCartHeader
      	.AddToken "footer", 3, vCartFooter

      	.parseTemplateFile
      end with

   End If

   ' Page is done. save and cleanup
   Session("Cart") = Cart.SaveCart
   Set Cart = Nothing

   'from old site. not sure yet what it's for
   Session("NavID") = ""
%>
