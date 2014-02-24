<!--#INCLUDE VIRTUAL="/includes/template_cls.asp" -->
<!--#INCLUDE VIRTUAL="/includes/common.asp" -->
<!--#INCLUDE VIRTUAL="/includes/cartconfig.asp" --><%
' get the template engine ready
dim pwdResetMsg, adjRebate, adjRebateTotal, adjRebateTotalAll
dim adCmdStoredProc, adVarChar, adInteger, adParamInput, adNumeric
adCmdStoredProc = 4
adVarChar = 200
adInteger = 3
adParamInput = 1
adNumeric = 131

call zeroRebateArray()
call getRebate()

Function newPWD()
    Randomize Timer
    newPWD = CStr(CInt(Rnd(20) * 1000)) & Mid("aljdflasjdfasdf", Int((10 * Rnd(50)) + 1), 2) & CStr(CInt(Rnd(400) * 1000))
End Function

sub saveCustomer(Conn2, vCustomerID)
    dim vSQL
    vSQL = "UPDATE Customers SET "
    vSQL = vSQL & "Name=" & CS(Cart.Info.Name, ",")
    vSQL = vSQL & "Company=" & CS(Cart.Info.Company, ",")
    vSQL = vSQL & "Address1=" & CS(Cart.Info.Address1, ",")
    vSQL = vSQL & "Address2=" & CS(Cart.Info.Address2, ",")
    vSQL = vSQL & "City=" & CS(Cart.Info.City, ",")
    vSQL = vSQL & "StateProvince=" & CS(Cart.Info.StateProvince, ",")
    vSQL = vSQL & "ZipPostal=" & CS(Cart.Info.ZipPostal, ",")
    vSQL = vSQL & "Country=" & CS(Cart.Info.Country, ",")
    vSQL = vSQL & "Phone=" & CS(Cart.Info.Phone, ",")
    vSQL = vSQL & "Fax=" & CS(Cart.Info.Fax, ",")
    vSQL = vSQL & "Email=" & CS(Cart.Info.Email, ",")
    vSQL = vSQL & "Custom1=" & CS(Cart.Info.Custom1, ",")
    vSQL = vSQL & "Custom2=" & CS(Cart.Info.Custom2, ",")
    vSQL = vSQL & "Custom3=" & CS(Cart.Info.Custom3, ",")
    vSQL = vSQL & "Custom4=" & CS(Cart.Info.Custom4, ",")
    vSQL = vSQL & "Custom5=" & CS(Cart.Info.Custom5, ",")
    vSQL = vSQL & "Custom6=" & CS(Cart.Info.Custom6, ",")
    vSQL = vSQL & "Custom7=" & CS(Cart.Info.Custom7, ",")
    vSQL = vSQL & "Custom8=" & CS(Cart.Info.Custom8, ",")

    vSQL = vSQL & "BillingSameAsShipping=" & Cart.Info.BillingSameAsShipping & ","

    vSQL = vSQL & "ShipName=" & CS(Cart.Info.ShipName, ",")
    vSQL = vSQL & "ShipCompany=" & CS(Cart.Info.ShipCompany, ",")
    vSQL = vSQL & "ShipAddress1=" & CS(Cart.Info.ShipAddress1, ",")
    vSQL = vSQL & "ShipAddress2=" & CS(Cart.Info.ShipAddress2, ",")
    vSQL = vSQL & "ShipCity=" & CS(Cart.Info.ShipCity, ",")
    vSQL = vSQL & "ShipStateProvince=" & CS(Cart.Info.ShipStateProvince, ",")
    vSQL = vSQL & "ShipZipPostal=" & CS(Cart.Info.ShipZipPostal, ",")
    vSQL = vSQL & "ShipCountry=" & CS(Cart.Info.ShipCountry, ",")
    vSQL = vSQL & "ShipPhone=" & CS(Cart.Info.ShipPhone, ",")
    vSQL = vSQL & "ShipFax=" & CS(Cart.Info.ShipFax, ",")
    vSQL = vSQL & "ShipEmail=" & CS(Cart.Info.ShipEmail, ",")
    vSQL = vSQL & "ShipCustom1=" & CS(Cart.Info.ShipCustom1, ",")
    vSQL = vSQL & "ShipCustom2=" & CS(Cart.Info.ShipCustom2, ",")
    vSQL = vSQL & "ShipCustom3=" & CS(Cart.Info.ShipCustom3, ",")
    vSQL = vSQL & "ShipCustom4=" & CS(Cart.Info.ShipCustom4, ",")
    vSQL = vSQL & "ShipCustom5=" & CS(Cart.Info.ShipCustom5, ",")
    vSQL = vSQL & "ShipCustom6=" & CS(Cart.Info.ShipCustom6, ",")
    vSQL = vSQL & "ShipCustom7=" & CS(Cart.Info.ShipCustom7, ",")
    vSQL = vSQL & "ShipCustom8=" & CS(Cart.Info.ShipCustom8, ",")

    vSQL = vSQL & "Resident=" & Cart.Info.IsStateResident &  ","
    vSQL = vSQL & "CountryResident=" & Cart.Info.IsCountryResident

    vSQL = vSQL & " WHERE CustomerID=" & vCustomerID
    Conn2.Execute(vSQL)
end sub

Sub SendEmailPWD( )
   dim msg,mEmail, mName
   dim mPWD, c, com, rs, p
   mPWD =  newPWD()
   mEmail = trim(Request("lemail"))
   if mEmail="" then
	pwdResetMsg = "<BR>YOU MUST RE-ENTER YOUR E-MAIL FIRST.<BR>" 
	vLoginErr="E-MAIL ADDRESS NOT TYPED"
	 exit sub
   end if
    Set c = Server.CreateObject("ADODB.Connection")
    set com = Server.CreateObject("ADODB.Command")
    
    c.Open "dsn=liidsn;uid=iiscart;pwd=iiscart"
     
    com.ActiveConnection = c
    com.CommandText = "getCustomer"
    com.CommandType = 4
    Set p = com.CreateParameter("@Email",200 , 1,50)
    p.value = trim(mEmail)
    com.Parameters.Append p
 
    Set p = com.CreateParameter("@pwd", 200, 1,50)
    p.value = trim(mPWD)
    com.Parameters.Append p

    Set p = com.CreateParameter("@setPassword", 3, 1)
    p.value = 1
    com.Parameters.Append p 
    
    set rs = com.execute
   
    if rs.eof then
       pwdResetMsg = "<BR><B>E-Mail address not found:  <BR>" & trim(mEmail) & "</B>"
    else
       mName = rs.fields("Name")
       msg = "NEW Password: " & mPWD
        
       eheader = "BICYCLEBUYS.COM PASSWORD REQUEST" & vbcrlf & vbcrlf

       eheader = eheader & "Date: " & Date & vbcrlf
       eheader = eheader & "Time: " & Time & vbcrlf & vbcrlf
       eheader = eheader & "Customer Name: " & mName & vbcrlf & vbcrlf

         
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
     
       ' ---- Send Email to Customer
           Cart.Mail.Host = "webserver"
       Cart.Mail.From = "Sales@BicycleBuys.com"
       Cart.Mail.FromName = "BicycleBuys.com Request"
       Cart.Mail.Subject = "BicycleBuys.com Password Reset"

       Cart.Mail.Body = eheader & vbcrlf &  msg & vbcrlf & vbcrlf & efoot

       ' Send it to the shipping email address too
        
       Cart.Mail.AddAddress  mEmail 
       Cart.Mail.Send ' send to buyer
       Cart.Mail.Reset
       pwdResetMsg = "<BR>An E-Mail has been sent to:<BR>" & mEmail
    end if
    rs.close
    set rs = nothing
    c.close
    set c = nothing
    
End Sub

Public Function saveCustomern()
    Dim c  
    Dim p  
    Dim com  
    Set c = Server.CreateObject("ADODB.Connection")
    set com = Server.CreateObject("ADODB.Command")
    
    c.Open "dsn=liidsn;uid=iiscart;pwd=iiscart"
     
    com.ActiveConnection = c
    com.CommandText = ""
    com.CommandType = adCmdStoredProc
    Set p = com.CreateParameter("@CustomerID", adInteger, adParamInput)
    com.Parameters.Append p
    Set p = com.CreateParameter("@Name", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Company", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Address1", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Address2", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@City", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@StateProvince", adVarChar, adParamInput, 20 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ZipPostal", adVarChar, adParamInput, 20 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Country", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Phone", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Fax", adVarChar, adParamInput, 20 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Email", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Custom1", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Custom2", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Custom3", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Custom4", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Custom5", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Custom6", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Custom7", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Custom8", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@BillingSameAsShipping", adNumeric, adParamInput)
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipName", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipCompany", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipAddress1", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipAddress2", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipCity", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipStateProvince", adVarChar, adParamInput, 20 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipZipPostal", adVarChar, adParamInput, 15 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipCountry", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipPhone", adVarChar, adParamInput, 20 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipFax", adVarChar, adParamInput, 20 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipEmail", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipCustom1", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipCustom2", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipCustom3", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipCustom4", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipCustom5", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipCustom6", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipCustom7", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@ShipCustom8", adVarChar, adParamInput, 50 )
    com.Parameters.Append p
    Set p = com.CreateParameter("@Resident", adNumeric, adParamInput)
    com.Parameters.Append p
    Set p = com.CreateParameter("@CountryResident", adNumeric, adParamInput)
    com.Parameters.Append p
    Set p = com.CreateParameter("@Password", adVarChar, adParamInput, 50 )
    com.Parameters.Append p

    com.parameters.item("CustomerID")=mCustomerID
    com.parameters.item("Name")="mName"
    com.parameters.item("Company")="mCompany"
    com.parameters.item("Address1")="mAddress1"
    com.parameters.item("Address2")="mAddress2"
    com.parameters.item("City")="mCity"
    com.parameters.item("StateProvince")="mStateProvince"
    com.parameters.item("ZipPostal")="mZipPostal"
    com.parameters.item("Country")="mCountry"
    com.parameters.item("Phone")="mPhone"
    com.parameters.item("Fax")="mFax"
    com.parameters.item("Email")="mEmail"
    com.parameters.item("Custom1")="mCustom1"
    com.parameters.item("Custom2")="mCustom2"
    com.parameters.item("Custom3")="mCustom3"
    com.parameters.item("Custom4")="mCustom4"
    com.parameters.item("Custom5")="mCustom5"
    com.parameters.item("Custom6")="mCustom6"
    com.parameters.item("Custom7")="mCustom7"
    com.parameters.item("Custom8")="mCustom8"
    com.parameters.item("BillingSameAsShipping")=mBillingSameAsShipping
    com.parameters.item("ShipName")="mShipName"
    com.parameters.item("ShipCompany")="mShipCompany"
    com.parameters.item("ShipAddress1")="mShipAddress1"
    com.parameters.item("ShipAddress2")="mShipAddress2"
    com.parameters.item("ShipCity")="mShipCity"
    com.parameters.item("ShipStateProvince")="mShipStateProvince"
    com.parameters.item("ShipZipPostal")="mShipZipPostal"
    com.parameters.item("ShipCountry")="mShipCountry"
    com.parameters.item("ShipPhone")="mShipPhone"
    com.parameters.item("ShipFax")="mShipFax"
    com.parameters.item("ShipEmail")="mShipEmail"
    com.parameters.item("ShipCustom1")="mShipCustom1"
    com.parameters.item("ShipCustom2")="mShipCustom2"
    com.parameters.item("ShipCustom3")="mShipCustom3"
    com.parameters.item("ShipCustom4")="mShipCustom4"
    com.parameters.item("ShipCustom5")="mShipCustom5"
    com.parameters.item("ShipCustom6")="mShipCustom6"
    com.parameters.item("ShipCustom7")="mShipCustom7"
    com.parameters.item("ShipCustom8")="mShipCustom8"
    com.parameters.item("Resident")=mResident
    com.parameters.item("CountryResident")=mCountryResident
    com.parameters.item("Password")="mPassword"

    com.Execute
End Function

set objTemplate = new template_cls

vOUT1 = ""
vOUT2 = ""
'   For each item in Cart.Items
'      response.write "<hr>IC4: " & Item.Custom4 & "<br>"
'   next

%><!--#INCLUDE VIRTUAL="/includes/cartdisplay_checkout.asp" --><%


' Put checkout and empty buttons on the display
' (disabled with < -1)
Dim vCheckEmpty, BillingErrString, vERR, vPlus4, vZipCode, vTMP, vSTFN, vSTMN, vSTLN, vCustomerID, LoginErrString
vCheckEmpty = ""
if cart.gridtotalquantity > 0 then
    vCheckEmpty = "<a href=""" &  vThisProto & vThisServer & "/checkout/""><img src=""/cartimages/checkout.gif"" alt=""Check out"" border=""0"" WIDTH=""100"" HEIGHT=""20""></a>" _
    & "<a href=""" &  vThisProto & vThisServer & "/emptycart/""><img src=""/cartimages/emptycart.gif"" alt=""Remove ALL items from the cart"" border=""0"" WIDTH=""100"" HEIGHT=""20""></a>"
end if

' commented out cuz it was doubling
'Cart.LoadCart(Session("Cart"))

Dim vOriginalValue, vLoginErr, vLoginErrString, vBillingErrString, vShippingErrString
Dim vShipAddress1, vPhoney, vPhoneyShip

if  IsNull(Cart.Info.ShipAddress1) then
    vShipAddress1=""
else
    vShipAddress1=trim(Cart.Info.ShipAddress1)
end if
'response.Write("xxx" & Request.Form("txtreset") & "XXX")
if Request.Form("txtreset") = "1" then
     SendEmailPWD()
     vLoginErrString = pwdResetMsg
end if
If Request("login") <> "" or  Request("txtreset")=1 and Request("lemail")="" then
'    response.Write(Request("lemail") & " " &  Request("lpassword") )
    vLoginErr = Cart.LoadCustomerDB(Request("lemail"), Request("lpassword"))

    if vLoginErr or Request("lemail")=""  then
        vLoginErrString = "<hr>Invalid email address or password.<BR>If you would like to reset your password <BR>RE-TYPE your e-mail address below then click Reset.<BR> <input type=hidden name=txtreset value=0><input type=""button"" name=""btnReset"" value=""Reset Password"" onclick=""javascript:document.myform.txtreset.value='1';document.myform.submit();"">"           
    end if
    if Not IsNumeric(Cart.Info.ShipCustom8) Then Cart.Info.ShipCustom8 = 0
End if

' save the entered info if login isnt happening
If Request("Save") <> "" and Request("login")="" Then

    Cart.Validate "address1,city,zippostal,country,phone,email"
    vBillingErrString = Cart.AcceptBillingInfo


    dim vBTFN, vBTMN, vBTLN
    ' get fn, mn, ln - put em back together and generate an error if need be
    vBTFN = request("FIRSTNAME")
    vBTMN = request("MIDDLENAME")
    vBTLN = request("LASTNAME")

    ' put these form values into cart before display mods happen
    ' billto
    Cart.Info.Custom3 = vBTFN
    Cart.Info.Custom4 = vBTMN
    Cart.Info.Custom5 = vBTLN

    ' put a space between fn/mn/ln for carts need for fullname
    if vBTMN <> "" Then vBTMN = " " & vBTMN
    if vBTLN <> "" Then vBTLN = " " & vBTLN

    ' put into cart as full name
    Cart.Info.Name = vBTFN & vBTMN & vBTLN

    ' if first or last name are blank issue an error
    if vBTLN = "" Then vBillingErrString = "LAST NAME; " &  vBillingErrString
    if vBTFN = "" Then vBillingErrString = "FIRST NAME; " &  vBillingErrString

    ' errors will occur on these cart required fields, but not required by us
    ' strip those errors out
    vBillingErrString = Replace(vBillingErrString, "CREDIT CARD TYPE; CREDIT CARD NUMBER; NAME ON CREDIT CARD; CREDIT CARD EXPIRATION MONTH; CREDIT CARD EXPIRATION YEAR; ", "")



    ' if the user enters US for country and nothing or other for state.
    if Cart.Info.Country = "US" and (Cart.Info.StateProvince = "NONE" or Cart.Info.StateProvince = "OTHER") Then vBillingErrString = vBillingErrString & "STATE; "

    ' if the user enters other for country but doesn't enter a province (custom1), display an error.
    if Cart.Info.Country = "OTHER" and (Cart.Info.StateProvince = "OTHER" or Cart.Info.StateProvince = "NONE") and (Cart.Info.Custom1="" or IsEmpty(Cart.Info.Custom1) or IsNULL(Cart.Info.Custom1)) Then vBillingErrString = vBillingErrString & "PROVINCE; "

    ' If the user enters "OTHER" for country, types in a country but sets state to anything but "OTHER", set the state to "OTHER" and display an error.
    if Cart.Info.Country = "OTHER" and Cart.Info.StateProvince <> "OTHER" then
        Cart.Info.StateProvince = "OTHER"
    '      BillingErrString = BillingErrString & "STATE SET TO ""OTHER"";"
    End If

    ' if the user enters other for country but doesn't pick a country (custom2), display an error.
    if Cart.Info.Country = "OTHER" then
    if (Cart.Info.Custom2="" or IsEmpty(Cart.Info.Custom2) or IsNULL(Cart.Info.Custom2)) Then
        vBillingErrString = vBillingErrString & "COUNTRY; "
    end if
    if (Cart.Info.ZipPostal) = "" then
        vBillingErrString = vBillingErrString & "PLEASE ENTER 'NONE' IF YOU DON'T HAVE A POSTAL CODE; "
    end if
end if

'''new as of05/15/07
if Cart.Info.Country = "US" then
    if Cart.Info.StateProvince = "ZZ" then
        vBillingErrString = vBillingErrString & "PLEASE Choose a state; "
    end if
end if
if Cart.Info.ShipCountry = "US" then
    if Cart.Info.ShipStateProvince = "ZZ" then
        vShippingErrString = vShippingErrString & "PLEASE Choose a shipping state; "
    end if
end if
' if user enters US as county, check format of zip code to make sure it's ##### or #####-####
' 1. clean up string, remove all spaces

'check email
'dim x, y, z
z = Cart.Info.Email
x = InStr(z,"@")
y = InStr(z,".")
if (x < 1) OR (y < 1)  then vBillingErrString = vBillingErrString & "Your Email is not formatted properly; "
z = request("ShipEmail")
x = InStr(z,"@")
y = InStr(z,".")
if (x < 1) OR (y < 1)  then vShippingErrString = vShippingErrString & "Your Shipping Email is not formatted properly; "

if Cart.Info.Country = "US" Then
    vERR = False
    vPlus4 = False

    vZipCode = Trim(Replace(Cart.Info.ZipPostal, "  ", " "))

    ' first 5 chars must be numeric
    vTMP = Left(vZipCode, 5)
    if NOT isnumeric(vTMP) OR len(vTMP) <> 5 Then
        vERR = True
    end if

    ' validate + 4 entries
    ' whole string must be 10 chars (xxxxx-xxxx)
    if len(vZipCode) = 10 AND vERR = False Then
        if instr(vZipCode, "-") > 0 Then
            ' found a +4 entry
            vTMP = Right(vZipCode, 4)
            ' make sure the +4 is numeric
            if NOT isnumeric(vTMP) Then
                vERR = True
            else
                vPlus4 = True
            end if
        end if
    end if

' make sure zip code is just zip code by checking its length and numeric-ness
if (len(vZipCode) > 5 AND NOT vPlus4) AND vERR = False Then vERR = True
    if vERR Then vBillingErrString = vBillingErrString & "INVALID ZIP CODE;"
    elseif Cart.Info.Country = "Canada" Then
        vERR = False
        vPlus4 = False
        if (len(Trim(Replace(Cart.Info.ZipPostal, "  ", " "))) < 6) then vBillingErrString = vBillingErrString & "INVALID ZIP CODE;"
    end if

    '''check phones
    vPhoney = trim(Cart.Info.Phone)
    vPhoneyShip = trim(request("ShipPhone"))
    if len(vPhoney) < 7 or not IsNumeric(vPhoney) then vBillingErrString = vBillingErrString & "INVALID Phone #;"
    if len(vPhoneyShip) < 7 or not IsNumeric(vPhoneyShip) then vShippingErrString = vShippingErrString & "INVALID Ship Phone #;"

    ' If password entries don't match...
    if Request("password1") <> Request("password2") then
        vBillingErrString = vBillingErrString & "PASSWORDS DON'T MATCH; "
    End If

    ' if there was an error, wrap that error in html
    if vBillingErrString <> "" Then vBillingErrString = "<hr><b>The following field(s) are required:</b> " & vBillingErrString

    ' handle shipping portion of form
    vShipSame = Request.Form("ORDERCUSTOM8")
    if vShipSame = "" then vShipSame = "OFF"
    Cart.Info.OrderCustom8 = vShipSame

    ' set shipping address same as billing if the check the same box
    if Cart.Info.OrderCustom8="ON" then
        Cart.Info.ShipName = Cart.Info.Name
        Cart.Info.ShipCompany = Cart.Info.Company
        Cart.Info.ShipAddress1 = Cart.Info.Address1
        Cart.Info.ShipAddress2 = Cart.Info.Address2
        Cart.Info.ShipCity = Cart.Info.City
        Cart.Info.ShipStateProvince = Cart.Info.StateProvince
        Cart.Info.ShipZipPostal = Cart.Info.ZipPostal
        Cart.Info.ShipCountry = Cart.Info.Country
        Cart.Info.ShipPhone = Cart.Info.Phone
        Cart.Info.ShipFax = Cart.Info.Fax
        Cart.Info.ShipEmail = Cart.Info.Email
        Cart.Info.ShipCustom1 = Cart.Info.Custom1
        Cart.Info.ShipCustom2 = Cart.Info.Custom2

        ' fn/mn/ln mods
        Cart.Info.ShipCustom3 = Cart.Info.Custom3
        Cart.Info.ShipCustom4 = Cart.Info.Custom4
        Cart.Info.ShipCustom5 = Cart.Info.Custom5

    Else
        ' shipping address not same as billing, validate it
        Cart.Validate "address1,city,zippostal,country,phone"
        vShippingErrString = vShippingErrString & Cart.AcceptShippingInfo

        ' if the user enters US for country and nothing or other for state then STATE is in error
        if Cart.Info.ShipCountry = "US" and (Cart.Info.ShipStateProvince = "NONE" or Cart.Info.ShipStateProvince = "OTHER") Then vShippingErrString = vShippingErrString & "STATE; "

        ' If the user enters "OTHER" for state, types in a province but sets country to "US",
        '    set the country to "OTHER" and display an error.
        if Cart.Info.ShipStateProvince = "OTHER" and (Cart.Info.ShipCustom1<>"" and NOT IsEmpty(Cart.Info.ShipCustom1) and NOT IsNULL(Cart.Info.ShipCustom1)) and Cart.Info.ShipCountry <> "OTHER" then
            Cart.Info.ShipCountry = "OTHER"
        End If

        ' if the user enters "OTHER" for state but doesn't enter a province (shipcustom1), display PROVINCE error.
        if Cart.Info.ShipCountry = "OTHER" and (Cart.Info.ShipCustom1="" or IsEmpty(Cart.Info.ShipCustom1) or IsNULL(Cart.Info.ShipCustom1)) Then vShippingErrString = vShippingErrString & "PROVINCE; "

        ' if the user enters "OTHER" for country but doesn't pick a country (shipcustom2), display COUNTRY error.
        if Cart.Info.ShipCountry = "OTHER" and (Cart.Info.ShipCustom2="" or IsEmpty(Cart.Info.ShipCustom2) or IsNULL(Cart.Info.ShipCustom2)) Then vShippingErrString = vShippingErrString & "COUNTRY; "

        ' If the user enters "OTHER" for country, picks a country but sets state to anything but "OTHER", set the state to "OTHER" and display an error.
        if Cart.Info.ShipCountry = "OTHER" and (Cart.Info.ShipCustom2 <> "" and NOT IsEmpty(Cart.Info.ShipCustom2) and NOT IsNULL(Cart.Info.ShipCustom2)) and Cart.Info.ShipStateProvince <> "OTHER" then
            Cart.Info.ShipStateProvince = "OTHER"
        End If


        ' if the user enters US for country and nothing or other for state.
        if Cart.Info.ShipCountry = "US" and (Cart.Info.ShipStateProvince = "NONE" or Cart.Info.ShipStateProvince = "OTHER") Then vShippingErrString = vShippingErrString & "STATE; "

        ' If the user enters "OTHER" for state, types in a province but sets country to "US", set the country to "OTHER" and display an error.
        if Cart.Info.ShipStateProvince = "OTHER" and (Cart.Info.ShipCustom1<>"" and NOT IsEmpty(Cart.Info.ShipCustom1) and NOT IsNULL(Cart.Info.ShipCustom1)) and Cart.Info.ShipCountry <> "OTHER" then
            Cart.Info.ShipCountry = "OTHER"
            '         ShippingErrString = ShippingErrString & "COUNTRY SET TO ""OTHER""; "
        End If

        ' if the user enters other for state but doesn't enter a province (shipcustom1), display an error.
        if Cart.Info.ShipCountry = "OTHER" and (Cart.Info.ShipCustom1="" or IsEmpty(Cart.Info.ShipCustom1) or IsNULL(Cart.Info.ShipCustom1)) Then vShippingErrString = vShippingErrString & "PROVINCE; "

        ' if the user enters other for country but doesn't pick a country (SHIPCOUNTRY), display an error.
        if Cart.Info.ShipCountry = "OTHER" and (Cart.Info.SHIPCOUNTRY="" or IsEmpty(Cart.Info.SHIPCOUNTRY) or IsNULL(Cart.Info.SHIPCOUNTRY)) Then vShippingErrString = vShippingErrString & "COUNTRY; "

        ' If the user enters "OTHER" for country, picks a country but sets state to anything but "OTHER", set the state to "OTHER" and display an error.
        if Cart.Info.ShipCountry = "OTHER" and (Cart.Info.SHIPCOUNTRY <> "" and NOT IsEmpty(Cart.Info.SHIPCOUNTRY) and NOT IsNULL(Cart.Info.SHIPCOUNTRY)) and Cart.Info.ShipStateProvince <> "OTHER" then
            Cart.Info.ShipStateProvince = "OTHER"
        '         ShippingErrString = ShippingErrString & "STATE SET TO ""OTHER""; "
        End If

        ' get names from form
        vSTFN = request("SHIPFIRSTNAME")
        vSTMN = request("SHIPMIDDLENAME")
        vSTLN = request("SHIPLASTNAME")

        ' shipto
        Cart.Info.ShipCustom3 = vSTFN
        Cart.Info.ShipCustom4 = vSTMN
        Cart.Info.ShipCustom5 = vSTLN
        Cart.Info.ShipName = trim(vSTFN) & " " & trim(vSTMN) & " " & trim(vSTLN)

        if len(Cart.Info.ShipName) < 1 then vShippingErrString = "No ship name NAME; " &  vShippingErrString
        if vSTLN = "" Then vShippingErrString = "SHIP-TO LAST NAME; " &  vShippingErrString
        if vSTFN = "" Then vShippingErrString = "SHIP-TO FIRST NAME; " &  vShippingErrString


    End If

' if browser is in this state then make sure we collect tax
    if Cart.Info.ShipStateProvince = vThisState then
        Cart.Info.IsStateResident = TRUE
        Cart.ShippingTaxRate = vStateTaxRate
    else
        Cart.Info.IsStateResident = 0
        Cart.ShippingTaxRate = "0%"
    End If

    if IsNumeric(Cart.Info.ShipCustom8) Then
        vShipType = Cart.ShippingType & ":" & Cart.Info.ShipCustom8
    else
        vShipType = "NONE:0"
    End If

    vShipTypeA = Split(vShipType, ":")
    Cart.ShippingType = vShipTypeA(0)
    Cart.Info.ShipCustom8 = vShipTypeA(1)

    ' make sure the comment fields are <=1000 in length
    Cart.Info.Comments = Left(Cart.Info.Comments, 1000)
    Cart.Info.ShipComments = Left(Cart.Info.ShipComments, 1000)

    ' if there was an error, wrap that error in html
    if vShippingErrString <> "" Then vShippingErrString = "<hr><b>The following field(s) are required:</b> " & vShippingErrString


    ' if user enters US as county, check format of zip code to make sure it's ##### or #####-####
    ' 1. clean up string, remove all spaces
    if Cart.Info.ShipCountry = "US" Then
        vERR = False
        vPlus4 = False
        vZipCode = Trim(Replace(Cart.Info.ShipZipPostal, "  ", " "))
        ' first 5 chars must be numeric
        vTMP = Left(vZipCode, 5)
        if NOT isnumeric(vTMP) Then
            vERR = True
        end if

        ' validate + 4 entries
        ' whole string must be 10 chars (xxxxx-xxxx)
        if len(vZipCode) = 10 AND vERR = False Then
            if instr(vZipCode, "-") > 0 Then
                ' found a +4 entry
                vTMP = Right(vZipCode, 4)
                ' make sure the +4 is numeric
                if NOT isnumeric(vTMP) Then
                    vERR = True
                else
                    vPlus4 = True
                end if
            end if
        end if

        ' make sure zip code is just zip code by checking its length and numeric-ness
        if (len(vZipCode) > 5 AND NOT vPlus4) AND vERR = False Then vERR = True
            if vERR Then vShippingErrString = vShippingErrString & "INVALID ZIP CODE;"
        elseif Cart.Info.ShipCountry = "Canada" Then
            vERR = False
            vPlus4 = False
            if (len(Trim(Replace(Cart.Info.ShipZipPostal, "  ", " "))) < 6) then vShippingErrString = vShippingErrString & "INVALID ZIP CODE;"
        end if
        dim sql, conn2
        Set Conn2 = Server.CreateObject("ADODB.Connection")
        Conn2.Open "dsn=liidsn;uid=iiscart;pwd=iiscart"
        sql = "exec getCustomer " &  trim(CS(Cart.Info.Email, "")) & "," & trim(CS(Request("password2"),""))  & ", 0"  ' SELECT * FROM Customers WHERE Email LIKE " & trim(CS(Cart.Info.Email, "")) & " FOR BROWSE;"            
        Set rs = Conn2.Execute(sql)
        if not rs.eof then 
            vCustomerID = rs("CustomerID")  
        else
            vCustomerID = -1
        end if
        ' save the cart to the session
	Cart.Info.OrderCustom7 = Request.ServerVariables("REMOTE_ADDR") 
	Cart.Info.OrderCustom6 = left(replace(replace(replace(Session("ReferredBy"),"'",""),"http://www.bicyclebuys.com/","B/"),"http://www.google.com/","G/"),50)
        Session("Cart") = Cart.SaveCart
        if (trim(Cart.Info.Address1)<>"" and Cart.Info.Address1 <> Cart.info.ShipAddress1 and trim(Cart.info.EMail)<>"") _
            or ( Request("password1") <> "" AND (vBillingErrString="" AND vShippingErrString= "")) then
            if (Request("password1") <> "") then
                Cart.Info.Password = Request("password1")
            else
                Cart.Info.Password = "xxJuNkYpWxx"
            end if

            ' make connection to db for save
           
            
            ' find other customers with same email address

            dim LoginErr

            ' This is a new email address, so we can save it.
            ' using cs function to keep data clean
            if rs.EOF then
                vSQL = "INSERT INTO Customers (Name, Company, Address1, Address2, City, StateProvince, ZipPostal, Country, Phone, Fax, Email, Custom1, Custom2, Custom3, Custom4, Custom5, Custom6, Custom7, Custom8, BillingSameAsShipping, "
                vSQL = vSQL & "ShipName, ShipCompany, ShipAddress1, ShipAddress2, ShipCity, ShipStateProvince, ShipZipPostal, ShipCountry, ShipPhone, ShipFax, ShipEmail, ShipCustom1, ShipCustom2, ShipCustom3, ShipCustom4, ShipCustom5, ShipCustom6, ShipCustom7, ShipCustom8, "
                vSQL = vSQL & "Resident, CountryResident, Password) "

                vSQL = vSQL & "VALUES ("
                vSQL = vSQL & CS(Cart.Info.Name, ",")
                vSQL = vSQL & CS(Cart.Info.Company, ",")
                vSQL = vSQL & CS(Cart.Info.Address1, ",")
                vSQL = vSQL & CS(Cart.Info.Address2, ",")
                vSQL = vSQL & CS(Cart.Info.City, ",")
                vSQL = vSQL & CS(Cart.Info.StateProvince, ",")
                vSQL = vSQL & CS(Cart.Info.ZipPostal, ",")
                vSQL = vSQL & CS(Cart.Info.Country, ",")
                vSQL = vSQL & CS(Cart.Info.Phone, ",")
                vSQL = vSQL & CS(Cart.Info.Fax, ",")
                vSQL = vSQL & CS(Cart.Info.Email, ",")
                vSQL = vSQL & CS(Cart.Info.Custom1, ",")
                vSQL = vSQL & CS(Cart.Info.Custom2, ",")
                vSQL = vSQL & CS(Cart.Info.Custom3, ",")
                vSQL = vSQL & CS(Cart.Info.Custom4, ",")
                vSQL = vSQL & CS(Cart.Info.Custom5, ",")
                vSQL = vSQL & CS(Cart.Info.Custom6, ",")
                vSQL = vSQL & CS(Cart.Info.Custom7, ",")
                vSQL = vSQL & CS(Cart.Info.Custom8, ",")

                vSQL = vSQL & Cart.Info.BillingSameAsShipping & ","

                vSQL = vSQL & CS(Cart.info.ShipName, ",")
                vSQL = vSQL & CS(Cart.info.ShipCompany, ",")
                vSQL = vSQL & CS(Cart.info.ShipAddress1, ",")
                vSQL = vSQL & CS(Cart.info.ShipAddress2, ",")
                vSQL = vSQL & CS(Cart.info.ShipCity, ",")
                vSQL = vSQL & CS(Cart.info.ShipStateProvince, ",")
                vSQL = vSQL & CS(Cart.info.ShipZipPostal, ",")
                vSQL = vSQL & CS(Cart.info.ShipCountry, ",")
                vSQL = vSQL & CS(Cart.info.ShipPhone, ",")
                vSQL = vSQL & CS(Cart.info.ShipFax, ",")
                vSQL = vSQL & CS(Cart.info.ShipEmail, ",")
                vSQL = vSQL & CS(Cart.info.ShipCustom1, ",")
                vSQL = vSQL & CS(Cart.info.ShipCustom2, ",")
                vSQL = vSQL & CS(Cart.info.ShipCustom3, ",")
                vSQL = vSQL & CS(Cart.info.ShipCustom4, ",")
                vSQL = vSQL & CS(Cart.info.ShipCustom5, ",")
                vSQL = vSQL & CS(Cart.info.ShipCustom6, ",")
                vSQL = vSQL & CS(Cart.info.ShipCustom7, ",")
                vSQL = vSQL & CS(Cart.info.ShipCustom8, ",")

                vSQL = vSQL & Cart.Info.IsStateResident &  ","
                vSQL = vSQL & Cart.Info.IsCountryResident & ","

                vSQL = vSQL & CS(Cart.Info.Password, "")
                vSQL = vSQL & ")"
                Conn2.Execute(vSQL)

                LoginErr = Cart.LoadCustomerDB(Cart.Info.Email, Cart.Info.Password)

                if LoginErr then
                    vLoginErrString = "Invalid email address or password. Confirm e-mail address. <input type=""submit"" name=""reset"" value=""Send Password"">"
                end if
                if Not IsNumeric(Cart.Info.ShipCustom8) Then Cart.Info.ShipCustom8 = 0
		Cart.Info.OrderCustom7 = Request.ServerVariables("REMOTE_ADDR") 
                Session("Cart") = Cart.SaveCart
            Else
               
                if (trim(Cart.Info.Address1)<>""  and trim(Cart.info.EMail)<>"") _
                    or ((request("password1") = rs("password") and request("password2") = rs("password"))) then                    
                    call saveCustomer(Conn2,vCustomerID)

                    LoginErr = Cart.LoadCustomerDB(Cart.Info.Email, Cart.Info.Password)

                    if LoginErr then
                        LoginErrString = "Invalid email address or password. Confirm e-mail address. <input type=""submit"" name=""reset"" value=""Send Password"">"
                    end if
                    if Not IsNumeric(Cart.Info.ShipCustom8) Then Cart.Info.ShipCustom8 = 0
			Cart.Info.OrderCustom7 = Request.ServerVariables("REMOTE_ADDR") 
                        Session("Cart") = Cart.SaveCart
                    Else
                        vBillingErrString = vBillingErrString & "If experiencing problems, please continue checkout without setting a password; "
                    End If
                End If                
            Else
                ' this customer wants nothing to do with passwords.
                ' so we set the pw to this junk now so we don't have to
                ' do it later when it's inconvenient since we still want this
                ' customers info saved to the db, just not with a blank pw.
                Cart.Info.Password = "xxJuNkYpWxx"
		Cart.Info.OrderCustom7 = Request.ServerVariables("REMOTE_ADDR") 
                Session("Cart") = Cart.SaveCart
            End If

            Cart.CalculateShipping
            Cart.Calculate
            'response.write "here:" & vBillingErrString & " : " & vShippingErrString
            ' if there are no errors then we're done -- move on to next page
            if vBillingErrString="" AND vShippingErrString= "" Then
		Cart.Info.OrderCustom7 = Request.ServerVariables("REMOTE_ADDR") 
                Session("Cart") = Cart.SaveCart
                call saveCustomer(Conn2,vCustomerID)
                rs.close
                set rs = nothing
                Conn2.Close
                set Conn2 = nothing
                Response.Redirect "/ship/"
            else
                rs.close
                set rs = nothing
                Conn2.Close
                set Conn2 = nothing            
            End If
        End if


        ''''''''''''''''''''''''''''''
        ' set up tmp vars for display
        ''''''''''''''''''''''''''''''

        vTMP1="": vTMP2="": vTMP3="": vTMP4="": vTMP5=""
        vTMP6="": vTMP7="": vTMP8="": vTMP9="": vTMP10=""

        ' US Selected
        if Cart.Info.Country="US" then vTMP1 = " SELECTED"

        ' Other Selected
        if Cart.Info.Country="OTHER" then vTMP2 = " SELECTED"

        ' State...  disable if country=other
        if Cart.Info.Country = "OTHER" Then vTMP3 = "disabled"

        ' Select OTHER state if state=other
        If Cart.Info.StateProvince = "OTHER" then vTMP4 = " SELECTED"

        For Each vState in vStates.Keys
            vSelected = ""
            If vState = Cart.Info.StateProvince then vSelected = " SELECTED"
            vTMP5 = vTMP5 & " <option value=""" & vState & """" & vSelected  & ">" & vStates.Item(vState) & "</option>" & chr(13)
        Next

        ' Country
        vCountry = Application("Country")
        vCountryCount = Application("CountryCount")

        if Cart.Info.Country = "US" OR Cart.Info.Country = "" Then
            vTMP6 = "disabled"
            vTMP7 = "International Only"
        else
            vTMP7 = Cart.Info.Custom1
        end if

        vTMP8 = ""
        for x = 0 to vCountryCount - 1
            vSelected = ""
            If Cart.Info.Custom2 = vCountry(x) then vSelected = " SELECTED"
            vTMP8 = vTMP8 & " <option value=""" & vCountry(x) & """ " & vSelected & ">" & vCountry(x) & "</option>" & Chr(13)
        next

        ' need checkbox set for shipping same as billin
        If Cart.Info.OrderCustom8 <> "OFF" Then vTMP9=" CHECKED"

        ' US Selected
        if Cart.Info.ShipCountry="US" then vTMP10 = " SELECTED"

        ' Other Selected
        if Cart.Info.ShipCountry="OTHER" then vTMP11 = " SELECTED"

        ' State...  disable if country=other
        if Cart.Info.ShipCountry = "OTHER" Then vTMP12 = "disabled"

        ' Select OTHER state if state=other
        If Cart.Info.ShipStateProvince = "OTHER" then vTMP13 = " SELECTED"

        For Each vState in vStates.Keys
            vSelected = ""
            If vState = Cart.Info.ShipStateProvince then vSelected = " SELECTED"
            vTMP14 = vTMP14 & " <option value=""" & vState & """" & vSelected  & ">" & vStates.Item(vState) & "</option>" & chr(13)
        Next

        ' Country
        if Cart.Info.ShipCountry = "US" OR Cart.Info.ShipCountry = "" Then
            vTMP15 = "disabled"
            vTMP16 = "International Only"
        else
            vTMP16 = Cart.Info.ShipCustom1
        end if

        vTMP8 = ""
        for x = 0 to vCountryCount - 1
            vSelected = ""
            If Cart.Info.Country = vCountry(x) then vSelected = " SELECTED"
            vTMP8 = vTMP8 & "               <option value=""" & vCountry(x) & """ " & vSelected & ">" & vCountry(x) & "</option>" & Chr(13)
        next
        vTMP17 = ""
        for x = 0 to vCountryCount - 1
            vSelected = ""
            If Cart.Info.ShipCountry = vCountry(x) then vSelected = " SELECTED"
            vTMP17 = vTMP17 & "               <option value=""" & vCountry(x) & """ " & vSelected & ">" & vCountry(x) & "</option>" & Chr(13)
        next


        vBTFN = Cart.Info.Custom3
        vBTMN = Cart.Info.Custom4
        vBTLN = Cart.Info.Custom5

        vSTFN = Cart.Info.ShipCustom3
        vSTMN = Cart.Info.ShipCustom4
        vSTLN = Cart.Info.ShipCustom5

        ' cart display built, now show it
        with objTemplate
        .TemplateFile = TMPLDIR & "checkout4.html"

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
                  adjRebateTotal=adjRebateTotal+adjRebate
                end if 
           next
          .AddToken "adjusttotal", 1, rebateHtml(adjRebateTotal) 
         
            
        
       
       
        .AddToken "displaycart", 1, vOUT2
        .AddToken "thisserver", 1, vThisServer
        .AddToken "checkempty", 1, vCheckEmpty
        .AddToken "sessionreferer", 1, Session("Referer")
        .AddToken "InfoFIRSTNAME", 1, vBTFN
        .AddToken "InfoMIDDLENAME", 1, vBTMN
        .AddToken "InfoLASTNAME", 1, vBTLN
        .AddToken "InfoCompany", 1, Cart.Info.Company
        .AddToken "USSelect", 1, vTMP1
        .AddToken "OTHERSelect", 1, vTMP2
        .AddToken "InfoCompany", 1, Cart.Info.Company
        .AddToken "StateDisabled", 1, vTMP3
        .AddToken "OtherStateSelected", 1, vTMP4
        .AddToken "StateSelect", 1, vTMP5
        .AddToken "Address1", 1, Cart.Info.Address1
        .AddToken "ZipPostal", 1, Cart.Info.ZipPostal
        .AddToken "City", 1, Cart.Info.City
        .AddToken "Email", 1, Cart.Info.Email
        .AddToken "Phone", 1, Cart.Info.Phone
        .AddToken "Comments", 1, Cart.Info.Comments
        .AddToken "Fax", 1, Cart.Info.Fax
        .AddToken "CountryDisabled", 1, vTMP6
        .AddToken "Country", 1, vTMP7
        .AddToken "Countries", 1, vTMP8

        .AddToken "shipsamechecked", 1, vTMP9
        .AddToken "ShipFirstName", 1, vSTFN
        .AddToken "ShipMiddleName", 1, vSTMN
        .AddToken "ShipLastName", 1, vSTLN
        .AddToken "ShipCompany", 1, Cart.Info.ShipCompany
        .AddToken "ShipUSSelect", 1, vTMP10
        .AddToken "ShipOTHERSelect", 1, vTMP11
        .AddToken "ShipInfoCompany", 1, Cart.Info.ShipCompany
        .AddToken "ShipStateDisabled", 1, vTMP12
        .AddToken "ShipOtherStateSelected", 1, vTMP13
        .AddToken "ShipStateSelect", 1, vTMP14
        .AddToken "ShipAddress1", 1, Cart.Info.ShipAddress1
        .AddToken "ShipZipPostal", 1, Cart.Info.ShipZipPostal
        .AddToken "ShipCity", 1, Cart.Info.ShipCity
        .AddToken "ShipEmail", 1, Cart.Info.ShipEmail
        .AddToken "ShipPhone", 1, Cart.Info.ShipPhone
        .AddToken "ShipComments", 1, Cart.Info.ShipComments
        .AddToken "ShipFax", 1, Cart.Info.ShipFax
        .AddToken "ShipCountryDisabled", 1, vTMP15
        ' .AddToken "ShipCountry", 1, Cart.Info.ShipCountry
        .AddToken "ShipCountry", 1, vTMP16
        .AddToken "ShipCountries", 1, vTMP17

        '         .AddToken "", 1, Cart.Info.

        ' .AddToken "continueshopping", 1, Session("Referer")
        .AddToken "continueshopping", 1, "/"

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

        if (vLoginErrString <> "") then
            vLoginErrString = "<b><font color=""#CC0000"">" & vLoginErrString &  "</font></b><br><br>"
        end if
        if (vBillingErrString <> "") then
            vBillingErrString = "<b><font color=""#CC0000"">" & vBillingErrString &  "</font></b><br><br>"
        end if
        if (vShippingErrString <> "") then
            vShippingErrString = "<b><font color=""#CC0000"">" & vShippingErrString &  "</font></b><br><br>"
        end if

        .AddToken "loginerrstring", 1, vLoginErrString
        .AddToken "billingerrstring", 1, vBillingErrString
        .AddToken "shippingerrstring", 1, vShippingErrString

        .AddToken "header", 3, vCartHeaderSummaryCheckout
        .AddToken "footer", 3, vCartFooterNoHelp
        
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
        
        .AddToken "adjusttotal", 1, rebateHtml(vAdjustTotal)   
             
        .parseTemplateFile
        end with
        
End If

' Page is done. save and cleanup
Cart.Info.OrderCustom7 = Request.ServerVariables("REMOTE_ADDR") 
Session("Cart") = Cart.SaveCart
Set Cart = Nothing

'from old site. not sure yet what it's for
Session("NavID") = "" 
%>
