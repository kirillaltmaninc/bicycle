<%


	'---- encryption key info
	Dim g_KeyLocation, g_Key, g_KeyString
	g_KeyLocation = "D:\root\new.bicyclebuys.com\crypt\crypt_key.txt"
	g_KeyString = ReadKeyFromFile(g_KeyLocation)

	Function EnCrypt(strCryptThis)
	   Dim strChar, iKeyChar, iStringChar, i, iCryptChar, strEncrypted
	   for i = 1 to Len(strCryptThis)
		  iKeyChar = Asc(mid(g_Key,i,1))
		  iStringChar = Asc(mid(strCryptThis,i,1))
		  iCryptChar = iKeyChar Xor iStringChar
		  strEncrypted =  strEncrypted & Chr(iCryptChar)
	   next
	   EnCrypt = strEncrypted
	End Function

	Function DeCrypt(strEncrypted)
	Dim strChar, iKeyChar, iStringChar, i, iDeCryptChar, strDecrypted
	   for i = 1 to Len(strEncrypted)
		  iKeyChar = (Asc(mid(g_Key,i,1)))
		  iStringChar = Asc(mid(strEncrypted,i,1))
		  iDeCryptChar = iKeyChar Xor iStringChar
		  strDecrypted =  strDecrypted & Chr(iDeCryptChar)
	   next
	   DeCrypt = strDecrypted
	End Function

	Function ReadKeyFromFile(strFileName)
	   Dim keyFile, fso, f, ts
	   set fso = Server.CreateObject("Scripting.FileSystemObject")
	   set f = fso.GetFile(strFileName)
	   set ts = f.OpenAsTextStream(1, -2)

	   Do While not ts.AtEndOfStream
		 keyFile = keyFile & ts.ReadLine
	   Loop

	   ReadKeyFromFile =  keyFile
	End Function

	dim FileDSN
	dim conn
	dim rs
	dim OrderID
	OrderID = request.querystring("OrderID")
	PaymentID = request.querystring("PaymentID")
	FileDSN = "Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webuserprod;Initial Catalog=BBC_PROD;Data Source=webserver"
	'FileDSN = "Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webuserprod;Initial Catalog=BBC_PROD;Data Source=deteljToshibaxp"
	'FileDSN = "DSN=bicyclebuys;Password=bbcwebUserprod;User ID=webuserprod;"

	Set Conn = Server.CreateObject("ADODB.Connection")

	Conn.Open FileDSN
	'sql= "SELECT Orders.OrderID, Customers.CustFN, Customers.CustMI, Customers.CustLN, Customers.CustAddr, Customers.CustCity, Customers.CustStateOrProvince, Customers.CustPostalCode, Customers.CustPhone, Customers.CustEmail, Customers.CustCountry"
	'sql = sql & ", Orders.ShipName, Orders.ShipCompanyName, "
	'sql = sql & " Orders.ShipAddress, Orders.ShipCity, Orders.ShipStateOrProvince, Orders.ShipPostalCode, Orders.ShipCountry, Orders.ShipPhoneNumber"
	'sql = sql & " FROM Orders INNER JOIN Customers ON Orders.OrdCustID = Customers.CustomerID"
	'sql = sql & " Where OrderID = " & OrderID
	sql = "SELECT Orders.*, Customers.* FROM Orders INNER JOIN Customers ON Orders.OrdCustID = Customers.CustomerID Where OrderID = " & OrderID
	set rs = conn.execute (sql)
	if not rs.eof then
		CustFN=rs.fields("CustFN")
		CustMI=rs.fields("CustMI")
		CustLN=rs.fields("CustLN")
		CustAddr=rs.fields("CustAddr")
		CustCity=rs.fields("CustCity")

		CustStateOrProvince=rs.fields("CustStateOrProvince")
		CustPostalCode=rs.fields("CustPostalCode")
		CustPhone=rs.fields("CustPhone")
		CustEmail=rs.fields("CustEmail")
		CustCountry=rs.fields("CustCountry")

	        if len(rs.fields("CustFN")) < 1 then
	        ShipName = rs.fields("CustFN") & " " & CustLN
	        else
	        ShipName = rs.fields("ShipName")
	        end if
	        ShipCompanyName = rs.Fields("ShipCompanyName")
	        ShipAddress = rs.Fields("ShipAddress")
	        ShipCity = rs.Fields("ShipCity")
	        ShipStateOrProvince = rs.Fields("ShipStateOrProvince")
	        ShipPostalCode = rs.Fields("ShipPostalCode")
           ShipPhoneNumber = rs.Fields("ShipPhoneNumber")
	        ShipEmail = rs.Fields("CustEmail")
        	  ShipCountry = rs.Fields("ShipCountry")
	end if
	'response.write(ShipName )
	rs.close
	sql= " exec spGetPaymentInfo " & PaymentID
	'response.write(sql)
 	set rs = conn.execute (sql)
	if not rs.eof then

		PaymentAmount=cstr(rs.fields("PaymentAmount"))
		if left(right(PaymentAmount,3),1)="." then
			'do nothing
		elseif left(right(PaymentAmount,2),1)="." then
			PaymentAmount=PaymentAmount & "0"
		elseif left(right(PaymentAmount,1),1)="." then
			PaymentAmount=PaymentAmount & "00"
		else
			PaymentAmount = PaymentAmount & ".00"
		end if
		CreditCardNumber=rs.fields("CreditCardNumber")
		CreditCardExpDate_Month=rs.fields("CreditCardExpDate_Month")
		CreditCardExpDate_Year=rs.fields("CreditCardExpDate_Year")
		CVV2=rs.fields("CVV2")
		CheckNumber=rs.fields("CheckNumber")
	end if
	rs.close


	Dim RegEx
	Set RegEx = New regexp
	RegEx.Pattern = "[0-9]{3}"
	RegEx.Global = True
	RegEx.IgnoreCase = True
	if RegEx.Test(CreditCardNumber) = FALSE then
		CreditCardNumber = Trim(CreditCardNumber)
		g_Key = mid(g_KeyString,1,Len(CreditCardNumber))
		CreditCardNumber = DeCrypt(CreditCardNumber)
	else
		CreditCardNumber = Trim(CreditCardNumber)
		g_Key = mid(g_KeyString,1,Len(CreditCardNumber))
		sql1 = "UPDATE Payments SET CreditCardNumber = '" & EnCrypt(CreditCardNumber) & "' WHERE PaymentID = " & PaymentID
		conn.Execute(sql1)
	end if
	Set RegEx = NOTHING


	set rs = nothing
	conn.close
	set conn = nothing



%>
<html>
<head>
</head>
<SCRIPT>
<!--
//--- 05.16.00: Added the Unique order Number
//--- Generate Unique Order Number
//--- (use 000658076426 for testing)
function GenerateOrderNumber()
{
  tmToday = new Date();
  return tmToday.getTime();
}
function StartOrder()
 {
  document.smallorder.custFN.focus();
 }
var sURL = unescape(window.location.pathname);

function refresh()
 {
	alert(sURL);
    window.location.replace( sURL );
 }
/*
Submit Once form validation-
*/

  function submitonce(theform)
    {
     //if IE 4+ or NS 6+
     if (document.all||document.getElementById)
      {
      //screen thru every element in the form, and hunt down "submit" and "reset"
      for (i=0;i<theform.length;i++)
        {
        var tempobj=theform.elements[i]
        if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
        //disable em
        tempobj.disabled=true
        }
      }
    }
//-->
</SCRIPT>

<form name="smallorder" action="ManualCS.asp" method="post" onSubmit="submitonce(this)">
<!-- to use this form on production change the HTML serial number to your number, and the post from 'developer' to 'www' -->
<font face="Verdana" size="2"><input type=hidden name="orderstring" value="ItemNum~ItemDesc~0.00~1~N~||">
<input type="hidden" name="serialnumber" value="000293293270"></font>
<div align="center">
  <center>
      <table border="0" cellpadding="2" cellspacing="0" width="800" bgcolor="#3399FF" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td colspan="2">
            <p align="center"><font face="Verdana" size="2"><b><br>
              </b></font><b><font face="Verdana" size="4"><i>BicycleBuys.com</i></font></b>
            <p align="center"> <font face="Verdana"><b>Order Entry Form</b></font>
          </td>
        </tr>
        <tr>
          <td align="right"><b><font face="Verdana" size="2">Order Number*:</font></b></td>
          <td><font face="Verdana" size="2">
            <input maxlength=25 size=25 name="ordernumber" tabindex="1" value="<%=OrderID%>">
            </font></td>
        </tr>
        <tr>
          <td align="right"><b><font face="Verdana" size="2">Name*:</font></b></td>
          <td><font face="Verdana" size="2">
            <input maxlength=84 size=30 name="sjname" tabindex="2" value="<%=CustFN & " "  & CustMID & " " & CustLN%>">

            <input type="hidden" maxlength=84 size=30 name="CustFN" tabindex="2" value="<%=CustFN%>">
			 <input type="hidden" maxlength=84 size=30 name="CustLN" tabindex="2" value="<%=CustLN%>">
            </font></td>
         </tr>
         <!--<tr>
            <td><font face="Verdana" size="2">-->



           <!-- </font></td>
        </tr>-->
        <tr>
          <td align="right"><b><font face="Verdana" size="2"> Address*:</font></b></td>
          <td><font face="Verdana" size="2">
            <input maxLength=95 size=32 name="streetaddress" tabindex="3" value="<%=CustAddr%>">
            </font></td>
        </tr>
        <tr>
          <td align="right"><b><font face="Verdana" size="2"> City*:</font></b></td>
          <td><font face="Verdana" size="2">
            <input maxLength=60 size=20 name="city" tabindex="4"  value="<%=CustCity%>">

            </font></td>
        </tr>
                <tr>
          <td align="right"><b><font face="Verdana" size="2">Province (if not
            Canadian): </font></b></td>
          <td><font face="Verdana" size="2">
            <input maxLength=40 size=40 name="province" tabindex="14" value="<%=custStateorProvince%>">
            </font></td>
        </tr>
        <!--<tr>
          <td align="right"><b><font face="Verdana" size="2"> State/Province*:</font></b></td>
          <td><font face="Verdana" size="2">
            <select name="state" size="1" tabindex="12">
              <option value selected> </option>
              <option value> United States </option>
              <option value="AL">Alabama </option>
              <option value="AK">Alaska </option>
              <option value="AZ">Arizona </option>
              <option value="AR">Arkansas </option>
              <option value="CA">California </option>
              <option value="CO">Colorado </option>
              <option value="CT">Connecticut </option>
              <option value="DE">Delaware </option>
              <option value="DC">District of Columbia </option>
              <option value="FL">Florida </option>
              <option value="GA">Georgia </option>
              <option value="HI">Hawaii </option>
              <option value="ID">Idaho </option>
              <option value="IL">Illinois </option>
              <option value="IN">Indiana </option>
              <option value="IA">Iowa </option>
              <option value="KS">Kansas </option>
              <option value="KY">Kentucky </option>
              <option value="LA">Louisiana </option>
              <option value="ME">Maine </option>
              <option value="MD">Maryland </option>
              <option value="MA">Massachusetts </option>
              <option value="MI">Michigan </option>
              <option value="MN">Minnesota </option>
              <option value="MS">Mississippi </option>
              <option value="MO">Missouri </option>
              <option value="MT">Montana </option>
              <option value="NE">Nebraska </option>
              <option value="NV">Nevada </option>
              <option value="NH">New Hampshire </option>
              <option value="NJ">New Jersey </option>
              <option value="NM">New Mexico </option>
              <option value="NY">New York </option>
              <option value="NC">North Carolina </option>
              <option value="ND">North Dakota </option>
              <option value="OH">Ohio </option>
              <option value="OK">Oklahoma </option>
              <option value="OR">Oregon </option>
              <option value="PA">Pennsylvania </option>
              <option value="RI">Rhode Island </option>
              <option value="SC">South Carolina </option>
              <option value="SD">South Dakota </option>
              <option value="TN">Tennessee </option>
              <option value="TX">Texas </option>
              <option value="UT">Utah </option>
              <option value="VT">Vermont </option>
              <option value="VA">Virginia </option>
              <option value="WA">Washington </option>
              <option value="WV">West Virginia </option>
              <option value="WI">Wisconsin </option>
              <option value="WY">Wyoming </option>
              <option value> </option>
              <option value> U.S. Territories </option>
              <option value="GU">Guam </option>
              <option value="AS">American Samoa </option>
              <option value="FM">Federated States of Micronesia </option>
              <option value="MP">Northern Mariana Islands </option>
              <option value="MH">Marshall Islands </option>
              <option value="PW">Palau Islands </option>
              <option value="PR">Puerto Rico </option>
              <option value="VI">US Virgin Islands </option>
              <option value> </option>
              <option value> Canadian Provinces </option>
              <option value="AB">Alberta </option>
              <option value="BC">British Columbia </option>
              <option value="MB">Manitoba </option>
              <option value="NB">New Brunswick </option>
              <option value="NF">Newfoundland </option>
              <option value="NT">Northwest Territories </option>
              <option value="NS">Nova Scotia </option>
              <option value="ON">Ontario </option>
              <option value="PE">Prince Edward Island </option>
              <option value="PQ">Quebec </option>
              <option value="SK">Saskatchewan </option>
              <option value="YT">Yukon Territory </option>
              <option value> </option>
              <option value> </option>
              <option value="XX">Other or None</option>
            </select>
            </font></td>
        </tr>-->
        <tr>
          <td align="right"><b><font face="Verdana" size="2"> Zip*:</font></b></td>
          <td><font face="Verdana" size="2">
            <input maxLength=10 size=10 name="zipcode" tabindex="5" value="<%=CustPostalCode%>">
            </font></td>
        </tr>
        <tr>
          <td align="right"><b><font face="Verdana" size="2"> Credit Card Number*:</font></b></td>
          <td><font face="Verdana" size="2">
            <input maxLength=22 size=22 name="accountnumber" tabindex="6" value="<%= CreditCardNumber %>" >
            </font></td>
        </tr>
        <tr>
          <td align="right"><b><font face="Verdana" size="2"> Expiration Month:
            (mm)*</font></b></td>
          <td><font face="Verdana" size="2" >
            <input maxLength=2 size=2 name="month" tabindex="7" value="<%=CreditCardExpDate_Month%>">
            </font></td>
        </tr>
        <tr>
          <td align="right"><b><font face="Verdana" size="2"> Expiration Year:
            (yy)*</font></b></td>
          <td><font face="Verdana" size="2">
            <input maxLength=4 size=2 name="year" tabindex="8" value="<%=CreditCardExpDate_Year%>" >
            </font></td>
        </tr>
        <tr>
          <td align="right"><b><font face="Verdana" size="2">Amount*: </font></b></td>
          <td><font face="Verdana" size="2">
            <input maxlength=8 size=8 name="transactionamount" tabindex="9" value="<%=PaymentAmount%>" >
            </font></td>
        </tr>
        <tr>
          <td align="right"><b><font face="Verdana" size="2">Phone*:</font></b></td>
          <td><font face="Verdana" size="2">
            <input maxlength=20 size=15 name="shiptophone" tabindex="10" value="<%=CustPhone%>">
            </font></td>
        </tr>
        <tr>
          <td align="right"><b><font face="Verdana" size="2"> E-mail*:</font></b></td>
          <td><font face="Verdana" size="2">
            <input maxlength=50 size=35 name="email" tabindex="11" value="<%=custEmail%>">
            </font></td>
        </tr>
        <tr>
          <td align="right"><b><font face="Verdana" size="2">Country: </font></b></td>
          <td><font face="Verdana" size="2">
            <input maxLength=40 size=40 name="country" tabindex="13" value="<%=CustCountry%>">
            </font></td>
        </tr>
        <!--<tr>
          <td align="right"><b><font face="Verdana" size="2">Province (if not
            Canadian): </font></b></td>
          <td><font face="Verdana" size="2">
            <input maxLength=40 size=40 name="province" tabindex="14" value="<%=custStateorProvince%>">
            </font></td>
        </tr>-->
        <tr>
          <td align="right"><b><font face="Verdana" size="2">Card Security Code:
            </font></b></td>
          <td><font face="Verdana" size="2">
            <input maxlength=8 size=8 name="cvv2" tabindex="15" value="<%=CVV2%>">
            </font></td>
        </tr>
        <tr>
          <td align="right"><b><font face="Verdana" size="2"> </font></b></td>
          <td><font face="Verdana" size="2"> </font></td>
        </tr>
        <tr>
          <td align="right">* Denotes Required Field</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td colspan="2">
            <p align="center"><font face="Verdana" size="2">
              <INPUT TYPE=submit VALUE="Click to process the Order" tabindex="16">
              <b>
              <input type="button" value="Reset Form" name="buttonRefresh" onClick="refresh()" tabindex="17">
              </b></font>
          </td>
        </tr>
      </table>

  </center>

<input type=input name="ShipName" tabindex="3" value="<%=ShipName%>">
<input type=input name="ShipCompany" tabindex="3" value="<%=ShipCompanyName%>">
<input type=input name="ShipAddress1" tabindex="3" value="<%=ShipAddress%>">
<input type=input name="ShipCity" tabindex="3" value="<%=ShipCity%>">
<input type=input name="ShipStateProvince" tabindex="3" value="<%=ShipStateOrProvince%>">
<input type=input name="ShipZipPostal" tabindex="3" value="<%=ShipPostalCode%>">
<input type=input name="ShipPhoneNumber" tabindex="3" value="<%=ShipPhoneNumber%>">
<input type=input name="ShipEmail" tabindex="3" value="<%=CustEmail%>">
<input type=input name="ShipCountry" tabindex="3" value="<%=ShipCountry%>">
</div>


<p><font face="Verdana" size="2">
<!-- END REQUIRED -->
 </font></p>

</form>
</html>
