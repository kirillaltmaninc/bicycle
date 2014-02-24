<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Bicycle Buys | BicycleBuys.com | Online Bike Shop | Bicycles | Bike Parts | Frames | Pedals</title>
<LINK rel=stylesheet type="text/css" href="bicyclebuys.css" title="bicyclebuys">
</head>

<body>
<div align="left">
<img src="images/smallheader.jpg" /><br />
  <%
   '%%%%% FUNCTIONS DEFINED FIRST %%%%%%%
   '%% CHECKSTRING - Make sure string doesn't contain any SQL munging characters
   FUNCTION CheckString (s, endchar)
   	pos = InStr(s, "'")
   	While pos > 0
   		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
   		pos = InStr(pos + 2, s, "'")
   	Wend
    CheckString="'" & s & "'" & endchar
   END FUNCTION
   

if request("saving") = "saving" then   
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'% Create the Conn Object and open it
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	'Filename = "d:\database\salemail.mdb"
	'FileDSN = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Filename & ";"
	' Modified on 3/6/04 by DD
	FileDSN = Application("FileDSN")

	Set Conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Conn.Open FileDSN


   if request.form("name") = "" or request.form("lastname") = "" or request.form("street") = "" or request.form("city") = "" or request.form("state") = "" or request.form("zip") = "" or request.form("country") = "" or request.form("email") = "" or request.form("telephone") = "" then
      vError = 1
      vRequiredString =  "<FONT ID=""body"" COLOR=""#FF0000"">Required</FONT>"
   End if

   If vError <> 1 Then
      vname = CheckString(request.form("name"), ",")
      vlastname = CheckString(request.form("lastname"), ",")
      vstreet = CheckString(request.form("street"), ",")
      vcity = CheckString(request.form("city"), ",")
      vstate = CheckString(request.form("state"), ",")
      vzip = CheckString(request.form("zip"), ",")
      vcountry = CheckString(request.form("country"), ",")
      vemail = CheckString(request.form("email"), ",")
      vtelephone = CheckString(request.form("telephone"), ",")
      vsalemail = CheckString(request.form("salemail"),"")
   
      sql = "INSERT INTO main (name,lastname,street,city,state,zip,country,email,telephone,salemail) "
      sql = sql & "VALUES ("
      sql = sql & vname &vlastname & vstreet & vcity & vstate & vzip & vcountry & vemail & vtelephone & vsalemail
      sql = sql & ")"
      
      Conn.Execute(sql)
      Conn.Close
      response.redirect "register_confirm.asp"
   End If
End If
%>
  <b>Registering at bicyclebuys.com gets you all these benefits:</b>
  <ul>
    <li> Special Sales Announcements 
    <li> Member Only Sale Events 
    <li> Hot Product Announcements 
    <li> Contests and Giveaways 
    <li> Other Special Extras    
  </ul>
  <br />
  And it only takes less than a minute!
</div>
<form action="register.asp" method="post">
	
	  <div align="left"><b>Register Now...</b><br>
	        <i>All fields are required</i>
        <br />
        <br />
        <table width="500" border="0" cellspacing="0" cellpadding="2">
          <tr>
            <td><font id="body"><b>First Name:</b></font></td>
          <td><input name="name" type="TEXT" size="35" maxlength="20" value="<%=request("name")%>" />
            <% if vError and Len(Request("name")) = 0 Then Response.write vRequiredString %></td>
        </tr>
          <tr>
            <td><font id="body"><b>Last Name:</b></font></td>
          <td><input name="lastname" type="TEXT" size="35" maxlength="20" value="<%=request("lastname")%>" />
            <% if vError and Len(Request("lastname")) = 0 Then Response.write vRequiredString %></td>
        </tr>
          <tr>
            <td><font id="body"><b>Address:</b></font></td>
          <td><input name="street" type="TEXT" size="35" maxlength="50" value="<%=request("street")%>" />
            <% if vError and Len(Request("street")) = 0 Then Response.write vRequiredString %></td>
        </tr>
          <tr>
            <td><font id="body"><b>City:</b></font></td>
          <td><input name="city" type="TEXT" size="35" maxlength="25" value="<%=request("city")%>" />
            <% if vError and Len(Request("city")) = 0 Then Response.write vRequiredString %></td>
        </tr>
          <tr>
            <td><font id="body"><b>State/Province:</b></font></td>
          <td><input name="state" type="TEXT" size="35" maxlength="40" value="<%=request("state")%>" />
            <% if vError and Len(Request("state")) = 0 Then Response.write vRequiredString %></td>
        </tr>
          <tr>
            <td><font id="body"><b>Zip/Postal Code:</b></font></td>
          <td><input name="zip" type="TEXT" size="35" maxlength="15" value="<%=request("zip")%>" />
            <% if vError and Len(Request("zip")) = 0 Then Response.write vRequiredString %></td>
        </tr>
          <tr>
            <td><font id="body"><b>Country:</b></font></td>
          <td><input name="country" type="TEXT" size="35" maxlength="40" value="<%=request("country")%>" />
            <% if vError and Len(Request("country")) = 0 Then Response.write vRequiredString %></td>
        </tr>
          <tr>
            <td><font id="body"><b>E-Mail Address:</b></font></td>
          <td><input name="email" type="TEXT" size="35" maxlength="60" value="<%=request("email")%>" />
            <% if vError and Len(Request("email")) = 0 Then Response.write vRequiredString %></td>
        </tr>
          <tr>
            <td><font id="body"><b>Phone Number:</b></font></td>
          <td><font id="body"><br />
          </font>
            <input name="telephone" type="TEXT" size="35" maxlength="20" value="<%=request("telephone")%>" />
            <% if vError and Len(Request("telephone")) = 0 Then Response.write vRequiredString %></td>
        </tr>
          <tr>
            <td><font id="body"><b>Add to Sale-Mail?</b></font></td>
          <td><input name="salemail" type="radio" value="YES" checked />
            <font id="body">Yes</font> &nbsp;
            <input name="salemail" type="radio" value="NO" />
            <font id="body">No</font></td>
        </tr>
          <tr>
            <td>&nbsp;</td>
          <td><input name="saving" type="hidden" value="saving" />
            <input name="save" type="image" src="/cartimages/registernow.gif" alt="Register Now" border="0" width="100" height="20" /></td>
        </tr>
        </table>
      </div>
</form>
	<div align="left"><br>
	  
	  
	    <!----------------BODY END CONTENT---------------->
	  
    </div>
</body>
</html>
