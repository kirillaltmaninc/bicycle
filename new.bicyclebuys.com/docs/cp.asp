<!--#INCLUDE VIRTUAL="/includes/template_cls.asp" -->
<!--#INCLUDE VIRTUAL="/includes/common.asp" -->
<!--#INCLUDE VIRTUAL="/includes/cartconfig.asp" -->
<html>
<head>
    <title>Change Password</title>
</head>
<body>
 <%
dim pwdResetMsg, ShowIt, valid

function ResetEmailPWD( )
    dim msg,mEmail, mName
    dim mPWD, c, com, rs, p
    ResetEmailPWD = true
    mEmail = trim(Request.Form("email"))
    mPWD = trim(Request.Form("currentPassword"))
   
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
    p.value = 0
    com.Parameters.Append p 
    
    set rs = com.execute
    if not rs.eof then   
        if rs.fields("ValidPassword") = "V" then 
            rs.close
            set rs = nothing
            com.parameters.item("@setPassword") = 1
            mPWD =   trim(Request.Form("newPassword")) 
            com.parameters.item("@pwd") = mPWD
            
            set rs = com.execute      
            if rs.eof then
               pwdResetMsg = "<BR><font style=""font-weight: bold; color: red;"">E-Mail address not found:  <BR>" & trim(mEmail) & "</B>"
               ResetEmailPWD = false
            else
               pwdResetMsg = "<BR>Password successfully changed for: <BR>" & mEmail
            end if
         else
            pwdResetMsg = "<BR><font style=""font-weight: bold; color: red;"">Invalid current password:  <BR>" & trim(mEmail) & "</B>"
            ResetEmailPWD = false
         end if
    else
        pwdResetMsg = "<BR><font style=""font-weight: bold; color: red;"">E-Mail address not found:  <BR>" & trim(mEmail) & "</B>"
        ResetEmailPWD = false
    end if
    rs.close
    set rs = nothing
    c.close
    set c = nothing    
End function

Function newPWD()
    Randomize Timer
    newPWD = CStr(CInt(Rnd(20) * 1000)) & Mid("aljdflasjdfasdf", Int((10 * Rnd(50)) + 1), 2) & CStr(CInt(Rnd(400) * 1000))
End Function

Sub SendEmailPWD( )
   dim msg,mEmail, mName
   dim mPWD, c, com, rs, p
   mPWD =  newPWD()
   mEmail = trim(Request.Form("email"))
  
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
       efoot = efoot & "/" & vbcrlf

       ' We're going to save the order into a file on the web server for
       ' order fulfillment by the BB team
     
       ' ---- Send Email to Customer
           Cart.Mail.Host = "webserver"
       Cart.Mail.From = "Sales@BicycleBuys.com"
       Cart.Mail.FromName = "BicycleBuys.com Request"
       Cart.Mail.Subject = "BicycleBuys.com Password Reset"

dim htm
htm="<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbcrlf
htm=htm & "<html xmlns=""http://www.w3.org/1999/xhtml"" >" & vbcrlf
htm=htm & "<head><title>BicycleBuys.com - Password Reset</title></head><body>" & vbcrlf
htm = htm & eheader & vbcrlf &  msg & vbcrlf & vbcrlf & efoot
htm=htm & "</body>" & vbcrlf
htm=htm & "</html>" & vbcrlf

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

'If request.servervariables("REQUEST_METHOD") = "POST" then
ShowIt = true
if request.servervariables("REQUEST_METHOD")= "POST" then
    if request.Form("txtreset") = "0" then
        ShowIt=false
        valid = true
        if trim(Request.Form("email")) <> trim(Request.Form("email2")) and Request.Form("email") <>"" then
            valid = false
            ShowIt = true 
            pwdResetMsg = "<font style=""font-weight: bold; color: red;"">E-mails are not the same.</font><BR>"
        end if
        if trim(Request.Form("newPassword")) <> trim(Request.Form("newPassword2")) then
            valid = false 
            ShowIt = true    
            pwdResetMsg = "<font style=""font-weight: bold; color: red;"">New Password are not the same.</font><BR>"
        end if
        if valid then 
            if not ResetEmailPWD() then 
                ShowIt = true 
            end if
        end if
        response.Write(pwdResetMsg)        
   else
        SendEmailPWD()
        response.Write(pwdResetMsg)
        ShowIt=false   
   end if
end if
if ShowIt = true then
%>

    <form name="myform" METHOD="post" ACTION="/cp.asp">
    <div>
        <table border="0" cellpadding="0" cellspacing="0" style="width: 392px">
            <tr>
                <td align="right" valign="middle">
                    E-mail Address:&nbsp;
                    <br />
                </td>
                <td valign="middle">
                    <input name="email" style="width: 200px" type="text"  value="<%=Request.Form("email")%>" /></td>
            </tr>
            <tr>
                <td style="width: 171px" align="right" valign="middle">
                    Confirm E-Mail Address:&nbsp;
                    <br />
                </td>
                <td valign="middle">
                    <input name="email2" style="width: 200px" type="text"   value="<%=Request.Form("email2")%>"/></td>
            </tr>
            <tr>
                <td style="width: 171px" align="right" valign="middle">
                    Current Password:&nbsp;
                    <br />
                    &nbsp;
                </td>
                <td valign="middle">
                    <input name="currentPassword" maxlength="50" style="width: 200px" type="password" /></td>
            </tr>
            <tr>
                <td style="width: 171px" align="right" valign="middle">
                    New Password:&nbsp;
                    <br />
                </td>
                <td valign="middle">
                    <input name="newPassword" maxlength="50" style="width: 200px" type="text" value="<%=Request.Form("newPassword")%>" /></td>
            </tr>
            <tr>
                <td style="height: 57px;" align="right" valign="middle">
                    Confirm new Password:&nbsp;
                    <br />
                    <br />
                </td>
                <td valign="middle" style="height: 57px">
                    <input name="newPassword2" maxlength="50" style="width: 200px" type="text" value="<%=Request.Form("newPassword2")%>"/></td>
            </tr>
        </table>
        <table border="0" style="width: 392px">
            <tr>
                <td>
        <input name="Button1" type="button" value="Cancel" onclick="javascript: self.close()"/></td>
                <td>
                </td>
                <td>
        <input name="Submit" type="submit" value="Change Password"></td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td>
                    <input type="button" name="btnReset" value="Reset Password" onclick="javascript:document.myform.txtreset.value='1';document.myform.submit();"></td>
            </tr>
        </table>
        <br />
        &nbsp;</div>
         <input type=hidden name=txtreset value=0>
    </form>

<% end if %>
</body>
</html>
