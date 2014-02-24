<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%

    dim subcatid, vendid,subcatid2, vendid2, email
    dim chkMonthly,chkQuarterly,chkHoliday,chkYearlyBlowOut, chkNone, msg, email2
    msg=""
    email = request.Form("email")
    email2 = request.Form("email2")
    chkMonthly = request.Form("chkMonthly")
    chkQuarterly = request.Form("chkQuarterly")
    chkHoliday = request.Form("chkHoliday")
    chkYearlyBlowOut = request.Form("chkYearlyBlowOut")
    chkNone=request.Form("chkNone")
    subcatid = request.Form("subcatid")
    vendid = request.Form("vendid")
    if subcatid= "" then subcatid = -1
    if vendid= "" then vendid = -1
    subcatid2 = request.Form("subcatid2")
    vendid2 = request.Form("vendid2")
    if subcatid2= "" then subcatid2 = -1
    if vendid2= "" then vendid2 = -1


    if request.Form("btnSave")="Save Settings" then
        if not SaveLoadUser(0) then
            msg = "***E-mail NOT FOUND***"
        else
            msg =  "Settings Updated Successfully"
        end if
    elseif request.Form("register")<>"" then
        if not SaveLoadUser(1) then
            msg = "***Succsessfully Subscribed***"
        else
            msg = "Loaded User setting for: " & email
        end if
    elseif request.Form("load")<>"" then
        if not SaveLoadUser(-1) then
            msg = "***E-mail NOT FOUND***"
        else
            msg = "Loaded User setting for: " & email
        end if
    elseif request.Form("btnCancel")<>"" then
        if not SaveLoadUser(-1) then
            msg = "***E-mail NOT FOUND***"
        else
            msg = "***Canceled Changes***"
        end if
    elseif request.Form("btnUnSub")<>"" then
        if not SaveLoadUser(-2) then
            msg = "***E-mail NOT FOUND***"
        else
            msg = "***Unsubscribed From All E-mails***"
        end if

    else
        if request.Form="" then setDefaults()
    end if
    if msg<>"" then
        msg = "<br><font class=""price"">" & msg & "</font><br>"
    end if

    function SaveLoadUser(createNew)
	    Dim dsn, conn, rs, sql
	    'dsn = Application("dsn")
	    dsn = "Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webUserprod;Initial Catalog=BBC_PROD;Data Source=10.0.0.66"

	    Set conn = Server.CreateObject("ADODB.Connection")
	    conn.Open dsn
	    set rs = Server.CreateObject("ADODB.Recordset")

        sql = "exec saveMain '" & email & "','"
        sql = sql & convertCheck(chkNone,"on","no", "yes") & "','"
        sql = sql & convertCheck(chkMonthly,"on","y", "n") & "','"
        sql = sql & convertCheck(chkQuarterly,"on","y", "n") & "','"
        sql = sql & convertCheck(chkHoliday,"on","y", "n") & "','"
        sql = sql & convertCheck(chkYearlyBlowOut,"on","y", "n") & "',"
        sql = sql &  VendID & ","
        sql = sql &  SubCatID & ","
        sql = sql &  VendID2 & ","
        sql = sql &  SubCatID2 & ",'"
        sql = sql &  "SKU1" & "','"
        sql = sql &  "SKU2" & "','"
        sql = sql &  "SKU3" & "','"
        sql = sql &  "SKU4" & "',"
        sql = sql &  createNew  & ", 0"

        rs.open sql, conn, 3
        SaveLoadUser = false

        if not rs.eof then
            SaveLoadUser = true
            chkNone = convertCheck(rs.fields("saleMail"),"no","on","")
            chkMonthly = convertCheck(rs.fields("MonthlyEmail"),"y","on","")
            chkQuarterly = convertCheck(rs.fields("QuarterlyEmail"),"y","on","")
            chkHoliday = convertCheck(rs.fields("HolidayEmail"),"y","on","")
            chkYearlyBlowOut = convertCheck(rs.fields("YearlyBlowOutEmail"),"y","on","")
            VendID=rs.fields("VendID")
            SubCatID=rs.fields("SubCatID")
            VendID2=rs.fields("VendID2")
            SubCatID2=rs.fields("SubCatID2")
            SKU1=rs.fields("SKU1")
            SKU2=rs.fields("SKU2")
            SKU3=rs.fields("SKU3")
            SKU4=rs.fields("SKU4")
        end if
        rs.close
        conn.close
        set rs = nothing
        set conn = nothing
    end function

    function convertCheck(val,chkVal,trueVal,falseVal)
        if cstr(val)=cstr(chkVal) then
            convertCheck= trueVal
        else
            convertCheck=falseVal
        end if
    end function

    function setDefaults()
        chkMonthly = "on"
        chkQuarterly = "on"
        chkHoliday = "on"
        chkYearlyBlowOut = "on"
        chkNone=""
    end function

    Function getVendors(subCatID, VendID)
	    Dim dsn, conn, rs, sql
	    'dsn = Application("dsn")
	    dsn = "Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webUserprod;Initial Catalog=BBC_PROD;Data Source=10.0.0.66"

	    Set conn = Server.CreateObject("ADODB.Connection")
	    conn.Open dsn
	    set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "exec getVendors " & subCatID
	    rs.open sql, conn, 3
	    getVendors="<option value=""-1""><-- Select Vendor/Show all categories --></option>"
	    while not rs.eof
		    getVendors=getVendors & "<option value=""" & rs.fields("VendID") & """" & Selected(VendID,rs.fields("VendID")) & ">" & rs.fields("Vendor") & "</option>"
		    rs.movenext
	    wend
	    rs.close
	    conn.close
	    set rs = nothing
	    set conn=nothing

    End Function

    Function getSubcats(subCatID, VendID)
	    Dim dsn, conn, rs, sql
	    'dsn = Application("dsn")
	    dsn = "Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webUserprod;Initial Catalog=BBC_PROD;Data Source=10.0.0.66"

	    Set conn = Server.CreateObject("ADODB.Connection")
	    conn.Open dsn
	    set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "exec getSubCats " & VendID
	    rs.open sql, conn, 3
	    getSubcats="<option value=""-1""><-- Select Category/Show all vendors --></option>"
	    while not rs.eof
		    getSubcats=getSubcats & "<option value=""" & rs.fields("SubCatID") & """" & Selected(SubCatID,rs.fields("SubCatID")) & ">" & rs.fields("Subcategory") & "</option>"
		    rs.movenext
	    wend
	    rs.close
	    conn.close
	    set rs = nothing
	    set conn=nothing

    End Function

    function Selected(a,b)
	    if cstr(a)=cstr(b) then
		    Selected =" selected "
	    else
		    Selected = ""
	    end if
    end function

    function checked(a,b)
        if cstr(a)=cstr(b) then
            checked=" checked=""CHECKED"" "
        else
            checked=""
        end if
    end function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>BicycleBuys.com E-Mail Settings/Price Notifications</title>
    <link href="bicyclebuys.css" rel="stylesheet" type="text/css" />
    <link href="menu.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <center>
        <form id="form1" action="mymail.asp" method="post">
            &nbsp;<table width="760px" border="0" cellpadding="0" cellspacing="0" class="border" style="text-align: left;">
                <tr>
                    <td style="margin-left: 100px">
                         <table style="text-align: left; width: 400px">
                            <tr>
                                <td colspan="2">
                                    <a href="http://www.bicyclebuys.com">
                                        <img src="images/bicycle_buys.jpg" /><br>
                                        <br>
                                    </a>
                                </td>
                            </tr>
				<tr>
                                <td colspan="2">
                                    To un-subscribe from all future e-mail enter your e-mail and click below...
                                 </td>
                                   
                            </tr>
                            <tr>
                                <td>
                                    E-Mail&nbsp;
                                </td>
                                <td>
                                    <input name="email" type="text" value="<%=email %>"></td>
                            </tr>
                            
                            <tr>
                                <td>
                                </td>
                                <td> 
				                                                                     
                                    <input name="btnUnSub" type="submit" value="Un-Subcribe All" style="width: 200px; position: relative" /><%=msg%></td>

               		 </tr>
				<tr>
                                <td colspan="2">
                                 <br>   OR if you prefer to decrease the frequency of e-mails you recieve;<br> enter your e-mail above, change your preferences below and click Save Settings.
                                 <br><br><br></td>
                                   
                            </tr>
                <tr>
                    <td colspan="2">
                        <input name="chkMonthly" type="checkbox"
                        <%=checked(chkMonthly,"on")%> >
                        Receive Monthly E-mails.</td>
                </tr>
                <tr>
                    <td colspan="2" style="height: 20px">
                        <input name="chkQuarterly" type="checkbox" <%=checked(chkQuarterly,"on")%>>
                        Receive Quarterly E-mails.</td>
                </tr>
                <tr>
                    <td colspan="2">
                        <input name="chkHoliday" type="checkbox" <%=checked(chkHoliday,"on")%>>
                        Receive Holiday E-mails.</td>
                </tr>
                <tr>
                    <td style="height: 36px" colspan="2">
                        <input name="chkYearlyBlowOut" type="checkbox" <%=checked(chkYearlyBlowOut,"on")%>>
                        Recieve One Yearly Blow-out E-mail.</td>
                </tr>
                <tr>
                    <td colspan="2">
                        <input name="chkNone" type="checkbox" <%=checked(chkNone,"on")%>>
                        I don't wish to recieve any e-mails.</td>
                </tr>
            </table>
            <br />
            <table style="width: 400px; text-align:left;">
                <tr>
                    <td colspan="2">
                        To participate in price alert e-mails, select a category and vendor combination below. 
			An e-mail notification will be sent to you informing you about limited time sales matching your selection.
	<br><br></td>
                </tr>
               <tr>
                    <td colspan="2">
                        Price Alert Settings 1</td>
                </tr>
                <tr>
                    <td>
                        Vendor:&nbsp;
                    </td>
                    <td>
                        <select id="vendid" name="vendid" onchange="this.form.submit()">
                            <%=getVendors(subcatid, vendid) %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td>
                        Category:&nbsp;
                    </td>
                    <td>
                        <select id="subcatid" name="subcatid" onchange="this.form.submit()">
                            <%=getSubCats(subcatid, vendid) %>
                        </select>
                    </td>
                </tr>
             <tr><td><br><br></td></tr>
                <tr>
                    <td colspan="2">
                        Price Alert Settings 2</td>
                </tr>
                <tr>
                    <td>
                        2nd Vendor:&nbsp;
                    </td>
                    <td>
                        <select id="vendid2" name="vendid2" onchange="this.form.submit()">
                            <%=getVendors(subcatid2, vendid2) %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td>
                        2nd Category:&nbsp;
                    </td>
                    <td>
                        <select id="subcatid2" name="subcatid2" onchange="this.form.submit()">
                            <%=getSubCats(subcatid2, vendid2) %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <%=msg%>
                        <input name="btnSave" type="submit" value="Save Settings">
                        <input name="btnCancel" type="submit" value="Cancel"></td>
                </tr>
            </table>
            </td></tr> </table>
        </form>
    </center>
</body>
</html>
