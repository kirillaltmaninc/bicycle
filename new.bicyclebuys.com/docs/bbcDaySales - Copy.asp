

<link rel="stylesheet" type="text/css" href="/index.css" title="index">
<form method="post" action="bbcDaySales.asp" >
Validate: <input type="password" id="checkIt" value="" name="checkIt">
<select id="minusDays" name="minusDays" onchange="this.form.submit();">
<option value="">Select Day</option>
<option value="0">Today</option>
<option value="-1">Yesterday</option>
<option value="-2">2 Days Ago</option>
<option value="-3">3 Days Ago</option>
<option value="-4">4 Days Ago</option>
<option value="-5">5 Days Ago</option>
<option value="-6">6 Days Ago</option>
<option value="-7">1 Week Ago</option>
<option value="-14">2 Weeks Ago</option>
<option value="-21">3 Weeks Ago</option>
<option value="-28">4 Weeks Ago</option>
</option>
</select></form>
<%

	dim minusDays
	minusdays = request.form("minusDays")
	if minusDays = "" then minusDays = 0



	function writeSales()
		dim x, v
	    Dim dsn, conn, rs, sql, f
	    'dsn = Application("dsn")
	    dsn = "Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webUserprod;Initial Catalog=BBC_PROD;Data Source=10.0.0.66"

	    Set conn = Server.CreateObject("ADODB.Connection")
	    conn.Open dsn
	    set rs = Server.CreateObject("ADODB.Recordset")

        sql = "exec spDaySales " & minusDays

        rs.open sql, conn, 3
	response.write "Direct Web Orders<BR><table border=1 class=""shipping"">"
		response.write "<tr>"
		x=1
		for each f in rs.fields
			response.write "<td>" & nz(f.name)  & "</td>"
			if x = 4 then
				response.write "<td> </td>"
				x = 0
			end if
			x = x+1
		next
		response.write "</tr>" & vbnewline

	while not rs.eof
		response.write "<tr>"
		x=0
		for each f in rs.fields
			if x = 4 then
				response.write "<td></td>"
				x = 0
			end if
			v = nz(f.value)
			if isnumeric(v) then
				response.write "<td>" &  FormatNumber(v,0)   & "</td>"
			else
				response.write "<td>" & v  & "</td>"
			end if
			x=x+1
		next
		response.write "</tr>" & vbnewline
		rs.movenext
	wend
	response.write "</table>"
        rs.close
        conn.close
        set rs = nothing
        set conn = nothing
    end function
    function nz(val)
	if isnull(val) then
		nz = " "
	else
		nz = val
	end if
    end function


  function writeCallines()
		dim x, v
	    Dim dsn, conn, rs, sql, f
	    'dsn = Application("dsn")
	    dsn = "Provider=SQLOLEDB.1;Password=bbcwebUserprod;Persist Security Info=True;User ID=webUserprod;Initial Catalog=BBC_PROD;Data Source=10.0.0.66"

	    Set conn = Server.CreateObject("ADODB.Connection")
	    conn.Open dsn
	    set rs = Server.CreateObject("ADODB.Recordset")

        sql = "exec spDayCallins " & minusDays

        rs.open sql, conn, 3
	response.write "<BR><BR>Phone Orders<BR><table border=1 class=""shipping"">"
		response.write "<tr>"
		x=1
		for each f in rs.fields
			response.write "<td>" & nz(f.name)  & "</td>"

			x = x+1
		next
		response.write "</tr>" & vbnewline

	while not rs.eof
		response.write "<tr>"
		x=0
		for each f in rs.fields
			v = nz(f.value)
			if isnumeric(v) then
				response.write "<td>" &  FormatNumber(v,0)   & "</td>"
			else
				response.write "<td>" & v  & "</td>"
			end if
			x=x+1
		next
		response.write "</tr>" & vbnewline
		rs.movenext
	wend
	response.write "</table>"
        rs.close
        conn.close
        set rs = nothing
        set conn = nothing
    end function

    function nz(val)
	if isnull(val) then
		nz = " "
	else
		nz = val
	end if
    end function


if request.form("checkIt")="bbccx" or session("checkIt") = "bbccx"  then
	session("checkIt") = "bbccx"
	call writeSales()
	call writeCallines()
end if

Response.write("<br><br>Count: " & Application("DayCount"))
Response.write("<br><br>Recent Referrers:<br>" & Application("Last10"))

%>