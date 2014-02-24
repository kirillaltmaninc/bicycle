

<link rel="stylesheet" type="text/css" href="/index.css" title="index">
<form method="post" action="custOrders.asp" >


</form>
<%


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


	response.write "</table>"
        rs.close
        conn.close
        set rs = nothing
        set conn = nothing
   

Response.write("<br><br>Count: " & Application("DayCount"))
Response.write("<br><br>Recent Referrers:<br>" & Application("Last10"))

%>