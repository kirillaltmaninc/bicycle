<%
' sitemap_gen_db.asp
' A simple script (using database) to automatically produce sitemaps for a 10.0.0.66, in the Google Sitemap Protocol (GSP)

MAXURLS_PER_SITEMAP = 50000

'modify this to change website, baseurl and table
baseurl="http://www.bicyclebuys.com"

xDb_Conn_Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath("db\yourdb.mdb") & ";"
strsql = "SELECT * FROM yourtable"

'see http://www.time.gov/ for utcOffset
utcOffset=1

response.ContentType = "text/xml"
response.write "<?xml version='1.0' encoding='UTF-8'?>"
response.write "<!-- generator='http://www.iteam5.net/francesco/sitemap_gen'-->"
response.write "<urlset xmlns='http://www.google.com/schemas/sitemap/0.84'>"

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open strsql, conn
Do while not rs.eof
	if URLS<MAXURLS_PER_SITEMAP then

		'modify this to change database field
		id_page=Server.URLEncode(rs("id"))
		filelmdate=rs("pagina_lastupdate")
		priority=rs("priority")

		if not isdate(filelmdate) then filelmdate=now()
		filedate=iso8601date(filelmdate,utcOffset)

		if priority="" or priority>1.0 then priority="1.0"

		response.write "<url><loc>"&server.htmlencode(baseurl&id_page)&"</loc><lastmod>"&filedate&"</lastmod><priority>"&priority&"</priority></url>"
		URLS=URLS+1
		Response.Flush
	rs.movenext
end if
Loop

response.write "</urlset>"

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing


Function iso8601date(dLocal,utcOffset)	
	Dim d
	' convert local time into UTC
	d = DateAdd("H",-1 * utcOffset,dLocal)

	' compose the date
	iso8601date = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2) & "T" & _
		Right("0" & Hour(d),2) & ":" & Right("0" & Minute(d),2) & ":" & Right("0" & Second(d),2) & "Z"
End Function
%>
