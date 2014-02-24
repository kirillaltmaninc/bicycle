<% response.Status="404 Not found" %>
<%
dim pageRequested 
pageRequested = mid(request.queryString, instr(request.queryString,";") + 1) 
response.write( pageRequested )

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>Indexing Error</title>
<meta name="robots" content="NOFOLLOW,NOINDEX">

<meta http-equiv="REFRESH" content="4;url=http://www.bicyclebuys.com"></HEAD>
<BODY>
Due to an indexing error you can find the product you are looking for in the site below.
<br>Sorry for the inconvenience.

You will be redirected in 4 seconds.
If not click <a href="http://www.bicyclebuys.com">here</a>.
<%response.write ASPError.property("File") 

%>
</BODY>
</HTML>

