<%

 'dim pageRequested 
'pageRequested = mid(request.queryString, instr(request.queryString,";") + 1) 
'response.write( pageRequested )
'response.status = "404 Not Found"
response.write("TEST")
response.end
%>

 