<%
	dim xx

	if Request.QueryString("c1")<>"" and Request.QueryString("c2")="xx" then 
		xx=server.urlencode(Request.QueryString("c1")) 
	elseif Request.QueryString("c1")<>"" then 
		xx=server.urlencode(Request.QueryString("c1")) & "/" & server.urlencode( Request.QueryString("c2"))   & "/" 
	else
		xx=""
	end if
     	response.write(xx)
                response.Status="301 Moved Permanently"
                response.AddHeader "Location", "/" & xx
                response.End
%>
 