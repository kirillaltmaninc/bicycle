<%
     
                response.Status="301 Moved Permanently"
                response.AddHeader "Location", "/"
                response.End
%>
 