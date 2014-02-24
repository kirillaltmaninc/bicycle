<% ' DISPLAY A LARGE ITEM IMAGE WITH 'CLOSE' BUTTON
vFN = request.querystring("fn")
vH = request.querystring("h")
vW = request.querystring("w")
%><html><body bgcolor="#FFFFFF">
<img src="<%=vFN%>" height="<%=vH%>" width="<%=vW%>"><BR><a href="#" onclick="window.close()"><img src="/images/close_button.gif" width="57" height="15" alt="" border="0" align="right"></a>
</body></html>