<html>
<head></head>
<body>

<%
dim FB
FB ="<iframe src="""" scrolling=""no"" frameborder=""0"" style=""border:none; overflow:hidden; width:450px; height:80px;"" allowTransparency=""true""></iframe>"
'response.Write FB
 if Request.ServerVariables("Https")="off" and Request.ServerVariables("HTTP_HOST") <>"www.bicyclebuys.net"  and Request.ServerVariables("HTTP_HOST") <>"bicyclebuys.net" then
    response.Write "HTTP"
 else
    response.Write "HTTPS"
 end if
 
 %>
 <script Language="JavaScript" type="text/JavaScript">
     function doit(obj) {
         obj.src = "http://www.facebook.com/plugins/like.php?href=" + location.href.toString() + "/FB&amp;layout=standard&amp;show_faces=true&amp;width=450&amp;action=like&amp;colorscheme=light&amp;height=80";
         alert(obj.src);
     }
     
 </script>
 <iframe name="test" src=""  oninit ="doit(this);" ></iframe>
 <script Language="JavaScript" type="text/JavaScript">
    doit(window.test);
 </script>
</body>
</html>