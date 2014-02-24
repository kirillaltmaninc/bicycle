<%

Dim Cart

set cart = server.createobject("iisCART2000.store")


if cart.register("2132.enc","DSN=liidsn;UID=iiscart;PWD=iiscart") then
'if cart.register("2132.enc") then  
  response.write "Key added sucessfully."
else
 response.write "Error adding key."
end if

set cart = nothing

%>