<%

Dim Cart

set cart = server.createobject("iisCART2000.store")


if cart.register("trial.enc","DSN=liidsn;UID=iiscart;PWD=iiscart") then
'if cart.register("trial.enc") then  
  response.write "Key added sucessfully."
else
 response.write "Error adding key."
end if

set cart = nothing

%>