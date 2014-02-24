<%
dim curUrl
'curUrl =Request.ServerVariables("ALL_HTTP")
'curUrl = Request.ServerVariables("HTTP_URL")
response.write(Request.ServerVariables("ALL_HTTP"))	

if curUrl="/index.asp?c=item&i=0701253PART&tk=louis-garneau-ghent-shoe-cover" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/clothing/socks/0701252-3PART"
end if

if curUrl="/index.asp?c=item&i=1500435&tk=ritchey-design-superlogic-titanium-bottom-bracket-68x103mm-english" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/drivetrain/bbbrackets/FREMRE"
end if

if curUrl="/index.asp?c=item&i=0402038PART&tk=crank-brothers-acid-1-clipless-pedals" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/pedals/pedalsmtb/0402055"
end if

if curUrl="/index.asp?c=item&i=0601546&tk=continental-rim-cement-tire-glue-for-tubular-carbon-rim" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/tires/tiresroad/0601510"
end if

if curUrl="/index.asp?c=item&i=1428834&tk=sram-red-double-tap-shifter-lever-set" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/drivetrain/shiftersroad/1428838"
end if

if curUrl="/index.asp?c=item&i=0790160PART&tk=louis-garneau-womens-santa-cruz-baggy-short" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/clothing/shorts/0701177PART"
end if

if curUrl="http://www.bicyclebuys.com/item/11Y0508000/shimano-dura-ace-wh-7850-c24-cl-carbon-clincher-road-wheelset" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/wheels/wheelroad/0508002"
end if

if curUrl="/index.asp?c=item&i=0790173PART&tk=louis-garneau-thermo-cool-insole" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/shoes/shoeroad/LGCA5PART"
end if

if curUrl="/index.asp?c=item&i=0509471&tk=alex-dm24-24-rear-wheel-black" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/wheels/wheelbmx/20X175ALWH1"
end if

if curUrl="/index.asp?c=item&i=1010016&tk=sram-double-tap-flat-bar-10sp-shifters" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/drivetrain/shiftersroad/1010180"
end if

if curUrl="http://www.bicyclebuys.com/item/1600011/gt-speed-series-saddle" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/seats/seatatb/0921650"
end if

if curUrl="/index.asp?c=item&i=0791499PART&tk=sugoi-cannondale-liquigas-pro-cycling-team-jersey-full-zipper" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/clothing/jerseys/0710131PART"
end if

if curUrl="/index.asp?c=item&i=1010229PART&tk=shimano-rd-2300-rear-derailleur" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/drivetrain/deraillrear/1010227PART"
end if

if curUrl="/index.asp?c=item&i=0740478PART&tk=pearl-izumi-womens-symphony-short" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/clothing/shorts/0740412PART"
end if

if curUrl="/index.asp?c=item&i=0601505&tk=continental-rim-cement-tire-glue-for-tubular-aluminum-rim" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/tires/tiresroad/0601510"
end if

if curUrl="/index.asp?c=item&i=0508653&tk=shimano-wh-rs30-a-road-wheelset-white" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/wheels/wheelroad/0508002"
end if

if curUrl="/index.asp?c=item&i=0740487PART&tk=pearl-izumi-womens-superstar-knicker" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/clothing/tights/0790014PART"
end if

if curUrl="/index.asp?c=item&i=1504337&tk=shimano-un54-bottom-bracket-68x107-shell" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/"
end if

if curUrl="/index.asp?c=bikes&d=bmxfreestylebikes" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/bikes/bmx-freestyle-bikes/"
end if

if curUrl="/index.asp?c=bikes&d=bikeshybrid" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/bikes/hybrid-bikes/"
end if

if curUrl="/index.asp?c=bikes&d=roadframe" then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location","http://www.bicyclebuys.com/bikes/road-frame/"
end if


%>