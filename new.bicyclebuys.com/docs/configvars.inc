<%

' --- Set these to the database values found in the shipping method table
vFreeShippingMethodID = 1
vFreeShippingMethod = "FEDEX Ground-US Mail"

' --- IISCart2000 Encryption key
vEncryptKey = "2002-WTC-HemLock-2003"

' --- Function to clean up strings for use in db
FUNCTION CS (s, endchar)
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
 CS="'" & s & "'" & endchar
END FUNCTION

' --- What state are we in?
Dim vThisState
vThisState = "NY"

' --- We need this for display purposes just about everywhere during the checkout process
Dim States
Set States = Server.CreateObject("Scripting.Dictionary")
States.Item("AP") = "APO/FPO"
States.Item("AL") = "Alabama"
States.Item("AK") = "Alaska"
States.Item("AZ") = "Arizona"
States.Item("AR") = "Arkansas"
States.Item("CA") = "California"
States.Item("CO") = "Colorado"
States.Item("CT") = "Connecticut"
States.Item("DE") = "Delaware"
States.Item("DC") = "District of Columbia"
States.Item("FL") = "Florida"
States.Item("GA") = "Georgia"
States.Item("HI") = "Hawaii"
States.Item("ID") = "Idaho"
States.Item("IL") = "Illinois"
States.Item("IN") = "Indiana"
States.Item("IA") = "Iowa"
States.Item("KS") = "Kansas"
States.Item("KY") = "Kentucky"
States.Item("LA") = "Louisiana"
States.Item("ME") = "Maine"
States.Item("MD") = "Maryland"
States.Item("MA") = "Massachusetts"
States.Item("MI") = "Michigan"
States.Item("MN") = "Minnesota"
States.Item("MS") = "Mississippi"
States.Item("MO") = "Missouri"
States.Item("MT") = "Montana"
States.Item("NE") = "Nebraska"
States.Item("NV") = "Nevada"
States.Item("NH") = "New Hampshire"
States.Item("NJ") = "New Jersey"
States.Item("NM") = "New Mexico"
States.Item("NY") = "New York"
States.Item("NC") = "North Carolina"
States.Item("ND") = "North Dakota"
States.Item("OH") = "Ohio"
States.Item("OK") = "Oklahoma"
States.Item("OR") = "Oregon"
States.Item("PA") = "Pennsylvania"
States.Item("PR") = "Puerto Rico"
States.Item("RI") = "Rhode Island"
States.Item("SC") = "South Carolina"
States.Item("SD") = "South Dakota"
States.Item("TN") = "Tennessee"
States.Item("TX") = "Texas"
States.Item("UT") = "Utah"
States.Item("VT") = "Vermont"
States.Item("VA") = "Virginia"
States.Item("WA") = "Washington"
States.Item("WV") = "West Virginia"
States.Item("WI") = "Wisconsin"
States.Item("WY") = "Wyoming"

' We need this variable so we can define a proper reference
' to the secure (httpS) vs. non-secure (http) URL's.
' Right now, only used in displaycart.asp on the checkout button
vThisServer = Request.ServerVariables("SERVER_NAME")
vThisPort = Request.ServerVariables("SERVER_PORT")

' Define where the order files are saved
vSaveOrderPath = "D:\JohnR\"
'vSaveOrderPath = "D:\PUBLIC\HTML\www.bicyclebuys.com\JohnR\"

' debuggin referral data
if vRemote_IP = "69.127.248.96"  Then response.write "<hr>Referred by: " & Session("ReferredBy")                                                                                                                                                                                                                                                                                                                                               
%>