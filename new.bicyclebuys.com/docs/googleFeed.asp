<% @ Language="VBScript" %>
<%
  ' Declare all variables.
  Dim objCN,objRS,objField
  Dim strSQL,strCN
  Dim strName,strValue

  ' Buffer and output as XML.
  Response.Buffer = True
  Response.ContentType = "text/xml"

  ' Start our XML document.
  Response.Write "<?xml version=""1.0""?>" & vbCrLf
  Response.Write "<rss version=""2.0"" xmlns:g=""http://base.google.com/ns/1.0"" xmlns:c=""http://base.google.com/cns/1.0"">" & vbCrLf

  ' Set SQL and database connection string.
  strSQL = "SELECT   * FROM googleFeed  "
  strCN = application("DSN")

  ' Open the database connection and recordset.
  Set objCN = Server.CreateObject("ADODB.Connection")
  objCN.Open strCN  
  Set objRS = objCN.Execute(strSQL)

  ' Output start of data.
  Response.Write "<channel><title>Google Shopping Feed</title><link>http://www.BicycleBuys.com</link><description>BicycleBuys.com Shopping Products</description>" & vbCrLf

  ' Loop through the data records.
  While Not objRS.EOF
    ' Output start of record.
    Response.Write "<item>" & vbCrLf
    ' Loop through the fields in each record.
    For Each objField in objRS.Fields
      strName  = objField.Name
      strValue = objField.Value
      If Len(strName)  > 0 Then strName = Server.HTMLEncode(strName)
      If Len(strValue) > 0 Then strValue = Server.HTMLEncode(strValue)
	if strName =  "image link" then strName = "image_link"
	if strName =  "payment notes" then strName = "payment_notes"

    if strName = "shipping" and strValue <>"" then
	      Response.Write "<g:shipping>"
	      Response.Write " <g:country>US</g:country>"
	      Response.Write " <g:region></g:region>"
	      Response.Write " <g:service>Ground</g:service>"
	      Response.Write " <g:price>0.00</g:price>"
	      Response.Write "</g:shipping>"
    elseif strName = "tax" then
        Response.Write "<g:tax>"
        Response.Write " <g:country>US</g:country>"
        Response.Write " <g:region>NY</g:region>"
        Response.Write " <g:rate>8.625</g:rate>"
        Response.Write " <g:tax_ship>y</g:tax_ship>"
        Response.Write "</g:tax>"
	elseif  strName = "price"  or strName = "condition"  or strName = "id" or strName = "payment_notes" then
	      Response.Write "<g:" & strName & ">"  
	      Response.Write   strValue   
	      Response.Write "</g:" & strName  & ">" & vbCrLf
 	else
	      Response.Write "<" & strName & ">"  
	      Response.Write   strValue   
	      Response.Write "</" & strName  & ">" & vbCrLf
	end if
    Next
    ' Move to next record in database.
    objRS.MoveNext
    ' Output end of record.
    Response.Write "</item>" & vbCrLf
  Wend
    objRS.close
    set objRS = nothing
  ' Output end of data.
  Response.Write "</channel></rss>" & vbCrLf 
%>
