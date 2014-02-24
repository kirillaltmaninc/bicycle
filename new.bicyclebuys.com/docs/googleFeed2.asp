<% @ Language="VBScript" %>
<%

  ' Buffer and output as XML.
  Response.Buffer = True
  Response.ContentType = "text/xml"
  response.write ("Started")

  ' Declare all variables.
  Dim objCN,objRS,objField
  Dim strSQL,strCN
  Dim strName,strValue

	dim path
	path= Server.MapPath("/") 
	Dim fso, vOldFile, vNewFile, vNewFile1
	Set fso = CreateObject("Scripting.FileSystemObject")

 response.write(path)
	If (fso.FileExists(path + "\googleFeed2.xml")) then fso.DeleteFile(path + "\googleFeed2.xml")

 response.write("X1")
	Set vNewFile = fso.OpenTextFile(path + "\googleFeed2.xml", 2, True)

 response.write("X2")


  ' Start our XML document.
  vNewfile.WriteLine "<?xml version=""1.0""?>" & vbCrLf

 response.write("X3")


  vNewfile.WriteLine "<rss version=""2.0"" xmlns:g=""http://base.google.com/ns/1.0"" xmlns:c=""http://base.google.com/cns/1.0"">" & vbCrLf

  ' Set SQL and database connection string.
  strSQL = "SELECT * FROM googleFeed where not [image link] is null"
  strCN = application("DSN")

  ' Open the database connection and recordset.
  Set objCN = Server.CreateObject("ADODB.Connection")
  objCN.Open strCN  
  Set objRS = objCN.Execute(strSQL)

  ' Output start of data.
  vNewfile.WriteLine "<channel><title>Google Shopping Feed</title><link>http://www.BicycleBuys.com</link><description>BicycleBuys.com Shopping Products</description>" & vbCrLf

  ' Loop through the data records.
  While Not objRS.EOF
    ' Output start of record.
    vNewfile.WriteLine "<item>" & vbCrLf
    ' Loop through the fields in each record.
    For Each objField in objRS.Fields
      strName  = objField.Name
      strValue = objField.Value
      If Len(strName)  > 0 Then strName = Server.HTMLEncode(strName)
	strValue = nz(strValue,"")
      If Len(strValue) > 0 Then strValue = Server.HTMLEncode(strValue)

	strValue = replace("""","""" & """", strValue)
	if strName =  "image link" then strName = "image_link"
	if strName =  "payment notes" then strName = "payment_notes"

      	if strName = "shipping" and strValue <>"" then
	      vNewfile.WriteLine "<g:shipping>"
	      vNewfile.WriteLine " <g:country>US</g:country>"
	      vNewfile.WriteLine " <g:region></g:region>"
	      vNewfile.WriteLine " <g:service>Ground</g:service>"
	      vNewfile.WriteLine " <g:price>0.00</g:price>"
	      vNewfile.WriteLine "</g:shipping>"
	elseif  strName = "price"  or strName = "condition"  or strName = "id" or strName = "payment_notes" then
	      vNewfile.WriteLine "<g:" & strName & ">"  
	      vNewfile.WriteLine   strValue   
	      vNewfile.WriteLine "</g:" & strName  & ">" & vbCrLf
 	else
	      vNewfile.WriteLine "<" & strName & ">"  
	      vNewfile.WriteLine   strValue   
	      vNewfile.WriteLine "</" & strName  & ">" & vbCrLf
	end if
    Next
    ' Move to next record in database.
    objRS.MoveNext
    ' Output end of record.
    vNewfile.WriteLine "</item>" & vbCrLf
  Wend

  ' Output end of data.
  vNewfile.WriteLine "</channel></rss>" & vbCrLf 
 response.write("X5")

vNewFile.Close

response.write ("Finished")
    objRS.close
    set objRS = nothing
%>
