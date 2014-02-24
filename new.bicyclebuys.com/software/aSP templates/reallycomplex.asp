<!--#INCLUDE FILE="includes/template_cls.asp"-->
<!--#INCLUDE FILE="includes/functions.asp"-->
<%
response.expires = -1
dim intID
intID = getID()
call doTemplate(intID)
response.end
function getID
	dim intValue
	getID = 0
	intValue = request.querystring("ID")
	if isnumeric(intValue) then
		if not(isempty(intValue)) then
			getID = intValue
		end if
	end if
end function
sub doTemplate(inID)
	dim objTemplate, strPath
	set objTemplate = new template_cls
	with objTemplate
		select case intID
			case 0
				strPath = "/templates/default/tmpl/"
				.TemplateFile = strPath & "main.html"
				.AddToken "imgpath", 1, "/templates/default/images/"
				.AddToken "content", 3, "/content/main.html"
				.AddToken "source", 1, strip(.loadFile("/reallycomplex.asp", 1))
			case 1
				strPath = "/templates/4guys/tmpl/"
				
				.AddToken "imgpath", 1, "/templates/4guys/images/"
				.AddToken "4guys.style", 2, strPath & "style.html"
				.AddToken "4guys.header", 2, strPath & "header.html"
				.AddToken "OAS.JS", 3, strPath & "OAS-JS.html"
				.AddToken "OAS.486X60_1", 2, strPath & "OAS486X60-1.html"
				.AddToken "OAS.486X60_2", 2, strPath & "OAS486X60-2.html"
				.AddToken "4guys.left", 2, strPath & "left.html"
				.AddToken "4guys.center", 2, strPath & "center.html"
				.AddToken "4guys.right", 2, strPath & "right.html"
				.AddToken "4guys.footer", 2, strPath & "footer.html"
				.AddToken "content", 1, FourGuysBox(.loadFile("/content/main.html", 1), "ASP Template v1.5")
				.AddToken "source", 1, FourGuysBox(strip(.loadFile("/reallycomplex.asp", 1)), "Source Code")
				.TemplateFile = strPath & "template.html"
		end select
		.AddToken "SELECT" & intID, 1, " SELECTED"
		.AddToken "tmplMenu", 2, "/includes/tmpl/form.html"
	response.write(.getParsedTemplateFile)
	end with
	set objTemplate = nothing
end sub
function strip(inItem)
	inItem = replace(inItem, "<", "&lt;")
	inItem = replace(inItem, ">", "&gt;")
	inItem = replace(inItem, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
	inItem = replace(inItem, VbCrLf, "<br>")
	strip = inItem
end function
function FourGuysBox(inContent, inTitle)
	dim objTemplate
	set objTemplate = new template_cls
	with objTemplate
		.AddToken "content", 1, inContent
		.AddToken "title", 1, inTitle
		.AddToken "date", 1, formatdatetime(now, 4)
		.TemplateFile = "/templates/4guys/tmpl/Box.html"
		FourGuysBox = .getParsedTemplateFile
	end with
	set objTemplate = nothing
end function
%>