<!--#INCLUDE file="includes/template_cls.asp"-->
<%
'Simple1.asp
dim objTemplate
const TMPLDIR = "/templates/Complex/tmpl/"
const IMGDIR = "/templates/Complex/images/"
set objTemplate = new template_cls
with objTemplate
	.TemplateFile = TMPLDIR & "template.html"
	.AddToken "date", 1, formatdatetime(now(), 3)
	.AddToken "title", 1, "ASP Templates [1.5] Complex"
	.AddToken "header", 2, TMPLDIR & "header.html"
	.AddToken "footer", 2, TMPLDIR & "footer.html"
	.AddToken "images", 1, IMGDIR
	.AddToken "body", 3, TMPLDIR & "text.html"
	.parseTemplateFile
end with
set objTemplate = nothing
%>