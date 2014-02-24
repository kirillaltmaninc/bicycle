<!--#INCLUDE file="includes/template_cls.asp"-->
<%
'Simple1.asp
dim objTemplate
const TMPLDIR = "/templates/Simple/tmpl/"
const IMGDIR = "/templates/Simple/images/"
set objTemplate = new template_cls
with objTemplate
	.TemplateFile = TMPLDIR & "template.html"
	.AddToken "date", 1, formatdatetime(now(), 3)
	.AddToken "title", 1, "ASP Templates [1.5] Simple"
	.AddToken "images", 1, IMGDIR
	.parseTemplateFile
end with
set objTemplate = nothing
%>