<!--#INCLUDE file="includes/template_cls.asp"-->
<%
'response.write "hey" & session("RecentlyViewed")
'response.end %>
<!--#INCLUDE file="includes/common.asp"-->
<%
   ' BICYCLEBUYS.COM
   '
   ' (c)2006 - Bicyclebuys.com all rights reserved
   '
   ' sizing.asp

   vSubmit = Trim(Escape(Left(request("submit"), 4)))
   vSearchTerm = Trim(Escape(Left(Request("searchterm"), 100)))
   vSearchVendID = Trim(Escape(Left(Request("v"), 4)))
   vSearchCategory =  Trim(Escape(Left(Request("searchcategory"), 4)))

   vSearchPage = getsearch

   with objTemplate
   	.TemplateFile = TMPLDIR & "instructions.html"

   	.AddToken "header", 3, vHeader
   	.AddToken "search_section", 1, vSearchPage
   	.AddToken "footer", 3, vFooter

   	.parseTemplateFile
   end with
   set objTemplate = nothing

%>
