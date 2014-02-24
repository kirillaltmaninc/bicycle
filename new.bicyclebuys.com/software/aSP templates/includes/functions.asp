<%
function processRows(inObj, inString)
	dim objT, intCount
	set objT = new template_cls
	with objT
	do while not(inObj.EOF)
		for intCount = 0 to inObj.Fields.Count - 1
			.AddToken inObj(intCount).Name, 1, inObj(intCount).Value
		next
		processRows = processRows & .getParsedTemplateString(inString)
		.RemoveAllTokens
		inObj.MoveNext
	loop
	end with
	set objT = nothing
end function
%>