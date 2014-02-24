<%

    Dim lastrundate

    lastrundate = session("lastrundate")

    if (lastrundate <> Date()) then

  

        session("lastrundate") = Date()
        Dim vSATitles(100), vSATitleColors(100), vSATitleBackgrounds(100), vSATexts(100), vSATextColors(100), vSATextBackgrounds(100), vSAImages(100)
        dim moving, sql, cnt, vSATitle, vSAText, vSALink, vSATarget, vSASequence, vSAActive, vSAStartDate, vSAEndDate, vSADisplay, vSAImage
        dim vSATextColor, vSATitleColor, vTMP, vSATitleBackground, vSATextBackground
        
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT * " _
        & "FROM SlideAdvertiseMent with (nolock)" _
        & "WHERE Active LIKE 'Y' " _
        & "AND DATEDIFF(DAY, StartDate, GetDate()) >= 0 " _
        & "AND DATEDIFF(DAY, EndDate, GetDate()) <= 0 " _
        & "AND (SlideTypeid = 0 or SlideTypeID is null) ORDER BY Sequence  For Browse"
        rs.open sql, Conn, 3

        cnt = 0
        if not rs.eof then
	    moving = ""


            while not rs.EOF 
               vSATitle = rs("Title") & ""
                vSAText = rs("Text") & ""
                vSALink = rs("Link") & ""
                vSATarget = rs("Target") & ""
                vSASequence = rs("Sequence") & ""
                vSAActive = rs("Active") & ""
                vSAStartDate = rs("StartDate") & ""
                vSAEndDate = rs("EndDate") & ""
                vSADisplay = rs("Display") & ""
                vSAImage = rs("Image") & ""
                vSATextColor = rs("TextColor") & ""
                vSATitleColor = rs("TitleColor") & ""
                if vSALink <> "" Then
                vTMP = "<a href=""" & vSALink & """"
                if vSATarget <> "" Then vTMP = vTMP & " target=""" & vSATarget & """"
                vTMP = vTMP & "><font color=""" & vSATextColor & """>" & vSAText & "</font></a>"

                if vSAImage <> "" Then
                    vTMP2 = "<a href=""" & vSALink & """"
                    if vSATarget <> "" Then vTMP2 = vTMP2 & " target=""" & vSATarget & """"
                        vTMP2 = vTMP2 & "><img src=""" & vSAImage & """ border=""0"" align=""right""></a>"
                    else
                        vTMP2 = ""
                    end if
                else
                    vTMP = vSAText
                end if
                vSAText = vTMP
                vSAImage = vTMP2

                ' set up the backgrounds
                vSATitleColor = rs("TitleColor") & ""
                vSATextColor = rs("TextColor") & ""
                vSATitleBackground = rs("TitleBackground") & ""
                vSATextBackground = rs("TextBackground") & ""

                cnt = cnt + 1 

                moving = moving & "<div id=""animatedtext" & cnt  & """  style=""position:relative; width:200px; height:115px; z-index:1; left: 20px;   visibility: visible;"">"
                moving = moving & "<TABLE WIDTH=248 BORDER=0 CELLPADDING=3 CELLSPACING=0>"
                moving = moving & "<TR><TD style=""font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 11px; color:#" & vSATitleColor & " ; background:#" & vSATitleBackground & ";"">"
                moving = moving & "<table width=100% cellpadding=2 cellspacing=0 border=0><tr><td align=left style=""{font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 11px; color:#" & vSATitleColor & " ; background:#" & vSATitleBackground & " ;}"">"
                moving = moving & "<font color=""#" & vSATitleColor & """>" & vSATitle & " </font>"
                moving = moving & "</td><td align=right><a href=""#"" onclick=""MM_showHideLayers('animatedtext" & cnt & "','','hide')""><img src=""/images/closex.gif"" border=0></a></td></tr></table>"
                moving = moving & "</TD></TR><TR>"
                moving = moving & "<TD style=""font-family: Verdana, Arial, Helvetica; font-size: 11px; font-style: bold; color:#" & vSATextColor & " ; background:#" & vSATextBackground & " ; border: solid; border-style: solid; border-width: 2px 2px 2px 2px; border-color: 000000; background:#" & vSATextBackground & ";"">"
                moving = moving & ""
                moving = moving & "      " & vSAImage & " " & vSAText & " "
                moving = moving & ""
                moving = moving & "</TD>"
                moving = moving & "</TR>"
                moving = moving & "</TABLE></div>"

                rs.movenext
            wend
            moving = moving & "</div>"
            moving = "<div id=""animateplace"" style=""position:absolute; top:-" & 115 * cnt & "px; left:375px; width:200px;visibility: visible;"">"  & moving

        end if
        rs.close
    End if
 


    function getSlider(sType)  
        dim  rs
        dim vSATitles(100), vSATitleColors(100), vSATitleBackgrounds(100), vSATexts(100), vSATextColors(100), vSATextBackgrounds(100), vSAImages(100)
        dim moving, sql, cnt, vSATitle, vSAText, vSALink, vSATarget, vSASequence, vSAActive, vSAStartDate, vSAEndDate, vSADisplay, vSAImage
        dim vSATextColor, vSATitleColor, vTMP, vSATitleBackground, vSATextBackground
        Dim lastrundate

        lastrundate = session("lastrundate" & sType) 
        if lastrundate = date() then 
        else 
            
            session("lastrundate" & sType) = Date()
            
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT * " _
            & "FROM SlideAdvertiseMent with (nolock)" _
            & "WHERE Active LIKE 'Y' " _
            & "AND DATEDIFF(DAY, StartDate, GetDate()) >= 0 " _
            & "AND DATEDIFF(DAY, EndDate, GetDate()) <= 0 " _
            & "AND (SlideTypeid = " & sType & ") ORDER BY Sequence  For Browse"
            rs.open sql, Conn, 3

            cnt = 0
            moving = ""
            if not rs.eof then                
                while not rs.EOF 
                   vSATitle = rs("Title") & ""
                    vSAText = rs("Text") & ""
                    vSALink = rs("Link") & ""
                    vSATarget = rs("Target") & ""
                    vSASequence = rs("Sequence") & ""
                    vSAActive = rs("Active") & ""
                    vSAStartDate = rs("StartDate") & ""
                    vSAEndDate = rs("EndDate") & ""
                    vSADisplay = rs("Display") & ""
                    vSAImage = rs("Image") & ""
                    vSATextColor = rs("TextColor") & ""
                    vSATitleColor = rs("TitleColor") & ""
                    if vSALink <> "" Then
                    vTMP = "<a href=""" & vSALink & """ class=""slider"""
                    if vSATarget <> "" Then vTMP = vTMP & " target=""" & vSATarget & """"
                    vTMP = vTMP & "><font color=""" & vSATextColor & """  >" & vSAText & "</font></a>"

                    if vSAImage <> "" Then
                            vTMP2 = "<a href=""" & vSALink & """"
                            if vSATarget <> "" Then vTMP2 = vTMP2 & " target=""" & vSATarget & """"
                            vTMP2 = vTMP2 & "><img src=""" & vSAImage & """ border=""0"" align=""right"" width=""150px""></a>"
                        else
                            vTMP2 = ""
                        end if
                    else
                        vTMP = vSAText
                    end if
                    vSAText = vTMP
                    vSAImage = vTMP2

                    ' set up the backgrounds
                    vSATitleColor = rs("TitleColor") & ""
                    vSATextColor = rs("TextColor") & ""
                    vSATitleBackground = rs("TitleBackground") & ""
                    vSATextBackground = rs("TextBackground") & ""

                    cnt = cnt + 1 

                    moving = moving & "<div id=""divScrollText" & cnt  & """  style=""position:relative; width:400px; height:115px; z-index:1; left: 20px;background:#ffffff;visibility: visible;"">"
                    moving = moving & "<TABLE BORDER=""0"" style=""background:#ffffff;"">"
                    moving = moving & "<TR><TD colspan=""5"" class="".slider"" style="" color:#" & vSATitleColor & " ; background:#" & vSATitleBackground & ";"">"
                    moving = moving & "<table width=100% cellpadding=2 cellspacing=0 border=0>"
                    moving = moving & "<tr><td class=""sliderT"" style=""{ color:#" & vSATitleColor & " ; background:#" & vSATitleBackground & " ;}"">"
                    moving = moving & "<font color=""#" & vSATitleColor & """>" & vSATitle & " </font>"
                    moving = moving & "</td><td align=right><a href=""#"" onclick=""MM_showHideLayers('divScrollText" & cnt & "','','hide')"">"
                    moving = moving & "<img src=""/images/closex.gif"" border=0></a></td></tr></table>"
                    moving = moving & "</TD></TR><TR>"
                    moving = moving & "<TD width=""60%""  style="" color:#" & vSATextColor & " ; background:#" & vSATextBackground & " ; background:#" & vSATextBackground & ";"" class="".sliderTD"">"                                
                    moving = moving & "<font class="".sliderTD"">" &  vSAText & "</font></td><td > " & vSAImage & " "                
                    moving = moving & "</TD>"
                    moving = moving & "</TR>"
                    moving = moving & "</TABLE></div>"

                    rs.movenext
                wend
                
                moving = "<div id=""divScroll"" style=""position:absolute;top:-" & 500 * cnt & "px; left:150px;visibility: visible;"">"  & moving  & "</div>"

            end if
            rs.close
        End if
        
        getSlider =moving
    end function 

    %>
