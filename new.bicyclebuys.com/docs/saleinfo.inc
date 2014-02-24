</TD><TD align="right" background="/images/sale_blank.gif">
	<img name="onsalethisweek" src="/images/onsalethisweek.gif" width="279" height="20" border="0"><BR>
	<img name="sale_space01" src="/images/sale_spacer01.gif" width="4" height="57" border="0"><% CornerSales %>
</TD><TD width="24" background="/images/right_sale.gif">
	<img name="right_sale" src="/images/right_sale.gif" width="24" height="84" border="0"><BR>
</TD></TR>
<%

Sub CornerSales()

   Set rscategories = Server.CreateObject("ADODB.Recordset")
   Set rshtmlspecials = Server.CreateObject("ADODB.Recordset")
   Set rsproduct = Server.CreateObject("ADODB.Recordset")

   FileDSN = Application("dsn")

   Set Conn1 = Server.CreateObject("ADODB.Connection")
   Conn1.Open FileDSN

   sql = "SELECT * FROM Categories WHERE Category LIKE 'special%'  For Browse"
   rscategories.open sql, conn1, 3
   
   if rscategories.EOF = 0 then
      do while not rscategories.eof
         sql = "SELECT * FROM HTML_Special_SaleItems WHERE NavTypeID=" & rscategories("CategoryID") & " For Browse"
'         response.write "<pre>" & sql & "</pre>"
         rshtmlspecials.open sql,Conn1,3

         sql = "SELECT * FROM vwWebproducts WHERE ProdID=" & rshtmlspecials("Col1_ProductID") & " For Browse"
'         response.write "<pre>" & sql & "</pre>"
         rsproduct.open sql,Conn1,3

         vImage = rsproduct("picture")
         if instr(vImage, "\") <> -1 then vImage = replace(vImage, "\", "/")
         vMiniWidth = rsproduct("Width_Large")
         vMiniHeight = rsproduct("Height_Large")

         vStaticWidth = 73
         vStaticHeight = 57

         'WIDTH SHOULD BE 39
         'HEIGHT SHOULD BE 57

         'vMiniWidth = 90
         'vMiniHeight = 130

         'vHeightDiff = 73
         'vHDiffPerc = 0.438461538461538461538461538462
         
         'vCalcHeight = 57
         'vCalcWidth = 50.53846153846153846153846153842

         vHeightDiff = vMiniHeight - vStaticHeight
         if vHeightDiff <> 0 then
            if vHeightDiff > 0 then
               vHDiffPerc = Abs(1 - (vHeightDiff / vMiniHeight))
               vCalcHeight = vStaticHeight
               vCalcWidth = int(vMiniWidth * vHDiffPerc)
               vWidthFullPadding = vStaticWidth - vCalcWidth
               vWidthFirstHalfPadding = int(vWidthFullPadding / 2)
               vWidthSecondHalfPadding = vWidthFullPadding - vWidthFirstHalfPadding
            End If
         End If
'         response.write "<pre>" & vCalcHeight & " x " & vCalcWidth & " - " & vHeightDiff & "</pre>"
         
         vURL = ""
         if vWidthFirstHalfPadding > 0 then
            vURL = "<img src=""/images/sale_spacer01.gif"" width=""" & vWidthFirstHalfPadding & """ height=""57"" border=""0"">"
         end if

         vURL = vURL & "<a href=""/Items01.asp?NavID=special&PID=" & rsproduct("ProdID") & "&pOP=" & rshtmlspecials("Col1_Original_Price") & """>"

'         vURL = vURL & "<img src=""/productimages/" & vImage & """ height=""57"" width=""73"" alt=""" & rsproduct("description") & " on sale now for just " & Formatcurrency(rsproduct("price")) & "!!"" border=""1"">"
         vURL = vURL & "<img src=""/productimages/" & vImage & """ height=""" & vCalcHeight & """ width=""" & vCalcWidth & """ alt=""" & rsproduct("description") & " on sale now for just " & Formatcurrency(rsproduct("price")) & "!!"" border=""1"">"

         vURL = vURL & "</A>"

         if vWidthSecondHalfPadding > 0 then
            vURL = vURL & "<img src=""/images/sale_spacer01.gif"" width=""" & vWidthSecondHalfPadding & """ height=""57"" border=""0"">"
         end if

         if rscategories("Category")= "special1" then vURL = vURL & "<img name=""sale_space02"" src=""/images/sale_spacer02.gif"" width=""8"" height=""57"" border=""0"">"
         if rscategories("Category")= "special2" then vURL = vURL & "<img name=""sale_space02"" src=""/images/sale_spacer02.gif"" width=""9"" height=""57"" border=""0"">"
         if rscategories("Category")= "special3" then vURL = vURL & "<img name=""sale_space03"" src=""/images/sale_spacer03.gif"" width=""3"" height=""57"" border=""0"">"
         response.write vURL
         rsproduct.close
         rshtmlspecials.close
         rscategories.movenext
      loop
      response.write "<BR><img name=""belowsale"" src=""/images/belowsale.gif"" width=""279"" height=""4"" border=""0""><BR>"
   End If
   rscategories.close
   Conn1.close
End Sub

%>