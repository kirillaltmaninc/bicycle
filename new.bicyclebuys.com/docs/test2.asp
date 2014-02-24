<%

Sub putRebate(RebateCode, amount,useOne)

   Dim vRebates
   Dim counter
   vRebates = Session("Rebates")   
     
   If 10 >= UBound(vRebates) Then
        ' add 10 elements!
        ReDim Preserve vRebates( 10, 2 )
    End If
   
   If vRebates(10, 2) < 0 Then 
        vRebates(10, 2) = 0
   end if
   RebateCode = Left(RebateCode, 20)
   counter = 0
   found = False   
   Do While counter < vRebates(10, 2) And counter < 10
        If vRebates(counter, 0) = RebateCode Then
            found = True
            Exit Do
        End If
        counter = counter + 1
   Loop
   If found Then
        vRebates(counter, 1) = vRebates(counter, 1) + amount
        if useOne = 1 then vRebates(counter, 2) = useOne   'or vRebates(counter, 2)
   Else 
        vRebates(10, 2) = vRebates(10, 2) + 1
        vRebates(counter, 0) = RebateCode
        vRebates(counter, 1) = amount
        vRebates(counter, 2) = useOne   'or vRebates(counter, 2)
   End If
    session("Rebates")=vRebates
       
End Sub

sub zeroRebateArray()
   Dim vRebates, counter, counter2   
    If IsArray(session("Rebates")) then         
        vRebates = Session("Rebates")            
    else        
        ReDim vRebates( 10, 2 )
    end if
   
   ReDim Preserve vRebates(10, 2)
   
   Do While counter < vRebates(10, 2)
        if vRebates(counter, 2) = 0 then
            counter2=counter
            Do While counter2 < vRebates(10, 2)
                vRebates(counter2, 0) = vRebates(counter2+1, 0)
                vRebates(counter2, 1) = vRebates(counter2+1, 1)
                vRebates(counter2, 2) = vRebates(counter2+1, 2)                
                counter2 = counter2 + 1
            loop
            vRebates(10, 2)=vRebates(10, 2)-1
        end if
        vRebates(counter, 1) = 0
        counter = counter + 1
   Loop   
   Session("Rebates")=vRebates
end sub

sub AddRebateCode(RebateCode)
   call putRebate(RebateCode, 0, 1)
end sub

sub ApplyRebate(RebateCode, amount)
   call putRebate(RebateCode, amount, 0)
end sub


function getRebates()
   Dim vRebates, counter 
   vRebates = Session("Rebates")
   ReDim Preserve vRebates( 10, 2 )
   counter = 0
   getRebates =""
   Do While counter < vRebates(10, 2)
        if vRebates(counter, 2) = 1 and vRebates(counter, 1) > 0 then
            getRebates = getRebates & "<tr><td>" & vRebates(counter, 0) &  "</td><td>" & vRebates(counter, 1) & "</td></tr>"
        end if   
        counter = counter + 1
   Loop
   
end function


Sub testRebate()
   Dim counter, vRebates
   'Zero out array
   zeroRebateArray()
 
    call AddRebateCode("d")
    call AddRebateCode("SIDI")
    call AddRebateCode("xxx")
    Call ApplyRebate("d", 1)
    Call ApplyRebate("d", 3.1)
    Call ApplyRebate("SIDI", 10)
    Call ApplyRebate("SIDI", 20)
          
   response.Write "<table>" & getRebates() & "</table>"
End Sub
dim FB
FB ="<iframe src="""" scrolling=""no"" frameborder=""0"" style=""border:none; overflow:hidden; width:450px; height:80px;"" allowTransparency=""true""></iframe>"
response.Write FB

call testRebate()
 %>
 <iframe name="test" src=""></iframe>
 <script Language="JavaScript" type="text/JavaScript">
     test.src = "http://www.facebook.com/plugins/like.php?href=" & location.href & "/FB&amp;layout=standard&amp;show_faces=true&amp;width=450&amp;action=like&amp;colorscheme=light&amp;height=80"
 
 </script>