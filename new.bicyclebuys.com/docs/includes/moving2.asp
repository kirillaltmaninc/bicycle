
<%
response.write("TEST")
    Dim lastrundate

    lastrundate = session("lastrundate")

    if (lastrundate <> Date()) or 1=1 then

   Dim dsn, conn
   dsn = Application("dsn")

   Set conn = Server.CreateObject("ADODB.Connection")
   conn.Open dsn

        session("lastrundate") = Date()
        Dim vSATitles(100), vSATitleColors(100), vSATitleBackgrounds(100), vSATexts(100), vSATextColors(100), vSATextBackgrounds(100), vSAImages(100)
        dim moving, sql, cnt, vSATitle, vSAText, vSALink, vSATarget, vSASequence, vSAActive, vSAStartDate, vSAEndDate, vSADisplay, vSAImage
        dim vSATextColor, vSATitleColor, vTMP, vSATitleBackground, vSATextBackground
        
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT * " _
        & "FROM SlideAdvertiseMent " _
        & "WHERE Active LIKE 'Y' " _
        & "AND DATEDIFF(DAY, StartDate, GetDate()) >= 1 " _
        & "AND DATEDIFF(DAY, EndDate, GetDate()) <= 1  or slideadvertid=99" _
        & "ORDER BY Sequence  For Browse"
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
                moving = moving & "</TABLE></div>x"

                rs.movenext
            wend
            moving = moving & "</div>"
            moving = "<div id=""animateplace"" style=""position:absolute; top:-" & 115 * cnt & "px; left:275px; width:200px;visibility: visible;"">"  & moving

        end if
        rs.close

    End if
 
response.write(moving)
    %>



<script language="JavaScript" type="text/javascript" >

document.MM_Time = 0;
function MM_timelinePlay(tmLnName, myID) { //v1.2
var the_timeout;
 if (document.MM_Time ==0 || getStyleObject("animateplace")!=null){
 	the_timeout=setTimeout("move();",50);
	 return false;
	document.MM_Time=1;
}
}


function moveDiv(end, x)
{
var the_timeout;
// get the stylesheet
//
var the_style = getStyleObject("animateplace");
if (the_style)
{

// get the current coordinate and add 5
//
var current_top = parseInt(the_style.top);
var new_top = current_top + 1 ;
var rate= 3*(1.00-(parseFloat(new_top) / parseFloat(end))) ;
 new_top = new_top  + parseInt(2*rate) ;


// set the left property of the DIV, add px at the
// end unless this is NN4
//
if (document.layers) 
{
the_style.top = new_top;
if ((parseInt(f_clientWidth() ) -100 ) > parseInt(the_style.left) + parseInt(the_style.width) +4) {
	the_style.left = parseInt(the_style.left) +1 + parseInt(2*rate) ;
	}
}
else 
{ 
the_style.top = new_top + "px";
if ((parseInt(f_clientWidth() ) -100 ) > parseInt(the_style.left) + parseInt(the_style.width) +4) {
	the_style.left = (parseInt(the_style.left) +1 + parseInt(2*rate)) + "px";
	}


}

// if we haven't gone to far, call moveDiv() again in a bit
// 
if (new_top < end)
{
//alert (the_style.top + "sd" + new_top);
the_timeout = setTimeout('moveDiv(' + end + ',' + x + ');',30);
} 
}
}


function MM_showHideLayers() { //v6.0

  var e,f,i,p,v,obj,args=MM_showHideLayers.arguments;
  obj=document.getElementById('animateplace');
  if (obj.style) { obj=obj.style;}
  e=parseInt(obj.top)/115;

  for (i=0; i<(args.length-2); i+=3) if ((obj=document.getElementById(args[0]))) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
	f=0;
  for (i=1;i<=e;i++){
	alert('animatedtext' + i);
	obj= document.getElementById('animatedtext' + i);
	if (obj.style) { obj=obj.style;}
	alert(obj.visibility);
	if (obj.visibility.substring(0,1)=='v'){
		f=1;	
		}

	}
  if (f==0) {
	obj=document.getElementById('animateplace')
	obj.style.top=-590;
	}
}



function getStyleObject(objectId) {
// cross-browser function to get an object's style object given its
if(document.getElementById && document.getElementById(objectId)) {
// W3C DOM
return document.getElementById(objectId).style;
} else if (document.all && document.all(objectId)) {
// MSIE 4 DOM
return document.all(objectId).style;
} else if (document.layers && document.layers[objectId]) {
// NN 4 DOM.. note: this won't find nested layers
return document.layers[objectId];
} else {
return false;
}
} // getStyleObject


function move()
{

var the_style = getStyleObject("animateplace");
if (the_style) {

var start = parseInt(the_style.top);
var end = 230 ; 
moveDiv(end, 200);
}

}

function f_filterResults(n_win, n_docel, n_body) {
	var n_result = n_win ? n_win : 0;
	if (n_docel && (!n_result || (n_result > n_docel)))
		n_result = n_docel;
	return n_body && (!n_result || (n_result > n_body)) ? n_body : n_result;
}


function f_clientWidth() {
	return f_filterResults (
		window.innerWidth ? window.innerWidth : 0,
		document.documentElement ? document.documentElement.clientWidth : 0,
		document.body ? document.body.clientWidth : 0
	);
}
function f_clientHeight() {
	return f_filterResults (
		window.innerHeight ? window.innerHeight : 0,
		document.documentElement ? document.documentElement.clientHeight : 0,
		document.body ? document.body.clientHeight : 0
	);
}
function f_scrollLeft() {
	return f_filterResults (
		window.pageXOffset ? window.pageXOffset : 0,
		document.documentElement ? document.documentElement.scrollLeft : 0,
		document.body ? document.body.scrollLeft : 0
	);
}
function f_scrollTop() {
	return f_filterResults (
		window.pageYOffset ? window.pageYOffset : 0,
		document.documentElement ? document.documentElement.scrollTop : 0,
		document.body ? document.body.scrollTop : 0
	);
}



move();


</script>




