
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
var new_top = current_top + 2;

// set the left property of the DIV, add px at the
// end unless this is NN4
//
if (document.layers) 
{
the_style.top = new_top;
}
else 
{ 
the_style.top = new_top + "px";
}

// if we haven't gone to far, call moveDiv() again in a bit
// 
if (new_top < end)
{
//alert (the_style.top + "sd" + new_top);
the_timeout = setTimeout('moveDiv(' + end + ',' + x + ');',40);
} 
}
}


function MM_showHideLayers() { //v6.0

  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=document.getElementById("animatedtext"))) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
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
var end = start + 200;
moveDiv(end, 200);
}

}
