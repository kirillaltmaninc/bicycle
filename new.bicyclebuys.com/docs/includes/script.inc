<% FUNCTION reqform(vFieldName)

   if Not isNull(request(vFieldName)) Then
      vTMP = Left(Server.HTMLEncode(Request(vFieldName)), 20)
   else
      vTMP = Left(Server.HTMLEncode(Request.QueryString(vFieldName)), 20)
   end if
   reqform = vTMP
   ' response.write "<hr>" & vfieldname & "/" & vTMP

END Function %><script language="JavaScript" TYPE="text/javascript">
<!-- hide this script from non-javascript-enabled browsers

if (document.images) {
shopcart_F1 = new Image(127,22); shopcart_F1.src = "/images/shopcart.gif";
shopcart_F2 = new Image(127,22); shopcart_F2.src = "/images/shopcart_F2.gif";
shopinfo_F1 = new Image(128,22); shopinfo_F1.src = "/images/shopinfo.gif";
shopinfo_F2 = new Image(128,22); shopinfo_F2.src = "/images/shopinfo_F2.gif";
homelogo_F1 = new Image(115,35); homelogo_F1.src = "/images/logo_07.gif";
homelogo_F2 = new Image(115,35); homelogo_F2.src = "/images/logo_07_F2.gif";
home_button_F1 = new Image(33,35); home_button_F1.src = "/images/home_button.gif";
home_button_F2 = new Image(33,35); home_button_F2.src = "/images/home_button_F2.gif";
backtotop_F1 = new Image(61,32); backtotop_F1.src = "/images/backtotop.gif";
backtotop_F2 = new Image(61,32); backtotop_F2.src = "/images/backtotop_F2.gif";
}

/* Function that swaps images. */

function di20(id, newSrc) {
    var theImage = FWFindImage(document, id, 0);
    if (theImage) {
        theImage.src = newSrc;
    }
}

/* Functions that track and set toggle group button states. */

function FWFindImage(doc, name, j) {
    var theImage = false;
    if (doc.images) {
        theImage = doc.images[name];
    }
    if (theImage) {
        return theImage;
    }
    if (doc.layers) {
        for (j = 0; j < doc.layers.length; j++) {
            theImage = FWFindImage(doc.layers[j].document, name, 0);
            if (theImage) {
                return (theImage);
            }
        }
    }
    return (false);
}

/* Function to automatically go to a new page when picking from a dropdown list */
function load1(form, win) {
  // vendorid - a reference to the select object
  // win - a reference to the window object
  win.location.href = form.vendorid.options[form.vendorid.selectedIndex].value
}

function load2(form, win) {
	// menu - a reference to the select object
	// win - a reference to the window object
	win.location.href = form.SHIPTYPE.options[form.SHIPTYPE.selectedIndex].value
}

function load3(form, win) {
	// menu - a reference to the select object
	// win - a reference to the window object
	win.location.href = form.SHIPSTATEPROVINCE.options[form.SHIPSTATEPROVINCE.selectedIndex].value
}

function load4(form, win) {
	// menu - a reference to the select object
	// win - a reference to the window object
	win.location.href = form.SHIPCOUNTRY.options[form.SHIPCOUNTRY.selectedIndex].value
}

function openpopwin(windowpage, popupwidth, popupheight) {
	window.open(windowpage, '', 'width=' + popupwidth + ',height=' + popupheight +
	',location=no,toolbar=no,menubar=no,scrollbars=yes,resizable=yes');
}

function openpopwin1(windowpage, popupwidth, popupheight) {
	window.open(windowpage, '', 'width=' + popupwidth + ',height=' + popupheight +
	',location=no,toolbar=yes,menubar=yes,scrollbars=yes,resizable=yes');
}

// stop hiding -->
</script>
<meta http-equiv="Content-Type"
content="text/html; charset=iso-8859-1">
 
<meta name="GENERATOR" content="Microsoft Notepad">



