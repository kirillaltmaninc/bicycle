<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Promotions.aspx.vb" Inherits="Promotions" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>BicycleBuys.com Promotions - Online Bike Shop</title>
    <link rel="stylesheet" type="text/css" href="/index.css" title="index">

    <script language="JavaScript" type="text/javascript"> 
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

    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta name="GENERATOR" content="Microsoft Notepad">
</head>
<body bgcolor="#E5E5F0" text="#000000" link="#3770A8" vlink="#3770A8" alink="#FFFFFF"
    topmargin="0" marginheight="0" leftmargin="0" marginwidth="0">
     
    <table width="400px" height="90%" border="0" cellpadding="0" cellspacing="0" id="tb100P">
        <tr>
            <td width="214" height="45" background="/cartimages/shiptop_bkg.gif">
                <br>
            </td>
            <td width="186" height="45" align="right" background="/cartimages/shiptop_bkg.gif">
                <a href="javascript:window.close()">
                    <img src="/cartimages/closewindow_top.gif" width="86" height="45" border="0"></a><br>
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF">
                <br>
                <table width="300px" border="0" cellpadding="0" cellspacing="0" id="tb90P">
                    <tr style="height: 20px">
                        <td align="center" colspan="3" >
                            <b>"too unbelievable to display"</b> 
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" style="text-align:justify">
                            <br /><b>Why aren't you able to see your cost?</b><br />
Certain manufacturers have required that BicyclesBuys.com publish only MSRP pricing, but we are able to give our customers any price we choose. The "too unbelievable to display" message means a discount is in effect for our customers, which we calculate for you once the item is placed in your shopping cart. Simply add the item to your cart and see your price!<br />
<br />
Of course, adding the item to your cart does not obligate you to buy it - you are able to remove it at any time.<br />

                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <table border="0">
                                <tr colspan="2">
                                    <td>
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td valign="top">
                                        &nbsp;
                                    </td>
                                    <td align="left">
                                        <font id="cartnormal"><b>Here is your promotional code to use in your cart or at checkout:</b><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label
                                            ID="Label1" runat="server" Text="Label" Font-Bold="True" 
                                            Font-Size="Medium" ForeColor="#333300"></asp:Label>
                                        </font>                                        
                                    </td>
                                </tr>
                            </table>
                </table>
            </td>
        </tr>
        <tr background="/cartimages/shipbottom_bkg.gif" style="background: /cartimages/shipbottom_bkg.gif">
            <td width="214" height="29" background="/cartimages/shipbottom_bkg.gif">
                <img src="/cartimages/bb_mini.gif" width="214" height="29" border="0">
            </td>
            <td height="29" align="right" background="/cartimages/shipbottom_bkg.gif">
                <a href="javascript:window.close()">
                    <img src="/cartimages/closewindow_bottom2.gif" width="86" height="29" border="0"></a><br>
            </td>
        </tr>
    </table>
</body>
</html>
