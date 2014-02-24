
document.MM_Time = 0;

function f_filterResults(n_win, n_docel, n_body) {
    var n_result = n_win ? n_win : 0;
    if (n_docel && (!n_result || (n_result > n_docel)))
        n_result = n_docel;
    return n_body && (!n_result || (n_result > n_body)) ? n_body : n_result;
}


function f_clientWidth() {
    return f_filterResults(
		window.innerWidth ? window.innerWidth : 0,
		document.documentElement ? document.documentElement.clientWidth : 0,
		document.body ? document.body.clientWidth : 0
	);
}
function f_clientHeight() {
    return f_filterResults(
		window.innerHeight ? window.innerHeight : 0,
		document.documentElement ? document.documentElement.clientHeight : 0,
		document.body ? document.body.clientHeight : 0
	);
}
function f_scrollLeft() {
    return f_filterResults(
		window.pageXOffset ? window.pageXOffset : 0,
		document.documentElement ? document.documentElement.scrollLeft : 0,
		document.body ? document.body.scrollLeft : 0
	);
}
function f_scrollTop() {
    return f_filterResults(
		window.pageYOffset ? window.pageYOffset : 0,
		document.documentElement ? document.documentElement.scrollTop : 0,
		document.body ? document.body.scrollTop : 0
	);
}


function MM_timelinePlay(tmLnName, myID, end, moveRightZero) { //v1.2
    var the_timeout;
    if (document.MM_Time == 0 || getStyleObject(myID.toString()) != null) {

        the_timeout = setTimeout("move('" + myID + "'," + end + ", " + moveRightZero + ");", 3);
        return false;
        document.MM_Time = 1;
    }
}


function move(myID, end, moveRightZero) {

    var the_style = getStyleObject(myID);
    if (the_style) {

        var start = parseInt(the_style.top);
        //        var end = 230;
        moveDiv(end, 200, myID, moveRightZero);
    }

}

function moveDiv(end, x, myID, moveRightZero) {
    var the_timeout;
    // get the stylesheet
    //
    var the_style = getStyleObject(myID);
    if (the_style) {

        // get the current coordinate and add 
        var current_top = parseInt(the_style.top);
        var new_top = current_top + 1;
        var rate = 2 * (1.00 - (parseFloat(new_top) / parseFloat(end)));
        new_top = new_top + parseInt(2 * rate);
        // alert("move('" + moveRightZero + "');");
        // set the left property of the DIV, add px at the
        // end unless this is NN4
        if (document.layers) {
            the_style.top = new_top;
            if (moveRightZero < 1 && (parseInt(f_clientWidth()) - 100) > parseInt(the_style.left) + parseInt(the_style.width) + 4) {
                the_style.left = parseInt(the_style.left) + 1 + parseInt(2 * rate);
            }
        } else {
            the_style.top = new_top + "px";
            if (moveRightZero < 1 && (parseInt(f_clientWidth()) - 100) > parseInt(the_style.left) + parseInt(the_style.width) + 4) {
                the_style.left = (parseInt(the_style.left) + 1 + parseInt(2 * rate)) + "px";
            }
        }

        // if we haven't gone to far, call moveDiv() again in a bit
        if (new_top < end) {
            //alert (the_style.top + "sd" + new_top);
            the_timeout = setTimeout('moveDiv(' + end + ',' + x + ',"' + myID + '",' + moveRightZero + ');', 30);
        }
    }
}



function MM_showHideLayers() { //v6.0

    var e, f, i, p, v, obj, args = MM_showHideLayers.arguments;
    obj = document.getElementById(args[0]);
    if (obj.style) { obj = obj.style; }
    e = parseInt(obj.top) / 115;

    /* for (i = 0; i < (args.length - 2); i += 3) if ((obj = document.getElementById(args[0]))) {
    v = args[i + 2];
    if (obj.style) { obj = obj.style; v = (v == 'show') ? 'visible' : (v == 'hide') ? 'hidden' : v; }
    obj.visibility = v;
    }
    */
    f = 0;

    //Set the current item to hidden 
    obj = document.getElementById(args[0]);
    if (obj.style) {
        obj = obj.style;
        obj.top = '-890px';
    } else {
        obj.top = -890;
    }

    //see if all children are hidden
    var objP = document.getElementById('animateplace');
    //alert(objP.childNodes.length);
    for (i = 1; i <= objP.childNodes.length; i++) {
        obj = document.getElementById('animatedtext' + i);
        if (obj.style) { obj = obj.style; }
        if (!(parseInt(obj.top) < 0)) { f = 1; }
    }
    //move div if all children are hidden 
    if (f == 0) {

        objP = document.getElementById('animateplace');
        if (objP.style) {
            objP = objP.style;
            objP.top = '-890px';
        } else {
            objP.top = -890;
        }
    }

}


function getStyleObject(objectId) {
    var obj = document.getElementById(objectId);

    if (obj != null) {
        if (obj.style) {
            obj = obj.style;
        }
    }
    return obj;

} // getStyleObject





function closeFloat() {
    var floatDiv = 'pfloatTXT';
    moveFL(-500, -500, getStyleObjectID(floatDiv));
}

function setmarkdesc(aval, vObj) {

    var obj = document.getElementById('pfloatTXT');
    var the_style = getStyleObjectID('pfloatTXT');

    // set the left property of the DIV, add px at the
    // end unless this is NN4
    var a = findPos(vObj);
    var ww = getWindowWidth();
    a.top = a.top + 40;

    if ((a.left + 450) < ww) {
        obj.innerHTML = '<b>' + vObj.alt + '</b><BR> ' + aMostPop[aval];
        moveFL(a.top, (a.left + 100), the_style);
    } else {
        obj.innerHTML = '<b>' + vObj.alt + '</b><BR> ' + aMostPop[aval];
        if (a.left - 450 < 0) {
            moveFL(a.top, 0, the_style);
        } else {
            moveFL(a.top, (a.left - 450), the_style);
        }
    }

    /*    if (a.left + 300 >  ww/ 2 && (a.left - 450) > 0 || (a.left + 450) > ww) {
    obj.innerHTML = '<b>' + vObj.alt + '</b><BR> ' + aMostPop[aval];
    if (a.left - 450 < 0) {
    moveFL(a.top, 0, the_style);
    } else {
    moveFL(a.top, (a.left - 450), the_style);
    }
    } else {
    obj.innerHTML = '<b>' + vObj.alt + '</b><BR> ' + aMostPop[aval];
    moveFL(a.top, (a.left + 100), the_style);
    }
    */
}
function getWindowWidth() {
    var myWidth = 0
    if (typeof (window.innerWidth) == 'number') {
        //Non-IE
        myWidth = window.innerWidth;
    } else if (document.documentElement && (document.documentElement.clientWidth || document.documentElement.clientHeight)) {
        //IE 6+ in 'standards compliant mode'
        myWidth = document.documentElement.clientWidth;
    } else if (document.body && (document.body.clientWidth || document.body.clientHeight)) {
        //IE 4 compatible
        myWidth = document.body.clientWidth;
    }
    return myWidth;
}

function findPos(obj) {
    var curleft = curtop = 0;
    if (obj.offsetParent) {
        curleft = obj.offsetLeft
        curtop = obj.offsetTop
        while (obj = obj.offsetParent) {
            curleft += obj.offsetLeft
            curtop += obj.offsetTop
        }
    }
    return { left: curleft, top: curtop };
}


function getStyleObjectID(objectId) {
    var obj = document.getElementById(objectId);

    if (obj != null) {
        if (obj.style) {
            obj = obj.style;
        }
    }
    return obj;

} // getStyleObject

function moveFL(top, left, the_style) {
    // set the left property of the DIV, add px at the
    // end unless this is NN4
    // var the_style = getStyleObject("pfloatTXT");
    //   obj.top = top;
    //    obj.left = left;

    if (document.layers) {
        obj.top = top;
        obj.left = left;
    } else {
        //the_style.top = top.toString() + "px";
        //the_style.left = left.toString() + "px";
    }
}


function moveIMG() {
    var myWidth = 0, myHeight = 0;
    if (typeof (window.innerWidth) == 'number') {
        //Non-IE
        myWidth = window.innerWidth;
        myHeight = window.innerHeight;
    } else if (document.documentElement && (document.documentElement.clientWidth || document.documentElement.clientHeight)) {
        //IE 6+ in 'standards compliant mode'
        myWidth = document.documentElement.clientWidth;
        myHeight = document.documentElement.clientHeight;
    } else if (document.body && (document.body.clientWidth || document.body.clientHeight)) {
        //IE 4 compatible
        myWidth = document.body.clientWidth;
        myHeight = document.body.clientHeight;
    }

    var w = myWidth * 1 / 2 ;
    
    moveFL(4, w, getStyleObjectID('adImg'))
}