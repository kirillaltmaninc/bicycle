function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_timelinePlay(tmLnName, myID) { //v1.2
  //Copyright 1998, 1999, 2000, 2001, 2002, 2003, 2004 Macromedia, Inc. All rights reserved.
  var i,j,tmLn,props,keyFrm,sprite,numKeyFr,firstKeyFr,propNum,theObj,firstTime=false;
  if (document.MM_Time == null) MM_initTimelines(); //if *very* 1st time
  tmLn = document.MM_Time[tmLnName];
  if (myID == null) { myID = ++tmLn.ID; firstTime=true;}//if new call, incr ID
  if (myID == tmLn.ID) { //if Im newest
    setTimeout('MM_timelinePlay("'+tmLnName+'",'+myID+')',tmLn.delay);
    fNew = ++tmLn.curFrame;
    for (i=0; i<tmLn.length; i++) {
      sprite = tmLn[i];
      if (sprite.charAt(0) == 's') {
        if (sprite.obj) {
          numKeyFr = sprite.keyFrames.length; firstKeyFr = sprite.keyFrames[0];
          if (fNew >= firstKeyFr && fNew <= sprite.keyFrames[numKeyFr-1]) {//in range
            keyFrm=1;
            for (j=0; j<sprite.values.length; j++) {
              props = sprite.values[j]; 
              if (numKeyFr != props.length) {
                if (props.prop2 == null) sprite.obj[props.prop] = props[fNew-firstKeyFr];
                else        sprite.obj[props.prop2][props.prop] = props[fNew-firstKeyFr];
              } else {
                while (keyFrm<numKeyFr && fNew>=sprite.keyFrames[keyFrm]) keyFrm++;
                if (firstTime || fNew==sprite.keyFrames[keyFrm-1]) {
                  if (props.prop2 == null) sprite.obj[props.prop] = props[keyFrm-1];
                  else        sprite.obj[props.prop2][props.prop] = props[keyFrm-1];
        } } } } }
      } else if (sprite.charAt(0)=='b' && fNew == sprite.frame) eval(sprite.value);
      if (fNew > tmLn.lastFrame) tmLn.ID = 0;
  } }
}

function MM_initTimelines() { //v4.0
    //MM_initTimelines() Copyright 1997 Macromedia, Inc. All rights reserved.
    var ns = navigator.appName == "Netscape";
    var ns4 = (ns && parseInt(navigator.appVersion) == 4);
    var ns5 = (ns && parseInt(navigator.appVersion) > 4);
    document.MM_Time = new Array(1);
    document.MM_Time[0] = new Array(1);
    document.MM_Time["Timeline1"] = document.MM_Time[0];
    document.MM_Time[0].MM_Name = "Timeline1";
    document.MM_Time[0].fps = 24;
    document.MM_Time[0][0] = new String("sprite");
    document.MM_Time[0][0].slot = 1;
    if (ns4)
        document.MM_Time[0][0].obj = document["animatedtext"];
    else if (ns5)
        document.MM_Time[0][0].obj = document.getElementById("animatedtext");
    else
        document.MM_Time[0][0].obj = document.all ? document.all["animatedtext"] : null;
    document.MM_Time[0][0].keyFrames = new Array(1, 55);
    document.MM_Time[0][0].values = new Array(2);
    if (ns5)
        document.MM_Time[0][0].values[0] = new Array("302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px", "302px");
    else
        document.MM_Time[0][0].values[0] = new Array(302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302,302);
    document.MM_Time[0][0].values[0].prop = "left";
    if (ns5)
        document.MM_Time[0][0].values[1] = new Array("-141px", "-137px", "-134px", "-130px", "-127px", "-123px", "-120px", "-116px", "-113px", "-109px", "-106px", "-102px", "-99px", "-95px", "-92px", "-88px", "-85px", "-81px", "-78px", "-74px", "-71px", "-67px", "-64px", "-60px", "-57px", "-53px", "-50px", "-46px", "-42px", "-39px", "-35px", "-32px", "-28px", "-25px", "-21px", "-18px", "-14px", "-11px", "-7px", "-4px", "0px", "3px", "7px", "10px", "14px", "17px", "21px", "24px", "28px", "31px", "35px", "38px", "42px", "45px", "49px");
    else
        document.MM_Time[0][0].values[1] = new Array(-141,-137,-134,-130,-127,-123,-120,-116,-113,-109,-106,-102,-99,-95,-92,-88,-85,-81,-78,-74,-71,-67,-64,-60,-57,-53,-50,-46,-42,-39,-35,-32,-28,-25,-21,-18,-14,-11,-7,-4,0,3,7,10,14,17,21,24,28,31,35,38,42,45,49);
    document.MM_Time[0][0].values[1].prop = "top";
    if (!ns4) {
        document.MM_Time[0][0].values[0].prop2 = "style";
        document.MM_Time[0][0].values[1].prop2 = "style";
    }
    document.MM_Time[0].lastFrame = 55;
    for (i=0; i<document.MM_Time.length; i++) {
        document.MM_Time[i].ID = null;
        document.MM_Time[i].curFrame = 0;
        document.MM_Time[i].delay = 1000/document.MM_Time[i].fps;
    }
}
MM_initTimelines(); 
