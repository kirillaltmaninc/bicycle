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
    var macIE5 = (navigator.platform ? (navigator.platform == "MacPPC") : false) && (navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4);
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
    document.MM_Time[0][0].keyFrames = new Array(1, 50);
    document.MM_Time[0][0].values = new Array(2);
    if (ns5 || macIE5)
        document.MM_Time[0][0].values[0] = new Array("20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px", "20px");
    else
        document.MM_Time[0][0].values[0] = new Array(20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20);
    document.MM_Time[0][0].values[0].prop = "left";
    if (ns5 || macIE5)
        document.MM_Time[0][0].values[1] = new Array("-400px", "-396px", "-392px", "-388px", "-384px", "-380px", "-376px", "-371px", "-367px", "-363px", "-359px", "-355px", "-351px", "-347px", "-343px", "-339px", "-335px", "-331px", "-327px", "-322px", "-318px", "-314px", "-310px", "-306px", "-302px", "-298px", "-294px", "-290px", "-286px", "-282px", "-278px", "-273px", "-269px", "-265px", "-261px", "-257px", "-253px", "-249px", "-245px", "-241px", "-237px", "-233px", "-229px", "-224px", "-220px", "-216px", "-212px", "-208px", "-204px", "-200px");
    else
        document.MM_Time[0][0].values[1] = new Array(-400,-396,-392,-388,-384,-380,-376,-371,-367,-363,-359,-355,-351,-347,-343,-339,-335,-331,-327,-322,-318,-314,-310,-306,-302,-298,-294,-290,-286,-282,-278,-273,-269,-265,-261,-257,-253,-249,-245,-241,-237,-233,-229,-224,-220,-216,-212,-208,-204,-200);
    document.MM_Time[0][0].values[1].prop = "top";
    if (!ns4) {
        document.MM_Time[0][0].values[0].prop2 = "style";
        document.MM_Time[0][0].values[1].prop2 = "style";
    }
    document.MM_Time[0].lastFrame = 50;
    for (i=0; i<document.MM_Time.length; i++) {
        document.MM_Time[i].ID = null;
        document.MM_Time[i].curFrame = 0;
        document.MM_Time[i].delay = 1000/document.MM_Time[i].fps;
    }
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
