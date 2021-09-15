var m_flag=1;			//¨º?¡¤???¨º?(1-??¨º?)

var message="?????¡Â3?¨ª???1¨¹¨¤¨ª?¦Ì¨ª3ENCHICMS ver 3.0.0¡ê???D¡ì¡ê?¨ª¨º?¨¤¦Ì?????¡ê?";

var delta=0.15
var collection;
function MM_reloadPage(init){//reloads the window if Nav4 resized
	if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
	document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
	else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_findObj(n, d){//v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
  d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_showHideLayers(){//v3.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
  if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
  obj.visibility=v; }
}
function floaters(){
	this.items=[];
	this.addItem=function(id,x,y,content){
		document.write('<DIV id='+id+' style="Z-INDEX: 10; POSITION: absolute;  width:80px; height:60px;left:'+(typeof(x)=='string'?eval(x):x)+';top:'+(typeof(y)=='string'?eval(y):y)+'">'+content+'</DIV>');
		var newItem				= {};
		newItem.object			= document.getElementById(id);
		newItem.x				= x;
		newItem.y				= y;
		this.items[this.items.length]		= newItem;
	}
	this.play=function(){
		collection=this.items
		setInterval('play()',10);
	}
}
function play(){
	for(var i=0;i<collection.length;i++){
		var followObj= collection[i].object;
		var followObj_x=(typeof(collection[i].x)=='string'?eval(collection[i].x):collection[i].x);
		var followObj_y=(typeof(collection[i].y)=='string'?eval(collection[i].y):collection[i].y);
		if(followObj.offsetLeft!=(document.body.scrollLeft+followObj_x)) {
			var dx=(document.body.scrollLeft+followObj_x-followObj.offsetLeft)*delta;
			dx=(dx>0?1:-1)*Math.ceil(Math.abs(dx));
			followObj.style.left=followObj.offsetLeft+dx;
		}
		if(followObj.offsetTop!=(document.body.scrollTop+followObj_y)) {
			var dy=(document.body.scrollTop+followObj_y-followObj.offsetTop)*delta;
			dy=(dy>0?1:-1)*Math.ceil(Math.abs(dy));
			followObj.style.top=followObj.offsetTop+dy;
		}
		followObj.style.display	= '';
	}
}	
		
if(m_flag==1){
	var theFloaters=new floaters();
	var lx=screen.width-400;
	theFloaters.addItem('followDiv1',lx,100,'<table width="260"  border="0" cellspacing="1" cellpadding="1" bgcolor="#93CBEC"><tr  bgcolor="#EEEEEE"><td><table width="100%"  border="0" cellspacing="0" cellpadding="0" bgcolor="ffffff"><tr class="text" bgcolor="#1DA4D0"><td width="8%">&nbsp;</td><td width="84%" height="22" align="center" style="color:#FFFFFF; font-size:12px; font-weight:bold">??¦Ì?1???</td><td width="8%" align="right"><label style="font-size:12px;cursor:hand;font-weight:bold;color: #FFFFFF;" onClick=MM_showHideLayers("followDiv1","","hide","followDiv2","","hide")>?¨¢</label></td></tr></table></td></tr><tr class=contbcon><td bgcolor="#FEFEFE"><table width="100%"  border="0" cellspacing="0" cellpadding="4"><tr><td><P style="font-size:12px; color:#000000;letter-spacing: 1pt;"><marquee direction=up scrollamount=1 height=130>'+message+'</marquee></P></td></tr></table></td></tr></table>');
	theFloaters.play();
}
