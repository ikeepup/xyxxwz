<%
'=====================================================================
' 软件名称：恩池网站管理系统
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Public Function UbbCode(strContent)
        Dim re
        Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
        '过滤危险脚本
        re.Pattern = "(</script>)"
        strContent = re.Replace(strContent, "&lt;/script&gt;")
        re.Pattern = "(script:)"
        strContent = re.Replace(strContent, "&#115 cript&#58")
	re.Pattern = "(script)"
        strContent = re.Replace(strContent, "&#115 cript")
        re.Pattern = "(js:)"
        strContent = re.Replace(strContent, "&#106s&#58")
        re.Pattern = "(value)"
        strContent = re.Replace(strContent, "&#118alue")
        re.Pattern = "(about:)"
        strContent = re.Replace(strContent, "about&#58")
        re.Pattern = "(file:)"
        strContent = re.Replace(strContent, "file&#58")
        re.Pattern = "(document.cookie)"
        strContent = re.Replace(strContent, "documents&#46cookie")
        re.Pattern = "(vbs:)"
        strContent = re.Replace(strContent, "&#118 bs&#58")
        re.Pattern = "(on(mouse|exit|error|click|key))"
        strContent = re.Replace(strContent, "&#111n$2")

	re.Pattern = "<IMG.[^>]*SRC(=| )(.[^>]*)>"
	'strContent = re.replace(strContent,"<IMG SRC=$2 onclick=""javascript:window.open(this.src);"" style=""CURSOR: pointer"" border=""0"">")
	strContent = re.replace(strContent,"<IMG SRC=$2 border=""0"">")
	
	're.Pattern="<img(.[^>]*)>"
	'strContent = re.replace(strContent,"<img$1 onload=""return imgzoom(this,550)"">")
	
	re.Pattern = "(\[i\])(.[^\[]*)(\[\/i\])"
        strContent = re.Replace(strContent, "<i>$2</i>")
        re.Pattern = "(\[u\])(.[^\[]*)(\[\/u\])"
        strContent = re.Replace(strContent, "<u>$2</u>")
        re.Pattern = "(\[b\])(.[^\[]*)(\[\/b\])"
        strContent = re.Replace(strContent, "<b>$2</b>")
        re.Pattern = "(\[fly\])(.*)(\[\/fly\])"
        strContent = re.Replace(strContent, "<marquee>$2</marquee>")

        re.Pattern = "\[size=([1-9])\](.[^\[]*)\[\/size\]"
        strContent = re.Replace(strContent, "<font size=$1>$2</font>")
        re.Pattern = "(\[center\])(.[^\[]*)(\[\/center\])"
        strContent = re.Replace(strContent, "<center>$2</center>")

        re.Pattern = "\[DIR=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/DIR]"
        strContent = re.Replace(strContent, "<embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed>")
        re.Pattern = "\[QT=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/QT]"
        strContent = re.Replace(strContent, "<embed src=$3 width=$1 height=$2 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#000000 kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>")
        re.Pattern = "\[MP=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/MP]"
        strContent = re.Replace(strContent, "<embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3  width=$1 height=$2></embed>")
        re.Pattern = "\[RM=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/RM]"
        strContent = re.Replace(strContent, "<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")

        re.Pattern = "(\[FLASH\])(.[^\[]*)(\[\/FLASH\])"
        'strContent = re.Replace(strContent, "<embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed>")
	strContent = re.Replace(strContent, "")
        re.Pattern = "(\[FLASH=*([0-9]*),*([0-9]*)\])(.[^\[]*)(\[\/FLASH\])"
        strContent = re.Replace(strContent, "<embed src=""$4"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed>")
	strContent = re.Replace(strContent, "")
        re.Pattern = "\[UPLOAD=(gif|jpg|jpeg|bmp|png)\](.[^\[]*)(gif|jpg|jpeg|bmp|png)\[\/UPLOAD\]"
        strContent = re.Replace(strContent, "<br><A HREF=""$2$1"" TARGET=_blank><IMG SRC=""$2$1"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")

        re.Pattern = "(\[UPLOAD=(.[^\[]*)\])(.[^\[]*)(\[\/UPLOAD\])"
        strContent = re.Replace(strContent, "<br><a href=""$3"">点击浏览该文件</a>")

        re.Pattern = "(\[URL\])(.[^\[]*)(\[\/URL\])"
        strContent = re.Replace(strContent, "<A HREF=""$2"" TARGET=_blank>$2</A>")
        re.Pattern = "(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
        strContent = re.Replace(strContent, "<A HREF=""$2"" TARGET=_blank>$3</A>")

        re.Pattern = "(\[EMAIL\])(.[^\[]*)(\[\/EMAIL\])"
        strContent = re.Replace(strContent, "<A HREF=""mailto:$2"">$2</A>")
        re.Pattern = "(\[EMAIL=(.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
        strContent = re.Replace(strContent, "<A HREF=""mailto:$2"" TARGET=_blank>$3</A>")

        re.Pattern = "(\[HTML\])(.[^\[]*)(\[\/HTML\])"
        strContent = re.Replace(strContent, "<table width='100%' border='0' cellspacing='0' cellpadding='6' bgcolor='#F6F6F6'><td><b>以下内容为程序代码:</b><br>$2</td></table>")
        re.Pattern = "(\[code\])(.[^\[]*)(\[\/code\])"
        strContent = re.Replace(strContent, "<table width='100%' border='0' cellspacing='0' cellpadding='6' bgcolor='#F6F6F6'><td><b>以下内容为程序代码:</b><br>$2</td></table>")

        re.Pattern = "(\[color=(.[^\[]*)\])(.[^\[]*)(\[\/color\])"
        strContent = re.Replace(strContent, "<font color=$2>$3</font>")
        re.Pattern = "(\[face=(.[^\[]*)\])(.[^\[]*)(\[\/face\])"
        strContent = re.Replace(strContent, "<font face=$2>$3</font>")
        re.Pattern = "\[align=(center|left|right)\](.*)\[\/align\]"
        strContent = re.Replace(strContent, "<div align=$1>$2</div>")

        re.Pattern = "(\[QUOTE\])(.*)(\[\/QUOTE\])"
        strContent = re.Replace(strContent, "<table cellpadding=0 cellspacing=0 border=1 WIDTH=94% bordercolor=#000000 bgcolor=#F2F8FF align=center  ><tr><td  ><table width=100% cellpadding=5 cellspacing=1 border=0><TR><TD BGCOLOR='#F6F6F6'>$2</table></table><br>")
        re.Pattern = "(\[move\])(.*)(\[\/move\])"
        strContent = re.Replace(strContent, "<MARQUEE scrollamount=3>$2</marquee>")
        re.Pattern = "\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
        strContent = re.Replace(strContent, "<table width=$1 style=""filter:glow(color=$2, strength=$3)"">$4</table>")
        re.Pattern = "\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
        strContent = re.Replace(strContent, "<table width=$1 style=""filter:shadow(color=$2, strength=$3)"">$4</table>")
        Set re = Nothing
	strContent = Replace(strContent, "{", "&#123;")
	strContent = Replace(strContent, "}", "&#125;")
	strContent = Replace(strContent, "$", "&#36;")
        UbbCode = strContent
End Function
%>
