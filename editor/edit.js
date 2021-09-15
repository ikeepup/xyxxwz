var strPath = "../editor/images/";
var strFilePath = "../editor/";
// 系统初试化 函数组开始
///////////////////////////////////////////////////////////////////////////////
  SEP_PADDING = 5;
  HANDLE_PADDING = 7; 
  
  var BrowserInfo = new Object() ;
  BrowserInfo.MajorVer = navigator.appVersion.match(/MSIE (.)/)[1] ;
  BrowserInfo.MinorVer = navigator.appVersion.match(/MSIE .\.(.)/)[1] ;
  BrowserInfo.IsIE55OrMore = BrowserInfo.MajorVer >= 6 || ( BrowserInfo.MajorVer >= 5 && BrowserInfo.MinorVer >= 5 ) ;

  var yToolbars =   new Array();
  var YInitialized = false;

document.writeln("<link href=\"" +strFilePath+ "editor.css\" type=\"text/css\" rel=\"stylesheet\">");
document.writeln("<table width=520 cellpadding=1 cellspacing=0 border=0 bgcolor=\"#E8E8E8\"><tr valign='top'><td colspan=2>");
document.writeln("<table width='100%' cellpadding=1 class='Toolbar' cellspacing=0 border=0><tr valign='top'><td><div class='yToolbar'><div class='TBHandle'></div><select ID=formatSelect CLASS=\"TBGen\" onChange=\"FormatText('FormatBlock',this[this.selectedIndex].value);\"><option selected>段落</option><option value=\"&lt;P&gt;\">正文</option><option value=\"&lt;H1&gt;\">标题一</option><option value=\"&lt;H2&gt;\">标题二</option><option value=\"&lt;H3&gt;\">标题三</option><option value=\"&lt;H4&gt;\">标题四</option><option value=\"&lt;H5&gt;\">标题五</option><option value=\"&lt;H6&gt;\">标题六</option><option value=\"&lt;PRE&gt;\">预设格式</option></select> <select name=\"selectFont\" CLASS=\"TBGen\" onChange=\"FormatText('fontname', selectFont.options[selectFont.selectedIndex].value);\"><option selected>字体<option value=\"removeFormat\">默认字体<option value=\"宋体\">宋体<option value=\"黑体\">黑体<option value=\"隶书\">隶书<option value=\"幼圆\">幼圆<option value=\"楷体_GB2312\">楷体<option value=\"仿宋_GB2312\">仿宋<option value=\"新宋体\">新宋体<option value=\"华文彩云\">华文彩云<option value=\"华文仿宋\">华文仿宋<option value=\"华文新魏\">华文新魏<option value=\"Arial\">Arial<option value=\"Arial Black\">Arial Black<option value=\"Arial Narrow\">Arial Narrow<option value=\"Century\">Century<option value=\"Courier New\">Courier New<option value=\"Georgia\">Georgia<option value=\"Impact\">Impact<option value=\"Lucida Console\">Lucida Console<option value=\"MS Sans Serif\">MS Sans Serif<option value=\"System\">System<option value=\"Symbol\">Symbol<option value=\"Tahoma\">Tahoma<option value=\"Verdana\">Verdana<option value=\"Webdings\">Webdings<option value=\"Wingdings\">Wingdings</option></select> <select CLASS=\"TBGen\" onChange=\"FormatText('fontsize',this[this.selectedIndex].value);\" name=\"D2\"><option class=\"heading\" selected>字体大小<option value=1>一号<option value=2>二号<option value=3>三号<option value=4>四号<option value=5>五号<option value=6>六号<option value=7>七号</option></select>");
document.writeln("<DIV CLASS=\"Btn\" title=\"突出颜色\" onClick=BackColor()><img class='Ico' src=" +strPath+ "bgcolor.gif></DIV>");
document.writeln("<DIV CLASS=\"Btn\" title=\"字体颜色\" onClick=foreColor()><img class='Ico' src=" +strPath+ "forecolor.gif></DIV>");
document.writeln("<DIV CLASS=\"Btn\" title=\"替换\" onClick=replace()><img class='Ico' src=" +strPath+ "findreplace.gif></DIV>");
document.writeln("<DIV CLASS=\"Btn\" title=\"插入图片\" onClick=img()><img class='Ico' src=" +strPath+ "img.gif></DIV>");
document.writeln("<DIV CLASS=\"Btn\" title=\"插入FLASH文件\" onclick=flash()><img class='Ico' src=" +strPath+ "Flash.gif></DIV>");
document.writeln("<DIV CLASS=\"Btn\" title=\"插入表格\" onclick=fortable()><img class='Ico' src=" +strPath+ "TableInsert.gif></DIV>");
document.writeln("<DIV CLASS=\"Btn\" title=\"引用样式\" onclick=WBTB_quote()><img class='Ico' src=" +strPath+ "quote.gif></DIV>");
var FormatTextlist="插入超链接 createLink|去掉超链接 Unlink|粗体 bold|倾斜 italic|下划线 underline|<br>|剪切 cut|复制 copy|粘贴 paste|撤消 undo|恢复 redo|全选 selectAll|取消选择 unselect|上标 superscript|下标 subscript|删除线 strikethrough|删除文字格式 RemoveFormat|左对齐 Justifyleft|居中 JustifyCenter|右对齐 JustifyRight|两端对齐 justifyfull|编号 insertorderedlist|项目符号 InsertUnorderedList|减少缩进量 Outdent|增加缩进量 indent|普通水平线 InsertHorizontalRule|删除当前选中区 Delete"
var list= FormatTextlist.split ('|'); 
for(i=0;i<list.length;i++) {
if (list[i]=="<br>"){document.write("</div></td></tr><tr><td><div class=yToolbar><div class='TBHandle'></div>");
}else{
var TextName= list[i].split (' '); 
document.write("<DIV CLASS=\"Btn\" title="+TextName[0]+" onClick=FormatText('"+TextName[1]+"')><img class='Ico' border=0 src=" +strPath+ ""+TextName[1]+".gif></DIV> ");
}
}
document.writeln("<div class='TBSep'></div><DIV CLASS=\"Btn\" title=\"清理代码\" onclick=CleanCode()><img class='Ico' src=" +strPath+ "CleanCode.gif></DIV>");
document.writeln("<div class='TBSep'></div><input id=EditMode class=\"TBGen\" title=\"HTML代码\" onclick=\"setMode(this.checked);\" type=checkbox value=\"ON\">"); ///HTML代码
document.writeln("</div></td></tr></table>");
document.writeln("<iframe class=Composition ID=Composition MARGINHEIGHT=5 MARGINWIDTH=5 width='100%' height='200' scrolling='yes'></iframe>");
//document.writeln("</td></tr><tr height=22 bgcolor=\"#F8F8F8\"><td></td><td>");
document.writeln("</td></tr></table>");

if (document.all){var IframeID=frames["Composition"];}else{var IframeID=document.getElementById("Composition").contentWindow;}
if (navigator.appVersion.indexOf("MSIE 6.0",0)==-1){IframeID.document.designMode="On"}
IframeID.document.open();
IframeID.document.write ('<script>i=0;function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13 && i==0){i=1;parent.document.myform.content.value=document.body.innerHTML;parent.document.myform.submit();parent.document.myform.Submit1.disabled=true;}}<\/script><style type=text/css>.quote{margin:5px 20px;border:1px solid #CCCCCC;padding:5px; background:#F3F3F3 }\nbody{boder:0px}.HtmlCode{margin:5px 20px;border:1px solid #CCCCCC;padding:5px;background:#FDFDDF;font-size:14px;font-family:Tahoma;font-style : oblique;line-height : normal ;font-weight:bold;}\nbody{boder:0px}</style><link href=\"" +strFilePath+ "EditorArea.css\" type=\"text/css\" rel=\"stylesheet\"><body onkeydown=ctlent()>');
IframeID.document.close();
IframeID.document.body.contentEditable = "True";
IframeID.document.body.innerHTML=document.getElementById("content").value;
IframeID.document.body.style.fontSize="10pt";

var bTextMode=false

//程序处始化
function document.onreadystatechange(){
  if (YInitialized) return;
  YInitialized = true;

  var i, s, curr;
  for (i=0; i<document.body.all.length; i++)
  {
    curr=document.body.all[i];
    if (curr.className == 'yToolbar')
    {
      InitTB(curr);
      yToolbars[yToolbars.length] = curr;
    }
  }
  
}

function InitBtn(btn)
{
    btn.onmouseover = BtnMouseOver;
    btn.onmouseout = BtnMouseOut;
    btn.onmousedown = BtnMouseDown;
    btn.onmouseup = BtnMouseUp;
    btn.ondragstart = YCancelEvent;
    btn.onselectstart = YCancelEvent;
    btn.onselect = YCancelEvent;
    btn.YUSERONCLICK = btn.onclick;
    btn.onclick = YCancelEvent;
    btn.YINITIALIZED = true;
    return true;
}

function InitTB(y)
{
    y.TBWidth = 0;

    if (! PopulateTB(y)) return false;

    y.style.posWidth = y.TBWidth;

    return true;
}

function YCancelEvent()
{
    event.returnValue=false;
    event.cancelBubble=true;
    return false;
}

function PopulateTB(y)
{
    var i, elements, element;

    elements = y.children;
    for (i=0; i<elements.length; i++) {
        element = elements[i];
        if (element.tagName == 'SCRIPT' || element.tagName == '!') continue;

        switch (element.className) {
            case 'Btn':
                if (element.YINITIALIZED == null)   {
                if (! InitBtn(element))
                    return false;
                }
                element.style.posLeft = y.TBWidth;
                y.TBWidth   += element.offsetWidth + 1;
            break;
      
            case 'BtnMenu':
                if (element.YINITIALIZED == null)   {
                if (! InitBtnMenu(element))
                    return false;
                }
                element.style.posLeft = y.TBWidth;
                y.TBWidth   += element.offsetWidth + 1;
                break;

            case 'TBGen':
                element.style.posLeft = y.TBWidth;
                y.TBWidth   += element.offsetWidth + 1;
                break;

            case 'TBSep':
                element.style.posLeft = y.TBWidth   + 2;
                y.TBWidth   += SEP_PADDING;
                break;

            case 'TBHandle':
                element.style.posLeft = 2;
                y.TBWidth   += element.offsetWidth + HANDLE_PADDING;
                break;

            default:
                return false;
        }
    }

    y.TBWidth += 1;
    return true;
}

function TemplateTBs()
{
    NumTBs = yToolbars.length;

    if (NumTBs == 0) return;

    var i;
    var ScrWid = (document.body.offsetWidth) - 6;
    var TotalLen = ScrWid;
    for (i = 0 ; i < NumTBs ; i++) {
        TB = yToolbars[i];
        if (TB.TBWidth > TotalLen) TotalLen = TB.TBWidth;
        }

    var PrevTB;
    var LastStart = 0;
    var RelTop = 0;
    var LastWid, CurrWid;
    var TB = yToolbars[0];
    TB.style.posTop = 0;
    TB.style.posLeft = 0;

    var Start = TB.TBWidth;
    for (i = 1 ; i < yToolbars.length ; i++) {
        PrevTB = TB;
        TB = yToolbars[i];
        CurrWid = TB.TBWidth;

    if ((Start + CurrWid) > ScrWid) {
        Start = 0;
        LastWid = TotalLen - LastStart;
    }
    else {
        LastWid =    PrevTB.TBWidth;
        RelTop -=    TB.offsetHeight;
    }

    TB.style.posTop = RelTop;
    TB.style.posLeft = Start;
    PrevTB.style.width = LastWid;

    LastStart = Start;
    Start += CurrWid;
  }

  TB.style.width = TotalLen - LastStart;

  i--;
  TB = yToolbars[i];
  var TBInd = TB.sourceIndex;
  var A = TB.document.all;
  var item;
  for (i in A) {
    item = A.item(i);
    if (! item) continue;
    if (! item.style) continue;
    if (item.sourceIndex <= TBInd) continue;
    if (item.style.position == 'absolute') continue;
    item.style.posTop = RelTop;
  }
}

function DoTemplate()
{
  TemplateTBs();
}

function BtnMouseOver()
{
  if (event.srcElement.tagName != 'IMG') return false;
  var image = event.srcElement;
  var element = image.parentElement;

  if (image.className == 'Ico') element.className = 'BtnMouseOverUp';
  else if (image.className == 'IcoDown') element.className = 'BtnMouseOverDown';

  event.cancelBubble = true;
}

function BtnMouseOut()
{
  if (event.srcElement.tagName != 'IMG') {
    event.cancelBubble = true;
    return false;
  }

  var image = event.srcElement;
  var element = image.parentElement;
  yRaisedElement = null;

  element.className = 'Btn';
  image.className = 'Ico';

  event.cancelBubble = true;
}

function BtnMouseDown()
{
  if (event.srcElement.tagName != 'IMG') {
    event.cancelBubble = true;
    event.returnValue=false;
    return false;
  }

  var image = event.srcElement;
  var element = image.parentElement;

  element.className = 'BtnMouseOverDown';
  image.className = 'IcoDown';

  event.cancelBubble = true;
  event.returnValue=false;
  return false;
}

function BtnMouseUp()
{
  if (event.srcElement.tagName != 'IMG') {
    event.cancelBubble = true;
    return false;
  }

  var image = event.srcElement;
  var element = image.parentElement;


if(navigator.appVersion.match(/8./i)=='8.')
    {
      if (element.YUSERONCLICK) eval(element.YUSERONCLICK + 'onclick(event)');  
   }
else

   {
     if (element.YUSERONCLICK) eval(element.YUSERONCLICK + 'anonymous()');
}


  element.className = 'BtnMouseOverUp';
  image.className = 'Ico';

  event.cancelBubble = true;
  return false;
}
// 系统初试化 函数组结速
///////////////////////////////////////////////////////////////////////////////

function validateMode()
{
  if (!	bTextMode) return true;
  alert("请取消“HTML 语法书写”选项再使用系统编辑功能!");
  IframeID.focus();
  return false;
}

function validateSubmit()
{
  if (!	bTextMode) return true;
  alert("HTML状态下不能提交数据，请取消“HTML 语法”选项！");
  IframeID.focus();
  return false;
}

function em(){
	if (!	validateMode())	return;
	var arr = showModalDialog("" +strFilePath+ "Emotion.htm", "", "dialogWidth:20em; dialogHeight:9.5em; status:0;help:0");
	if (arr != null){
	IframeID.focus()
	sel=IframeID.document.selection.createRange();
	sel.pasteHTML(arr);
	}
}

function CleanCode(){
	if (!	validateMode())	return;
	var body = IframeID.document.body;
	var html = IframeID.document.body.innerHTML;
	html = html.replace(/\<p>/gi,"[$p]");
	html = html.replace(/\<\/p>/gi,"[$\/p]");
	html = html.replace(/\<br>/gi,"[$br]");
	html = html.replace(/\<[^>]*>/g,"");
	html = html.replace(/\[\$p\]/gi,"<p>");
	html = html.replace(/\[\$\/p\]/gi,"<\/p>");
	html = html.replace(/\[\$br\]/gi,"<br>");
	IframeID.document.body.innerHTML = html;
}


function FormatText(command,option){
	if (!	validateMode())	return;
	IframeID.focus();IframeID.document.execCommand(command,true,option);
}

function CheckLength(){
	alert("\n您的内容已有 "+IframeID.document.body.innerHTML.length+" 字节");
}

function emoticon(theSmilie){
	if (!	validateMode())	return;
	IframeID.focus();
	sel=IframeID.document.selection.createRange();
	sel.pasteHTML("<img src=../images/emotion/"+theSmilie+".gif>");
}

function DoTitle(addTitle) {
var revisedTitle;var currentTitle = document.myform.topic.value;revisedTitle = addTitle+currentTitle;document.myform.topic.value=revisedTitle;document.myform.topic.focus();
}

function Gopreview()
{
	if (!	validateMode())	return;
	document.preview.content.value=IframeID.document.body.innerHTML; 
	window.open('', 'preview_page', 'resizable,scrollbars,width=750,height=450');
	document.preview.submit()
}

function BackColor()
{
	if (!	validateMode())	return;
	var arr = showModalDialog("" +strFilePath+ "selcolor.htm", "", "dialogWidth:18em; dialogHeight:17.5em; status:0;help:0");
	if (arr != null) FormatText('BackColor', arr);
	else IframeID.focus();
}

function foreColor()
{
	if (!	validateMode())	return;
	var arr = showModalDialog("" +strFilePath+ "selcolor.htm", "", "dialogWidth:18em; dialogHeight:17.5em; status:0;help:0");
	if (arr != null) FormatText('forecolor', arr);
	else IframeID.focus();
}

//////替换内容
function replace()
{
	if (!	validateMode())	return;
  var arr = showModalDialog("" +strFilePath+ "replace.htm", "", "dialogWidth:22em;dialogHeight:10em;status:0;help:0");
	if (arr != null){
		var ss;
		ss = arr.split("*")
		a = ss[0];
		b = ss[1];
		i = ss[2];
		con = IframeID.document.body.innerHTML;
		if (i == 1)
		{
			con = newasp_rCode(con,a,b,true);
		}else{
			con = newasp_rCode(con,a,b);
		}
		IframeID.document.body.innerHTML = con;
	}
	else IframeID.focus();
}
function newasp_rCode(s,a,b,i){
	a = a.replace("?","\\?");
	if (i==null)
	{
		var r = new RegExp(a,"gi");
	}else if (i) {
		var r = new RegExp(a,"g");
	}
	else{
		var r = new RegExp(a,"gi");
	}
	return s.replace(r,b); 
}
//////替换内容结束

function img(){
	if (!	validateMode())	return;
	url=prompt("请输入图片文件地址:","http://");
	if(!url || url=="http://") return;
	IframeID.focus();
	sel=IframeID.document.selection.createRange();
	sel.pasteHTML("<img border=0 src="+url+">");
}

function RealPlay(){
	if (!	validateMode())	return;
	var arr = showModalDialog("" +strFilePath+ "RealPlay.htm", "", "dialogWidth:22em; dialogHeight:10.5em; status:0;help:0");
	if (arr != null){
	IframeID.focus()
	sel=IframeID.document.selection.createRange();
	sel.pasteHTML(arr);
	}
}

function MediaPlayer(){
	if (!	validateMode())	return;
	var arr = showModalDialog("" +strFilePath+ "MediaPlayer.htm", "", "dialogWidth:22em; dialogHeight:10.5em; status:0;help:0");
	if (arr != null){
	IframeID.focus()
	sel=IframeID.document.selection.createRange();
	sel.pasteHTML(arr);
	}
}

function flash(){
	if (!	validateMode())	return;
	var arr = showModalDialog("" +strFilePath+ "flash.htm", "", "dialogWidth:22em; dialogHeight:9em; status:0;help:0");
	if (arr != null){
	//IframeID.focus()
	sel=IframeID.document.selection.createRange();
	sel.pasteHTML(arr);
	}
}

function WBTB_quote()
{
	WBTB_specialtype("<div style=\"margin:5px; padding:5px; border: 1px Dotted #CCCCCC; TABLE-LAYOUT: fixed; background:#F8F8F8\"><font style=\"color: #990000;font-weight:bold\">以下是引用片段：</font><br>","</div>");	

}
function WBTB_code()
{
	WBTB_specialtype("<div style=\"margin:5px;border:1px solid #CCCCCC;padding:5px;background:#FDFDDF;font-size:14px;font-family:Tahoma;font-style : oblique;line-height : normal ;font-weight:bold;\"><font style=\"color: #990000;font-weight:bold\">以下是代码片段：</font><br>","</div>");	

}
var WBTB_bIsIE5 = (navigator.userAgent.indexOf("IE 5")  > -1) || (navigator.userAgent.indexOf("IE 6")  > -1);
var WBTB_edit;	//selectRang
var WBTB_RangeType;
var WBTB_selection;

//应用html
function WBTB_specialtype(Mark1, Mark2){
	var strHTML;
	if (!	validateMode())	return;
	if (WBTB_bIsIE5) WBTB_selectRange();
	if (WBTB_RangeType == "Text"){
		if (Mark2==null)
		{
			strHTML = "<" + Mark1 + ">" + WBTB_edit.htmlText + "</" + Mark1 + ">"; 
		}else{
			strHTML = Mark1 + WBTB_edit.htmlText +  Mark2; 
		}
		WBTB_edit.pasteHTML(strHTML);
		Composition.focus();
		WBTB_edit.select();
	}else{window.alert("请选择相应内容！")}		
}
//选择内容替换文本
function WBTB_InsertSymbol(str1)
{
	Composition.focus();
	if (WBTB_bIsIE5) WBTB_selectRange();	
	WBTB_edit.pasteHTML(str1);
}


function WBTB_selectRange(){
	WBTB_selection = IframeID.document.selection;
	WBTB_edit = IframeID.document.selection.createRange();
	WBTB_RangeType =  IframeID.document.selection.type;
}

//////设置编辑器模式
function setMode(newMode)
{
  bTextMode = newMode;
  var content;
  if (bTextMode) {
    cleanHtml();

    content=IframeID.document.body.innerHTML;
    IframeID.document.body.innerText=content;
  } else {
    content=IframeID.document.body.innerText;
    IframeID.document.body.innerHTML=content;
  }

  Composition.focus();
}

function cleanHtml()
{
  var fonts = IframeID.document.body.all.tags("FONT");
  var curr;
  for (var i = fonts.length - 1; i >= 0; i--) {
    curr = fonts[i];
    if (curr.style.backgroundColor == "#ffffff") curr.outerHTML	= curr.innerHTML;
  }
}
//插入表格
function fortable()
{
	if (!	validateMode())	return;
	IframeID.focus();
	var arr = showModalDialog("" +strFilePath+ "table.htm", window, "dialogWidth:22em; dialogHeight:19em; status:0; help:0;scroll:no;");
	if (arr)
	{
		IframeID.document.body.innerHTML+=arr;
	}
	IframeID.focus();
}

// 改变编辑区高度
function sizeChange(size){
	//if (!BrowserInfo.IsIE55OrMore){
		//alert("此功能需要IE5.5版本以上的支持！");
		//return false;
	//}
	for (var i=0; i<parent.frames.length; i++){
		if (parent.frames[i].document==self.document){
			var obj=parent.frames[i].frameElement;
			var height = parseInt(obj.offsetHeight);
			if (height+size>=100){
				obj.height=height+size;
			}
			break;
		}
	}
}

//-------------------
function Newasp_formatimg()
{
	if (BrowserInfo.IsIE55OrMore){
		var tmp=IframeID.document.body.all.tags("IMG");
	}else{
		var tmp=IframeID.document.getElementsByTagName("IMG");
	}
	for(var i=0;i<tmp.length;i++){
		var tempstr='';
		if(tmp[i].align!=''){tempstr=" align="+tmp[i].align;}
		if(tmp[i].border!=''){tempstr=tempstr+" border="+tmp[i].border;}
		tmp[i].outerHTML="<IMG src=\""+tmp[i].src+"\""+tempstr+">"
	}
}

//清理多余HTML代码
function Newasp_cleanHtml(content)
{
	content = content.replace(/<p>&nbsp;<\/p>/gi,"")
	content = content.replace(/<p><\/p>/gi,"<p>")
	content = content.replace(/<div><\/\1>/gi,"")
	content = content.replace(/<p>/,"<br>")
	content = content.replace(/(<(meta|iframe|frame|span|tbody|layer)[^>]*>|<\/(iframe|frame|meta|span|tbody|layer)>)/gi, "");
	content = content.replace(/<\\?\?xml[^>]*>/gi, "") ;
	content = content.replace(/o:/gi, "");
return content;
}
// 取编辑器的内容
function getHTML() {
	var html;
	Newasp_formatimg();
	html = IframeID.document.body.innerHTML;
	html = Newasp_cleanHtml(html);
	return html;
}