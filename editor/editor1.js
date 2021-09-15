var strPath = "../editor/images/";
var strFilePath = "../editor/";
//var _maxCount = '25000';
var bTextMode=false
if (_maxCount == 0 || _maxCount == ''){
	_maxCount = 64000;
}
 
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
document.writeln("<table width='100%' cellpadding=1 class='Toolbar' cellspacing=0 border=0><tr valign='top'><td><div class='yToolbar'><div class='TBHandle'></div>");
document.writeln("<select name=\"selectFont\" CLASS=\"TBGen\" onChange=\"FormatText('fontname', selectFont.options[selectFont.selectedIndex].value);\"><option selected>字体<option value=\"removeFormat\">默认字体<option value=\"宋体\">宋体<option value=\"黑体\">黑体<option value=\"隶书\">隶书<option value=\"幼圆\">幼圆<option value=\"楷体_GB2312\">楷体<option value=\"仿宋_GB2312\">仿宋<option value=\"新宋体\">新宋体<option value=\"华文彩云\">华文彩云<option value=\"华文仿宋\">华文仿宋<option value=\"华文新魏\">华文新魏<option value=\"Arial\">Arial<option value=\"Arial Black\">Arial Black<option value=\"Arial Narrow\">Arial Narrow<option value=\"Century\">Century<option value=\"Courier New\">Courier New<option value=\"Georgia\">Georgia<option value=\"Impact\">Impact<option value=\"Lucida Console\">Lucida Console<option value=\"MS Sans Serif\">MS Sans Serif<option value=\"System\">System<option value=\"Symbol\">Symbol<option value=\"Tahoma\">Tahoma<option value=\"Verdana\">Verdana<option value=\"Webdings\">Webdings<option value=\"Wingdings\">Wingdings</option></select>");
document.writeln("<select CLASS=\"TBGen\" onChange=\"FormatText('fontsize',this[this.selectedIndex].value);\" name=\"D2\"><option class=\"heading\" selected>字体大小<option value=1>一号<option value=2>二号<option value=3>三号<option value=4>四号<option value=5>五号<option value=6>六号<option value=7>七号</option></select>");
document.writeln("<div class='TBSep'></div>");
var FormatTextlist="粗体 bold|倾斜 italic|下划线 underline|左对齐 Justifyleft|居中 JustifyCenter|<br>|撤消 undo|恢复 redo|全选 selectAll|删除文字格式 RemoveFormat"
var list= FormatTextlist.split ('|'); 
for(i=0;i<list.length;i++) {
if (list[i]=="<br>"){document.write("<DIV CLASS=\"Btn\" title=\"字体颜色\" onClick=foreColor()><img class='Ico' src=" +strPath+ "forecolor.gif></DIV>");
document.writeln("<DIV CLASS=\"Btn\" title=\"更多字体\" onClick=showToolMenu('font')><img class='Ico' src=" +strPath+ "fontmenu.gif></DIV><div class='TBSep'></div>");
}else{
var TextName= list[i].split (' '); 
document.write("<DIV CLASS=\"Btn\" title="+TextName[0]+" onClick=FormatText('"+TextName[1]+"')><img class='Ico' border=0 src=" +strPath+ ""+TextName[1]+".gif></DIV> ");
}
}
document.writeln("<DIV CLASS=\"Btn\" title=\"更多文本格式\" onclick=\"showToolMenu('edit')\"><img class='Ico' src=" +strPath+ "FormMenu.gif></DIV><div class='TBSep'></div>");
document.writeln("<DIV CLASS=\"Btn\" title=\"替换\" onClick=replace()><img class='Ico' src=" +strPath+ "findreplace.gif></DIV>");
document.writeln("<DIV CLASS=\"Btn\" title=\"清理代码\" onclick=CleanCode()><img class='Ico' src=" +strPath+ "CleanCode.gif></DIV>");
document.writeln("</div></td></tr></table>");
document.writeln("<iframe class=Composition ID=Composition MARGINHEIGHT=5 MARGINWIDTH=5 width='100%' height='200' scrolling='yes'></iframe>");
document.writeln("</td></tr><tr><td>");
document.writeln("<table width='100%' height=20 cellpadding=1 cellspacing=0 border=0 class=StatusBar><tr><td align=center>");
document.writeln("<span style=\"color:#606060;font-size:12px;filter: dropshadow(color=#ffffff,offx=-1,offy=1,positive=1);width: 100%; line-height: 20px\" id=wordCount></span>");
document.writeln("</td></tr></table>");
document.writeln("</td></tr></table>");

if (document.all){var IframeID=frames["Composition"];}else{var IframeID=document.getElementById("Composition").contentWindow;}
if (navigator.appVersion.indexOf("MSIE 6.0",0)==-1){IframeID.document.designMode="On"}
IframeID.document.open();
IframeID.document.write ('<script>i=0;function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13 && i==0){i=1;parent.document.myform.content1.value=document.body.innerHTML;parent.document.myform.submit();parent.document.myform.Submit1.disabled=true;}}<\/script><style type=text/css>.quote{margin:5px 20px;border:1px solid #CCCCCC;padding:5px; background:#F3F3F3 }\nbody{boder:0px}.HtmlCode{margin:5px 20px;border:1px solid #CCCCCC;padding:5px;background:#FDFDDF;font-size:14px;font-family:Tahoma;font-style : oblique;line-height : normal ;font-weight:bold;}\nbody{boder:0px}</style><link href=\"" +strFilePath+ "EditorArea.css\" type=\"text/css\" rel=\"stylesheet\"><body onkeydown=ctlent()>');
IframeID.document.close();
calcWordCount();
IframeID.document.body.contentEditable = "True";
IframeID.document.body.innerHTML=document.getElementById("content1").value;
IframeID.document.body.style.fontSize="10pt";

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

  //////if (element.YUSERONCLICK) eval(element.YUSERONCLICK + 'anonymous()');
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

// 替换特殊字符
function HTMLEncode(text){
	text = text.replace(/&/g, "&amp;") ;
	text = text.replace(/"/g, "&quot;") ;
	text = text.replace(/</g, "&lt;") ;
	text = text.replace(/>/g, "&gt;") ;
	text = text.replace(/'/g, "&#146;") ;
	text = text.replace(/\ /g,"&nbsp;");
	text = text.replace(/\n/g,"<br>");
	text = text.replace(/\t/g,"&nbsp;&nbsp;&nbsp;&nbsp;");
	return text;
}
function emot(){
	if (!	validateMode())	return;
	var arr = showModalDialog("" +strFilePath+ "Emotion.htm", "", "dialogWidth:20em; dialogHeight:9.5em; status:0;help:0");
	if (arr != null){
	IframeID.focus()
	sel=IframeID.document.selection.createRange();
	sel.pasteHTML(arr);
	}
}
function FormatText(command,option){
	if (!	validateMode())	return;
	IframeID.focus();IframeID.document.execCommand(command,true,option);
}
//内容长度
function CheckLength(){
	alert("\n您的内容已有 "+IframeID.document.body.innerHTML.length+" 字节");
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

// 系统下拉菜单开始
///////////////////////////////////////////////////////////////////////////////
// 菜单常量
var sMenuHr="<tr><td align=center valign=middle height=2><TABLE border=0 cellpadding=0 cellspacing=0 width=128 height=2><tr><td height=1 class=HrShadow><\/td><\/tr><tr><td height=1 class=HrHighLight><\/td><\/tr><\/TABLE><\/td><\/tr>";
var sMenu1="<TABLE border=0 cellpadding=0 cellspacing=0 class=Menu width=150><tr><td width=18 valign=bottom align=center style='background:url(sysimage/contextmenu.gif);background-position:bottom;'><\/td><td width=132 class=RightBg><TABLE border=0 cellpadding=0 cellspacing=0>";
var sMenu2="<\/TABLE><\/td><\/tr><\/TABLE>";
var StyleMenuHeader = "<head><link href=\""+strFilePath+"MenuArea.css\" type=\"text/css\" rel=\"stylesheet\"></head><body scroll=\"no\" onConTextMenu=\"event.returnValue=false;\">";
// 菜单
var oPopupMenu = null;
if (BrowserInfo.IsIE55OrMore){
	oPopupMenu = window.createPopup();
}

// 取菜单行
function getMenuRow(s_Disabled, s_Event, s_Image, s_Html) {
	var s_MenuRow = "";
	s_MenuRow = "<tr><td align=center valign=middle><TABLE border=0 cellpadding=0 cellspacing=0 width=132><tr "+s_Disabled+"><td valign=middle height=20 class=MouseOut onMouseOver=this.className='MouseOver'; onMouseOut=this.className='MouseOut';";
	if (s_Disabled==""){
		s_MenuRow += " onclick=\"parent."+s_Event+";parent.oPopupMenu.hide();\"";
	}
	s_MenuRow += ">"
	if (s_Image !=""){
		s_MenuRow += "&nbsp;<img border=0 src='"+strPath+"/"+s_Image+"' width=20 height=20 align=absmiddle "+s_Disabled+">&nbsp;";
	}else{
		s_MenuRow += "&nbsp;";
	}
	s_MenuRow += s_Html+"<\/td><\/tr><\/TABLE><\/td><\/tr>";
	return s_MenuRow;

}

// 取标准的format菜单行
function getFormatMenuRow(menu, html, image){
	var s_Disabled = "";
	if (!IframeID.document.queryCommandEnabled(menu)){
		s_Disabled = "disabled";
	}
	var s_Event = "FormatText('"+menu+"')";
	var s_Image = menu+".gif";
	if (image){
		s_Image = image;
	}
	return getMenuRow(s_Disabled, s_Event, s_Image, html)
}

// 工具栏菜单
function showToolMenu(menu){
	if (!	validateMode())	return;
	var sMenu = ""
	var width = 150;
	var height = 0;

	var lefter = event.clientX;
	var leftoff = event.offsetX
	var topper = event.clientY;
	var topoff = event.offsetY;

	var oPopDocument = oPopupMenu.document;
	var oPopBody = oPopupMenu.document.body;

	switch(menu){
	case "font":		// 字体菜单
		sMenu += getMenuRow("", "foreColor()", "forecolor.gif", "字体颜色");
		sMenu += getMenuRow("", "BackColor()", "bgcolor.gif", "字体背景颜色");
		sMenu += sMenuHr;
		sMenu += getFormatMenuRow("bold", "粗体", "bold.gif");
		sMenu += getFormatMenuRow("italic", "斜体", "italic.gif");
		sMenu += getFormatMenuRow("underline", "下划线", "underline.gif");
		sMenu += getFormatMenuRow("strikethrough", "中划线", "strikethrough.gif");
		sMenu += sMenuHr;
		sMenu += getFormatMenuRow("superscript", "上标", "superscript.gif");
		sMenu += getFormatMenuRow("subscript", "下标", "subscript.gif");
		sMenu += sMenuHr;
		sMenu += getFormatMenuRow("JustifyLeft", "左对齐", "JustifyLeft.gif");
		sMenu += getFormatMenuRow("JustifyCenter", "居中对齐", "JustifyCenter.gif");
		sMenu += getFormatMenuRow("JustifyRight", "右对齐", "JustifyRight.gif");
		sMenu += getFormatMenuRow("JustifyFull", "两端对齐", "JustifyFull.gif");
		sMenu += sMenuHr;
		sMenu += getFormatMenuRow("createLink", "插入超链接", "createLink.gif");
		sMenu += getFormatMenuRow("Unlink", "去掉超链接", "Unlink.gif");
		height = 288;
		break;
	case "edit":		// 编辑菜单
		sMenu += getFormatMenuRow("Cut", "剪切", "cut.gif");
		sMenu += getFormatMenuRow("Copy", "复制", "copy.gif");
		sMenu += getFormatMenuRow("Paste", "粘贴", "paste.gif");
		sMenu += sMenuHr;
		sMenu += getFormatMenuRow("delete", "删除", "delete.gif");
		sMenu += getFormatMenuRow("RemoveFormat", "删除文字格式", "removeformat.gif");
		sMenu += sMenuHr;
		sMenu += getFormatMenuRow("SelectAll", "全部选中", "selectall.gif");
		sMenu += getFormatMenuRow("Unselect", "取消选择", "unselect.gif");
		sMenu += sMenuHr;
		sMenu += getFormatMenuRow("insertorderedlist", "编号", "insertorderedlist.gif");
		sMenu += getFormatMenuRow("insertunorderedlist", "项目符号", "insertunorderedlist.gif");
		sMenu += getFormatMenuRow("indent", "增加缩进量", "indent.gif");
		sMenu += getFormatMenuRow("outdent", "减少缩进量", "outdent.gif");
		sMenu += getFormatMenuRow("insertparagraph", "插入段落", "insertparagraph.gif");
		sMenu += sMenuHr;
		sMenu += getMenuRow("", "replace()", "findreplace.gif", "查找替换");
		sMenu += getMenuRow("", "emot()", "emot.gif", "插入表情图标");
		height = 288;
		break;
	}
	
	sMenu = sMenu1 + sMenu + sMenu2;
	
	oPopDocument.open();
	oPopDocument.write(StyleMenuHeader+sMenu);
	oPopDocument.close();

	height+=2;
	if(lefter+width > document.body.clientWidth) lefter=lefter-width;
	//if(topper+height > document.body.clientHeight) topper=topper-height;

	oPopupMenu.show(lefter - leftoff - 2, topper - topoff + 22, width, height, document.body);

	return false;
}

//calc count
var _calcCountTimer;
function calcWordCount() {
	if (!	validateMode())	return;
	var s_current = '当前 ';
	var s_word = ' 个字符';
	var s_maxword = '最多 ';
	var t = document.getElementById('wordCount');
	var t1 = document.getElementById('wordCount1');
	if (t) {
		t.innerHTML = '['+s_current+ IframeID.document.body.innerHTML.length + s_word + (_maxCount > 0 ? ','+s_maxword+ _maxCount + s_word : '') + ']';
	}
	if (t1) {
		t1.innerHTML = '['+s_current + IframeID.value.length +  s_word + (_maxCount > 0 ? ','+s_maxword + _maxCount + s_word : '') + ']';
	}
	if (_calcCountTimer) {
		window.clearTimeout(_calcCountTimer);
	}
	_calcCountTimer = window.setTimeout('calcWordCount()', 1000);
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
//代码过滤及JS提取
function Newasp_FilterScript(content)
{
	content = content.replace(/<(\w[^div|>]*) class\s*=\s*([^>|\s]*)([^>]*)/gi,"<$1$3") ;
	content = content.replace(/<(\w[^font|>]*) style\s*=\s*\"[^\"]*\"([^>]*>)/gi,"<$1 $2") ;
	content = content.replace(/<(\w[^>]*) lang\s*=\s*([^>|\s]*)([^>]*)/gi,"<$1$3") ;
	var RegExp = /<(script[^>]*)>((.|\n)*)<\/script>/gi;
	content = content.replace(RegExp, "[code]&lt;$1&gt;<br>$2<br>&lt;\/script&gt;[\/code]");
	RegExp = /<(\w[^>|\s]*)([^>]*)(on(finish|mouse|Exit|error|click|key|load|change|focus|blur))(.[^>]*)/gi;
	content = content.replace(RegExp, "<$1")
	RegExp = /<(\w[^>|\s]*)([^>]*)(&#|window\.|javascript:|js:|about:|file:|Document\.|vbs:|cookie| name| id)(.[^>]*)/gi;
	content = content.replace(RegExp, "<$1")
	return content;
}

// 取编辑器的内容
function getHTML() {
	var html;
	Newasp_formatimg();
	html = IframeID.document.body.innerHTML;
	html = Newasp_cleanHtml(html);
	html = Newasp_FilterScript(html);
	return html;
}