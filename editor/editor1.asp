<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%
If enchiasp.CheckPost = False Then
	Call OutAlertScript("您提交的数据不合法，请不要从外部提交。")
End If
If Session("AdminName") = "" Then Response.End
Dim ChannelID,AutoRemote
ChannelID = CInt(Request("ChannelID"))
AutoRemote = 0     '是否自动保存远程图片,1=是,0=否
%>
<html>
<head>
<title> 在线编辑器</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Editor.css" type="text/css" rel="stylesheet">
<Script Language=Javascript>
var sPath = document.location.pathname;
sPath = sPath.substr(0, sPath.length-10);
var sLinkFieldName = "content" ;
var sLinkOriginalFileName = "originalfilename" ;
var sLinkSaveFileName = "savefilename" ;
var sLinkSavePathFileName = "UploadFileList" ;
// 全局设置对象
var config = new Object() ;
config.License = "" ;
config.StyleName = "s_light";
config.StyleMenuHeader = "<head><link href=\"MenuArea.css\" type=\"text/css\" rel=\"stylesheet\"></head><body scroll=\"no\" onConTextMenu=\"event.returnValue=false;\">";
config.StyleDir = "";
config.StyleUploadDir = "UploadFile/";
config.InitMode = "EDIT";
config.AutoDetectPasteFromWord = true;
config.BaseUrl = "1";
config.BaseHref = "";
config.AutoRemote = "<%=AutoRemote%>";
config.ShowBorder = "0";
config.ChannelID = "<%=ChannelID%>";
var sBaseHref = "";
if(config.BaseHref!=""){
	sBaseHref = "<base href=\"" + document.location.protocol + "//" + document.location.host + config.BaseHref + "\">";
}
config.StyleEditorHeader = "<head><link href=\"EditorArea.css\" type=\"text/css\" rel=\"stylesheet\">" + sBaseHref + "</head><body MONOSPACE>" ;
</Script>
<Script Language=Javascript src="editor.js"></Script>
<Script Language=Javascript src="table.js"></Script>
<Script Language=Javascript src="menu.js"></Script>
<script language="javascript" event="onerror(msg, url, line)" for="window">
return true ;	 // 隐藏错误
</script>
</head>
<body SCROLLING=no onConTextMenu="event.returnValue=false;">
<table border=0 cellpadding=0 cellspacing=0 width='100%' height='100%'>
<tr><td>
	<table border=0 cellpadding=0 cellspacing=0 width='100%' class='Toolbar' id='enchicms_Toolbar'><tr><td><div class=yToolbar><DIV CLASS="TBHandle"></DIV><SELECT CLASS="TBGen" onchange="format('FormatBlock',this[this.selectedIndex].value);this.selectedIndex=0"><option selected>段落样式</option>
<option value="&lt;P&gt;">普通</option>
<option value="&lt;H1&gt;">标题一</option>
<option value="&lt;H2&gt;">标题二</option>
<option value="&lt;H3&gt;">标题三</option>
<option value="&lt;H4&gt;">标题四</option>
<option value="&lt;H5&gt;">标题五</option>
<option value="&lt;H6&gt;">标题六</option>
<option value="&lt;p&gt;">段落</option>
<option value="&lt;dd&gt;">定义</option>
<option value="&lt;dt&gt;">术语定义</option>
<option value="&lt;dir&gt;">目录列表</option>
<option value="&lt;menu&gt;">菜单列表</option>
<option value="&lt;PRE&gt;">已编排格式</option></SELECT><SELECT CLASS="TBGen" onchange="format('fontname',this[this.selectedIndex].value);this.selectedIndex=0"><option selected>字体</option>
<option value="宋体">宋体</option>
<option value="黑体">黑体</option>
<option value="楷体_GB2312">楷体</option>
<option value="仿宋_GB2312">仿宋</option>
<option value="隶书">隶书</option>
<option value="幼圆">幼圆</option>
<option value="Arial">Arial</option>
<option value="Arial Black">Arial Black</option>
<option value="Arial Narrow">Arial Narrow</option>
<option value="Brush Script	MT">Brush Script MT</option>
<option value="Century Gothic">Century Gothic</option>
<option value="Comic Sans MS">Comic Sans MS</option>
<option value="Courier">Courier</option>
<option value="Courier New">Courier New</option>
<option value="MS Sans Serif">MS Sans Serif</option>
<option value="Script">Script</option>
<option value="System">System</option>
<option value="Times New Roman">Times New Roman</option>
<option value="Verdana">Verdana</option>
<option value="Wide Latin">Wide Latin</option>
<option value="Wingdings">Wingdings</option></SELECT><SELECT CLASS="TBGen" onchange="format('fontsize',this[this.selectedIndex].value);this.selectedIndex=0"><option selected>字号</option>
<option value="7">一号</option>
<option value="6">二号</option>
<option value="5">三号</option>
<option value="4">四号</option>
<option value="3">五号</option>
<option value="2">六号</option>
<option value="1">七号</option></SELECT><SELECT CLASS="TBGen" onchange="doZoom(this[this.selectedIndex].value)"><option value="10">10%</option>
<option value="25">25%</option>
<option value="50">50%</option>
<option value="75">75%</option>
<option value="100" selected>100%</option>
<option value="150">150%</option>
<option value="200">200%</option>
<option value="500">500%</option>
</SELECT>
<DIV CLASS="Btn" TITLE="粗体" onclick="format('bold')"><IMG CLASS="Ico" SRC="images/bold.gif"></DIV>
<DIV CLASS="Btn" TITLE="斜体" onclick="format('italic')"><IMG CLASS="Ico" SRC="images/italic.gif"></DIV>
<DIV CLASS="Btn" TITLE="下划线" onclick="format('underline')"><IMG CLASS="Ico" SRC="images/underline.gif"></DIV>
<DIV CLASS="Btn" TITLE="中划线" onclick="format('StrikeThrough')"><IMG CLASS="Ico" SRC="images/strikethrough.gif"></DIV>
<DIV CLASS="Btn" TITLE="上标" onclick="format('superscript')"><IMG CLASS="Ico" SRC="images/superscript.gif"></DIV>
<DIV CLASS="Btn" TITLE="下标" onclick="format('subscript')"><IMG CLASS="Ico" SRC="images/subscript.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="左对齐" onclick="format('justifyleft')"><IMG CLASS="Ico" SRC="images/JustifyLeft.gif"></DIV>
<DIV CLASS="Btn" TITLE="居中对齐" onclick="format('justifycenter')"><IMG CLASS="Ico" SRC="images/JustifyCenter.gif"></DIV>
<DIV CLASS="Btn" TITLE="右对齐" onclick="format('justifyright')"><IMG CLASS="Ico" SRC="images/JustifyRight.gif"></DIV>
<DIV CLASS="Btn" TITLE="两端对齐" onclick="format('JustifyFull')"><IMG CLASS="Ico" SRC="images/JustifyFull.gif"></DIV></div></td>
</tr>
<tr>
<td><div class=yToolbar><DIV CLASS="TBHandle"></DIV>
<DIV CLASS="Btn" TITLE="剪切" onclick="format('cut')"><IMG CLASS="Ico" SRC="images/cut.gif"></DIV>
<DIV CLASS="Btn" TITLE="复制" onclick="format('copy')"><IMG CLASS="Ico" SRC="images/copy.gif"></DIV>
<DIV CLASS="Btn" TITLE="常规粘贴" onclick="format('paste')"><IMG CLASS="Ico" SRC="images/paste.gif"></DIV>
<DIV CLASS="Btn" TITLE="纯文本粘贴" onclick="PasteText()"><IMG CLASS="Ico" SRC="images/pastetext.gif"></DIV>
<DIV CLASS="Btn" TITLE="从Word中粘贴" onclick="PasteWord()"><IMG CLASS="Ico" SRC="images/pasteword.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="查找替换" onclick="findReplace()"><IMG CLASS="Ico" SRC="images/findreplace.gif"></DIV>
<DIV CLASS="Btn" TITLE="删除" onclick="format('delete')"><IMG CLASS="Ico" SRC="images/delete.gif"></DIV>
<DIV CLASS="Btn" TITLE="删除文字格式" onclick="format('RemoveFormat')"><IMG CLASS="Ico" SRC="images/RemoveFormat.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="撤消" onclick="goHistory(-1)"><IMG CLASS="Ico" SRC="images/undo.gif"></DIV>
<DIV CLASS="Btn" TITLE="恢复" onclick="goHistory(1)"><IMG CLASS="Ico" SRC="images/redo.gif"></DIV><DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="全部选中" onclick="format('SelectAll')"><IMG CLASS="Ico" SRC="images/selectAll.gif"></DIV>
<DIV CLASS="Btn" TITLE="取消选择" onclick="format('Unselect')"><IMG CLASS="Ico" SRC="images/unselect.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="编号" onclick="format('insertorderedlist')"><IMG CLASS="Ico" SRC="images/insertorderedlist.gif"></DIV>
<DIV CLASS="Btn" TITLE="项目符号" onclick="format('insertunorderedlist')"><IMG CLASS="Ico" SRC="images/insertunorderedlist.gif"></DIV>
<DIV CLASS="Btn" TITLE="增加缩进量" onclick="format('indent')"><IMG CLASS="Ico" SRC="images/indent.gif"></DIV>
<DIV CLASS="Btn" TITLE="减少缩进量" onclick="format('outdent')"><IMG CLASS="Ico" SRC="images/outdent.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="字体颜色" onclick="ShowDialog('dialog/selcolor.htm?action=forecolor', 280, 250, true)"><IMG CLASS="Ico" SRC="images/forecolor.gif"></DIV>
<DIV CLASS="Btn" TITLE="对象背景颜色" onclick="ShowDialog('dialog/selcolor.htm?action=bgcolor', 280, 250, true)"><IMG CLASS="Ico" SRC="images/bgcolor.gif"></DIV>
<DIV CLASS="Btn" TITLE="字体背景颜色" onclick="ShowDialog('dialog/selcolor.htm?action=backcolor', 280, 250, true)"><IMG CLASS="Ico" SRC="images/backcolor.gif"></DIV>
<DIV CLASS="Btn" TITLE="背景图片" onclick="ShowDialog('dialog/backimage.htm', 350, 210, true)"><IMG CLASS="Ico" SRC="images/bgpic.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="绝对或相对位置" onclick="absolutePosition()"><IMG CLASS="Ico" SRC="images/abspos.gif"></DIV>
<DIV CLASS="Btn" TITLE="上移一层" onclick="zIndex('forward')"><IMG CLASS="Ico" SRC="images/forward.gif"></DIV>
<DIV CLASS="Btn" TITLE="下移一层" onclick="zIndex('backward')"><IMG CLASS="Ico" SRC="images/backward.gif"></DIV></div></td>
</tr>
<tr>
<td><div class=yToolbar><DIV CLASS="TBHandle"></DIV>
<DIV CLASS="Btn" TITLE="插入或修改图片" onclick="ShowDialog('dialog/img.htm', 350, 315, true)"><IMG CLASS="Ico" SRC="images/img.gif"></DIV>
<DIV CLASS="Btn" TITLE="插入Flash动画" onclick="ShowDialog('dialog/flash.htm', 350, 200, true)"><IMG CLASS="Ico" SRC="images/flash.gif"></DIV>
<DIV CLASS="Btn" TITLE="插入自动播放的媒体文件" onclick="ShowDialog('dialog/media.htm', 350, 200, true)"><IMG CLASS="Ico" SRC="images/Media.gif"></DIV>
<DIV CLASS="Btn" TITLE="插入其他文件" onclick="ShowDialog('dialog/file.htm', 350, 150, true)"><IMG CLASS="Ico" SRC="images/file.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="表格菜单" onclick="showToolMenu('table')"><IMG CLASS="Ico" SRC="images/tablemenu.gif"></DIV>
<DIV CLASS="Btn" TITLE="表单菜单" onclick="showToolMenu('form')"><IMG CLASS="Ico" SRC="images/FormMenu.gif"></DIV>
<DIV CLASS="Btn" TITLE="显示或隐藏指导方针" onclick="showBorders()"><IMG CLASS="Ico" SRC="images/ShowBorders.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="插入或修改栏目框" onclick="ShowDialog('dialog/fieldset.htm', 350, 170, true)"><IMG CLASS="Ico" SRC="images/fieldset.gif"></DIV>
<DIV CLASS="Btn" TITLE="插入或修改网页帧" onclick="ShowDialog('dialog/iframe.htm', 350, 200, true)"><IMG CLASS="Ico" SRC="images/iframe.gif"></DIV>
<DIV CLASS="Btn" TITLE="插入水平尺" onclick="format('InsertHorizontalRule')"><IMG CLASS="Ico" SRC="images/InsertHorizontalRule.gif"></DIV>
<DIV CLASS="Btn" TITLE="插入或修改字幕" onclick="ShowDialog('dialog/marquee.htm', 395, 150, true)"><IMG CLASS="Ico" SRC="images/Marquee.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="插入或修改超级链接" onclick="createLink()"><IMG CLASS="Ico" SRC="images/CreateLink.gif"></DIV>
<DIV CLASS="Btn" TITLE="图形热点链接" onclick="mapEdit()"><IMG CLASS="Ico" SRC="images/map.gif"></DIV>
<DIV CLASS="Btn" TITLE="取消超级链接或标签" onclick="format('UnLink')"><IMG CLASS="Ico" SRC="images/Unlink.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="插入特殊字符" onclick="ShowDialog('dialog/symbol.htm', 350, 220, true)"><IMG CLASS="Ico" SRC="images/symbol.gif"></DIV>
<DIV CLASS="Btn" TITLE="插入表情图标" onclick="ShowDialog('dialog/emot.htm', 300, 180, true)"><IMG CLASS="Ico" SRC="images/emot.gif"></DIV>
<DIV CLASS="Btn" TITLE="插入Excel表格" onclick="insert('excel')"><IMG CLASS="Ico" SRC="images/excel.gif"></DIV>
<DIV CLASS="Btn" TITLE="插入当前日期" onclick="insert('nowdate')"><IMG CLASS="Ico" SRC="images/date.gif"></DIV>
<DIV CLASS="Btn" TITLE="插入当前时间" onclick="insert('nowtime')"><IMG CLASS="Ico" SRC="images/time.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="引用样式" onclick="insert('quote')"><IMG CLASS="Ico" SRC="images/quote.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="全屏编辑" onclick="Maximize()"><IMG CLASS="Ico" SRC="images/maximize.gif"></DIV>
<DIV CLASS="Btn" TITLE="查看使用帮助" onclick="ShowDialog('dialog/help.htm','400','300')"><IMG CLASS="Ico" SRC="images/help.gif"></DIV>
<DIV CLASS="Btn" TITLE="插入分页符" onclick="insert('page')"><IMG CLASS="Ico" SRC="images/insertpage.gif"></DIV></div></td>
</tr>
</table>
</td></tr>
<tr><td height='100%'>
	<table border=0 cellpadding=0 cellspacing=0 width='100%' height='100%'>
	<tr><td height='100%'>
	<input type="hidden" ID="ContentEdit" value="">
	<input type="hidden" ID="ModeEdit" value="">
	<input type="hidden" ID="ContentLoad" value="">
	<input type="hidden" ID="ContentFlag" value="0">
<input name="image" type='hidden' id="image">
	<iframe class="Composition" ID="enchicms" MARGINHEIGHT="1" MARGINWIDTH="1" width="100%" height="100%" scrolling="yes"> 
	</iframe>
	</td></tr>
	</table>
</td></tr>
<tr><td height=25>
	<TABLE border="0" cellPadding="0" cellSpacing="0" width="100%" class=StatusBar height=25>
	<TR valign=middle>
	<td>
		<table border=0 cellpadding=0 cellspacing=0 height=20>
		<tr>
		<td width=10></td>
		<td class=StatusBarBtnOff id=enchicms_CODE onclick="setMode('CODE')"><img border=0 src="images/modecode.gif" width=50 height=15 align=absmiddle></td>
		<td width=5></td>
		<td class=StatusBarBtnOff id=enchicms_EDIT onclick="setMode('EDIT')"><img border=0 src="images/modeedit.gif" width=50 height=15 align=absmiddle></td>
		<td width=5></td>
		<td class=StatusBarBtnOff id=enchicms_TEXT onclick="setMode('TEXT')"><img border=0 src="images/modetext.gif" width=50 height=15 align=absmiddle></td>
		<td width=5></td>
		<td class=StatusBarBtnOff id=enchicms_VIEW onclick="setMode('VIEW')"><img border=0 src="images/modepreview.gif" width=50 height=15 align=absmiddle></td>
		</tr>
		</table>
	</td>
	<td align=center style="font-size:9pt"><input type=checkbox name=AutoRemote value='1' onClick="remoteUpload();"> 自动保存远程图片</td>        
	<td align=right>
		<table border=0 cellpadding=0 cellspacing=0 height=20>
		<tr>
		<td style="cursor:pointer;" onclick="sizeChange(300)"><img border=0 SRC="images/sizeplus.gif" width=20 height=20 alt="增高编辑区"></td>
		<td width=5></td>
		<td style="cursor:pointer;" onclick="sizeChange(-300)"><img border=0 SRC="images/sizeminus.gif" width=20 height=20 alt="减小编辑区"></td>
		<td width=40></td>
		</tr>
		</table>
	</td>
	</TR>
	</Table>
</td></tr>
</table>
<div id="enchicms_Temp_HTML" style="VISIBILITY: hidden; OVERFLOW: hidden; POSITION: absolute; WIDTH: 1px; HEIGHT: 1px"></div>
<form id="enchicms_UploadForm" action="upload.asp?action=remote&type=remote&ChannelID=<%=ChannelID%>" method="post" target="enchicms_UploadTarget">
<input type="hidden" name="enchicms_UploadText">
</form>
<iframe name="enchicms_UploadTarget" width=0 height=0></iframe>
<div id=divProcessing style="width:200px;height:30px;position:absolute;display:none">
<table border=0 cellpadding=0 cellspacing=1 bgcolor="#000000" width="100%" height="100%"><tr><td bgcolor=#0650D2><marquee align="middle" behavior="alternate" scrollamount="5" style="font-size:9pt"><font color=#FFFFFF>...正在保存数据...请等待...</font></marquee></td></tr></table>
</div>
</body>
</html>
