<!--#include file="../conn.asp"-->
<!--#include file="../inc/const.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%
If enchiasp.CheckPost = False Then
	Call OutAlertScript("���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ��")
End If
If Session("AdminName") = "" Then Response.End
Dim ChannelID,AutoRemote
ChannelID = CInt(Request("ChannelID"))
AutoRemote = 0     '�Ƿ��Զ�����Զ��ͼƬ,1=��,0=��
%>
<html>
<head>
<title> ���߱༭��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Editor.css" type="text/css" rel="stylesheet">
<Script Language=Javascript>
var sPath = document.location.pathname;
sPath = sPath.substr(0, sPath.length-10);
var sLinkFieldName = "content" ;
var sLinkOriginalFileName = "originalfilename" ;
var sLinkSaveFileName = "savefilename" ;
var sLinkSavePathFileName = "UploadFileList" ;
// ȫ�����ö���
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
return true ;	 // ���ش���
</script>
</head>
<body SCROLLING=no onConTextMenu="event.returnValue=false;">
<table border=0 cellpadding=0 cellspacing=0 width='100%' height='100%'>
<tr><td>
	<table border=0 cellpadding=0 cellspacing=0 width='100%' class='Toolbar' id='enchicms_Toolbar'><tr><td><div class=yToolbar><DIV CLASS="TBHandle"></DIV><SELECT CLASS="TBGen" onchange="format('FormatBlock',this[this.selectedIndex].value);this.selectedIndex=0"><option selected>������ʽ</option>
<option value="&lt;P&gt;">��ͨ</option>
<option value="&lt;H1&gt;">����һ</option>
<option value="&lt;H2&gt;">�����</option>
<option value="&lt;H3&gt;">������</option>
<option value="&lt;H4&gt;">������</option>
<option value="&lt;H5&gt;">������</option>
<option value="&lt;H6&gt;">������</option>
<option value="&lt;p&gt;">����</option>
<option value="&lt;dd&gt;">����</option>
<option value="&lt;dt&gt;">���ﶨ��</option>
<option value="&lt;dir&gt;">Ŀ¼�б�</option>
<option value="&lt;menu&gt;">�˵��б�</option>
<option value="&lt;PRE&gt;">�ѱ��Ÿ�ʽ</option></SELECT><SELECT CLASS="TBGen" onchange="format('fontname',this[this.selectedIndex].value);this.selectedIndex=0"><option selected>����</option>
<option value="����">����</option>
<option value="����">����</option>
<option value="����_GB2312">����</option>
<option value="����_GB2312">����</option>
<option value="����">����</option>
<option value="��Բ">��Բ</option>
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
<option value="Wingdings">Wingdings</option></SELECT><SELECT CLASS="TBGen" onchange="format('fontsize',this[this.selectedIndex].value);this.selectedIndex=0"><option selected>�ֺ�</option>
<option value="7">һ��</option>
<option value="6">����</option>
<option value="5">����</option>
<option value="4">�ĺ�</option>
<option value="3">���</option>
<option value="2">����</option>
<option value="1">�ߺ�</option></SELECT><SELECT CLASS="TBGen" onchange="doZoom(this[this.selectedIndex].value)"><option value="10">10%</option>
<option value="25">25%</option>
<option value="50">50%</option>
<option value="75">75%</option>
<option value="100" selected>100%</option>
<option value="150">150%</option>
<option value="200">200%</option>
<option value="500">500%</option>
</SELECT>
<DIV CLASS="Btn" TITLE="����" onclick="format('bold')"><IMG CLASS="Ico" SRC="images/bold.gif"></DIV>
<DIV CLASS="Btn" TITLE="б��" onclick="format('italic')"><IMG CLASS="Ico" SRC="images/italic.gif"></DIV>
<DIV CLASS="Btn" TITLE="�»���" onclick="format('underline')"><IMG CLASS="Ico" SRC="images/underline.gif"></DIV>
<DIV CLASS="Btn" TITLE="�л���" onclick="format('StrikeThrough')"><IMG CLASS="Ico" SRC="images/strikethrough.gif"></DIV>
<DIV CLASS="Btn" TITLE="�ϱ�" onclick="format('superscript')"><IMG CLASS="Ico" SRC="images/superscript.gif"></DIV>
<DIV CLASS="Btn" TITLE="�±�" onclick="format('subscript')"><IMG CLASS="Ico" SRC="images/subscript.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="�����" onclick="format('justifyleft')"><IMG CLASS="Ico" SRC="images/JustifyLeft.gif"></DIV>
<DIV CLASS="Btn" TITLE="���ж���" onclick="format('justifycenter')"><IMG CLASS="Ico" SRC="images/JustifyCenter.gif"></DIV>
<DIV CLASS="Btn" TITLE="�Ҷ���" onclick="format('justifyright')"><IMG CLASS="Ico" SRC="images/JustifyRight.gif"></DIV>
<DIV CLASS="Btn" TITLE="���˶���" onclick="format('JustifyFull')"><IMG CLASS="Ico" SRC="images/JustifyFull.gif"></DIV></div></td>
</tr>
<tr>
<td><div class=yToolbar><DIV CLASS="TBHandle"></DIV>
<DIV CLASS="Btn" TITLE="����" onclick="format('cut')"><IMG CLASS="Ico" SRC="images/cut.gif"></DIV>
<DIV CLASS="Btn" TITLE="����" onclick="format('copy')"><IMG CLASS="Ico" SRC="images/copy.gif"></DIV>
<DIV CLASS="Btn" TITLE="����ճ��" onclick="format('paste')"><IMG CLASS="Ico" SRC="images/paste.gif"></DIV>
<DIV CLASS="Btn" TITLE="���ı�ճ��" onclick="PasteText()"><IMG CLASS="Ico" SRC="images/pastetext.gif"></DIV>
<DIV CLASS="Btn" TITLE="��Word��ճ��" onclick="PasteWord()"><IMG CLASS="Ico" SRC="images/pasteword.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="�����滻" onclick="findReplace()"><IMG CLASS="Ico" SRC="images/findreplace.gif"></DIV>
<DIV CLASS="Btn" TITLE="ɾ��" onclick="format('delete')"><IMG CLASS="Ico" SRC="images/delete.gif"></DIV>
<DIV CLASS="Btn" TITLE="ɾ�����ָ�ʽ" onclick="format('RemoveFormat')"><IMG CLASS="Ico" SRC="images/RemoveFormat.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="����" onclick="goHistory(-1)"><IMG CLASS="Ico" SRC="images/undo.gif"></DIV>
<DIV CLASS="Btn" TITLE="�ָ�" onclick="goHistory(1)"><IMG CLASS="Ico" SRC="images/redo.gif"></DIV><DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="ȫ��ѡ��" onclick="format('SelectAll')"><IMG CLASS="Ico" SRC="images/selectAll.gif"></DIV>
<DIV CLASS="Btn" TITLE="ȡ��ѡ��" onclick="format('Unselect')"><IMG CLASS="Ico" SRC="images/unselect.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="���" onclick="format('insertorderedlist')"><IMG CLASS="Ico" SRC="images/insertorderedlist.gif"></DIV>
<DIV CLASS="Btn" TITLE="��Ŀ����" onclick="format('insertunorderedlist')"><IMG CLASS="Ico" SRC="images/insertunorderedlist.gif"></DIV>
<DIV CLASS="Btn" TITLE="����������" onclick="format('indent')"><IMG CLASS="Ico" SRC="images/indent.gif"></DIV>
<DIV CLASS="Btn" TITLE="����������" onclick="format('outdent')"><IMG CLASS="Ico" SRC="images/outdent.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="������ɫ" onclick="ShowDialog('dialog/selcolor.htm?action=forecolor', 280, 250, true)"><IMG CLASS="Ico" SRC="images/forecolor.gif"></DIV>
<DIV CLASS="Btn" TITLE="���󱳾���ɫ" onclick="ShowDialog('dialog/selcolor.htm?action=bgcolor', 280, 250, true)"><IMG CLASS="Ico" SRC="images/bgcolor.gif"></DIV>
<DIV CLASS="Btn" TITLE="���屳����ɫ" onclick="ShowDialog('dialog/selcolor.htm?action=backcolor', 280, 250, true)"><IMG CLASS="Ico" SRC="images/backcolor.gif"></DIV>
<DIV CLASS="Btn" TITLE="����ͼƬ" onclick="ShowDialog('dialog/backimage.htm', 350, 210, true)"><IMG CLASS="Ico" SRC="images/bgpic.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="���Ի����λ��" onclick="absolutePosition()"><IMG CLASS="Ico" SRC="images/abspos.gif"></DIV>
<DIV CLASS="Btn" TITLE="����һ��" onclick="zIndex('forward')"><IMG CLASS="Ico" SRC="images/forward.gif"></DIV>
<DIV CLASS="Btn" TITLE="����һ��" onclick="zIndex('backward')"><IMG CLASS="Ico" SRC="images/backward.gif"></DIV></div></td>
</tr>
<tr>
<td><div class=yToolbar><DIV CLASS="TBHandle"></DIV>
<DIV CLASS="Btn" TITLE="������޸�ͼƬ" onclick="ShowDialog('dialog/img.htm', 350, 315, true)"><IMG CLASS="Ico" SRC="images/img.gif"></DIV>
<DIV CLASS="Btn" TITLE="����Flash����" onclick="ShowDialog('dialog/flash.htm', 350, 200, true)"><IMG CLASS="Ico" SRC="images/flash.gif"></DIV>
<DIV CLASS="Btn" TITLE="�����Զ����ŵ�ý���ļ�" onclick="ShowDialog('dialog/media.htm', 350, 200, true)"><IMG CLASS="Ico" SRC="images/Media.gif"></DIV>
<DIV CLASS="Btn" TITLE="���������ļ�" onclick="ShowDialog('dialog/file.htm', 350, 150, true)"><IMG CLASS="Ico" SRC="images/file.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="���˵�" onclick="showToolMenu('table')"><IMG CLASS="Ico" SRC="images/tablemenu.gif"></DIV>
<DIV CLASS="Btn" TITLE="���˵�" onclick="showToolMenu('form')"><IMG CLASS="Ico" SRC="images/FormMenu.gif"></DIV>
<DIV CLASS="Btn" TITLE="��ʾ������ָ������" onclick="showBorders()"><IMG CLASS="Ico" SRC="images/ShowBorders.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="������޸���Ŀ��" onclick="ShowDialog('dialog/fieldset.htm', 350, 170, true)"><IMG CLASS="Ico" SRC="images/fieldset.gif"></DIV>
<DIV CLASS="Btn" TITLE="������޸���ҳ֡" onclick="ShowDialog('dialog/iframe.htm', 350, 200, true)"><IMG CLASS="Ico" SRC="images/iframe.gif"></DIV>
<DIV CLASS="Btn" TITLE="����ˮƽ��" onclick="format('InsertHorizontalRule')"><IMG CLASS="Ico" SRC="images/InsertHorizontalRule.gif"></DIV>
<DIV CLASS="Btn" TITLE="������޸���Ļ" onclick="ShowDialog('dialog/marquee.htm', 395, 150, true)"><IMG CLASS="Ico" SRC="images/Marquee.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="������޸ĳ�������" onclick="createLink()"><IMG CLASS="Ico" SRC="images/CreateLink.gif"></DIV>
<DIV CLASS="Btn" TITLE="ͼ���ȵ�����" onclick="mapEdit()"><IMG CLASS="Ico" SRC="images/map.gif"></DIV>
<DIV CLASS="Btn" TITLE="ȡ���������ӻ��ǩ" onclick="format('UnLink')"><IMG CLASS="Ico" SRC="images/Unlink.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="���������ַ�" onclick="ShowDialog('dialog/symbol.htm', 350, 220, true)"><IMG CLASS="Ico" SRC="images/symbol.gif"></DIV>
<DIV CLASS="Btn" TITLE="�������ͼ��" onclick="ShowDialog('dialog/emot.htm', 300, 180, true)"><IMG CLASS="Ico" SRC="images/emot.gif"></DIV>
<DIV CLASS="Btn" TITLE="����Excel���" onclick="insert('excel')"><IMG CLASS="Ico" SRC="images/excel.gif"></DIV>
<DIV CLASS="Btn" TITLE="���뵱ǰ����" onclick="insert('nowdate')"><IMG CLASS="Ico" SRC="images/date.gif"></DIV>
<DIV CLASS="Btn" TITLE="���뵱ǰʱ��" onclick="insert('nowtime')"><IMG CLASS="Ico" SRC="images/time.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="������ʽ" onclick="insert('quote')"><IMG CLASS="Ico" SRC="images/quote.gif"></DIV>
<DIV CLASS="TBSep"></DIV>
<DIV CLASS="Btn" TITLE="ȫ���༭" onclick="Maximize()"><IMG CLASS="Ico" SRC="images/maximize.gif"></DIV>
<DIV CLASS="Btn" TITLE="�鿴ʹ�ð���" onclick="ShowDialog('dialog/help.htm','400','300')"><IMG CLASS="Ico" SRC="images/help.gif"></DIV>
<DIV CLASS="Btn" TITLE="�����ҳ��" onclick="insert('page')"><IMG CLASS="Ico" SRC="images/insertpage.gif"></DIV></div></td>
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
	<td align=center style="font-size:9pt"><input type=checkbox name=AutoRemote value='1' onClick="remoteUpload();"> �Զ�����Զ��ͼƬ</td>        
	<td align=right>
		<table border=0 cellpadding=0 cellspacing=0 height=20>
		<tr>
		<td style="cursor:pointer;" onclick="sizeChange(300)"><img border=0 SRC="images/sizeplus.gif" width=20 height=20 alt="���߱༭��"></td>
		<td width=5></td>
		<td style="cursor:pointer;" onclick="sizeChange(-300)"><img border=0 SRC="images/sizeminus.gif" width=20 height=20 alt="��С�༭��"></td>
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
<table border=0 cellpadding=0 cellspacing=1 bgcolor="#000000" width="100%" height="100%"><tr><td bgcolor=#0650D2><marquee align="middle" behavior="alternate" scrollamount="5" style="font-size:9pt"><font color=#FFFFFF>...���ڱ�������...��ȴ�...</font></marquee></td></tr></table>
</div>
</body>
</html>
