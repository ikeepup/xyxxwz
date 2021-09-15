<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
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
Call InnerLocation("发布文章")
Dim Rs,SQL
dim temp
dim i
dim tt
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID < 1 Then ChannelID = 1
ChannelID = CLng(ChannelID)

If CInt(GroupSetting(7)) = 0 Then
	ErrMsg = ErrMsg + "<li>对不起！您没有发布文章的权限，如需要该权限请联系管理员。</li>"
	Founderr = True
End If
tt=false
temp=split(GroupSetting(35),"$$$")
for i=0 to ubound(temp)
	If temp(i) = ChannelID  Then
		tt=true
		exit for
	End If
next

if tt=false then
	ErrMsg = ErrMsg + "<li>对不起！您没有该频道发布文章的权限，如需要该权限请联系管理员。</li>"
	Founderr = True
end if


Dim Action:Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveArticle
	Case "view"
		Call ArticleView
	Case Else
		Call showmain
End Select
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
Sub showmain()
	If Founderr = True Then Exit Sub
%>
<script language=JavaScript>
var _maxCount = '<%=CLng(GroupSetting(16))%>';
function doChange(objText, objDrop){
	if (!objDrop) return;
	if(document.myform.BriefTopic.selectedIndex<2){
		document.myform.BriefTopic.selectedIndex+=1;
	}
	var str = objText.value;
	var arr = str.split("|");
	var nIndex = objDrop.selectedIndex;
	objDrop.length=1;
	for (var i=0; i<arr.length; i++){
		objDrop.options[objDrop.length] = new Option(arr[i], arr[i]);
	}
	objDrop.selectedIndex = nIndex;
}
function doSubmit(){
	if (document.myform.title.value==""){
		alert("文章标题不能为空！");
		return false;
	}
	if (document.myform.Author.value==""){
		alert("文章作者不能为空！");
		return false;
	}
	if (document.myform.ComeFrom.value==""){
		alert("文章来源不能为空！");
		return false;
	}
	if (document.myform.ClassID.value==""){
		alert("该一级分类已经有下属分类，请选择其下属分类！");
		return false;
	}
	if (document.myform.ClassID.value=="0"){
		alert("该分类是外部连接，不能添加内容！");
		return false;
	}
	<%If CInt(GroupSetting(2)) = 1 Then%>
	if (document.myform.codestr.value==""){
		alert("请填写验证码！");
		return false;
	}
	<%End If%>
	myform.content.value = getHTML(); 
	MessageLength = Composition.document.body.innerHTML.length;
	if(MessageLength < 2){
		alert("文章内容不能小于2个字符！");
		return false;
	}
	if(MessageLength > _maxCount){
		alert("文章的内容不能超过"+_maxCount+"个字符！");
		return false;
	}
	document.myform.Submit1.disabled = true;
	document.myform.submit();
}
</script>

<div onkeydown=CtrlEnter()>
<table  border="0" align="center" cellpadding="3" cellspacing="1" class="Usertableborder">
        <tr>
          <th colspan="2">&gt;&gt;发布文章&lt;&lt;</th>
        </tr>
	<form method=Post name="myform" action="ArticlePost.Asp">
	<input type="Hidden" name="action" value="save">
	<input type=hidden name=ChannelID value='<%=ChannelID%>'>
        <tr>
          <td width="15%" align="right" nowrap class="usertablerow2"><strong>所属分类</strong></td>
          <td width="85%" class="usertablerow1">
<%
	Dim sClassSelect
	Response.Write "<select name=""ClassID"" id=""ClassID"">"
	sClassSelect = enchiasp.LoadSelectClass(ChannelID)
	'sClassSelect = Replace(sClassSelect, "{ClassID=" & ClassID & "}", "selected")
	Response.Write sClassSelect
	Response.Write "</select>"
%>
	  </td>
        </tr>
        <tr>
          <td align="right" noWrap class="usertablerow2"><strong>文章标题</strong></td>
          <td class="usertablerow1"><select name="BriefTopic" id="BriefTopic">
            <option value="0">选择话题</option>
			<option value="1">[图文]</option>
			<option value="2">[组图]</option>
			<option value="3">[新闻]</option>
			<option value="4">[推荐]</option>
			<option value="5">[注意]</option>
			<option value="6">[转载]</option>
          </select> <input name="title" type="text" id="title" size="60" value="">       
          <font color=red>*</font></td>      
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>相关文章</strong></td>
          <td class="usertablerow1"><input name="Related" type="text" id="Related" size="60" value=""> <font color=red>*</font></td>      
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>文章作者</strong></td>
          <td class="usertablerow1"><input name="Author" type="text" size="30" value="">
		    <select name=font2 onChange="Author.value=this.value;">
			<option selected value="">选择作者</OPTION>
			<option value=佚名>佚名</option>
			<option value=本站>本站</option>
			<option value=不详>不详</option>
			<option value=未知>未知</option>
			</select></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>文章来源</strong></td>
          <td class="usertablerow1"><input name="ComeFrom" type="text" size="30" value="">
		  	<select name=font1 onChange="ComeFrom.value=this.value;">
			<option selected value="">选择来源</OPTION>
			<option value=本站原创>本站原创</option>
			<option value=本站整理>本站整理</option>
			<option value=不详>不详</option>
			<option value=转载>转载</option>
			</select></td>
        </tr>
        <tr>
	  <td align="right" class="usertablerow2"><strong>文章内容</strong></td>
          <td class="usertablerow1"><textarea name='content' id='content' style='display:none' rows="1" cols="20"></textarea>
		<script Language=Javascript src="../editor/post.js"></script></td>      
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>首页图片</strong></td>
          <td class="usertablerow1"><input name="ImageUrl" type="text" id="ImageUrl" size="60" value="">
			<input type=hidden name=UploadFileList id=UploadFileList onchange="doChange(this,document.myform.ImageFileList)" value="">
			<br>直接从上传图片中选择
			<select name="ImageFileList" id="ImageFileList" onChange="ImageUrl.value=this.value;">
			<option value=''>不选择首页推荐图片</option>
			</select>
			</td>
        </tr>
	<tr>
	  <td align="right" class="usertablerow2"><strong>文件上传</strong></td>
          <td class="usertablerow1"><iframe name="image1" frameborder=0 width='100%' height=55 scrolling=no src=upload.asp?ChannelID=<%=ChannelID%>></iframe></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>文章星级</strong></td>
          <td class="usertablerow1"><select name="star">
		<option value=5>★★★★★</option>
          	<option value=4>★★★★</option>
          	<option value=3 selected>★★★</option>
		<option value=2>★★</option>
		<option value=1>★</option>
          </select></td>
        </tr>
<%
	If CInt(GroupSetting(2)) = 1 Then
%>
	<tr>
		<td class=Usertablerow2 align="right"><strong>验证码</strong></td>
		<td class=Usertablerow1><input type="text" name="codestr" maxlength="4" size="4">&nbsp;<img src="../inc/getcode.asp"></td>
	</tr>
<%
	End If
%>
	<tr>
          <td align="right" class="usertablerow2"><strong>禁止评论</strong></td>
          <td class="usertablerow1"><input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"></td>
        </tr>
	<tr>
          <td align="right" class="usertablerow2"><strong>立即发布</strong></td>
          <td class="usertablerow1"><input name="isAccept" type="checkbox" id="isAccept" value="1" disabled> 是（<font color=blue>如果选中的话将直接发布，否则审核后才能发布。</font>）</td>
        </tr>
        <tr align="center">
          <td colspan="2" class="usertablerow2">
	  <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="返回上一页" class=Button>
	  <input type="button" name="Submit1" value="现在发布" class=Button onclick="doSubmit();"></td>
        </tr></form>
      </table></div>
<%
End Sub
Private Sub SaveArticle()
	Dim TextContent,ForbidEssay,isAccept,i,ArticleID
	If CLng(UserToday(3)) => CLng(GroupSetting(10)) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>您每天最多只能发布<font color=red><b>" & GroupSetting(10) & "</b></font>篇文章，如果还要继续发布请明天再来吧！</li>"
	End If
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If Trim(Request.Form("title")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>文章标题不能为空！</li>"
	End If
	If Len(Request.Form("title")) => 100 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>文章标题不能超过100个字符！</li>"
	End If
	If Len(Request.Form("Related")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>相关文章不能超过200个字符！</li>"
	End If
	If Trim(Request.Form("Author")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>文章作者不能为空！</li>"
	End If
	If Trim(Request.Form("ComeFrom")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>文章来源不能为空！</li>"
	End If
	If Not IsNumeric(Request.Form("star")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>文章星级不能为空。</li>"
	End If

	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该一级分类已经有下属分类，不能添加内容！</li>"
	End If
	If Trim(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>该分类是外部连接，不能添加内容！</li>"
	End If
	
	If CInt(GroupSetting(2)) = 1 Then
		If Not enchiasp.CodeIsTrue() Then
			ErrMsg = ErrMsg + "<meta http-equiv=refresh content=""2;URL="&Request.ServerVariables("HTTP_REFERER")&"""><li>验证码校验失败，请返回刷新页面再试。两秒后自动返回</li>"
			Founderr = True
		End If
		Session("GetCode") = ""
	End If
	TextContent = ""
	For i = 1 To Request.Form("content").Count
		TextContent = TextContent & Request.Form("content")(i)
	Next
	If Len(TextContent) < 2 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>文章内容不能为空！</li>"
	End If
	If Len(TextContent) > CLng(GroupSetting(16)) Then
		ErrMsg = ErrMsg + "<li>文章内容不能大于" & GroupSetting(16) & "字符！</li>"
		Founderr = True
	End If
	If CInt(Request("isAccept")) = 1 Then
		isAccept = 1
	Else
		isAccept = 0
	End If
	If CInt(Request.Form("ForbidEssay")) = 1 Then
		ForbidEssay = 1
	Else
		ForbidEssay = 0
	End If
	If Founderr = True Then Exit Sub
	On Error Resume Next
	Call PreventRefresh  '防刷新
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Article where (ArticleID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = enchiasp.ChkNumeric(Request.Form("ClassID"))
		Rs("SpecialID") = 0
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("ColorMode") = 0
		Rs("FontMode") = 0
		Rs("content") = Html2Ubb(TextContent)
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Author") = enchiasp.ChkFormStr(Request.Form("Author"))
		Rs("ComeFrom") = enchiasp.ChkFormStr(Request.Form("ComeFrom"))
		Rs("star") = enchiasp.ChkNumeric(Request.Form("star"))
		Rs("isTop") = 0
		Rs("AllHits") = 0
		Rs("DayHits") = 0
		Rs("WeekHits") = 0
		Rs("MonthHits") = 0
		Rs("HitsTime") = Now()
		Rs("WriteTime") = Now()
		Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
		Rs("username") = Trim(enchiasp.MemberName)
		Rs("isBest") = 0
		Rs("BriefTopic") = enchiasp.ChkNumeric(Request.Form("BriefTopic"))
		Rs("ImageUrl") = enchiasp.ChkFormStr(Request.Form("ImageUrl"))
		Rs("UploadImage") = enchiasp.ChkFormStr(Request.Form("UploadFileList"))
		Rs("UserGroup") = 0
		Rs("PointNum") = 0
		Rs("isUpdate") = 1
		Rs("isAccept") = isAccept
		Rs("ForbidEssay") = ForbidEssay
	Rs.update
	Rs.Close
	Rs.Open "SELECT TOP 1 ArticleID FROM ECCMS_Article WHERE ChannelID=" & ChannelID & " ORDER BY ArticleID DESC", Conn, 1, 1
	ArticleID = Rs("ArticleID")
	Rs.Close:Set Rs = Nothing
	Dim strUserToday
	strUserToday = UserToday(0) &","& UserToday(1) &","& UserToday(2) &","& UserToday(3)+1 &","& UserToday(4) &","& UserToday(5)
	UpdateUserToday(strUserToday)
	Call Returnsuc("<li>恭喜您！提交成功。请等待管理员验证后正式发布。</li><li><a href=ArticlePost.Asp?action=view&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & ">点击此处查看该文章</a></li>")
End Sub

Sub ArticleView()
	If Request("ArticleID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！您选择了错误的系统参数。</li>"
		Exit Sub
	End If
	If ChannelID = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>请指定频道。</li>"
		Exit Sub
	End If
	SQL = "select ArticleID,title,content,ColorMode,FontMode,Author,ComeFrom,WriteTime,username from ECCMS_Article where ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And ArticleID=" & Request("ArticleID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！没有找到任何文章。或者您选择了错误的系统参数！</li>"
		Exit Sub
	Else
%>
<script language=javascript>
var enchiasp_fontsize=9;
var enchiasp_lineheight=12;
</script>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="UserTableBorder" style="table-layout:fixed;word-break:break-all">
	<tr>
	  <th>&gt;&gt;查看文章内容&lt;&lt;</th>
	</tr>
	<tr>
	  <td align="center" class="usertablerow2"><a href=ArticleList.Asp?action=edit&ChannelID=<%=ChannelID%>&ArticleID=<%=Rs("ArticleID")%>><font size=4><%=enchiasp.ReadFontMode(Rs("title"),Rs("ColorMode"),Rs("FontMode"))%></font></a></td>
	</tr>
	<tr>
	  <td align="center" class="usertablerow1">作者：<%=Rs("Author")%>     
        　来源于：<%=Rs("ComeFrom")%>  
        　发布时间：<%=Rs("WriteTime")%>       　发布人：<font color=blue><%=Rs("username")%></font></td>
	</tr>
	<tr>
	  <td class="usertablerow1"><p align="right"><a style="CURSOR: hand; POSITION: relative" onclick="if(enchiasp_fontsize&gt;8){enchiaspContentLabel.style.fontSize=(--enchiasp_fontsize)+&quot;pt&quot;;enchiaspContentLabel.style.lineHeight=(--enchiasp_lineheight)+&quot;pt&quot;;}" title="减小字体"><img src="../images/1.gif" border="0" width="15" height="15"><font color="#FF6600">减小字体</font></a> 
                    <a style="CURSOR: hand; POSITION: relative" onclick="if(enchiasp_fontsize&lt;64){enchiaspContentLabel.style.fontSize=(++enchiasp_fontsize)+&quot;pt&quot;;enchiaspContentLabel.style.lineHeight=(++enchiasp_lineheight)+&quot;pt&quot;;}" title="增大字体"><img src="../images/2.gif" border="0" width="15" height="15"><font color="#FF6600">增大字体</font></a></p>
					<div id="enchiaspContentLabel"><%=Replace(enchiasp.ReadContent(Rs("content")), "[page_break]", "", 1, -1, 1)%></div></td>
	</tr>
	<tr>
	  <td align="center" class="usertablerow2"><input type="button" onclick="javascript:window.close()" value="关闭本窗口" name="B2" class=Button>&nbsp;&nbsp;      
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="返回顶部" class=button></td>
	</tr>
</table>
<%
	End If
	Rs.Close
	Set Rs = Nothing 
End Sub
%>
<!--#include file="foot.inc"-->