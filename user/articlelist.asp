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
Call InnerLocation("我发布的文章")
Dim Action,SQL,Rs,i
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))
If ChannelID < 1 Then ChannelID = 1
ChannelID = CLng(ChannelID)

Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveArticle
	Case "edit"
		Call EditArticle
	Case "del"
		Call DeleteArticle
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
<script language="JavaScript">
<!--
function myuser_articlelist_top(accept){
	document.write ('<th valign=middle>');
	if (accept==1)
	{
		document.write ('我的文章列表--已审核的文章');
	}else{
		document.write ('我的文章列表--未审核的文章');
	}
	document.write ('</th>');
	document.write ('<th valign=middle noWrap>审核</th>');
	document.write ('<th valign=middle noWrap>发布日期</th>');
	document.write ('<th valign=middle noWrap>管理操作</th>');
	document.write ('</tr>');
}
function myuser_articlelist_not(){
	document.write ('<tr>');
	document.write ('<td class=Usertablerow1 align=center valign=middle colspan=5>没有找到任何文章。</td>');
	document.write ('</tr>');
}
function myuser_articlelist_loop(channelid,ArticleID,accept,title,classname,dated,hits,style){
	var tablebody;
	if (style==1)
	{
		tablebody="Usertablerow1";
	}else{
		tablebody="Usertablerow2";
	}
	document.write ('<tr>');
	document.write ('<td class='+tablebody+' valign=middle>['+classname+'] ');
	document.write ('<a href="articlelist.asp?action=view&channelid='+channelid+'&ArticleID='+ArticleID+'">'+title+'</a></td>');
	document.write ('<td class='+tablebody+' align=center valign=middle>');
	if (accept==1)
	{
		document.write ('<font color=blue>已审</font>');
	}else{
		document.write ('<font color=red>未审</font>');
	}
	document.write ('</td>');
	document.write ('<td class='+tablebody+' align=center valign=middle>'+dated+'</td>');
	document.write ('<td class='+tablebody+' align=center valign=middle>');
	document.write ('<a href="articlelist.asp?action=edit&channelid='+channelid+'&ArticleID='+ArticleID+'">修改</a> | ');
	document.write ('<a href="articlelist.asp?action=del&channelid='+channelid+'&ArticleID='+ArticleID+'" onClick="return confirm(\'确定要删除此文章吗？\')">删除</a>');
	document.write ('</td>');
	document.write ('</tr>');
}
-->
</script>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr height=20 align=center>
		<td class=Usertablerow2 colspan=5><a href="?ChannelID=<%=ChannelID%>&Accept=1">已审核的文章</a> | 
		<a href="?ChannelID=<%=ChannelID%>">未审核的文章</a> | 
		<a href="articlepost.asp?ChannelID=<%=ChannelID%>">发布新的文章</a> </td>
	</tr>
<%
	Dim CurrentPage,page_count,totalrec,Pcount,maxperpage
	Dim isAccept,s
	maxperpage = 20 '###每页显示数
	
	If Trim(Request("Accept")) <> "" Then
		isAccept = 1
	Else
		isAccept = 0
	End If
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CInt(CurrentPage)
	End If
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	Response.Write "<script>myuser_articlelist_top("& isAccept &")</script>" & vbNewLine
	totalrec = enchiasp.Execute("SELECT COUNT(ArticleID) FROM ECCMS_Article WHERE ChannelID = " & ChannelID & " And username='" & enchiasp.MemberName & "' And isAccept="& isAccept)(0)
	Pcount = CInt(totalrec / maxperpage)  '得到总页数
	If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT A.ArticleID,A.title,A.WriteTime,A.AllHits,A.isAccept,C.ClassName FROM [ECCMS_Article] A INNER JOIN [ECCMS_Classify] C on A.ClassID=C.ClassID WHERE A.ChannelID = " & ChannelID & "  And A.username='" & enchiasp.MemberName & "' And isAccept="& isAccept &" ORDER BY A.isTop DESC, A.WriteTime DESC ,A.ArticleID DESC"
	Rs.Open SQL, Conn, 1, 1
	If Rs.EOF And Rs.BOF Then
		Response.Write "<script>myuser_articlelist_not()</script>" & vbNewLine
	Else
		If Pcount > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		page_count = 0
		If Rs.EOf Then Exit Sub
		SQL = Rs.GetRows(maxperpage)
		For i=0 To Ubound(SQL,2)
			If (i mod 2) = 0 Then
				s = 2
			Else
				s = 1
			End If
			Response.Write VbCrLf
			Response.Write "<script>myuser_articlelist_loop("
			Response.Write ChannelID
			Response.Write ","
			Response.Write SQL(0,i)
			Response.Write ","
			Response.Write SQL(4,i)
			Response.Write ",'"
			Response.Write EncodeJS(SQL(1,i))
			Response.Write "','"
			Response.Write EncodeJS(SQL(5,i))
			Response.Write "','"
			Response.Write FormatDated(SQL(2,i),4)
			Response.Write "',"
			Response.Write SQL(3,i)
			Response.Write ","
			Response.Write s
			Response.Write ")</script>"
			Response.Write VbCrLf
			page_count = page_count + 1
		Next
		SQL=Null
	End If
	Rs.Close:Set Rs = Nothing
	Response.Write "<tr align=right><td class=Usertablerow2 colspan=5>"
	Response.Write ShowPages(CurrentPage,Pcount,totalrec,maxperpage,"&ChannelID="& ChannelID &"&Accept="& Request("Accept"))
	Response.Write "</td></tr>" & vbNewLine
	Response.Write "</table>"
End Sub
Function EncodeJS(str)
	str = enchiasp.HtmlEncode(str)
	str = Replace(Replace(Replace(Replace(str,"\","\\"),"'","\'"),VbCrLf,"\n"),Chr(13),"")
	EnCodeJs = str
End Function
Sub DeleteArticle()
	If CInt(GroupSetting(8)) = 0 Then
		ErrMsg = ErrMsg + "<li>对不起！您没有删除文章的权限，如需要该权限请联系管理员。</li>"
		Founderr = True
		Exit Sub
	End If
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If Request("ArticleID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！您选择了错误的系统参数。</li>"
		Exit Sub
	End If
	SQL = "SELECT isAccept FROM ECCMS_Article WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And isAccept=0 And ArticleID=" & CLng(Request("ArticleID"))
	Set Rs = enchiasp.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！此文章已经通过审核,您没有权限删除,如有什么问题请联系管理员。</li>"
		Set Rs = Nothing
		Exit Sub
	Else
		enchiasp.Execute("DELETE FROM ECCMS_Article WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And isAccept=0 And ArticleID=" & CLng(Request("ArticleID")))
	End If
	Set Rs = Nothing
	Response.Redirect  Request.ServerVariables("HTTP_REFERER")
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
	'response.write sql
	'response.end
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
	  <td align="center" class="usertablerow1">作者：<%=Rs("Author")%> 　来源于：<%=Rs("ComeFrom")%> 　发布时间：<%=Rs("WriteTime")%> 　发布人：<font color=blue><%=Rs("username")%></font></td>
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
Sub SaveArticle()
	If CInt(GroupSetting(8)) = 0 Then
		ErrMsg = ErrMsg + "<li>对不起！您没有修改文章的权限，如需要该权限请联系管理员。</li>"
		Founderr = True
		Exit Sub
	End If
	Dim TextContent,ForbidEssay,isAccept,i,ArticleID
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
	If Trim(Request.Form("content")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>文章内容不能为空！</li>"
	End If
	If Len(Request.Form("content")) > CLng(GroupSetting(16)) Then
		ErrMsg = ErrMsg + "<li>文章内容不能大于" & GroupSetting(16) & "字符！</li>"
		Founderr = True
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
	SQL = "SELECT * FROM ECCMS_Article WHERE username='" & enchiasp.MemberName & "' And ArticleID=" & CLng(Request("ArticleID"))
	Rs.Open SQL,Conn,1,3
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = enchiasp.CheckNumeric(Request.Form("ClassID"))
		Rs("title") = enchiasp.ChkFormStr(Request.Form("title"))
		Rs("content") = Html2Ubb(TextContent)
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Author") = enchiasp.ChkFormStr(Request.Form("Author"))
		Rs("ComeFrom") = enchiasp.ChkFormStr(Request.Form("ComeFrom"))
		Rs("star") = enchiasp.ChkNumeric(Request.Form("star"))
		Rs("username") = Trim(enchiasp.MemberName)
		Rs("BriefTopic") = enchiasp.ChkNumeric(Request.Form("BriefTopic"))
		Rs("ImageUrl") = enchiasp.ChkFormStr(Request.Form("ImageUrl"))
		Rs("UploadImage") = enchiasp.ChkFormStr(Request.Form("UploadFileList"))
		Rs("isAccept") = isAccept
		Rs("ForbidEssay") = ForbidEssay
	Rs.update
	ArticleID = Rs("ArticleID")
	Rs.Close:Set Rs = Nothing
	Call Returnsuc("<li>恭喜您！修改文章成功。</li><li><a href=?action=view&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & ">点击此处查看该文章</a></li>")
End Sub
Function InitSelect(UploadFileList, ImageUrl)
	Dim i
	InitSelect = "<select name='ImageFileList' onChange=""ImageUrl.value=this.value;"">"
	InitSelect = InitSelect & "<option value=''>不选择首页推荐图片</option>"
	If Not IsNull(UploadFileList) Then
		UploadFileList = Split(UploadFileList, "|")
		For i = 0 To UBound(UploadFileList)
			If UploadFileList(i) <> "" Then
				InitSelect = InitSelect & "<option value=""" & UploadFileList(i) & """"
				If UploadFileList(i) = ImageUrl Then
					InitSelect = InitSelect & " selected"
				End If
				InitSelect = InitSelect & ">" & UploadFileList(i) & "</option>"
			End If
		Next
	End If
	InitSelect = InitSelect & "</select>"
End Function
Sub EditArticle()
	Dim ClassID
	If CInt(GroupSetting(8)) = 0 Then
		ErrMsg = ErrMsg + "<li>对不起！您没有修改文章的权限，如需要该权限请联系管理员。</li>"
		Founderr = True
		Exit Sub
	End If
	If Founderr = True Then Exit Sub
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
	SQL = "SELECT * FROM ECCMS_Article WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And ArticleID=" & Request("ArticleID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！没有找到任何文章。或者您选择了错误的系统参数！</li>"
		Exit Sub
	End If
	ClassID = Rs("ClassID")
	If Rs("isAccept") <> 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>此文章已经通过审核,您没有权限修改,如有什么问题请联系管理员。</li>"
		Set Rs = Nothing
		Exit Sub
	End If
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
	<form method=Post name="myform" action="Articlelist.Asp">
	<input type="Hidden" name="action" value="save">
	<input type=hidden name=ChannelID value='<%=ChannelID%>'>
	<input type=hidden name=ArticleID value='<%=Rs("ArticleID")%>'>
        <tr>
          <td width="15%" align="right" nowrap class="usertablerow2"><strong>所属分类</strong></td>
          <td width="85%" class="usertablerow1">
<%
	Dim sClassSelect
	Response.Write "<select name=""ClassID"" id=""ClassID"">"
	sClassSelect = enchiasp.LoadSelectClass(ChannelID)
	sClassSelect = Replace(sClassSelect, "{ClassID=" & ClassID & "}", "selected")
	Response.Write sClassSelect
	Response.Write "</select>"
%>
	  </td>
        </tr>
        <tr>
          <td align="right" noWrap class="usertablerow2"><strong>文章标题</strong></td>
          <td class="usertablerow1"><select name="BriefTopic" id="BriefTopic">
            <option value="0">选择话题</option>
			<option value="0"<%If Rs("BriefTopic") = 0 Then Response.Write (" selected")%>>选择话题</option>
			<option value="1"<%If Rs("BriefTopic") = 1 Then Response.Write (" selected")%>>[图文]</option>
			<option value="2"<%If Rs("BriefTopic") = 2 Then Response.Write (" selected")%>>[组图]</option>
			<option value="3"<%If Rs("BriefTopic") = 3 Then Response.Write (" selected")%>>[新闻]</option>
			<option value="4"<%If Rs("BriefTopic") = 4 Then Response.Write (" selected")%>>[推荐]</option>
			<option value="5"<%If Rs("BriefTopic") = 5 Then Response.Write (" selected")%>>[注意]</option>
			<option value="6"<%If Rs("BriefTopic") = 6 Then Response.Write (" selected")%>>[转载]</option>
          </select> <input name="title" type="text" id="title" size="60" value="<%=Rs("title")%>"> 
          <font color=red>*</font></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>相关文章</strong></td>
          <td class="usertablerow1"><input name="Related" type="text" id="Related" size="60" value="<%=Rs("Related")%>"> <font color=red>*</font></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>文章作者</strong></td>
          <td class="usertablerow1"><input name="Author" type="text" size="30" value="<%=Rs("Author")%>">
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
          <td class="usertablerow1"><input name="ComeFrom" type="text" size="30" value="<%=Rs("ComeFrom")%>">
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
          <td class="usertablerow1"><textarea name='content' id='content' style='display:none'><%=Server.HTMLEncode(Rs("content"))%></textarea>
		<script Language=Javascript src="../editor/post.js"></script></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>首页图片</strong></td>
          <td class="usertablerow1"><input name="ImageUrl" type="text" id="ImageUrl" size="60" value="<%=Rs("ImageUrl")%>">
			<input type=hidden name=UploadFileList id=UploadFileList onchange="doChange(this,document.myform.ImageFileList)" value="<%=Rs("UploadImage")%>">
			<br>直接从上传图片中选择
			<%
			Response.Write InitSelect(Rs("UploadImage"),Rs("ImageUrl"))
			%>
			</td>
        </tr>
	<tr>
	  <td align="right" class="usertablerow2"><strong>文件上传</strong></td>
          <td class="usertablerow1"><iframe name="image1" frameborder=0 width='100%' height=55 scrolling=no src=upload.asp?ChannelID=<%=ChannelID%>></iframe></td>
        </tr>
        <tr>
          <td align="right" class="usertablerow2"><strong>文章星级</strong></td>
          <td class="usertablerow1"><select name="star">
          	<option value=5<%If Rs("star") = 5 Then Response.Write (" selected")%>>★★★★★</option>
          	<option value=4<%If Rs("star") = 4 Then Response.Write (" selected")%>>★★★★</option>
          	<option value=3<%If Rs("star") = 3 Then Response.Write (" selected")%>>★★★</option>
		<option value=2<%If Rs("star") = 2 Then Response.Write (" selected")%>>★★</option>
		<option value=1<%If Rs("star") = 1 Then Response.Write (" selected")%>>★</option>
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
          <td class="usertablerow1"><input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1"<%If Rs("ForbidEssay") <> 0 Then Response.Write (" checked")%>></td>
        </tr>
	<tr>
          <td align="right" class="usertablerow2"><strong>立即发布</strong></td>
          <td class="usertablerow1"><input name="isAccept" type="checkbox" id="isAccept" value="1" disabled> 是（<font color=blue>如果选中的话将直接发布，否则审核后才能发布。</font>）</td>
        </tr>
        <tr align="center">
          <td colspan="2" class="usertablerow2">
	  <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="返回上一页" class=Button>
	  <input type="button" name="Submit1" value="修改文章" class=Button onclick="doSubmit();"></td>
        </tr></form>
      </table></div>
<%
End Sub
%>
<!--#include file="foot.inc"-->












