<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
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
Call InnerLocation("用户短信服务")

Dim Rs,SQL,i,Action
Dim smsincept,smscontent,smstopic,sid,sendername,Chatloglist

If CInt(GroupSetting(22)) = 0 Then
	ErrMsg = ErrMsg + "<li>对不起！您没有使用短信服务的权限，如有什么问题请联系管理员。</li>"
	Founderr = True
End If
If Trim(Request("touser")) <> "" Then
	sendername = enchiasp.CheckbadStr(Request("touser"))
	smsincept =  enchiasp.CheckbadStr(Request("touser"))
Else
	sendername = enchiasp.CheckbadStr(Request("sender"))
End If
Chatloglist = ""
Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "del"
		Call DelMessage
	Case "alldel"
		Call DelAllMessage
	Case "save"
		Call SaveMessage
	Case "read"
		Call ReadMessage
	Case "outread"
		Call ReadMessage
	Case "new"
		Call SendMessage
	Case "fw"
		Call SendMessage
	Case "删除收件箱"
		Call Delinbox
	Case "清空收件箱"
		Call DelAllinbox
	Case "删除发件箱"
		Call DelSendbox
	Case "清空发件箱"
		Call DelAllSendbox
	Case Else
		ErrMsg = ErrMsg + "<li>错误的系统参数~!</li>"
		Founderr = True
End Select
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If

Sub SendMessage()
	Call UserMessage
	If Founderr = True Then Exit Sub
%>
<script language=JavaScript>
var _maxCount = '<%=CLng(GroupSetting(23))%>';
function DoTitle(addTitle) {  
 var revisedTitle;  
 var currentTitle = document.myform.incept.value;
 if(currentTitle=="") revisedTitle = addTitle; 
 else { 
  var arr = currentTitle.split(","); 
  for (var i=0; i < arr.length; i++) { 
   if( addTitle.indexOf(arr[i]) >=0 && arr[i].length==addTitle.length ) return; 
  } 
  revisedTitle = currentTitle+","+addTitle; 
 } 

 document.myform.incept.value=revisedTitle;  
 document.myform.incept.focus(); 
 return; 
} 

function doSubmit(){
	if (document.myform.incept.value==""){
		alert("收件人不能为空！");
		return false;
	}
	if (document.myform.topic.value==""){
		alert("短信标题不能为空！");
		return false;
	}
	<%If CInt(GroupSetting(2)) = 1 Then%>
	if (document.myform.codestr.value==""){
		alert("请填写验证码！");
		return false;
	}
	<%End If%>
	myform.content1.value = getHTML(); 
	MessageLength = Composition.document.body.innerHTML.length;
	if(MessageLength < 2){
		alert("短信内容不能小于2个字符！");
		return false;
	}
	if(MessageLength > _maxCount){
		alert("短信的内容不能超过"+_maxCount+"个字符！");
		return false;
	}
	document.myform.Submit1.disabled = true;
	document.myform.submit();
}
</script>

<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<form name=myform method=post action="message.asp">
	<input type="hidden" name="action" value="save">
	<tr>
		<th colspan=2>站内短消息</th>
	</tr>
<%
	Call MessageTop
%>
	<tr>
		<td class=Usertablerow1>收件人</td>
		<td class=Usertablerow1><input type=text name="incept" value="<%=smsincept%>" size=50>
		<select name=friend onchange="DoTitle(this.options[this.selectedIndex].value)">
		<option selected value="">选择</option>
		<%=Option_Friend%> 
		</select></td>
	</tr>
	<tr>
		<td class=Usertablerow1>标题</td>
		<td class=Usertablerow1><input type="text" name="topic" maxlength="70" size="70" value="<%=smstopic%>"></td>
	</tr>
<%
	If CInt(GroupSetting(2)) = 1 Then
%>
	<tr>
		<td class=Usertablerow1>验证码</td>
		<td class=Usertablerow1><input type="text" name="codestr" maxlength="4" size="4">&nbsp;<img src="../inc/getcode.asp"></td>
	</tr>
<%
	End If
%>
	<tr>
		<td class=Usertablerow1 noWrap>短信内容</td>
		<td class=Usertablerow1><textarea name='content1' id='content1' style='display:none'><%=Server.HTMLEncode(smscontent)%></textarea>
		<script Language=Javascript src="../editor/editor1.js"></script></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 colspan=2><b>说明：</b>标题最多50个字符，内容最多<%=CLng(GroupSetting(23))%>个字符。</td>
	</tr>
	<tr align=center height=20>
		<td class=Usertablerow2 colspan=2><input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="返回上一页" class=Button>&nbsp;
		<input type="reset" name="submit2" value=" 清除 " class=button>&nbsp;
<SCRIPT LANGUAGE="JavaScript">
<!--
var reaction='<%=enchiasp.CheckStr(Request("reaction"))%>';
var action='new';
if (action=='new')
{
if (reaction=='chatlog')
{
document.write ('<input class=button type=button value="关闭聊天记录" name="chatlog" onclick="location.href=\'?action=new&sid=<%=Request("sid")%>&touser=<%=sendername%>\'">');
}
else{
document.write ('<input class=button type=button value="查看聊天记录" name="chatlog" onclick="location.href=\'?action=new&sid=<%=Request("sid")%>&touser=<%=sendername%>&reaction=chatlog\'">');
}
}
//-->
</SCRIPT>
		<input type="button" name="Submit1" value=" 发送 " onclick="doSubmit();" class=button></td>
	</tr>
<SCRIPT LANGUAGE="JavaScript">
<!--
var reaction='<%=enchiasp.CheckStr(Request("reaction"))%>';
var chatloglist='<%=Chatloglist%>';
var myname='<%=enchiasp.MemberName%>';
var action='new';
if (action=='new')
{
if (reaction=='chatlog')
{
	document.write ('<tr>');
	document.write ('<th colspan=2>我与<%=sendername%>的聊天记录</th>');
	document.write ('</tr>');
	if (myname=='')
	{
		document.write ('<tr>');
		document.write ('<td class=Usertablerow1 colspan=2>自己跟自己的聊天记录没什么好看的：）</td>');
		document.write ('</tr>');
	}
	else{
		document.write (chatloglist);
	}
}
}
//-->
</SCRIPT>
	</form>
</table>
<%
End Sub

Sub MessageTop()
%>
	<tr align=center height=20>
		<td class=Usertablerow1 colspan=2><a href="message.asp?action=del&sid=<%=Request("sid")%>" onclick=showClick('您确定要删除此短信吗?')><img src="images/m_delete.gif" border=0 alt="删除消息"></a> &nbsp; 
		<a href="message.asp?action=new"><img src="images/m_write.gif" border=0 alt="发送消息"></a> &nbsp;
		<a href="message.asp?action=new&touser=<%=sendername%>&sid=<%=Request("sid")%>"><img src="images/replypm.gif" border=0 alt="回复消息"></a>&nbsp;
		<a href="message.asp?action=fw&sid=<%=Request("sid")%>"><img src="images/m_fw.gif" border=0 alt=转发消息></a></td>
	</tr>
<%
End Sub

Sub ReadMessage()
	If Founderr = True Then Exit Sub
	If Action = "outread" Then
		Set Rs = enchiasp.Execute("select * from ECCMS_Message where sender='"&enchiasp.MemberName&"' And delSend=0 And id="& CLng(Request("sid")))
	Else
		Set Rs = enchiasp.Execute("select * from ECCMS_Message where (incept='"&enchiasp.MemberName&"' Or flag=1) And id="& CLng(Request("sid")))
	End If
	If Rs.BOF And Rs.EOF Then
		ErrMsg = ErrMsg + "<li>错误的系统参数~！</li>"
		Founderr = True
		Set Rs = Nothing
		Exit Sub
	End If
	Dim smsnumber
	If Rs("isRead") = 0 And Action="read" Then
		smsnumber = newincept(enchiasp.membername) - 1
		if smsnumber < 0 Then smsnumber = 0
		SQL = "Update ECCMS_User Set usermsg=" & smsnumber & " where username='"&enchiasp.membername&"'"
		enchiasp.Execute(SQL)
		if Rs("flag") = 0 Then
			SQL = "Update ECCMS_Message Set isRead=1 where id="& CLng(Request("sid"))
			enchiasp.Execute(SQL)
		End If
	End If
%>
<table cellspacing=1 align=center cellpadding=3 bgcolor=#cccccc border=0 class=Usertableborder>
	<tr>
		<th>阅读短消息</th>
	</tr>
<%
	Call MessageTop
%>
	<tr height=20>
		<td class=Usertablerow2>　在<b><%=Rs("SendTime")%></b>，
<%
	If Action = "outread" Then
		Response.Write "您给<b>" & Server.HTMLEncode(Rs("incept")) & "</b>发送的消息！"
	Else
		Response.Write "<b>" & Server.HTMLEncode(Request("sender")) & "</b>给您发送的消息！"
	End If
%>
		</td>
	</tr>
	<tr>
		<td class=Usertablerow1><b>短信标题：</b><%=Rs("title")%><hr size=1><%=ubbcode(Rs("content"))%></td>
	</tr>
	<tr align=center height=20>
		<td class=Usertablerow2 colspan=2><input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="返回上一页" class=Button>&nbsp;</td>
	</tr>
</table>
<%
	Set Rs = Nothing
End Sub
Sub UserMessage()
	If Founderr = True Then Exit Sub
	If Not IsNumeric(Request("sid")) And Trim(Request("sid")) <> "" Then
		ErrMsg = ErrMsg + "错误的系统参数!ID请输入整数"
		Founderr = True
		Exit Sub
	End If
	If Trim(Request("sid")) <> "" Then
		sid = CLng(Request("sid"))
	End If
	If Action = "fw" And IsNumeric(Request("sid"))  Then
		Set Rs = enchiasp.Execute("select * from ECCMS_Message where (sender='"&enchiasp.MemberName&"' Or incept='"&enchiasp.MemberName&"') And id="& CLng(Request("sid")))
		If Rs.BOF And Rs.EOF Then
			ErrMsg = ErrMsg + "<li>错误的系统参数~！</li>"
			Founderr = True
			Set Rs = Nothing
			Exit Sub
		End If
		smsincept = ""
		smscontent = "=================== 下面是转发信息 =================== <br>" & Rs("content") & "<br>====================================================<br>"
		smstopic = "FW：" & Rs("title")
		sendername = Rs("sender")
		Set Rs = Nothing
	End If
	If Trim(Request("touser")) <> "" And Request("sid") <> "" Then
		Set Rs = enchiasp.Execute("select * from ECCMS_Message where id="& CLng(Request("sid")) &" And incept='"&enchiasp.MemberName&"'")
		If Rs.BOF And Rs.EOF Then
			ErrMsg = ErrMsg + "<li>错误的系统参数~！</li>"
			Founderr = True
			Set Rs = Nothing
			Exit Sub
		End If
		smsincept = Rs("incept")
		smscontent = "============在 " & Rs("SendTime") & " 您来信中写道：============<br>" & Rs("content") & "<br>======================================================<br>"
		smstopic = "RW：" & Rs("title")
		sendername = Rs("sender")
		Set Rs = Nothing
	End If
	Dim Touser,temp_chat1,temp_chat2
	If Request("reaction")="chatlog" Then
		Touser=enchiasp.CheckStr(Request("touser"))
		SQL="SELECT top 30 sender,incept,title,content,sendtime FROM ECCMS_Message WHERE ((sender='"&enchiasp.MemberName&"' And incept='"&Touser&"') or (sender='"&Touser&"' And incept='"&enchiasp.MemberName&"')) And delSend=0 ORDER BY sendtime DESC"
		Set Rs=enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			Chatloglist="<tr><td class=Usertablerow1 colspan=2>还没有任何聊天记录！</td></tr>"
		Else
			SQL=Rs.GetRows(-1)
			Rs.close:Set Rs=nothing

			For i=0 to Ubound(SQL,2)
				chatloglist=chatloglist & "<tr><td class=Usertablerow2 height=25 colspan=2>"
				If Trim(SQL(0,i))=enchiasp.MemberName Then
					temp_chat1 = "在" & SQL(4,i)
					temp_chat1 = temp_chat1 & "，您发送此消息给" & enchiasp.HtmlEncode(SQL(1,i))
					chatloglist=chatloglist & temp_chat1
				Else
					temp_chat2 = "在" & SQL(4,i) & "，"
					temp_chat2 = temp_chat2 & enchiasp.HtmlEncode(SQL(0,i)) & "给您发送的消息！"
					chatloglist=chatloglist & temp_chat2
				End If
				chatloglist=chatloglist & "</td></tr><tr><td class=Usertablerow1 valign=top align=left colspan=2><b>消息标题："
				chatloglist=chatloglist & enchiasp.HtmlEncode(SQL(2,i))
				chatloglist=chatloglist & "</b><hr size=1>"
				chatloglist=chatloglist & UbbCode(SQL(3,i))
				chatloglist=chatloglist & "</td></tr>"
			Next
		End If
	End If
End Sub
Sub DelMessage()
	If Founderr = True Then Exit Sub
	If Not IsNumeric(Request("sid")) Then
		ErrMsg = ErrMsg + "<li>对不起！错误的系统参数。</li>"
		Founderr = True
		Exit Sub
	End If
	SQL="SELECT incept FROM ECCMS_Message WHERE (sender='"&enchiasp.MemberName&"' Or incept='"&enchiasp.MemberName&"') And id="& Request("sid")
	Set Rs=enchiasp.Execute(SQL)
	If Rs.EOF And Rs.BOF Then
		ErrMsg = ErrMsg + "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Exit Sub
		Set Rs = Nothing
	Else
		If Rs(0) = enchiasp.MemberName Then
			enchiasp.Execute("Delete From ECCMS_Message where flag=0 And incept='"&enchiasp.MemberName&"' And id="& Request("sid"))
		Else
			enchiasp.Execute("Update ECCMS_Message Set delsend=1 where sender='"&enchiasp.MemberName&"' And id="& Request("sid"))
		End If
	End If
	Rs.Close:Set Rs = Nothing
	Call Returnsuc("<li>删除短消息完成！</li>")
End Sub
Sub DelAllMessage()
	If Founderr = True Then Exit Sub
	enchiasp.Execute("Delete From ECCMS_Message where flag=0 And incept='"&enchiasp.MemberName&"'")
	enchiasp.Execute("Update ECCMS_Message Set delsend=1 where sender='"&enchiasp.MemberName&"'")
	Call Returnsuc("<li>您的短消息已经全部清除！</li>")
End Sub
'================================================
' 函数名：Option_Friend
' 作  用：用户好友下拉名单
'================================================
Function Option_Friend()
	DIM i
	SQL = "select friend from ECCMS_Friend where grouping<>2 And userid="& enchiasp.memberid &" order by addtime desc"
	Set Rs = enchiasp.Execute(Sql)
	If Not Rs.EOF Then
		SQL = Rs.GetRows(-1)
		Rs.Close:Set Rs=Nothing
	End if
	If IsArray(SQL) Then
		For i=0 To Ubound(SQL,2)
		Option_Friend = Option_Friend & "<option value="""& SQL(0,i) &""">"& SQL(0,i) &"</option> "
		Next
	Else
		Option_Friend = ""
	End If
End Function
'================================================
' 函数名：newincept
' 作  用：统计短信
'================================================
Function newincept(iusername)
	Dim Rs
	Rs = enchiasp.Execute("Select Count(id) from ECCMS_Message where isRead=0 And flag=0 And incept='"& iusername &"'")
	newincept = Rs(0)
	Set Rs=Nothing
	If IsNull(newincept) Then newincept = 0
End Function
'================================================
' 函数名：ChkHateName
' 作  用：黑名单验证
'================================================
Function ChkHateName(sName)
	DIM SQL,Rs
	ChkHateName = False
	SQL="Select friend From ECCMS_Friend Where (userid="& enchiasp.memberid &" Or username='"& sName &"') And grouping=2"
	Set Rs = enchiasp.Execute(SQL)
	If Not Rs.EOF Then
		SQL=Rs.GetString(,, ",", "", "")
		Rs.Close:Set Rs=Nothing
		If Instr(SQL,sName) Or Instr(SQL,enchiasp.membername) Then ChkHateName = True
	End If
End Function
'================================================
' 函数名：CheckID
' 作  用：验证短信ID
'================================================
Function CheckID(CHECK_ID)
	Dim Delid,Fixid
	CheckID=True
	Delid=replace(CHECK_ID,"'","")
	Delid=replace(Delid,";","")
	Delid=replace(Delid,"--","")
	Delid=replace(Delid,")","")
	Fixid=replace(Delid,",","")
	Fixid=Trim(replace(fixid," ",""))
	If Delid="" or isnull(Delid) Then  CheckID=False
	If Not IsNumeric(fixid) Then CheckID=False
End Function
'================================================
' 过程名：SaveMessage
' 作  用：保存短消息
'================================================
Sub SaveMessage()
	Dim strIncept,strContent,strTitle,InceptName,n
	If CLng(UserToday(4)) => CLng(GroupSetting(29)) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>您每天最多只能发布<font color=red><b>" & GroupSetting(29) & "</b></font>篇文章，如果还要继续发布请明天再来吧！</li>"
	End If
	If Trim(Request.Form("incept")) = "" Then
		ErrMsg = ErrMsg + "<li>请填写收件人姓名！</li>"
		Founderr = True
	Else
		strIncept = enchiasp.CheckbadStr(Request.Form("incept"))
		strIncept = split(strIncept,",")
	End If
	If Trim(Request.Form("topic")) = "" Then
		ErrMsg = ErrMsg + "<li>请填写短信标题！</li>"
		Founderr = True
	Else
		strTitle = Left(enchiasp.ChkFormStr(Request.Form("topic")),50)
	End If
	If Trim(Request.Form("content1")) = "" Then
		ErrMsg = ErrMsg + "<li>请填写短信内容！</li>"
		Founderr = True
	Else
		strContent = Html2Ubb(Request.Form("content1"))
	End If
	If Len(Request.Form("content1")) > CLng(GroupSetting(23)) Then
		ErrMsg = ErrMsg + "<li>短信内容不能大于" & GroupSetting(23) & "字符！</li>"
		Founderr = True
	End If
	If CInt(GroupSetting(2)) = 1 Then
		If Not enchiasp.CodeIsTrue() Then
			ErrMsg = ErrMsg + "<meta http-equiv=refresh content=""2;URL="&Request.ServerVariables("HTTP_REFERER")&"""><li>验证码校验失败，请返回刷新页面再试。两秒后自动返回</li>"
			Founderr = True
		End If
		Session("GetCode") = ""
	End If
	If Founderr = True Then Exit Sub
	On Error Resume Next
	Call PreventRefresh  '防刷新
	n=0
	For i = 0 To Ubound(strIncept)
		If i >= 5 Then Exit For
		n = n + 1
		InceptName = Trim(strIncept(i))
		SQL = "select username from [ECCMS_User] where username='"&InceptName&"'"
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			ErrMsg = ErrMsg + "<li>没有找到<font color=red>" & InceptName & "</font>这个用户，短信发送不成功~！</li>"
			Founderr = True
			Rs.Close:Set Rs = Nothing
			Exit Sub
		Else
			InceptName = Rs(0)
		End If
		Rs.Close:Set Rs = Nothing
		If ChkHateName(InceptName) Then
			ErrMsg = ErrMsg + "由于对方<font color=red>" & InceptName & "</font>已将你列入黑名单，或<font color=red>" & InceptName & "</font>存在你的黑名单中，因此短信发送被终止！"
			Founderr = True
			Exit Sub
		Else
			SQL = "Insert into ECCMS_Message (sender,incept,title,content,flag,SendTime,isRead,delSend) values ('"& enchiasp.membername &"','"& InceptName &"','"& strTitle &"','"& strContent &"',0,"& NowString &",0,0) "
			enchiasp.Execute(SQL)
			SQL = "Update ECCMS_User Set usermsg=usermsg+1 where username='"&InceptName&"'"
			enchiasp.Execute(SQL)
		End If
		
	Next
	Dim strUserToday
	strUserToday = UserToday(0) &","& UserToday(1) &","& UserToday(2) &","& UserToday(3) &","& UserToday(4)+n &","& UserToday(5)
	UpdateUserToday(strUserToday)
	Call Returnsuc("<li>恭喜您！发送短信成功。</li>")
End Sub
'删除收件箱
Sub Delinbox()
	If Not CheckID(Request("id")) Then
		ErrMsg = ErrMsg + "<li>错误的系统参数！</li>"
		Founderr = True
	End If
	If Founderr = True Then Exit Sub
	enchiasp.Execute("Delete From ECCMS_Message where flag=0 And incept='"&enchiasp.MemberName&"' And id in (" & enchiasp.CheckBadstr(Request("id")) & ")")
	Call Returnsuc("<li>删除收件箱中的短信成功！</li>")
End Sub
'清空收件箱
Sub DelAllinbox()
	If Founderr = True Then Exit Sub
	enchiasp.Execute("Delete From ECCMS_Message where flag=0 And incept='"&enchiasp.MemberName&"'")
	Call Returnsuc("<li>您的收件箱已成功清空！</li>")
End Sub
'删除发件箱
Sub DelSendbox()
	If Not CheckID(Request("id")) Then
		ErrMsg = ErrMsg + "<li>错误的系统参数！</li>"
		Founderr = True
	End If
	If Founderr = True Then Exit Sub
	enchiasp.Execute("Update ECCMS_Message Set delsend=1 where sender='"&enchiasp.MemberName&"' And id in (" & enchiasp.CheckBadstr(Request("id")) & ")")
	Call Returnsuc("<li>删除发件箱中的短信成功！</li>")
End Sub
'清空发件箱
Sub DelAllSendbox()
	If Founderr = True Then Exit Sub
	enchiasp.Execute("Update ECCMS_Message Set delsend=1 where sender='"&enchiasp.MemberName&"'")
	Call Returnsuc("<li>您的发件箱已成功清空！</li>")
End Sub
%>
<!--#include file="foot.inc"-->




















