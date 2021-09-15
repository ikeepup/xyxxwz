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
Call InnerLocation("好友管理")

Dim Rs,SQL,i

If CInt(GroupSetting(4)) = 0 Then
	ErrMsg = ErrMsg + "<li>对不起！您没有使用好友管理的权限，如有什么问题请联系管理员。</li>"
	Founderr = True
End If
Dim Action:Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "add"
		Call FriendAdd
	Case "移动"
		Call MoveFriend
	Case "删除"
		Call FriendDel
	Case "清空好友"
		Call DelAllFriend
	Case Else
		Call showmain
End Select
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
Sub showmain()
	If Founderr = True Then Exit Sub
	Dim PageListNum,totalrec,Pcount,CurrentPage,page_count
	PageListNum = 20
	page_count = 0
	If Not IsNumeric(Request("page")) And Trim(Request("page")) <> "" Then
		Response.Write ("错误的系统参数!请输入整数")
		Response.End
	End If
	If Not IsEmpty(Request("page")) And Trim(Request("page")) <> "" Then
		CurrentPage = Clng(Request("page"))
	Else
		CurrentPage = 1
	End If
	totalrec = enchiasp.Execute("Select Count(FriendID) from ECCMS_Friend where username='"& enchiasp.CheckStr(enchiasp.membername) &"'")(0)
	If totalrec Mod PageListNum = 0 Then
		Pcount = totalrec \ PageListNum
	Else
		Pcount = totalrec \ PageListNum + 1
	End If
	If CurrentPage > Pcount Then CurrentPage = Pcount
	If CurrentPage < 1 Then CurrentPage = 1
%>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr>
		<th colspan=6>>> 我的好友 <<</th>
	</tr>
	<form action="friend.asp" method=post name=inbox>
	<tr>
		<td width="15%" align=center class=Usertablerow2><b class=userfont2>组 别</b></td>
		<td width="25%" align=center class=Usertablerow2><b class=userfont2>用户名</b></td>
		<td width="30%" align=center class=Usertablerow2><b class=userfont2>邮 箱</b></td>
		<td width="10%" align=center class=Usertablerow2><b class=userfont2>OICQ</b></td>
		<td width="10%" align=center class=Usertablerow2><b class=userfont2>发短信</b></td>
		<td width="10%" align=center class=Usertablerow2><b class=userfont2>操 作</b></td>
	</tr>
<%
	Set Rs=Server.Createobject("adodb.recordset")
	SQL = "select F.FriendID,F.userid,F.Friend,F.grouping,U.usermail,U.HomePage,U.oicq From [ECCMS_Friend] F inner join [ECCMS_User] U on F.Friend=U.username where F.userid="&enchiasp.memberid
	SQL = SQL+" order by F.addtime desc"
	Rs.Open SQL,Conn,1,1
	If Rs.EOF And Rs.BOF Then
		Rs.Close:Set Rs = Nothing
	Else
		Rs.Move (CurrentPage - 1) * Cint(PageListNum)
		SQL = Rs.GetRows(PageListNum)
		Rs.Close:Set Rs = Nothing
		For i=0 To Ubound(SQL,2)
%>
	<tr>
		<td align=center class=Usertablerow1><b class=userfont2><%
		If CInt(SQL(3,i)) = 0 Then
			Response.Write "陌生人"
		ElseIf CInt(SQL(3,i)) = 1 Then
			Response.Write "我的好友"
		ElseIf CInt(SQL(3,i)) = 2 Then
			Response.Write "黑名单"
		Else
			Response.Write "黑名单"
		End If
		%></b></td>
		<td align=center class=Usertablerow1><a href="dispuser.asp?name=<%=SQL(2,i)%>" target=_blank title="浏览 <%=SQL(2,i)%> 的个人资料"><%=SQL(2,i)%></a></td>
		<td align=center class=Usertablerow1><a href="mailto:<%=SQL(4,i)%>"><%=SQL(4,i)%></a></td>
		<td align=center class=Usertablerow1><a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<%=SQL(6,i)%>" title="<%=SQL(2,i)%> 的 Oicq:<%=SQL(6,i)%>" target=_blank><img src=images/oicq.gif border=0></a></td>
		<td align=center class=Usertablerow1><a href="message.asp?action=new&touser=<%=SQL(2,i)%>" title="给 <%=SQL(2,i)%> 发短信"><img src=images/message.gif border=0></a></td>
		<td align=center class=Usertablerow1><input type=checkbox name=id value="<%=SQL(0,i)%>"></td>
	</tr>
<%
			page_count = page_count+1
		Next
	End If
%>
	<tr>
		<td colspan=6 align=center class=Usertablerow1><%Response.Write ShowPages (CurrentPage,Pcount,totalrec,PageListNum,"")%></td>
	</tr>
	<tr>
		<td colspan=6 align=center class=Usertablerow2><input type=checkbox name=chkall value=on onclick="CheckAll2(this.form)">选中所有显示记录&nbsp;
		<select name="grouping">
		<option value="" selected>批量移动到...</option>
		<option value="0" >陌生人</option>
		<option value="1" >我的好友</option>
		<option value="2" >黑名单</option>
		</select>&nbsp;
		<input type=submit name=action onclick="{if(confirm('确定移动选定的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="移动" class=button>&nbsp;
		<input type=button name=action onclick="showsub('addfriend')" value="添加好友" class=button>&nbsp;
		<input type=submit name=action onclick="{if(confirm('确定删除选定的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="删除" class=button>&nbsp;
		<input type=submit name=action onclick="{if(confirm('确定清除所有的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="清空好友" class=button></td>
	</tr></form>
</table>
<div id=addfriend style="display:none">
<br style="overflow: hidden; line-height: 10px">
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr>
		<th>>> 添加好友 <<</th>
	</tr>
	<form name=myform method=post action=?action=add>
	<tr>
		<td align=center class=Usertablerow1><b class=userfont2>好友：</b><input type="text" name="friend" size=45>
		<b class=userfont2>组别：</b><select name="grouping">
		<option value="0" selected>请选择....</option>
		<option value="0" >陌生人</option>
		<option value="1" >我的好友</option>
		<option value="2" >黑名单</option>
		</select>
		<input type=submit value="添加" class=button>&nbsp;<input type="reset" name="Clear" value="清除" class=button><br>
		<div><b>注意：</b><%If CLng(GroupSetting(6)) <> 0 Then%>你最多只能添加 <b class=userfont1><%=GroupSetting(6)%></b> 位好友，<%End If%>黑名单组，拒收所有来自黑名单的短信。 </div></td>
	</tr>
	</form>
</table>
</div>

<%
End Sub
'================================================
' 过程名：FriendDel
' 作  用：批量删除好友
'================================================
Sub FriendDel()
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	Dim FriendID,fixid
	FriendID = Replace(Request.form("id"),"'","")
	FriendID = Replace(FriendID,";","")
	FriendID = Replace(FriendID,"--","")
	FriendID = Replace(FriendID,")","")
	fixid = Replace(FriendID,",","")
	fixid = Trim(Replace(fixid," ",""))
	If FriendID = "" Or IsNull(FriendID) Then
		ErrMsg = ErrMsg + "<li>无效的系统参数。</li>"
		Founderr = True
		Exit Sub
	ElseIf Not IsNumeric(fixid) Then
		ErrMsg = ErrMsg + "<li>无效的系统参数。</li>"
		Founderr = True
		Exit Sub
	Else
		enchiasp.Execute("Delete From ECCMS_Friend where userid="&enchiasp.memberid&" And FriendID in ("&FriendID&")")
		Call Returnsuc("<li>好友删除成功！</li>")
	End If
End Sub
'================================================
' 过程名：DelAllFriend
' 作  用：删除所有好友
'================================================
Sub DelAllFriend()
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	enchiasp.Execute("Delete From ECCMS_Friend where userid="& enchiasp.memberid)
	Call Returnsuc("<li>好友清空成功！</li>")
End Sub
'================================================
' 过程名：FriendAdd
' 作  用：添加好友
'================================================
Sub FriendAdd()
	Call PreventRefresh
	Dim grouping,strIncept,FriendName,TotalFriend
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	If Trim(Request("friend")) = "" Then
		ErrMsg = ErrMsg + "<li>请选择要添加好友的名称！</li>"
		Founderr = True
	Else
		strIncept = enchiasp.CheckBadstr(Request("friend"))
		strIncept = split(strIncept,",")
	End If
	If Trim(Request("grouping"))<>"" And IsNumeric(Request("grouping")) then 
		grouping = CInt(Request("grouping"))
	Else
		grouping = 0
	End If
	If Founderr = True Then Exit Sub
	For i = 0 To Ubound(strIncept)
		If i >= 5 Then Exit For
		FriendName = Trim(strIncept(i))
		SQL="select username from [ECCMS_User] where username='"&FriendName&"'"
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			ErrMsg = ErrMsg + "<li>没有找到<font color=red>" & FriendName & "</font>这个用户，操作未成功。</li>"
			Founderr = True
			Exit Sub
		Else
			FriendName = Rs(0)
		End If
		Rs.close
		If enchiasp.membername = Trim(FriendName) Then
			ErrMsg = ErrMsg + "<li>对不起！不能把自已添加为好友。</li>"
			Founderr = True
			Exit Sub
		End If
		If CLng(GroupSetting(6)) <> 0 Then
			TotalFriend = enchiasp.Execute("Select Count(FriendID) from ECCMS_Friend where userid="& enchiasp.memberid &" And username='"& enchiasp.CheckStr(enchiasp.membername) &"'")(0)
			If CLng(TotalFriend) >= CLng(GroupSetting(6)) Then
				ErrMsg = ErrMsg + "<li>对不起！你最多只能添加 <font color=red><b>" & GroupSetting(6) & "</b></font> 位好友。</li>"
				Founderr = True
				Exit Sub
			End If
		End  If
		SQL = "Select FriendID From ECCMS_Friend Where userid="& enchiasp.memberid &" And friend='"& FriendName &"'"
		Set Rs = enchiasp.Execute(SQL)
		If Rs.EOF And Rs.BOF Then
			SQL = "Insert into ECCMS_Friend (userid,UserName,Friend,addTime,grouping) values ("& enchiasp.memberid &",'"& enchiasp.membername &"','"& FriendName &"',"& NowString &","& grouping &") "
			enchiasp.Execute(SQL)
		Else
			ErrMsg = ErrMsg + "<li><font color=red>" & FriendName & "</font>这个用户已经添加过了，请不要重复添加，谢谢！。</li>"
			Founderr = True
			Exit Sub
		End If
	Next
	Call Returnsuc("<li>恭喜您！添加好友成功。</li>")
End Sub
'================================================
' 过程名：MoveFriend
' 作  用：移动好友到其它组
'================================================
Sub MoveFriend()
	If enchiasp.CheckPost=False Then
		ErrMsg = Postmsg
		Founderr = True
		Exit Sub
	End If
	Dim grouping
	Dim FriendID,fixid
	If Trim(Request("grouping"))<>"" And IsNumeric(Request("grouping")) Then
		grouping = CInt(Request("grouping"))
	Else
		ErrMsg = ErrMsg + "<li>好友分组不能为空。</li>"
		Founderr = True
		Exit Sub
	End If
	FriendID = Replace(Request.form("id"),"'","")
	FriendID = Replace(FriendID,";","")
	FriendID = Replace(FriendID,"--","")
	FriendID = Replace(FriendID,")","")
	fixid = Replace(FriendID,",","")
	fixid = Trim(Replace(fixid," ",""))
	If FriendID = "" Or IsNull(FriendID) Then
		ErrMsg = ErrMsg + "<li>无效的系统参数。</li>"
		Founderr = True
		Exit Sub
	ElseIf Not IsNumeric(fixid) Then
		ErrMsg = ErrMsg + "<li>无效的系统参数。</li>"
		Founderr = True
		Exit Sub
	Else
		enchiasp.Execute("Update ECCMS_Friend set grouping = "&grouping&" where userid="&enchiasp.memberid&" And FriendID in ("&FriendID&")")
		Call Returnsuc("<li>恭喜您！移动好友分组成功。</li>")
	End If
End Sub
%>
<!--#include file="foot.inc"-->















