<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/chkinput.asp"-->
<!--#include file="../api/cls_api.asp"-->
<%
Admin_header
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
Dim Action
Dim i,ii,RsObj
Dim keyword,findword,strClass,sUserGroup,foundsql
Dim seluserid,UserGrade,UserGroupStr,UserPassWord,username
Dim maxperpage,CurrentPage,totalnumber,TotalPageNum,userlock
Action = LCase(Request("action"))

Select Case Trim(Action)
	Case "save"
		If Not ChkAdmin("AddUser") Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
		Call SaveUser
	Case "modify"
		Call ModifyUser
	Case "add"
		If Not ChkAdmin("AddUser") Then
			Server.Transfer("showerr.asp")
			Response.End
		End If
		Call AddUser
	Case "edit"
		Call EditUser
	Case "del"
		Call BatDelUser
	Case Else
		Call MainPage
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Sub PageTop()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th colspan=2>会员管理</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr><form method=Post name=myform action=admin_user.asp onSubmit='return JugeQuery(this);'>"
	Response.Write "	  <td class=TableRow1>搜索："
	Response.Write "	  <input name=keyword type=text size=20>"
	Response.Write "	  条件："
	Response.Write "	  <select name=queryopt>"
	Response.Write "		<option value=1 selected>会员名称</option>"
	Response.Write "		<option value=2>真实姓名</option>"
	Response.Write "		<option value=3>用户昵称</option>"
	Response.Write "	  </select> <input type=submit name=Submit value='开始搜索' class=Button></td>"
	Response.Write "	  <td class=TableRow1>"
	Response.Write "	  </td></form>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td colspan=2 class=TableRow2><strong>操作选项：</strong> <a href='admin_user.asp'>会员管理首页</a> | "
	Response.Write "	  <a href='admin_user.asp?action=add'><font color=blue>添加会员</font></a> | "
	Response.Write "	  <a href='admin_user.asp?lock=1'><font color=blue>等待验证的会员</font></a> "
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup where Grades <> 0 order by Groupid")
	Do While Not RsObj.EOF
		Response.Write " | <a href=admin_user.asp?UserGrade="
		Response.Write RsObj("Grades")
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</a>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
	Response.Write "	  </td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Response.Write "<br>"
End Sub
Sub MainPage()
	Call PageTop
	If Not ChkAdmin("AdminUser") Then
		Server.Transfer("showerr.asp")
		Response.End
	End If
	If Not IsEmpty(Request("seluserid")) Then
		seluserid = Request("seluserid")
		Select Case enchiasp.CheckStr(Request("act"))
			Case "删除用户"
				Call BatDelUser
			Case "激活用户"
				Call NolockUser
			Case "锁定用户"
				Call IslockUser
			Case "转移用户"
				Call MoveUser
			Case Else
				Response.Write "无效参数！"
		End Select
	End If
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th width='5%' nowrap>选择</th>
	<th width='20%' nowrap>用户名</th>
	<th width='10%' nowrap>用户身份证</th>
	<th width='10%' nowrap>会员类型</th>
	<th width='5%' nowrap>邮箱</th>
	<th width='5%' nowrap>性别</th>
	<th width='20%' nowrap>操作选项</th>
	<th width='15%' nowrap>最后登陆时间</th>
	<th width='5%' nowrap>登陆次数</th>
	<th width='5%' nowrap>状态</th>
</tr>
<%
	If Trim(Request("UserGrade")) <> "" Then
		SQL = "SELECT GroupName,Grades FROM [ECCMS_UserGroup] WHERE Grades=" & Request("UserGrade")
		Set Rs = enchiasp.Execute(SQL)
		If Rs.Bof And Rs.EOF Then
			Response.Write "Sorry！没有找到任何会员。或者您选择了错误的系统参数!"
			Response.End
		Else
			sUserGroup = Rs("GroupName")
			UserGrade = Rs("Grades")
		End If
		Rs.Close
	Else
		sUserGroup = "全部会员"
		UserGrade = 0
	End If
	maxperpage = 20 '###每页显示数
	
	If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
		Response.Write ("错误的系统参数!请输入整数")
		Response.End
	End If
	If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
		CurrentPage = CInt(Request("page"))
	Else
		CurrentPage = 1
	End If
	userlock =0
	If CInt(CurrentPage) = 0 Then CurrentPage = 1
	If Not IsNull(Request("keyword")) And Request("keyword") <> "" Then
		keyword = enchiasp.ChkQueryStr(Request("keyword"))
		If CInt(Request("queryopt")) = 1 Then
			findword = "where username like '%" & keyword & "%'"
		ElseIf CInt(Request("queryopt")) = 2 Then
			findword = "where TrueName like '%" & keyword & "%'"
		Else
			findword = "where nickname like '%" & keyword & "%'"
		End If
		foundsql = findword
		sUserGroup = "查询会员"
	Else
		If Trim(Request("UserGrade")) <> "" Then
			foundsql = "where UserGrade = " & Request("UserGrade")
			UserGrade = Request("UserGrade")
		Else
			If Trim(Request("lock")) <> "" Then
				foundsql = "where userlock =1"
				userlock =1
			Else
				foundsql = ""
			End If
		End If
	End If
	TotalNumber = enchiasp.Execute("SELECT COUNT(userid) FROM ECCMS_User "& foundsql &"")(0)
	TotalPageNum = CInt(TotalNumber / maxperpage)  '得到总页数
	If TotalPageNum < TotalNumber / maxperpage Then TotalPageNum = TotalPageNum + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT userid,username,nickname,UserGrade,UserGroup,UserClass,UserLock,userpoint,usermoney,TrueName,UserSex,usermail,HomePage,oicq,JoinTime,ExpireTime,LastTime,userlogin FROM [ECCMS_User] "& foundsql &" ORDER BY JoinTime DESC ,userid DESC"
	If IsSqlDataBase=1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = enchiasp.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=10 class=TableRow1>还没有找到任何会员信息！</td></tr>"
	Else
		Rs.MoveFirst
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0

		Response.Write "<tr>"
		Response.Write "	<td colspan=10 class=tablerow2>"
		Call showpage()
		Response.Write "</td>"
		Response.Write "	<form name=selform method=post action="""">"
		Response.Write "</tr>"

		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.End
			If (i mod 2) = 0 Then
				strClass = "class=TableRow1"
			Else
				strClass = "class=TableRow2"
			End If
			Response.Write "<tr align=center>"
			Response.Write "	<td " & strClass & "><input type=checkbox name=seluserid value='" & Rs("userid") & "'></td>"
			Response.Write "	<td " & strClass & ">"
			Response.Write "<a href='?action=edit&userid=" & Rs("userid") & "' title='用户昵称：" & Rs("nickname") & "'>"
			If Rs("UserGrade") = 999 Then
				Response.Write "<span class=style2>"
			Else
				Response.Write "<span>"
			End If
			Response.Write Rs("username")
			Response.Write "</span></a>"
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			Response.Write Rs("UserGroup")
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			If Rs("UserGrade") = 999 Then
				Response.Write "管理员"
			Else
				If Rs("UserClass") = 0 Then
					Response.Write "计点会员"
				ElseIf Rs("UserClass") = 1 Then
					Response.Write "计时会员"
				Else
					Response.Write "计时到期"
				End If
			End If
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			Response.Write "<a href='admin_mailist.asp?action=mail&useremail="
			Response.Write Rs("usermail")
			Response.Write "'><img src='images/email.gif' border=0 alt='给用户发邮件'></a>"
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			Response.Write Rs("UserSex")
			Response.Write "	</td>"
			Response.Write "	<td nowrap " & strClass & ">"
			Response.Write "<a href='?action=edit&userid=" & Rs("userid") & "'>编辑</a> | "
			Response.Write "<a href='?action=del&userid=" & Rs("userid") & "'>删除</a>"
			Response.Write "	</td>"
			Response.Write "	<td nowrap " & strClass & ">"
			If Rs("LastTime") >= Date Then
				Response.Write "<font color=red>"
				Response.Write Rs("LastTime")
				Response.Write "</font>"
			Else
				Response.Write Rs("LastTime")
			End If
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			Response.Write Rs("userlogin")
			Response.Write "	</td>"
			Response.Write "	<td " & strClass & ">"
			If Rs("UserLock") = 0 Then
				Response.Write "<font color=blue>√</font>"
			Else
				Response.Write "<font color=red>×</font>"
			End If
			Response.Write "	</td>"
			Response.Write "</tr>"
			Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td colspan=10 class=tablerow1>
	<input class=Button type=button name=chkall value='全选' onClick=CheckAll(this.form)><input class=Button type=button name=chksel value='反选' onClick=ContraSel(this.form)>&nbsp;&nbsp;管理选项：&nbsp;
	 <input class=button onClick="{if(confirm('确定删除选定的用户吗?')){this.document.form.submit();return true;}return false;}" type=submit value='删除用户' name=act> 
	 <input class=button onClick="{if(confirm('确定激活选定的用户吗?')){this.document.form.submit();return true;}return false;}" type=submit value='激活用户' name=act> 
	 <input class=button onClick="{if(confirm('确定锁定选定的用户吗?')){this.document.form.submit();return true;}return false;}" type=submit value='锁定用户' name=act> 
	 <input class=button onClick="{if(confirm('确定转移选定的用户吗?')){this.document.form.submit();return true;}return false;}" type=submit value='转移用户' name=act> → 
	 <select name='sUserGrade'>
	 <option value=''>↓请选择用户组↓</option>
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup where Grades <> 0 order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & "," & RsObj("GroupName") & """>"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>
	</select>
	</td>
</tr></form>
<tr>
	<td colspan=10 class=tablerow1><%Call showpage()%></td>
</tr>
</table>

<%
End Sub

Sub AddUser()
	Call PageTop
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan="2">添加会员</th>
</tr>
<form name=myform method=post action=?action=save>
<tr>
	<td width='30%' align=right class=tablerow1><strong>登陆名称：</strong></td>
	<td width='70%' class=tablerow1><input type=text name=username size=20 value=''></td>
</tr>
<tr>
	<td align=right class=tablerow2><strong>用户密码：</strong></td>
	<td class=tablerow2><input type=password name=password1 size=20></td>
</tr>
<tr>
	<td align=right class=tablerow1><strong>确认密码：</strong></td>
	<td class=tablerow1><input type=password name=password2 size=20></td>
</tr>
<tr>
	<td align=right class=tablerow2><strong>用户昵称：</strong></td>
	<td class=tablerow2><input type=text name=nickname size=20 value=''></td>
</tr>
<tr>
	<td align=right class=tablerow1><strong>用户邮箱：</strong></td>
	<td class=tablerow1><input type=text name=usermail size=30 value='<%=enchiasp.MasterMail%>'></td>
</tr>
<tr>
	<td align=right class=tablerow2><strong>用户姓别：</strong></td>
	<td class=tablerow2><select name='UserSex'>
		<option value='男'>帅哥</option>
		<option value='女'>美女</option>
	</select></td>
</tr>
<tr>
	<td align=right class=tablerow1><strong>所属用户组：</strong></td>
	<td class=tablerow1><select name='UserGrade'>
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup where Grades <> 0 order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & "," & RsObj("GroupName") & """"
		If RsObj("Grades") = 1 Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>
	</select></td>
</tr>
<tr>
	<td align=right class=tablerow2><strong>用户点数：</strong></td>
	<td class=tablerow2><input type=text name=userpoint size=10 value='50'></td>
</tr>
<tr align=center>
	<td colspan=2 class=tablerow1>
	<input type=button name=Submit2 onclick="javascript:history.go(-1)" value='返回上一页' class=Button>
	<input type=Submit name=Submit1 value='添加用户' class=Button></td>
</tr>
</form>
</table>

<%
End Sub

Sub EditUser()
	Call PageTop
	Dim userid,username
	userid = enchiasp.ChkNumeric(Request("userid"))
	username = Replace(Request("username"), "'", "")
	If userid = 0 Then
		SQL = "SELECT * FROM ECCMS_user WHERE username='" & username & "'"
	Else
		SQL = "SELECT * FROM ECCMS_user WHERE userid=" & userid
	End If
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！没有找到任何会员。或者您选择了错误的系统参数!</li>"
		Exit Sub
	End If
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan=4>查看/修改会员资料</th>
</tr>
<form name=myform method=post action=?action=modify>
<input type=hidden name=userid value='<%=Rs("userid")%>'>
<tr>
	<td width='10%' class=tablerow1>会员名称</td>
	<td width='40%' class=tablerow1><input type=text name=username size=20 value='<%=Rs("username")%>' disabled></td>
	<td width='10%' class=tablerow1>真实姓名</td>
	<td width='40%' class=tablerow1><input type=text name=TrueName size=20 value='<%=Rs("TrueName")%>'></td>
</tr>
<tr>
	<td class=tablerow2>用户密码</td>
	<td class=tablerow2><input type=password name=password size=20> <font color=blue>如果不修改密码请留空</font></td>
	<td class=tablerow2>用户邮箱</td>
	<td class=tablerow2><input type=text name=usermail size=30 value='<%=Rs("usermail")%>'></td>
</tr>
<tr>
	<td class=tablerow1>交易密码</td>
	<td class=tablerow1><input type=text name=BuyCode size=20> <font color=blue>如果不修改密码请留空</font></td>
	<td class=tablerow1>用户状态</td>
	<td class=tablerow1>
	<input type=radio name=UserLock value='0'<%If Rs("UserLock") = 0 Then Response.Write " checked"%>> 激活&nbsp;&nbsp;
	<input type=radio name=UserLock value='1'<%If Rs("UserLock") <> 0 Then Response.Write " checked"%>> 锁定&nbsp;&nbsp;
	</td>
</tr>
<tr>
	<td class=tablerow2>用户等级</td>
	<td class=tablerow2><select name='UserGrade'>
<%
	Set RsObj = enchiasp.Execute("Select GroupName,Grades From ECCMS_UserGroup where Grades <> 0 order by Groupid")
	Do While Not RsObj.EOF
		Response.Write Chr(9) & Chr(9) & "<option value=""" & RsObj("Grades") & "," & RsObj("GroupName") & """"
		If RsObj("Grades") = Rs("UserGrade") Then Response.Write " selected"
		Response.Write ">"
		Response.Write RsObj("GroupName")
		Response.Write "</option>" & vbCrLf
		RsObj.movenext
	Loop
	Set RsObj = Nothing
%>
	</select></td>
	<td class=tablerow2>会员类型</td>
	<td class=tablerow2><select name='UserClass'>
		<option value='0'<%If Rs("UserClass") = 0 Then Response.Write " selected"%>>计点会员</option>
		<option value='1'<%If Rs("UserClass") = 1 Then Response.Write " selected"%>>计时会员</option>
		<option value='999'<%If Rs("UserClass") = 999 Then Response.Write " selected"%>>到期会员</option>
	</select></td>
</tr>
<tr>
	<td class=tablerow1>用户点数</td>
	<td class=tablerow1><input type=text name=userpoint size=10 value='<%=Rs("userpoint")%>'></td>
	<td class=tablerow1>账户余额</td>
	<td class=tablerow1><input type=text name=usermoney size=10 value='<%=Rs("usermoney")%>'> 元</td>
</tr>
<tr>
	<td class=tablerow2 nowrap>用户经验值</td>
	<td class=tablerow2><input type=text name=experience size=10 value='<%=Rs("experience")%>'></td>
	<td class=tablerow2 nowrap>用户魅力值</td>
	<td class=tablerow2><input type=text name=charm size=10 value='<%=Rs("charm")%>'></td>
</tr>
<tr>
	<td class=tablerow1>身分证号码</td>
	<td class=tablerow1><input type=text name=UserIDCard size=35 value='<%=Rs("UserIDCard")%>'></td>
	<td class=tablerow1>姓别</td>
	<td class=tablerow1><select name='UserSex'>
		<option value='男'<%If Rs("UserSex") = "男" Then Response.Write " selected"%>>帅哥</option>
		<option value='女'<%If Rs("UserSex") = "女" Then Response.Write " selected"%>>美女</option>
	</select></td>
</tr>
<tr>
	<td class=tablerow2>用户电话</td>
	<td class=tablerow2><input type=text name=phone size=20 value='<%=Rs("phone")%>'></td>
	<td class=tablerow2>用户QQ</td>
	<td class=tablerow2><input type=text name=oicq size=20 value='<%=Rs("oicq")%>'></td>
</tr>
<tr>
	<td class=tablerow1>邮政编码</td>
	<td class=tablerow1><input type=text name=postcode size=20 value='<%=Rs("postcode")%>'></td>
	<td class=tablerow1>联系地址</td>
	<td class=tablerow1><input type=text name=address size=45 value='<%=Rs("address")%>'></td>
</tr>
<tr>
	<td class=tablerow2>密码问题</td>
	<td class=tablerow2><input type=text name=question size=20 value='<%=Rs("question")%>'></td>
	<td class=tablerow2>密码答案</td>
	<td class=tablerow2><input type=text name=answer size=20> <font color=blue>如果不修改密码请留空</font></td>
</tr>
<tr>
	<td class=tablerow1 nowrap>最后登陆时间</td>
	<td class=tablerow1><input type=text name=LastTime size=30 value='<%=Rs("LastTime")%>'></td>
	<td class=tablerow1>最后登陆IP</td>
	<td class=tablerow1><input type=text name=userlastip size=20 value='<%=Rs("userlastip")%>'></td>
</tr>
<tr>
	<td class=tablerow2>注册时间</td>
	<td class=tablerow2><input type=text name=JoinTime size=30 value='<%=Rs("JoinTime")%>'></td>
	<td class=tablerow2>到期时间</td>
	<td class=tablerow2><input type=text name=ExpireTime size=30 value='<%=Rs("ExpireTime")%>'></td>
</tr>
<tr>
	<td class=tablerow1>用户图像</td>
	<td class=tablerow1><input type=text name=UserFace size=30 value='<%=Rs("UserFace")%>'></td>
	<td class=tablerow1>登陆次数</td>
	<td class=tablerow1><input type=text name=userlogin size=10 value='<%=Rs("userlogin")%>'></td>
</tr>
<tr>
	<td class=tablerow1>密码保护</td>
	<td class=tablerow1>
	<input type=radio name=Protect value='0'<%If Rs("Protect") = 0 Then Response.Write " checked"%>> 未申请&nbsp;&nbsp;
	<input type=radio name=Protect value='1'<%If Rs("Protect") <> 0 Then Response.Write " checked"%>> 已申请&nbsp;&nbsp;</td>
	<td class=tablerow1>用户昵称</td>
	<td class=tablerow1><input type=text name=nickname size=20 value='<%=Rs("nickname")%>'></td>
</tr>
<tr align=center>
	<td colspan=4 class=tablerow2>
	<input type=button name=Submit2 onclick="javascript:history.go(-1)" value='返回上一页' class=Button>
	<input type=Submit name=Submit1 value='确认修改' class=Button></td>
</tr></form>
</table>

<%
End Sub

Sub CheckSave()
	If Trim(Request.Form("usermail")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>用户邮箱不能为空！</li>"
	End If
	If IsValidEmail(Trim(Request.Form("usermail"))) = False Then
		ErrMsg = ErrMsg + "<li>您的Email有错误。</li>"
		FoundErr = True
	End If
	If Not IsNumeric(Request.Form("userpoint")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>用户点数必需是数字！</li>"
	End If
	If Trim(Request.Form("nickname")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>用户昵称不能为空！</li>"
	End If
	If enchiasp.IsValidStr(Request("nickname")) = False Then
		ErrMsg = ErrMsg + "<li>用户昵称中含有非法字符。</li>"
		Founderr = True
	End If
	UserGroupStr = Split(Request.Form("UserGrade"), ",")
End Sub

Sub SaveUser()
	CheckSave
	Dim Password,Question,Answer
	Dim usersex,sex
	If Trim(Request.Form("username")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>用户名不能为空！</li>"
	End If
	If enchiasp.IsValidStr(Request("username")) = False Then
		ErrMsg = ErrMsg + "<li>用户名中含有非法字符。</li>"
		Founderr = True
	Else
		username = enchiasp.CheckBadstr(Request("username"))
	End If
	If Trim(Request.Form("password1")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>用户密码不能为空！</li>"
	End If
	If Trim(Request.Form("password2")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>确认密码不能为空！</li>"
	End If
	If Request.Form("password1") <> Request.Form("password2") Then
		ErrMsg = ErrMsg + "<li>您输入的密码和确认密码不一致。</li>"
		FoundErr = True
	End If
	If enchiasp.IsValidPassword(Request("password2")) = False Then
		ErrMsg = ErrMsg + "<li>密码中含有非法字符。</li>"
		Founderr = True
	Else
		Password = Trim(Request.Form("password2"))
		UserPassWord =  md5(Password)
	End If
	If Trim(Request.Form("usersex")) = "" Then
		ErrMsg = ErrMsg + "<li>您的姓别不能为空！</li>"
		Founderr = True
	Else
		usersex = enchiasp.CheckBadstr(Request.Form("usersex"))
	End If
	If usersex = "女" Then
		sex = 0
	Else
		sex = 1
	End If
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_User WHERE username = '" & username & "'")
	If Not (Rs.bof And Rs.EOF) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！此用户已经存在,请换一个用户名再试！</li>"
		Exit Sub
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = enchiasp.Execute("SELECT username FROM ECCMS_Admin WHERE username='" & UserName & "'")
	If Not (Rs.BOF And Rs.EOF) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry！此用户已经存在,请换一个用户名再试！</li>"
		Exit Sub
	End If
	Rs.Close:Set Rs = Nothing
	Question = Trim(Request.Form("question"))
	Answer = Trim(Request.Form("answer"))
	If Question = "" Then Question = enchiasp.GetRandomCode
	If Answer = "" Then Answer = enchiasp.GetRandomCode
	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	Dim API_enchiasp,API_SaveCookie,SysKey
	If API_Enable Then
		Set API_enchiasp = New API_Conformity
		API_enchiasp.NodeValue "action","reguser",0,False
		API_enchiasp.NodeValue "username",UserName,1,False
		Md5OLD = 1
		SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
		Md5OLD = 0
		API_enchiasp.NodeValue "syskey",SysKey,0,False
		API_enchiasp.NodeValue "password",Password,0,False
		API_enchiasp.NodeValue "email",enchiasp.CheckStr(Request.Form("usermail")),1,False
		API_enchiasp.NodeValue "question",Question,1,False
		API_enchiasp.NodeValue "answer",Answer,1,False
		API_enchiasp.NodeValue "gender",sex,0,False
		API_enchiasp.SendHttpData
		If API_enchiasp.Status = "1" Then
			Founderr = True
			ErrMsg =  ErrMsg & API_enchiasp.Message
			Exit Sub
		End If
		Set API_enchiasp = Nothing
	End If
	'-----------------------------------------------------------------
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_User WHERE (userid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("username") = username
		Rs("password") = UserPassWord
		Rs("nickname") = Trim(Request.Form("nickname"))
		Rs("UserGrade") = CInt(UserGroupStr(0))
		Rs("UserGroup") = Trim(UserGroupStr(1))
		Rs("UserClass") = 0
		Rs("UserLock") = 0
		Rs("UserFace") = "face/1.gif"
		Rs("userpoint") = Trim(Request.Form("userpoint"))
		Rs("usermoney") = 0
		Rs("savemoney") = 0
		Rs("prepaid") = 0
		Rs("experience") = 10
		Rs("charm") = 10
		Rs("TrueName") = Trim(Request.Form("username"))
		Rs("usersex") = enchiasp.CheckStr(Request.Form("usersex"))
		Rs("usermail") = enchiasp.CheckStr(Request.Form("usermail"))
		Rs("oicq") = ""
		Rs("question") = Question
		Rs("answer") = md5(Answer)
		Rs("JoinTime") = Now()
		Rs("ExpireTime") = Now()
		Rs("LastTime") = Now()
		Rs("Protect") = 0
		Rs("usermsg") = 0
		Rs("userlastip") = ""
		Rs("userlogin") = 0
		Rs("usersetting") = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
	Rs.update
	Rs.Close:Set Rs = Nothing
	Call RemoveCache
	Succeed("<li>恭喜您！添加会员[<font color=blue>" & Request("username") & "</font>]成功。</li>")
End Sub

Sub ModifyUser()
	CheckSave
	Dim sex
	If Trim(Request.Form("usersex")) = "女" Then
		sex = 0
	Else
		sex = 1
	End If
	If enchiasp.IsValidPassword(Request("password")) = False And Trim(Request("password")) <> "" Then
		ErrMsg = ErrMsg + "<li>密码中含有非法字符。</li>"
		Founderr = True
	End If
	If enchiasp.IsValidPassword(Request("BuyCode")) = False And Trim(Request("BuyCode")) <> "" Then
		ErrMsg = ErrMsg + "<li>交易密码中含有非法字符。</li>"
		Founderr = True
	End If
	If enchiasp.IsValidPassword(Request("answer")) = False And Trim(Request("answer")) <> "" Then
		ErrMsg = ErrMsg + "<li>问题答案中含有非法字符。</li>"
		Founderr = True
	End If
	If Not IsDate(Request.Form("JoinTime")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>注册时间参数错误！</li>"
	End If
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_User WHERE userid = " & CLng(Request("userid"))
	Rs.Open SQL,Conn,1,3
		'Rs("username") = Trim(Request.Form("username"))
		Rs("nickname") = Trim(Request.Form("nickname"))
		If Trim(Request.Form("password")) <> "" Then Rs("password") = md5(Request.Form("password"))
		If Trim(Request.Form("BuyCode")) <> "" Then Rs("BuyCode") = md5(Request.Form("BuyCode"))
		Rs("UserGrade") = CInt(UserGroupStr(0))
		Rs("UserGroup") = Trim(UserGroupStr(1))
		Rs("UserClass") = Trim(Request.Form("UserClass"))
		Rs("UserLock") = Trim(Request.Form("UserLock"))
		Rs("UserFace") = Trim(Request.Form("UserFace"))
		Rs("userpoint") = Trim(Request.Form("userpoint"))
		Rs("usermoney") = Trim(Request.Form("usermoney"))
		Rs("experience") = Trim(Request.Form("experience"))
		Rs("charm") = Trim(Request.Form("charm"))
		Rs("TrueName") = Trim(Request.Form("TrueName"))
		Rs("UserIDCard") = Trim(Request.Form("UserIDCard"))
		Rs("usersex") = Trim(Request.Form("usersex"))
		Rs("usermail") = Trim(Request.Form("usermail"))
		Rs("phone") = Trim(Request.Form("phone"))
		Rs("oicq") = Trim(Request.Form("oicq"))
		Rs("postcode") = Trim(Request.Form("postcode"))
		Rs("address") = Trim(Request.Form("address"))
		Rs("question") = Trim(Request.Form("question"))
		If Trim(Request.Form("answer")) <> "" Then Rs("answer") = md5(Request.Form("answer"))
		Rs("Protect") = Trim(Request.Form("Protect"))
		Rs("JoinTime") = Trim(Request.Form("JoinTime"))
		Rs("ExpireTime") = Trim(Request.Form("ExpireTime"))
		Rs("LastTime") = Trim(Request.Form("LastTime"))
		Rs("userlastip") = Trim(Request.Form("userlastip"))
		Rs("userlogin") = Trim(Request.Form("userlogin"))
	Rs.update
	username = Rs("username")
	Rs.Close:Set Rs = Nothing
	If Founderr = False Then
		'-----------------------------------------------------------------
		'系统整合
		'-----------------------------------------------------------------
		Dim API_enchiasp,API_SaveCookie,SysKey
		If API_Enable Then
			Set API_enchiasp = New API_Conformity
			API_enchiasp.NodeValue "action","update",0,False
			API_enchiasp.NodeValue "username",UserName,1,False
			Md5OLD = 1
			SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
			Md5OLD = 0
			API_enchiasp.NodeValue "syskey",SysKey,0,False
			API_enchiasp.NodeValue "password",Trim(Request.form("password")),1,False
			API_enchiasp.NodeValue "answer",Trim(Request.Form("answer")),1,False
			API_enchiasp.NodeValue "question",Trim(Request.Form("question")),1,False
			API_enchiasp.NodeValue "email",Trim(Request.Form("usermail")),1,False
			API_enchiasp.NodeValue "gender",sex,0,False
			API_enchiasp.SendHttpData
			If API_enchiasp.Status = "1" Then
				ErrMsg = API_enchiasp.Message
			End If
			Set API_enchiasp = Nothing
		End If
		'-----------------------------------------------------------------
	End If
	Call RemoveCache
	Succeed("<li>恭喜您！修改会员[<font color=blue>" & username & "</font>]的资料成功。</li>" & ErrMsg)
End Sub

Sub BatDelUser()
	Dim AllUserID,AllUserName
	If Trim(Request("userid")) <> "" Then
		seluserid = Request("userid")
	End If
	If Len(seluserid) = 0 Then seluserid = "0"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT userid,username FROM [ECCMS_User] WHERE userid in (" & seluserid & ")"
	Rs.Open SQL,Conn,1,1
	If Not (Rs.Bof And Rs.EOF) Then
		Do While Not Rs.EOF
			AllUserID = AllUserID & Rs(0) & ","
			AllUserName = AllUserName & Rs(1) & ","
			enchiasp.Execute("UPDATE ECCMS_Message SET delsend=1 WHERE sender='"& enchiasp.CheckStr(Rs(1)) &"'")
			enchiasp.Execute("DELETE FROM ECCMS_Message WHERE flag=0 And incept='"& enchiasp.CheckStr(Rs(1)) &"'")
		Rs.movenext
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	If AllUserID <> "" Then
		If Right(AllUserID,1) = "," Then AllUserID = Left(AllUserID,Len(AllUserID)-1)
		If Right(AllUserName,1) = "," Then AllUserName = Left(AllUserName,Len(AllUserName)-1)
		enchiasp.Execute ("DELETE FROM ECCMS_User WHERE userid in (" & AllUserID & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_Favorite WHERE userid in (" & AllUserID & ")")
		enchiasp.Execute ("DELETE FROM ECCMS_Friend WHERE userid in (" & AllUserID & ")")
	
		'-----------------------------------------------------------------
		'系统整合
		'-----------------------------------------------------------------
		Dim API_enchiasp,API_SaveCookie,SysKey
		If API_Enable Then
			Set API_enchiasp = New API_Conformity
			API_enchiasp.NodeValue "action","delete",0,False
			API_enchiasp.NodeValue "username",AllUserName,1,False
			Md5OLD = 1
			SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
			Md5OLD = 0
			API_enchiasp.NodeValue "syskey",SysKey,0,False
			API_enchiasp.SendHttpData
			Set API_enchiasp = Nothing
		End If
		'-----------------------------------------------------------------
		OutHintScript ("批量删除操作成功！")
	End If
	Call RemoveCache
	'OutHintScript ("批量删除操作成功！")
End Sub

Sub IslockUser()
	enchiasp.Execute ("UPDATE ECCMS_User SET UserLock=1 WHERE userid in (" & seluserid & ")")
	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	Dim API_enchiasp,API_SaveCookie,SysKey
	If API_Enable Then
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT userid,username FROM [ECCMS_User] WHERE userid in (" & seluserid & ")"
		Rs.Open SQL,Conn,1,1
		If Not (Rs.Bof And Rs.EOF) Then
			Do While Not Rs.EOF
				UserName = Rs(1)
				Set API_enchiasp = New API_Conformity
				API_enchiasp.NodeValue "action","lock",0,False
				API_enchiasp.NodeValue "username",UserName,1,False
				Md5OLD = 1
				SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
				Md5OLD = 0
				API_enchiasp.NodeValue "syskey",SysKey,0,False
				API_enchiasp.NodeValue "userstatus",1,0,False
				API_enchiasp.SendHttpData
				Set API_enchiasp = Nothing
			Rs.movenext
			Loop
		End If
		Rs.Close:Set Rs = Nothing
	End If
	'-----------------------------------------------------------------
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Sub NolockUser()
	enchiasp.Execute ("UPDATE ECCMS_User SET UserLock=0 WHERE userid in (" & seluserid & ")")
	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	Dim API_enchiasp,API_SaveCookie,SysKey
	If API_Enable Then
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT userid,username FROM [ECCMS_User] WHERE userid in (" & seluserid & ")"
		Rs.Open SQL,Conn,1,1
		If Not (Rs.Bof And Rs.EOF) Then
			Do While Not Rs.EOF
				UserName = Rs(1)
				Set API_enchiasp = New API_Conformity
				API_enchiasp.NodeValue "action","lock",0,False
				API_enchiasp.NodeValue "username",UserName,1,False
				Md5OLD = 1
				SysKey = Md5(API_enchiasp.XmlNode("username") & API_ConformKey)
				Md5OLD = 0
				API_enchiasp.NodeValue "syskey",SysKey,0,False
				API_enchiasp.NodeValue "userstatus",0,0,False
				API_enchiasp.SendHttpData
				Set API_enchiasp = Nothing
			Rs.movenext
			Loop
		End If
		Rs.Close:Set Rs = Nothing
	End If
	'-----------------------------------------------------------------
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Sub MoveUser()
	If Request("sUserGrade") = "" Then
		OutAlertScript("请选择正确的系统参数！")
		Exit Sub
	End If
	UserGroupStr = Split(Request("sUserGrade"), ",")
	enchiasp.Execute ("update ECCMS_User set UserGrade=" & CInt(UserGroupStr(0)) & ", UserGroup='" & UserGroupStr(1) & "' where userid in (" & seluserid & ")")
	Response.redirect (Request.ServerVariables("HTTP_REFERER"))
End Sub

Sub showpage()
	Dim n
	If totalnumber Mod maxperpage = 0 Then
		n = totalnumber \ maxperpage
	Else
		n = totalnumber \ maxperpage + 1
	End If
	Response.Write "<table cellspacing=1 width='100%' border=0><form method=Post action=?UserGrade=" & Request("UserGrade") & "&lock=" & Request("lock") & "><tr><td align=center> " & vbCrLf
	Response.Write "<font color='red'>" & sUserGroup & "</font> "
	If CurrentPage < 2 Then
		Response.Write "共有会员 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 位&nbsp;首 页&nbsp;上一页&nbsp;|&nbsp;"
	Else
		Response.Write "共有会员 <font COLOR=#FF0000><strong>" & totalnumber & "</strong></font> 位&nbsp;<a href=?page=1&UserGrade=" & Request("UserGrade") & "&lock=" & Request("lock") & ">首 页</a>&nbsp;"
		Response.Write "<a href=?page=" & CurrentPage - 1 & "&UserGrade=" & Request("UserGrade") & "&lock=" & Request("lock") & ">上一页</a>&nbsp;|&nbsp;"
	End If
	If n - CurrentPage < 1 Then
		Response.Write "下一页&nbsp;尾 页" & vbCrLf
	Else
		Response.Write "<a href=?page=" & (CurrentPage + 1) & "&UserGrade=" & Request("UserGrade") & "&lock=" & Request("lock") & ">下一页</a>"
		Response.Write "&nbsp;<a href=?page=" & n & "&UserGrade=" & Request("UserGrade") & "&lock=" & Request("lock") & ">尾 页</a>" & vbCrLf
	End If
	Response.Write "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
	Response.Write "&nbsp;转到："
	Response.Write "<input name=page size=3 value='" & CurrentPage & "'> <input class=Button type=submit name=Submit value='转到'>"
	Response.Write "</td></tr></FORM></table>" & vbCrLf
End Sub
Sub RemoveCache()
	enchiasp.DelCahe "RenewStatistics"
	enchiasp.DelCahe "TotalStatistics"
End Sub
%>












