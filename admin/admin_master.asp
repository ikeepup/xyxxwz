<!--#include file="setup.asp" -->
<!--#include file="check.asp"-->
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
Dim selAdminID
Dim i,Action,strClass
Admin_header
If Not ChkAdmin("999") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Response.Write "<table cellpadding=2 cellspacing=1 border=0 class=tableBorder align=center>" & vbCrLf
Response.Write " <tr>" & vbCrLf
Response.Write " <th height=22 colspan=6>管理员操作</th>" & vbCrLf
Response.Write " </tr>" & vbCrLf
Response.Write " <tr>" & vbCrLf
Response.Write " <td class=TableRow1> <b>管理选项：</b> <a href=admin_master.asp>管理首页</a> &nbsp;<a href=admin_master.asp?action=add>添加管理员</a>"
Response.Write " </td>" & vbCrLf
Response.Write " </tr>" & vbCrLf
Response.Write " </table><br>" & vbCrLf
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "renew"
	Call UpdateFlag
Case "del"
	Call del
Case "pasword"
	Call pasword
Case "newpass"
	Call newpass
Case "add"
	Call addadmin
Case "edit"
	Call userinfo
Case "savenew"
	Call savenew
Case "active"
	Call ActiveLock
Case Else
	Call userlist
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
private function instr2(ByVal old,byval str)
	'Adminflag, "Add" & strModules & ChanID
	on error resume next
	dim temp
	dim i
	dim tempb 
	tempb=0
	if old<>"" then
		
		temp=split(old,",")
		for i=0 to ubound(temp)
		
		if str=temp(i) then
				tempb=1
				exit for
			end if
		next 
		instr2=tempb
	else
		instr2=0
	end if
	
	
	
End Function
Private Sub userlist()
	Response.Write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th height=22 colspan=6>管理员管理(点击用户名进行操作)</th>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr align=center>" & vbCrLf
	Response.Write "<td height=22 class=TableTitle><B>用户名</B></td>" & vbCrLf
	Response.Write "<td class=TableTitle><B>上次登陆时间</B></td>" & vbCrLf
	Response.Write "<td class=TableTitle><B>上次登陆IP</B></td>" & vbCrLf
	Response.Write "<td class=TableTitle><B>操作</B></td>" & vbCrLf
	Response.Write "<td class=TableTitle><B>状态</B></td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Set Rs = enchiasp.Execute("select * from ECCMS_Admin order by Logintime desc")
	i = 0
	Do While Not Rs.EOF
		If (i mod 2) = 0 Then
			strClass = "class=TableRow1"
		Else
			strClass = "class=TableRow2"
		End If
		Response.Write " <tr>" & vbCrLf
		Response.Write " <td " & strClass & "><a href=""?id="
		Response.Write Rs("id")
		Response.Write "&action=pasword"" title='点击此处修改管理员信息'>"
		Response.Write Rs("username")
		Response.Write "</a></td>" & vbCrLf
		Response.Write "<td align=center " & strClass & ">"
		Response.Write Rs("Logintime")
		Response.Write "</td>" & vbCrLf
		Response.Write "<td align=center " & strClass & ">"
		Response.Write Rs("Loginip")
		Response.Write "</td>" & vbCrLf
		Response.Write "<td align=center " & strClass & "><a href=""?action=Active&id=" & Rs("id") & "&lock="
		If Rs("isLock") = 0 Then
			Response.Write "1"" onclick=""return confirm('您确定要锁定此管理员吗?')"">锁定管理员</a> | "
		Else
			Response.Write "0"" onclick=""return confirm('您确要激活此管理员吗?')"">激活管理员</a> | "
		End If
		Response.Write "<a href=""?action=del&id="
		Response.Write Rs("id")
		Response.Write "&name="
		Response.Write Rs("username")
		Response.Write """ onclick=""return confirm('此操作将删除该管理员\n 您确定执行此操作吗?')"">删除</a>&nbsp;|&nbsp;<a href=""?id="
		Response.Write Rs("id")
		Response.Write "&action=edit"">编辑权限</a>" & vbCrLf
		
		response.write "| "
		Response.Write "<a href='?action=pasword&id="
		Response.Write Rs("id")
		response.write "'>"
		response.write "二次码规则 "
		response.write "</a>"
		
		response.write "</td>"
		Response.Write "<td align=center " & strClass & ">"
		If Rs("isLock") = 0 Then
			Response.Write "正常"
		Else
			Response.Write "<font color=red>锁定<font>"
		End If
		Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf
		Rs.movenext
		i = i + 1
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td colspan=""6"" align=center Class=TableRow1>" & vbCrLf
	Response.Write " <input class=""button"" type=button name=""Submit"" value=""添加管理员"" onClick=""self.location='admin_master.asp?action=add'"" >" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " </table>" & vbCrLf
End Sub

Private Sub del()
	If Trim(Request("id")) <> "" Then
		enchiasp.Execute ("delete from ECCMS_Admin where username<>'" & AdminName & "' And id=" & Request("id"))
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		ErrMsg = "<li>错误的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
End Sub

Private Sub ActiveLock()
	If Trim(Request("lock")) <> "" And Trim(Request("id")) <> "" Then
		enchiasp.Execute ("update ECCMS_Admin set isLock="&Request("lock")&" where username<>'" & AdminName & "' And id=" & Request("id"))
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		ErrMsg = "<li>错误的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
End Sub


Private Sub pasword()
	Dim oldpassword
	If Not ChkAdmin("9999") Then
		ErrMsg = "<li>您没有此操作权限!</li><li>如有什么问题请联系站长？</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("select * from ECCMS_Admin where id=" & Request("id"))
	oldpassword = Rs("password")
	Response.Write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write "<form action=""?action=newpass"" method=post>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th colspan=2 height=23>管理员资料管理－－密码修改" & vbCrLf
	Response.Write " </th>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>后台登陆名称：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=hidden name=""oldusername"" value="""
	Response.Write Rs("username")
	Response.Write """>" & vbCrLf
	Response.Write " <input type=text size=25 name=""username2"" value="""
	Response.Write Rs("username")
	Response.Write """>" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>后台登陆密码：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=""password"" size=25 name=""password2"">"
	Response.Write " (如果不修改密码请留空)" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>管理员级别：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=AdminGrade value='0' "
	If Rs("AdminGrade") = 0 Then Response.Write " checked"
	Response.Write " > 普通管理员&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=AdminGrade value='999' "
	If Rs("AdminGrade") = 999 Then Response.Write " checked"
	Response.Write " > 高级管理员 （拥有最高权限）" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td align=""right"" class=tablerow1>是否激活管理员：</td>" & vbCrLf
	Response.Write " <td class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isLock value='1' "
	If Rs("isLock") = 1 Then Response.Write " checked"
	Response.Write " > 否&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=isLock value='0' "
	If Rs("isLock") = 0 Then Response.Write " checked"
	Response.Write " > 是" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td align=""right"" class=tablerow1>限制一个管理员登陆：</td>" & vbCrLf
	Response.Write " <td class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isAloneLogin value='0' "
	If Rs("isAloneLogin") = 0 Then Response.Write " checked"
	Response.Write " > 否&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=isAloneLogin value='1' "
	If Rs("isAloneLogin") = 1 Then Response.Write " checked"
	Response.Write " > 是" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
		Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>是否启用密码规则：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isuseercima value='1' "
	If Rs("isuseercima") = 1 Then Response.Write " checked"
	Response.Write " > 是" & vbCrLf
	Response.Write " <input type=radio name=isuseercima value='0'"
	If Rs("isuseercima") = 0 Then Response.Write " checked"
	Response.Write " > 否" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>密码规则：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=jiafa value='1' "
	If Rs("jiafa") = 1 Then Response.Write " checked"
	Response.Write " > 加法" & vbCrLf
	Response.Write " <input type=radio name=jiafa value='0' "
	If Rs("jiafa") = 0 Then Response.Write " checked"
	Response.Write " > 乘法" & vbCrLf
	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>第1个运算的位置：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name='weizhi1' value='"
	response.write rs("weizhi1")
	response.write "'>" & vbCrLf	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>第2个运算的位置：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name='weizhi2' value='"
	response.write rs("weizhi2")
	response.write "'>" & vbCrLf	
	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>插入二次基码的位置：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name='jimaweizhi' value='"
	response.write rs("jimaweizhi")
	response.write "'>" & vbCrLf	
	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>密码规则说明：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " 密码规则用于有效保护管理员的密码，不同的管理员可以设置不同的密码规则，二次密码在二次基码基础上按照一定的规则生成。例如开启密码规则后，如果基码为liuyunfan，规则为加法，验证码为2365，取第1个位置和第3个位置，插入基码的第4个位置，那么二次密码为liuy8unfan，如果要取消或更改二次密码基码，则在[基本设置]中修改。" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>绑定IP：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name='useip' value='"
	response.write rs("useip")
	response.write "'>" & vbCrLf	
	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	
	
	Response.Write " <tr align=""center"">" & vbCrLf
	Response.Write " <td colspan=""2"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=hidden name=id value="""
	Response.Write Request("id")
	Response.Write """>" & vbCrLf
	Response.Write " <input type=button name=Submit4 onclick='javascript:history.go(-1)' value='返回上一页' class=Button> <input type=""submit"" name=""Submit"" value=""更 新"" class=""button"">" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " </form>" & vbCrLf
	Response.Write " </table>" & vbCrLf
Rs.Close
Set Rs = Nothing
End Sub

Private Sub newpass()
	Dim passnw
	Dim usernw
	Dim aduser
	Dim oldpassword
	If Not ChkAdmin("9999") Then
		ErrMsg = "<li>您没有此操作权限!</li><li>如有什么问题请联系站长？</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("select * from ECCMS_Admin where id=" & Request("id"))
	oldpassword = Rs("password")
	If Request("username2") = "" Then
		ErrMsg = "<li>请输入管理员名字。<a href=?>［ <font color=red>返回</font> ］</a>"
		Founderr = True
		Exit Sub
	Else
		usernw = Trim(Request("username2"))
	End If
	
	if Request.Form("isuseercima") = "1" then
		If Request.Form("weizhi1") = "" or Request.Form("weizhi2") = "" or Request.Form("jimaweizhi") = "" Then
			ErrMsg = "请输入二次密码的相关内容相关内容！"
			Founderr = True
			Exit Sub
		else
			if not (IsNumeric(Trim(Request.Form("weizhi1"))) and IsNumeric(Trim(Request.Form("weizhi2"))) and IsNumeric(Trim(Request.Form("jimaweizhi")))) then
				FoundErr = True
				ErrMsg = ErrMsg + "<li>二次密码的相关内容只能输入数字！</li>"
				exit sub
			else
				if cint(Trim(Request.Form("weizhi1")))>4 or cint(Trim(Request.Form("weizhi2")))>4 then
					FoundErr = True
					ErrMsg = ErrMsg + "<li>二次密码的相关数字超过验证码的长度4位，请修改！</li>"
					exit sub
				end if
			end if

		End If
	end if


	If Request("password2") = "" Then
		passnw = "没有修改"
	Else
		passnw = Request("password2")
	End If
	Set Rs = CreateObject("adodb.recordset")
	SQL = "select * from ECCMS_Admin where username='" & Trim(Request("oldusername")) & "'"
	Rs.Open SQL, conn, 1, 3
	If Not Rs.EOF And Not Rs.bof Then
		Rs("username") = usernw
		If Request("password2") <> "" Then Rs("password") = md5(Request("password2"))
		If CInt(Request.Form("AdminGrade")) = 999 Then
			Rs("status") = "高级管理员"
		Else
			Rs("status") = "普通管理员"
		End If
		Rs("AdminGrade") = Request.Form("AdminGrade")
		Rs("isLock") = Request.Form("isLock")
		Rs("isAloneLogin") = Request.Form("isAloneLogin")
		rs("isuseercima")= Request.Form("isuseercima")
		rs("jiafa")= Request.Form("jiafa")
		rs("weizhi1")= Request.Form("weizhi1")
		rs("weizhi2")= Request.Form("weizhi2")
		rs("jimaweizhi")= Request.Form("jimaweizhi")
		'if Request.Form("useip")<>"" then
		rs("useip")= Request.Form("useip")
		'end if



		Succeed ("<li>管理员资料更新成功，请记住更新信息。<br> 管理员：" & Request("username2") & " <BR> 密   码：" & passnw & "")
		Rs.update
	End If
	Rs.Close
	Set Rs = Nothing	
End Sub

Private Sub addadmin()
	If Not ChkAdmin("9999") Then
		ErrMsg = "<li>您没有此操作权限!</li><li>如有什么问题请联系站长？</li>"
		Founderr = True
		Exit Sub
	End If
	Response.Write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write "<form action=""?action=savenew"" method=post>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <th colspan=2 height=23>管理员管理－－添加管理员" & vbCrLf
	Response.Write " </th>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>后台登陆名称：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name=""username2"">" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>后台登陆密码：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=""password"" name=""password2"">" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>管理员级别：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=AdminGrade value='0' checked> 普通管理员&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=AdminGrade value='999'> 高级管理员 （拥有最高权限）" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td align=""right"" class=tablerow1>限制一个管理员登陆：</td>" & vbCrLf
	Response.Write " <td class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isAloneLogin value='0' checked> 否&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=isAloneLogin value='1'> 是" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>是否激活管理员：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isLock value='1' checked> 否&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=isLock value='0'> 是" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>是否启用密码规则：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=isuseercima value='1' checked> 是&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=isuseercima value='0'> 否" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>密码规则：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=radio name=jiafa value='1' checked> 加法&nbsp;&nbsp;" & vbCrLf
	Response.Write " <input type=radio name=jiafa value='0'> 乘法" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>第1个运算的位置：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name=""weizhi1"">" & vbCrLf	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>第2个运算的位置：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name=""weizhi2"">" & vbCrLf	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>插入二次基码的位置：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name=""jimaweizhi"">" & vbCrLf	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	
	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>密码规则说明：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " 密码规则用于有效保护管理员的密码，不同的管理员可以设置不同的密码规则，二次密码在二次基码基础上按照一定的规则生成。例如开启密码规则后，如果基码为liuyunfan，规则为加法，验证码为2365，取第1个位置和第3个位置，插入基码的第4个位置，那么二次密码为liuy8unfan，如果要取消或更改二次密码基码，则在[基本设置]中修改。" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	Response.Write " <tr >" & vbCrLf
	Response.Write " <td width=""26%"" align=""right"" class=tablerow1>用户IP绑定：</td>" & vbCrLf
	Response.Write " <td width=""74%"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=text name=""useip"">" & vbCrLf	
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf

	
	Response.Write " <tr align=""center"">" & vbCrLf
	Response.Write " <td colspan=""2"" class=tablerow1>" & vbCrLf
	Response.Write " <input type=button name=Submit4 onclick='javascript:history.go(-1)' value='返回上一页' class=Button>  <input type=""submit"" name=""Submit"" value=""添 加"" class=""button"">" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Response.Write " </form>" & vbCrLf
	Response.Write " </table>" & vbCrLf
End Sub

Private Sub savenew()
	Dim adminuserid
	If Not ChkAdmin("9999") Then
		ErrMsg = "<li>您没有此操作权限!</li><li>如有什么问题请联系站长？</li>"
		Founderr = True
		Exit Sub
	End If
	If Request.Form("username2") = "" Then
		ErrMsg = "请输入后台登陆用户名！"
		Founderr = True
		Exit Sub
	Else
		adminuserid = Request.Form("username2")
	End If
	If Request.Form("password2") = "" Then
		ErrMsg = "请输入后台登陆密码！"
		Founderr = True
		Exit Sub
	End If
	if Request.Form("isuseercima") = "1" then
		If Request.Form("weizhi1") = "" or Request.Form("weizhi2") = "" or Request.Form("jimaweizhi") = "" Then
			ErrMsg = "请输入二次密码的相关内容相关内容！"
			Founderr = True
			Exit Sub
		else
			if not (IsNumeric(Trim(Request.Form("weizhi1"))) and IsNumeric(Trim(Request.Form("weizhi2"))) and IsNumeric(Trim(Request.Form("jimaweizhi")))) then
				FoundErr = True
				ErrMsg = ErrMsg + "<li>二次密码的相关内容只能输入数字！</li>"
				exit sub
			else
				if cint(Trim(Request.Form("weizhi1")))>4 or cint(Trim(Request.Form("weizhi2")))>4 then
					FoundErr = True
					ErrMsg = ErrMsg + "<li>二次密码的相关数字超过验证码的长度4位，请修改！</li>"
					exit sub
				end if
			end if

		End If
	end if

	
	
	
	Set Rs = enchiasp.Execute("select username from ECCMS_Admin where username='" & Replace(Request.Form("username2"), "'", "") & "'")
	If Not (Rs.EOF And Rs.bof) Then
		ErrMsg = "您输入的用户名已经在管理用户中存在！"
		Founderr = True
		Exit Sub
	End If
	Set Rs = Server.CreateObject("adodb.recordset")
	SQL = "select * from ECCMS_Admin where (id is null)"
	Rs.open SQL,conn,1,3
	Rs.addnew
		Rs("username") = Replace(Request.Form("username2"), "'", "")
		If CInt(Request.Form("AdminGrade")) = 999 Then
			Rs("status") = "高级管理员"
		Else
			Rs("status") = "普通管理员"
		End If
		Rs("password") = md5(Request.Form("password2"))
		Rs("isLock") = Request.Form("isLock")
		Rs("AdminGrade") = Request.Form("AdminGrade")
		Rs("Adminflag") = ",,,,,,,,,,,,,,,"
		Rs("LoginTime") = Now()
		Rs("Loginip") = enchiasp.GetUserIP
		Rs("RandomCode") = enchiasp.GetRandomCode
		Rs("isAloneLogin") = Request.Form("isAloneLogin")
		rs("isuseercima")= Request.Form("isuseercima")
		rs("jiafa")= Request.Form("jiafa")
		rs("weizhi1")= Request.Form("weizhi1")
		rs("weizhi2")= Request.Form("weizhi2")
		rs("jimaweizhi")= Request.Form("jimaweizhi")
		'if Request.Form("useip")<>"" then
		rs("useip")= Request.Form("useip")
		'end if


	Rs.update
	Rs.close:set Rs=Nothing
	Succeed ("用户ID:" & adminuserid & " 添加成功，请到管理员管理给予相应的权限，如需修改请返回管理员管理！")
End Sub

Private Sub userinfo()
	Dim Adminflag,rsChannel
	Dim ChanID,ModuleName,strModules
	Set Rs = enchiasp.Execute("SELECT id,Adminflag FROM ECCMS_Admin WHERE id=" & Request("id"))
	Adminflag = Rs("Adminflag")
	Rs.Close
	Set Rs = Nothing
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan=6>管理员权限管理(请选择相应的权限分配给管理员)</th>
</tr>
<form name=myform method=post action=?action=renew>
<input type=hidden name=id value="<%=Request("id")%>">
<tr>
	<td class=tablerow2 colspan=6>&nbsp;<b>常规设置</b></td>
</tr>
<tr>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="SiteConfig" <%If InStr2(Adminflag, "SiteConfig") <> 0 Then Response.Write "checked"%>> 基本设置</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="Advertise" <%If InStr2(Adminflag, "Advertise") <> 0 Then Response.Write "checked"%>> 广告管理</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="Channel" <%If InStr2(Adminflag, "Channel") <> 0 Then Response.Write "checked"%>> 频道设置</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="Template" <%If InStr2(Adminflag, "Template") <> 0 Then Response.Write "checked"%>> 模板管理</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="TemplateLoad" <%If InStr2(Adminflag, "TemplateLoad") <> 0 Then Response.Write "checked"%>> 模板导入、导出</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="Announce" <%If InStr2(Adminflag, "Announce") <> 0 Then Response.Write "checked"%>> 公告管理</td>
</tr>
<tr>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminLog" <%If InStr2(Adminflag, "AdminLog") <> 0 Then Response.Write "checked"%>> 日志管理</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="SendMessage" <%If InStr2(Adminflag, "SendMessage") <> 0 Then Response.Write "checked"%>> 发送信息</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="CreateIndex" <%If InStr2(Adminflag, "CreateIndex") <> 0 Then Response.Write "checked"%>> 生成首页</td>
	<td class=tablerow1></td>
	<td class=tablerow1></td>
	<td class=tablerow1></td>
</tr>
<%
	Set rsChannel = enchiasp.Execute("SELECT ChannelID,ChannelName,modules,ModuleName FROM ECCMS_Channel WHERE StopChannel = 0 And ChannelID <> 4 And ChannelType < 2 Order By orders Asc")
	Do While Not rsChannel.EOF
	ChanID = rsChannel("ChannelID")
	Select Case rsChannel("modules")
		Case 1:strModules = "Article"
		Case 2:strModules = "Soft"
		Case 3:strModules = "Shop"
		Case 5:strModules = "Flash"
		Case 6:strModules = "yemian"
		Case 7:strModules = "job"
	Case Else
		strModules = "Article"
	End Select
%>
<tr>
	<td class=tablerow2 colspan=6>&nbsp;<b><%=rsChannel("ChannelName")%></b></td>

</tr>

<%
select case rsChannel("modules")
	case 6
		'单页面图文频道
		%>
		<tr>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="Add<%=strModules & ChanID%>" <%If InStr2(Adminflag, "Add" & strModules & ChanID) <> 0 Then Response.Write "checked"%>> 添加内容</td> 
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="Admin<%=strModules & ChanID%>" <%If InStr2(Adminflag, "Admin" & strModules & ChanID) <> 0 Then Response.Write "checked"%>> 内容管理</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminClass<%=ChanID%>" <%If InStr2(Adminflag, "AdminClass" & ChanID) <> 0 Then Response.Write "checked"%>> 栏目管理</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminUpload<%=ChanID%>" <%If InStr2(Adminflag, "AdminUpload" & ChanID) <> 0 Then Response.Write "checked"%>> 上传文件管理</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminSelect<%=ChanID%>" <%If InStr2(Adminflag, "AdminSelect" & ChanID) <> 0 Then Response.Write "checked"%>> 选择上传文件</td>
			<td class=tablerow1></td>
		</tr>
		<%
	
	case else
		%>
		<tr>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="Add<%=strModules & ChanID%>" <%If InStr2(Adminflag, "Add" & strModules & ChanID) <> 0 Then Response.Write "checked"%>> 添加<%=rsChannel("ModuleName")%></td> 
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="Admin<%=strModules & ChanID%>" <%If InStr2(Adminflag, "Admin" & strModules & ChanID) <> 0 Then Response.Write "checked"%>> <%=rsChannel("ModuleName")%>管理</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminClass<%=ChanID%>" <%If InStr2(Adminflag, "AdminClass" & ChanID) <> 0 Then Response.Write "checked"%>> <%=rsChannel("ModuleName")%>分类管理</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminUpload<%=ChanID%>" <%If InStr2(Adminflag, "AdminUpload" & ChanID) <> 0 Then Response.Write "checked"%>> 上传文件管理</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminSelect<%=ChanID%>" <%If InStr2(Adminflag, "AdminSelect" & ChanID) <> 0 Then Response.Write "checked"%>> 选择上传文件</td>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminJsFile<%=ChanID%>" <%If InStr2(Adminflag, "AdminJsFile" & ChanID) <> 0 Then Response.Write "checked"%>> JS文件管理</td> 

		</tr>
		<tr>
			<td class=tablerow1><input type="checkbox" name="Adminflag" value="Auditing<%=ChanID%>" <%If InStr2(Adminflag, "Auditing" & ChanID) <> 0 Then Response.Write "checked"%>>  <%=rsChannel("ModuleName")%>审核管理</td>
			<td class=tablerow1><%If rsChannel("modules") = 2 Or rsChannel("modules") = 5 Then%><input type="checkbox" name="Adminflag" value="DownServer<%=ChanID%>" <%If InStr2(Adminflag, "DownServer" & ChanID) <> 0 Then Response.Write "checked"%>> 下载服务器管理<%End If%></td>
			<td class=tablerow1><%If rsChannel("modules") = 2 Then%><input type="checkbox" name="Adminflag" value="ErrorSoft<%=ChanID%>" <%If InStr2(Adminflag, "ErrorSoft" & ChanID) <> 0 Then Response.Write "checked"%>> 错误软件报告<%End If%></td>
			<td class=tablerow1></td> 			
			<td class=tablerow1></td>
			<td class=tablerow1></td>
		</tr>
		<%
end select
		rsChannel.movenext
	Loop
	Set rsChannel = Nothing
%>


<tr>
	<td class=tablerow2 colspan=6>&nbsp;<b>其它管理</b></td>
</tr>
<tr>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="Vote" <%If InStr2(Adminflag, "Vote") <> 0 Then Response.Write "checked"%>> 投票管理</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="FriendLink" <%If InStr2(Adminflag, "FriendLink") <> 0 Then Response.Write "checked"%>> 友情连接管理</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="UploadFile" <%If InStr2(Adminflag, "UploadFile") <> 0 Then Response.Write "checked"%>> 上传文件</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="GuestBook" <%If InStr2(Adminflag, "GuestBook") <> 0 Then Response.Write "checked"%>> 留言管理</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="rizhi" <%If InStr2(Adminflag, "rizhi") <> 0 Then Response.Write "checked"%>> 日志管理</td>
	<td class=tablerow1></td>

</tr>
<tr>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian" <%If InStr2(Adminflag, "flashtupian") <> 0 Then Response.Write "checked"%>> 新闻图片变换</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian2" <%If InStr2(Adminflag, "flashtupian2") <> 0 Then Response.Write "checked"%>> 青春之窗图片变换</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian3" <%If InStr2(Adminflag, "flashtupian3") <> 0 Then Response.Write "checked"%>> 党的建设图片变换</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian4" <%If InStr2(Adminflag, "flashtupian4") <> 0 Then Response.Write "checked"%>> 职工之家图片变换</td>
<td class=tablerow1><input type="checkbox" name="Adminflag" value="gundong" <%If InStr2(Adminflag, "gundong") <> 0 Then Response.Write "checked"%>> 图片左右滚动管理</td>
	<td class=tablerow1></td>

</tr>

<tr>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian5" <%If InStr2(Adminflag, "flashtupian5") <> 0 Then Response.Write "checked"%>>人力资源图片变换</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="flashtupian6" <%If InStr2(Adminflag, "flashtupian6") <> 0 Then Response.Write "checked"%>>科技创新图片变换</td>
	<td class=tablerow1><input type="checkbox" name="Adminflag" value="AdminSelect0" <%If InStr2(Adminflag, "AdminSelect0") <> 0 Then Response.Write "checked"%>>从图片中选择文件</td>
	<td class=tablerow1></td>
<td class=tablerow1></td>
	<td class=tablerow1></td>

</tr>




<tr>
	<td class=tablerow2 colspan=6 align=center><input type=button name=Submit4 onclick='javascript:history.go(-1)' value='返回上一页' class=Button> 　　<input class=Button type=button name=chkall value='全选' onClick='CheckAll(this.form)'><input class=Button type=button name=chksel value='反选' onClick='ContraSel(this.form)'>
	<input type="submit" name="Submit" value="更新管理员权限" class=button></td>
</tr>
</form>
</table>

<%
End Sub

Private Sub UpdateFlag()
	Set Rs = Server.CreateObject("adodb.recordset")
	SQL = "SELECT * FROM ECCMS_Admin WHERE id=" & Request("id")
	Rs.Open SQL, conn, 1, 3
	If Not (Rs.EOF And Rs.BOF) Then
		Rs("Adminflag") = Replace(Replace(Request("Adminflag"), "'", ""), " ", "")
		Rs.update
	End If
	Rs.Close
	Set Rs = Nothing
	Sucmsg = "<li>管理员更新成功，请记住更新信息。"
	Succeed (Sucmsg)
End Sub
%>