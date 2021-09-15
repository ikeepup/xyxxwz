<!--#include file="config.asp"-->
<!--#include file="../inc/chkinput.asp"-->
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
Call InnerLocation("交费确认")

Dim Action,SQL,Rs
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "save"
	Call SaveConfirm
Case Else
	Call showmain
End Select
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
Sub showmain()
%>
<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>
	<tr height=20>
		<th colspan=2>交费确认</th>
	</tr>
	<tr height=20 align=center>
		<td class=Usertablerow2 colspan=2><font color=red>注意：</font><font color=blue>请一定要正确填写以下含*的选项，以方便我们核对！</font></td>
	</tr>
	<form name=form2 method=post action=?action=save>
	<tr height=20>
		<td class=Usertablerow1 width="20%" align=right>汇款日期：</td>
		<td class=Usertablerow1 width="80%"><input type="text" name="PayDate" size=15 value="<%=date()%>"> <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>汇款金额：</td>
		<td class=Usertablerow1><input type="text" name="PayMoney" size=15 onkeyup=if(isNaN(this.value))this.value=''> 元 <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>定 单 号：</td>
		<td class=Usertablerow1><input type="text" name="indent" size=30> <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>汇款方式：</td>
		<td class=Usertablerow1>
		<input type=radio name=paymode value="银行汇款" checked> 电汇&nbsp;&nbsp;
		<input type=radio name=paymode value="邮局汇款"> 邮汇&nbsp;&nbsp;
		<input type=radio name=paymode value="网上支付"> 网上支付
		</td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>用户名：</td>
		<td class=Usertablerow1><input type="text" name="username" size=15 value="<%=enchiasp.MemberName%>"> <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>汇款人名称：</td>
		<td class=Usertablerow1><input type="text" name="customer" size=30> <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>汇款人邮箱：</td>
		<td class=Usertablerow1><input type="text" name="Email" size=30> <font color=red>*</font></td>
	</tr>
	<tr height=20>
		<td class=Usertablerow1 align=right>其它说明：</td>
		<td class=Usertablerow1><textarea name=readme rows=5 cols=50></textarea> <font color=red>*</font></td>
	</tr>
	<tr height=20 align=center>
		<td class=Usertablerow2 colspan=2><input type=submit value=" 确认提交 "  class=Button></td>
	</tr>
	</form>
<%
	Response.Write "</table>"
End Sub
Sub SaveConfirm()
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If Not IsDate(Request.Form("PayDate")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>日期输入错误。</li>"
	End If
	If Not IsNumeric(Request.Form("PayMoney")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>汇款金额输入错误。</li>"
	End If
	If Trim(Request.Form("indent")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>你的定单号没有填咧？</li>"
	End If
	If IsValidEmail(Request.Form("Email")) = False Then
		ErrMsg = ErrMsg + "<li>您的Email有错误！</li>"
		Founderr = True
	End If
	If Trim(Request.Form("customer")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>汇款人名称不能为空。</li>"
	End If
	If Trim(Request.Form("username")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>用户名不能为空？</li>"
	End If
	If Founderr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Confirm where (id is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("paymode").Value =enchiasp.CheckBadstr( Trim(Request.Form("paymode")))
		Rs("PayDate").Value = Trim(Request.Form("PayDate"))
		Rs("PayMoney").Value = Trim(Request.Form("PayMoney"))
		Rs("indent").Value = Left(enchiasp.ChkFormStr(Request.Form("indent")),35)
		Rs("Email").Value = Trim(Request.Form("Email"))
		Rs("customer").Value = Left(enchiasp.ChkFormStr(Request.Form("customer")),30)
		Rs("username").Value = Left(enchiasp.ChkFormStr(Request.Form("username")),30)
		Rs("readme").Value = Left(enchiasp.ChkFormStr(Request.Form("readme")),200)
		Rs("isPass").Value = 0
	Rs.Update
	Rs.close:set Rs = Nothing
	Call Returnsuc("<li>恭喜您！确认信息提交成功，我们会在一个工作日内处理你的定单。")
End Sub

%>
<!--#include file="foot.inc"-->











