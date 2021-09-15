<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/cls_payment.asp"-->
<!--#include file="head.inc"-->
<%
'=====================================================================
' 软件名称：恩池网站管理系统---帐号充值
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Call InnerLocation("帐号充值")

Dim Rs,SQL,strChinaeBank
strChinaeBank = Split(enchiasp.ChinaeBank, "|||")
Dim Action:Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "add"
		Call AddMoney
	Case "pay"
		Call PayMoney
	Case "view"
		Call ViewRecord
	Case "del"
		Call DelRecord
	Case Else
		Call showmain
End Select
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If

Sub showmain()
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=2>会员帐号充值</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=Usertablerow1 colspan=2><a href=""?action=view"">查看历史定单</a></td>"
	Response.Write "	</tr>"
	Response.Write "	<form name=addform method=post action=?action=add>"
	Response.Write "	<tr>"
	Response.Write "		<td class=Usertablerow1 align=right><b>充值金额：</b></td>"
	Response.Write "		<td class=Usertablerow1><input type=text name=money size=20 onkeyup=if(isNaN(this.value))this.value='' value=''> <font color=blue>元</font>"
	Response.Write "		<input type=submit value="" 确定 "" class=Button> <input type=reset value="" 重填 "" class=Button></td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
End Sub
Sub AddMoney()
	Response.Write vbNewLine
	Response.Write "<script language=JavaScript>" & vbNewLine
	Response.Write "function chkaddform(myform1){" & vbNewLine
	Response.Write "	if (myform1.codestr.value==''){" & vbNewLine
	Response.Write "		alert('请填写验证码！');" & vbNewLine
	Response.Write "		return false;" & vbNewLine
	Response.Write "	}" & vbNewLine
	Response.Write "}" & vbNewLine
	Response.Write "</script>" & vbNewLine
	If Not IsNumeric(Request.Form("money")) Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>请输入你要充值的金额，或者你输入的金额有错误！</li>"
		Exit Sub
	End If
	If FormatNumber(Request.Form("money")) <= 0 Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>您输入的充值金额总要大于 0 元吧？</li>"
		Exit Sub
	End If
	Dim OrderForm,curdate
	Dim sRnd
	Randomize
	sRnd = Int(9000 * Rnd) + 1000
	curdate=now()                                               
	OrderForm = Year(curdate) & Month(curdate) & Day(curdate) &"-"& sRnd &"-"& Hour(curdate) & Minute(curdate) & Second(curdate)
	Call PreventRefresh  '防刷新
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=2>会员帐号充值</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=addform method=post action=?action=pay onSubmit=""return chkaddform(this);"">"
	Response.Write "	<input type=hidden name=title value=""会员帐号充值"">"
	Response.Write "	<input type=hidden name=userid value=""" & Trim(enchiasp.memberid) & """>"
	Response.Write "	<input type=hidden name=username value=""" & Trim(enchiasp.membername) & """>"
	Response.Write "	<tr>"
	Response.Write "		<td class=Usertablerow1 align=right><b>您要充值的金额：</b></td>"
	Response.Write "		<td class=Usertablerow1><font color=red>" & FormatCurrency(Request.Form("money"),2,-1) & "</font> 元"
	Response.Write "		<input type=hidden name=addmoney value=""" & CCur(Request.Form("money")) & """></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=Usertablerow1 align=right><b>充值定单号：</b></td>"
	Response.Write "		<td class=Usertablerow1><font color=red>" & OrderForm & "</font>"
	Response.Write "		<input type=hidden name=OrderForm value=""" & OrderForm & """> （请牢记您的订单号）</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=Usertablerow1 align=right><b>其它说明：</b><br>最多200个字符</td>"
	Response.Write "		<td class=Usertablerow1><textarea name=readme rows=5 cols=50></textarea></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=Usertablerow1 align=right><b>支付方式：</b></td>"
	Response.Write "		<td class=Usertablerow1><select name=paytype>"
	Response.Write "		<option value=0>银行汇款</option>"
	Response.Write "		<option value=1>在线支付</option>"
	Response.Write "		<option value=2>邮局汇款</option>"
	Response.Write "		<option value=3>上门交费</option>"
	Response.Write "	</select>"
	If CInt(strChinaeBank(2)) > 0 Then
		Response.Write "&nbsp;&nbsp;<b>注意：</b> 在线支付需要收取<font color=red>" & strChinaeBank(2) & "%</font>的手续费"
	Else
		Response.Write "&nbsp;&nbsp;<b>注意：</b> <font color=red>在线支付可以实时充值</font>"
	End If
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=Usertablerow1 align=right><b>验证码：</b></td>"
	Response.Write "		<td class=Usertablerow1><input type=""text"" name=""codestr"" maxlength=""4"" size=""4"">&nbsp;<img src=""../inc/getcode.asp""></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=Usertablerow1 colspan=2><input type=submit value="" 确定支付 "" class=Button></td>"
	Response.Write "	</tr>"
	Response.Write "	</form>"
	Response.Write "</table>"
End Sub
Sub ViewRecord()
	Dim CurrentPage,page_count,totalrec,Pcount,PageListNum
	PageListNum = 20
	CurrentPage = Request("page")
	If CurrentPage = "" Or Not IsNumeric(CurrentPage) Then
		CurrentPage = 1
	Else
		CurrentPage = CLng(CurrentPage)
	End If
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th>定 单 号</th>"
	Response.Write "		<th>支付金额</th>"
	Response.Write "		<th>支付标题</th>"
	Response.Write "		<th>提交日期</th>"
	Response.Write "		<th>付款方式</th>"
	Response.Write "		<th>状 态</th>"
	Response.Write "		<th>操 作</th>"
	Response.Write "	</tr>"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_AddMoney WHERE userid=" & enchiasp.memberid & " And username='" & enchiasp.membername & "' And deletion=0 ORDER BY id DESC"
	Rs.Open SQL,conn,1,1
	If Not (Rs.BOF And Rs.EOF) Then
		Rs.PageSize = PageListNum
		Rs.AbsolutePage = CurrentPage
		page_count = 0
		totalrec = Rs.recordcount
		Do While Not Rs.EOF And (Not page_count = Rs.PageSize)
			Response.Write "	<tr align=center>"
			Response.Write "		<td class=Usertablerow1><font color=red>" & Rs("OrderForm") & "</font></td>"
			Response.Write "		<td class=Usertablerow1>" & FormatCurrency(Rs("addmoney"),2,-1) & " 元</td>"
			Response.Write "		<td class=Usertablerow1>" & Rs("title") & "</td>"
			Response.Write "		<td class=Usertablerow1>" & Rs("addtime") & "</td>"
			Response.Write "		<td class=Usertablerow1>" & Rs("paytype") & "</td>"
			Response.Write "		<td class=Usertablerow1>"
			If Rs("finished") > 0 Then
				Response.Write "<font color=blue>已处理</font>"
			Else
				Response.Write "<font color=red>未处理</font>"
			End If
			Response.Write "</td>"
			Response.Write "		<td class=Usertablerow1>"
			If Rs("finished")>0 Then
				Response.Write "<a href=""?action=del&id=" & Rs("id") & """ onClick=""return confirm('确定要删除此定单吗？')"">删 除</a>"
			Else
				Response.Write "<a onClick=""return confirm('此定单还未处理，不能删除！')"">删 除</a>"
			End If
			Response.Write "</td>"
			Response.Write "	</tr>"
			page_count = page_count + 1
		Rs.movenext
		Loop
	Else
		Response.Write "	<tr align=center>"
		Response.Write "		<td class=Usertablerow1 colspan=7>没有任何定单！</td>"
		Response.Write "	</tr>"
	End If
	If totalrec Mod PageListNum = 0 Then
		Pcount =  totalrec \ PageListNum
	Else
		Pcount =  totalrec \ PageListNum+1
	End If
	If page_count = 0 Then CurrentPage = 0
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=Usertablerow2 colspan=7>"
	Response.Write ShowPages(CurrentPage,Pcount,totalrec,PageListNum,"action="& Request("action"))
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	Set Rs = Nothing
End Sub
Sub DelRecord()
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If Not IsNumeric(Request("id")) Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>请正确输入ID！</li>"
		Exit Sub
	End If
	enchiasp.Execute ("UPDATE ECCMS_AddMoney SET deletion=1 WHERE userid=" & enchiasp.memberid & " And finished>0 And id="& CLng(Request("id")))
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub PayMoney()
	If enchiasp.CheckPost=False Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If Not IsNumeric(Request.Form("addmoney")) Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>请输入你要充值的金额，或者你输入的金额有错误！</li>"
		Exit Sub
	End If
	If FormatNumber(Request.Form("addmoney")) <= 0 Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>您输入的充值金额总要大于 0 元吧？</li>"
		Exit Sub
	End If
	If Not IsNumeric(Request.Form("paytype")) Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>请选择支付方式！</li>"
		Exit Sub
	End If
	If Trim(Request.Form("OrderForm")) = Empty Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>你的定单号错误！</li>"
		Exit Sub
	End If
	If Not enchiasp.CodeIsTrue() Then
		ErrMsg = ErrMsg + "<meta http-equiv=refresh content=""2;URL=addmoney.asp""><li>验证码校验失败，请返回刷新页面再试。两秒后自动返回</li>"
		Session("BankPayCode") = ""
		Founderr = True
		Exit Sub
	End If
	Session("GetCode") = ""
	Set Rs = enchiasp.Execute("SELECT id FROM ECCMS_AddMoney WHERE OrderForm='"& enchiasp.CheckbadStr(Request.Form("OrderForm")) &"'")
	If Not (Rs.BOF And Rs.EOF) Then
		ErrMsg = ErrMsg + "<li>您已经提交了表单，请不要重复提交！！！</li>"
		Session("BankPayCode") = ""
		Founderr = True
		Set Rs = Nothing
		Exit Sub
	End If
	Set Rs = Nothing
	Dim strpaytype
	Select Case CInt(Request.Form("paytype"))
	Case 0
		strpaytype = "银行汇款"
	Case 1
		strpaytype = "在线支付"
	Case 2
		strpaytype = "邮局汇款"
	Case 3
		strpaytype = "上门交费"
	Case Else
		strpaytype = "其它汇款"
	End Select
	If CInt(Request.Form("paytype")) = 1 Then
		
		Call Web_Payment
	Else
		If Founderr = True Then Exit Sub
		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM ECCMS_AddMoney WHERE (id is null)"
		Rs.Open SQL,Conn,1,3
		Rs.AddNew
			Rs("userid").Value = enchiasp.memberid
			Rs("username").Value = enchiasp.membername
			Rs("title").Value = enchiasp.CheckBadstr(Request.Form("title"))
			Rs("OrderForm").Value =enchiasp.CheckBadstr(Request.Form("OrderForm"))

			Rs("addmoney").Value = CCur(Request.Form("addmoney"))
			Rs("addtime").Value = Now()
			Rs("readme").Value = enchiasp.ChkbadStr(Request.Form("readme"))
			Rs("paytype").Value = strpaytype
			Rs("finished").Value = 0
			Rs("deletion").Value = 0
		Rs.Update
		Rs.Close:Set Rs = Nothing
		Call Returnsuc("<li>恭喜您！充值信息提交成功。</li>")
	End If
	
End Sub

Sub Web_Payment()
	
	Response.Write "<table cellspacing=1 align=center cellpadding=3 border=0 class=Usertableborder>"
	Response.Write "	<tr>"
	Response.Write "		<th colspan=2>确认在线支付信息</th>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=Usertablerow1 align=right><b>用户名称：</b></td>"
	Response.Write "		<td class=Usertablerow1>" & enchiasp.membername & "</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=Usertablerow1 align=right><b>定单号：</b></td>"
	Response.Write "		<td class=Usertablerow1><font color=red>" & Trim(Request.Form("OrderForm")) & "</font></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=Usertablerow1 align=right><b>实际支付金额：</b></td>"
	Response.Write "		<td class=Usertablerow1>￥ " & enchiasp.ReadPayMoney(Request.Form("addmoney"),False) & " 元</td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "		<td class=Usertablerow1 colspan=2>&nbsp;&nbsp;<font color=blue>如果以上信息正确，请您前往在线支付平台交费。</font></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr align=center>"
	Response.Write "		<td class=Usertablerow2 colspan=2>"
	Dim wp
	Set wp = New WebPayment_Cls
	wp.PayPlatform = CInt(enchiasp.StopBankPay)
	wp.Paymentid = Trim(strChinaeBank(0))
	wp.Paymentkey = Trim(strChinaeBank(1))
	wp.Percent = enchiasp.CheckNumeric(strChinaeBank(2))
	wp.Returnurl = enchiasp.GetSiteUrl &"/user/receive.asp"
	wp.Orderid = enchi.CheckBadstr(Request.Form("OrderForm"))
	wp.Paymoney = enchiasp.CheckNumeric(Request.Form("addmoney"))
	wp.Consignee = MemberName
	wp.Consigner = MemberName
	wp.Email = MemberEmail
	wp.Address = enchiasp.SiteUrl
	wp.PaymentPlatform
	Select Case CInt(wp.ErrNumber)
	Case 1
		ErrMsg = wp.Description
		Founderr = True
	Case 2
		ErrMsg = wp.Description
		Founderr = True
	Case 6
		ErrMsg = wp.Description
		Founderr = True
	Case 8
		ErrMsg = wp.Description
		Founderr = True
	End Select
	Set wp = Nothing
	Response.Write "</td>"
	Response.Write "	</tr>"
	Response.Write "</table>"
	If Founderr = True Then Exit Sub
End Sub
%><!--#include file="foot.inc"-->