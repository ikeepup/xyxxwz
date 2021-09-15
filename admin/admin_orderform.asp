<!--#include file="setup.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
<!--#include file="check.asp"-->
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

Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
Response.Write "	<tr>"
Response.Write "	  <th>" & sModuleName & "管理选项</th>"
Response.Write "	</tr>"
Response.Write "	<tr><form method=Post name=myform action='' onSubmit='return JugeQuery(this);'>"
Response.Write "	<td class=TableRow1>搜索："
Response.Write "	  <input name=keyword type=text size=30>"
Response.Write "	  条件："
Response.Write "	  <select name='field'>"
Response.Write "		<option value='1' selected>订单号</option>"
Response.Write "		<option value='2'>收 货 人</option>"
Response.Write "		<option value='3'>用 户 名</option>"
Response.Write "	  </select> <input type=submit name=Submit value='开始查询' class=Button><br>"
Response.Write "	  <b>说明：</b>点击订单号查看和处理订单</td></form>"
Response.Write "	</tr></form>"
Response.Write "	<tr>"
Response.Write "	  <td colspan=2 class=TableRow2><strong>操作选项：</strong> <a href='admin_orderform.asp'>管理首页</a> | "
Response.Write "	  <a href='admin_orderform.asp?finish=1'>已处理订单</a> | "
Response.Write "	  <a href='admin_orderform.asp?finish=0'>未处理订单</a> | "
Response.Write "	  <a href='admin_orderform.asp?Cancel=1'>回收站管理</a></td>"
Response.Write "	</tr>"
Response.Write "</table>"
Response.Write "<br>"
If Not ChkAdmin("adminorder") Then
	Server.Transfer("showerr.asp")
	Response.End
End If

Dim Action,i
If CInt(ChannelID) = 0 Then ChannelID = 3
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "save"
	Call SaveOrderForm
Case "view"
	Call ViewOrderForm
Case "del"
	Call DelOrderForm
Case "cancel"
	Call ReclaimOrder
Case "finish"
	Call FinishOrderForm
Case "pay"
	Call PaymentState
Case Else
	Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn

Sub showmain()
	Dim finish,Cancel
	Dim keyword,findword,foundsql
	Dim maxperpage,CurrentPage,Pcount,totalrec,totalnumber
	Dim strList,strName,strRowstyle

	maxperpage = 30		'--每页显示列表数
	
	finish = enchiasp.ChkNumeric(Request("finish"))
	Cancel = enchiasp.ChkNumeric(Request("Cancel"))
	Response.Write "<script language=""JavaScript"" src=""include/showpage.js""></script>" & vbNewLine
	Response.Write "<table  border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th width='5%' nowrap>选择</th>"
	Response.Write "	  <th width='23%'>定 单 号</th>"
	Response.Write "	  <th width='12%' nowrap>收 货 人</th>"
	Response.Write "	  <th width='15%' nowrap>合 计 金 额</th>"
	Response.Write "	  <th width='19%' nowrap>定 购 时 间</th>"
	Response.Write "	  <th width='10%' nowrap>付 款 方 式</th>"
	Response.Write "	  <th width='8%' nowrap>付款状态</th>"
	Response.Write "	  <th width='8%' nowrap>订单处理</th>"
	Response.Write "	</tr>"
	
	If Request("keyword") <> "" Then
		keyword = enchiasp.ChkQueryStr(Request("keyword"))
		If CInt(Request("field")) = 1 Then
			foundsql = " And OrderID like '%" & keyword & "%'"
		ElseIf CInt(Request("field")) = 2 Then
			foundsql = " And Consignee like '%" & keyword & "%'"
		ElseIf CInt(Request("field")) = 3 Then
			foundsql = " And username like '%" & keyword & "%'"
		Else
			foundsql = " And OrderID like '%" & keyword & "%'"
		End If
		strName = "订单查询"
		strList = "&keyword=" & keyword
	Else
		If Request("finish") <> "" Then
			foundsql = " And finish=" & finish
			strList = "&finish=" & finish
			If finish = 0 Then
				strName = "未处理订单"
			Else
				strName = "已处理订单"
			End If
		Else
			If Cancel = 0 Then
				strName = "所有订单"
			Else
				strName = "已经删除订单"
			End If
		End If
	End If
	strList = strList & "&Cancel=" & Cancel
	CurrentPage = enchiasp.ChkNumeric(Request("page"))
	If CurrentPage = 0 Then CurrentPage = 1
	totalrec = enchiasp.Execute("SELECT COUNT(id) FROM [ECCMS_OrderForm] WHERE Cancel="& Cancel & foundsql &"")(0)
	Pcount = CLng(totalrec / maxperpage)  '得到总页数
	If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage > Pcount Then CurrentPage = Pcount
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT id,userid,username,ProductID,OrderID,Surcharge,totalmoney,Consignee,Email,PayMode,addTime,invoice,finish,Cancel,PayDone FROM [ECCMS_OrderForm] WHERE Cancel="& Cancel & foundsql &" ORDER BY id DESC"
	If IsSqlDataBase = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	enchiasp.SqlQueryNum = enchiasp.SqlQueryNum + 1
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=""center"" colspan=""9"" class=""TableRow2"">还没有找到任何订单！</td></tr>"
	Else
		If Pcount > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0

		Response.Write "	<tr>"
		Response.Write "	  <td colspan=""8"" class=""TableRow2"">"
		ShowListPage CurrentPage,Pcount,totalrec,maxperpage,strList,strName
		Response.Write "</td>"
		Response.Write "	</tr>"
		Response.Write "	<form name=selform method=post action=""admin_orderform.asp"">"
		Response.Write "	<input type=hidden name=action value='del'>"
		Do While Not Rs.EOF And i < CInt(maxperpage)
			If Not Response.IsClientConnected Then Response.End
			If (i mod 2) = 0 Then
				strRowstyle = "class=""TableRow1"""
			Else
				strRowstyle = "class=""TableRow2"""
			End If
			Response.Write "	<tr align=""center"">"
			Response.Write "	  <td align=center " & strRowstyle & "><input type=checkbox name=id value=" & Rs("id") & "></td>"
			Response.Write "	  <td " & strRowstyle & " title=""点击此处查看订单详细信息""><a href='?action=view&id=" & Rs("id") & "' class=""showlink"">"
			Response.Write Rs("OrderID")
			Response.Write "</a></td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & " title=""点击用户名查看此用户信息"">"
			If Rs("userid") > 0 Then
				Response.Write "<a href='admin_user.asp?action=edit&userid=" & Rs("userid") & "'>"
				Response.Write Rs("Consignee")
				Response.Write "</a>"
			Else
				Response.Write Rs("Consignee")
			End If
			Response.Write "</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & " align=""left"">￥"
			Response.Write FormatNumber(Rs("totalmoney"))
			Response.Write " 元</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & " align=""left"">"
			If Datediff("d",Rs("addTime"),Now()) = 0 Then
				Response.Write "<font color=""red"">" & Rs("addTime") & "</font>"
			Else
				Response.Write "<font color=""#808080"">" & Rs("addTime") & "</font>"
			End If
			Response.Write "</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & ">"
			Response.Write Rs("PayMode")
			Response.Write "</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & ">"
			If Rs("PayDone") > 0 Then
				Response.Write "<a href='?action=pay&sid=0&id=" & Rs("id") & "' title=""点击此处改变支付状态"">"
				Response.Write "<font color=""blue"">已支付</font>"
				Response.Write "</a>"
			Else
				Response.Write "<a href='?action=pay&sid=1&id=" & Rs("id") & "' title=""点击此处改变支付状态"">"
				Response.Write "<font color=""red"">未支付</font>"
				Response.Write "</a>"
			End If
			Response.Write "</td>" & vbNewLine
			Response.Write "	  <td " & strRowstyle & ">"
			If Rs("finish") > 0 Then
				'Response.Write "<a href='?action=finish&fid=0&id=" & Rs("id") & "' title=""点击此处直接处理订单"">"
				Response.Write "<font color=""blue"">已处理</font>"
				'Response.Write "</a>"
			Else
				'Response.Write "<a href='?action=finish&fid=1&id=" & Rs("id") & "' title=""点击此处取消订单"">"
				Response.Write "<font color=""red"">未处理</font>"
				'Response.Write "</a>"
			End If
			Response.Write "</td>" & vbNewLine
			Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
	<tr>
	  <td colspan="8" class="TableRow1">
	  <input class=Button type="button" name="chkall" value="全选" onClick="CheckAll(this.form)"><input class=Button type="button" name="chksel" value="反选" onClick="ContraSel(this.form)">
	  <input class=Button type="submit" name="Submit2" value="彻底删除" onclick="return confirm('订单删除后将不能恢复\n您确定执行该操作吗?');">
	  <%
	  If Cancel = 0 Then
	  %>
	  <input type=hidden name=can value='1'>
	  <input class=Button type="submit" name="Submit3" value="放入回收站" onclick="document.selform.action.value='cancel';return confirm('您确定要将这些订单放入回收站吗?');">
	  <%
	  Else
	  %>
	  <input type=hidden name=can value='0'>
	  <input class=Button type="submit" name="Submit4" value="还原回收站" onclick="document.selform.action.value='cancel';return confirm('您确定还原订单吗?');">
	  <%
	  End If
	  %>
	  </td>
	</tr>
</form>
	<tr>
	  <td colspan="8" align="right" class="TableRow2"><%ShowListPage CurrentPage,Pcount,totalrec,maxperpage,strList,strName %></td>
	</tr>
</table>
<%
End Sub
Sub ReclaimOrder()
	If Request("id") <> "" And Request("can") <> "" Then
		If CInt(Request("can")) = 0 Then
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET Cancel=0 WHERE id in (" & Request("id") & ")")
			OutHintScript("您选择的订单已成功还原！")
		Else
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET Cancel=1 WHERE id in (" & Request("id") & ")")
			OutHintScript("您选择的订单已成功放入回收站！")
		End If
	Else
		ErrMsg = "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
End Sub
Sub DelOrderForm()
	If Request("id") <> "" Then
		enchiasp.Execute ("DELETE FROM [ECCMS_OrderForm] WHERE id in (" & Request("id") & ")")
		enchiasp.Execute ("DELETE FROM [ECCMS_Buy] WHERE orderid in (" & Request("id") & ")")
	Else
		ErrMsg = "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	OutHintScript("您选择的订单已成功删除！")
End Sub
Sub FinishOrderForm()
	If Request("id") <> "" And Request("fid") <> "" Then
		If Request("fid") = 0 Then
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET finish=0 WHERE id=" & CLng(Request("id")))
		Else
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET finish=1 WHERE id=" & CLng(Request("id")))
		End If
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		ErrMsg = "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
End Sub

Sub PaymentState()
	If Request("id") <> "" And Request("sid") <> "" Then
		If Request("sid") = 0 Then
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET PayDone=0 WHERE id=" & CLng(Request("id")))
		Else
			enchiasp.Execute ("UPDATE [ECCMS_OrderForm] SET PayDone=1 WHERE id=" & CLng(Request("id")))
		End If
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		ErrMsg = "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
End Sub

Sub ViewOrderForm()
	Dim id,totalmoney
	id = enchiasp.ChkNumeric(Request("id"))
	If id = 0 Then
		ErrMsg = "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT * FROM [ECCMS_OrderForm] WHERE id=" & id)
	If Rs.BOF And Rs.EOF Then
		ErrMsg = "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Set Rs = Nothing
		Exit Sub
	End If
	totalmoney = FormatNumber(Rs("totalmoney"))
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="tableborder">
<tr>
	<th colspan="4">订单查看/处理</th>
</tr>
<form name="subform" method="post" action="admin_orderform.asp">
<input type="hidden" name="action" value="save">
<input type="hidden" name="id" value="<%=Rs("id")%>">
<tr>
	<td width='15%' class="tablerow1" align="right">订 单 号：</td>
	<td width='42%' class="tablerow1"><font color=red><%=Rs("OrderID")%></font></td>
	<td width='15%' class="tablerow1" align="right">用 户 名：</td>
	<td width='28%' class="tablerow1"><%
	If Rs("userid") > 0 Then
		Response.Write "<a href='admin_user.asp?action=edit&userid=" & Rs("userid") & "'>"
		Response.Write Rs("username")
		Response.Write "</a>"
	Else
		Response.Write "匿名用户"
	End If
	%></td>
</tr>
<tr>
	<td class="tablerow2" align="right">合计金额：</td>
	<td class="tablerow2"><font color=blue>￥<%=FormatNumber(Rs("totalmoney"))%> 元</font></td>
	<td class="tablerow2" align="right">附加费用：</td>
	<td class="tablerow2">￥<%=FormatNumber(Rs("Surcharge"),,-1)%> 元</td>
</tr>
<tr>
	<td class="tablerow1" align="right">订购时间：</td>
	<td class="tablerow1"><font color=red><%=Rs("addTime")%></font></td>
	<td class="tablerow1" align="right">付款方式：</td>
	<td class="tablerow1"><%=Rs("PayMode")%></td>
</tr>
<tr>
	<td class="tablerow2" align="right">收 货 人：</td>
	<td class="tablerow2"><font color=blue><%=Rs("Consignee")%></font></td>
	<td class="tablerow2" align="right">收货单位：</td>
	<td class="tablerow2"><%=Rs("Company")%></td>
</tr>
<tr>
	<td class="tablerow1" align="right">收货人电话：</td>
	<td class="tablerow1"><%=Rs("phone")%></td>
	<td class="tablerow1" align="right">收货人邮编：</td>
	<td class="tablerow1"><%=Rs("postcode")%></td>
</tr>
<tr>
	<td class="tablerow2" align="right">收货人邮箱：</td>
	<td class="tablerow2"><%=Rs("Email")%></td>
	<td class="tablerow2" align="right">收货人QQ号：</td>
	<td class="tablerow2"><%=enchiasp.ChkNull(Rs("oicq"))%></td>
</tr>
<tr>
	<td class="tablerow1" align="right">收货人地址：</td>
	<td class="tablerow1"><%=Rs("Address")%></td>
	<td class="tablerow2" align="right">是否开发票：</td>
	<td class="tablerow2"><%
	If Rs("invoice") > 0 Then
		Response.Write "<font color=""red"">是</font>"
	Else
		Response.Write "<font color=""#808080"">否</font>"
	End If
%></td>
</tr>
<tr>
	<td class="tablerow2" align="right">其它说明：</td>
	<td class="tablerow2" colspan="3">&nbsp;&nbsp;<%=enchiasp.HTMLEncode(Rs("Readme"))%></td>
</tr>
<tr>
	<td class="tablerow1" align="right">订单处理：</td>
	<td class="tablerow1"><font color=red><%
	If Rs("finish") > 0 Then
		Response.Write "<font color=""blue"">已处理</font>"
	Else
		Response.Write "<font color=""red"">未处理</font>"
	End If
%></font></td>
	<td class="tablerow1" align="right">付款状态：</td>
	<td class="tablerow1"><%
	If Rs("PayDone") > 0 Then
		Response.Write "已支付 <input type=radio name=PayDone value=""1"" checked>"
	Else
		Response.Write "未支付 <input type=radio name=PayDone value=""0"" checked>&nbsp;&nbsp;"
		Response.Write "已支付 <input type=radio name=PayDone value=""1"">"
	End If
	Rs.Close:Set Rs = Nothing
%></td>
</tr>
<tr>
	<th>订购数量</td>
	<th>产品名称</td>
	<th>单 价</td>
	<th>合 计</td>
</tr>
<%
	SQL = "SELECT * FROM ECCMS_Buy WHERE orderid=" & id & " ORDER BY ID ASC"
	Set Rs = enchiasp.Execute(SQL)
	If Not (Rs.BOF And Rs.EOF) Then
	Do While Not Rs.EOF
%>
<tr>
	<td class="tablerow1" align="center"><%=Rs("Amount")%></td>
	<td class="tablerow1"><a href="admin_shop.asp?action=view&shopid=<%=Rs("shopid")%>"><%=Rs("TradeName")%></a></td>
	<td class="tablerow1" align="center"><font color=blue>￥<%=FormatNumber(Rs("Price"))%> 元</font></td>
	<td class="tablerow1"  align="center"><font color=red>￥<%=FormatNumber(Rs("totalmoney"))%> 元</font></td>
</tr>
<%
		Rs.movenext
	Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr align="center">
	<td class="tablerow2" colspan="4"><input type="button" onclick="javascript:history.go(-1)" value="返回上一页" name="B1" class=Button>&nbsp;&nbsp; 
		<input type=hidden name=can value='1'>
		<input class=Button type="submit" name="Submit3" value="取消订单" onclick="document.subform.action.value='cancel';return confirm('您确定要将这些订单放入回收站吗?');">&nbsp;&nbsp;
		<input class=Button type="submit" name="Submit4" value="删除订单" onclick="document.subform.action.value='del';return confirm('订单删除后将不能恢复\n您确定执行该操作吗?');">&nbsp;&nbsp;
		<input class=Button type="submit" name="Submit2" value="处理此订单" onclick="return confirm('确定处理此订单吗?');">
	</td>
</tr></form>
<tr>
	<td class="tablerow2" colspan="4"><b>说明：</b><br>
	&nbsp;&nbsp;请确定用户已经汇款，再处理订单，订单一但处理，就证明货已发出。
	</td>
</tr>
</table>
<%
End Sub

Sub SaveOrderForm()
	Dim id,totalmoney,Consignee,Readme,PayDone
	id = enchiasp.ChkNumeric(Request("id"))
	If id = 0 Then
		ErrMsg = "<li>请选择正确的系统参数！</li>"
		Founderr = True
		Exit Sub
	End If
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_OrderForm WHERE finish=0 And id=" & id
	Rs.Open SQL,Conn,1,3
	If Rs.BOF And Rs.EOF Then
		ErrMsg = "<li>此订单已经处理，请不要重复处理订单！</li>"
		Founderr = True
		Set Rs = Nothing
		Exit Sub
	Else
		PayDone = Rs("PayDone")
		Rs("finish") = 1
		Rs("PayDone") = enchiasp.ChkNumeric(Request.Form("PayDone"))
		Rs.update
		totalmoney = Rs("totalmoney")
		Consignee = Rs("Consignee")
		Readme = enchiasp.ChkNull(Rs("Readme"))
	End If
	Rs.Close
	'-- 开始添加交易明细表
	If PayDone = 0 Then
		SQL = "SELECT * FROM ECCMS_Account WHERE (AccountID is null)"
		Rs.Open SQL,Conn,1,3
		Rs.addnew
			Rs("payer").Value = Consignee
			Rs("payee").Value = enchiasp.CheckRequest(enchiasp.SiteName,20)
			Rs("product").Value = "网上购物"
			Rs("Amount").Value = 1
			Rs("unit").Value = "次"
			Rs("price").Value = totalmoney
			Rs("TotalPrices").Value = totalmoney
			Rs("DateAndTime").Value = Now()
			Rs("Accountype").Value = 0
			Rs("Explain").Value = Readme
			Rs("Reclaim").Value = 0
		Rs.update
		Rs.Close:Set Rs = Nothing
	End If
	Succeed("<li>恭喜您！订单处理成功。请赶快给用户发货去吧!</li>")
End Sub
%>