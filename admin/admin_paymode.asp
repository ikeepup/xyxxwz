<!--#include file="setup.asp"-->
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
Dim Action
If Not ChkAdmin("adminpaymode") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "del"
	Call DelPaymode
Case "save"
	Call SavePaymode
Case "modify"
	Call ModifyPaymode
Case "edit"
	Call EditPaymode
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
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th colspan=4 noWrap>增加付款方式</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action=?action=save>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow1 align=right noWrap><b>付款标题：</b></td>"
	Response.Write "	  <td class=TableRow1><input type=text name=modename size=35></td>"
	Response.Write "	  <td class=TableRow2 rowspan=6 align=right noWrap><b>其它说明：</b><br>支持HTML</td>"
	Response.Write "	  <td class=TableRow1 rowspan=6><textarea name=readme rows=10 cols=45></textarea></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow2 align=right><b>开户银行：</b></td>"
	Response.Write "	  <td class=TableRow2><input type=text name=site size=35></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow1 align=right><b>银行帐号：</b></td>"
	Response.Write "	  <td class=TableRow1><input type=text name=code size=35></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow2 align=right><b>收 款 人：</b></td>"
	Response.Write "	  <td class=TableRow2><input type=text name=payee size=20></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow1 align=right noWrap><b>银行 URL：</b></td>"
	Response.Write "	  <td class=TableRow1><input type=text name=url size=35 value='http://'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow2 align=center colspan=4><input type=submit value="" 确认提交 "" class=Button></td>"
	Response.Write "	</tr><form>"
	Response.Write "</table>"
	Response.Write "<br>"
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th colspan=3>付款方式</th>"
	Response.Write "	</tr>"
	Set Rs = enchiasp.Execute("SELECT modeid,modename,site,code,payee,url,readme FROM ECCMS_Paymode ORDER BY modeid")
	If Rs.BOF And Rs.EOF Then
		Response.Write "<tr><td align=center colspan=3 class=TableRow1>还没有添加付款方式！</td></tr>"
	Else
		Do While Not Rs.EOF
			Response.Write "	<tr>"
			Response.Write "	  <td colspan=3 class=TableTitle><a href='"& Rs("url") &"' target=_blank class=showtitle>"& Rs("modename") &"</a></td>"
			Response.Write "	</tr>"
			Response.Write "	<tr>"
			Response.Write "	  <td class=TableRow1 width='10%' align=right noWrap><b>开户银行：</b></td>"
			Response.Write "	  <td class=TableRow1 width='35%'>" & Rs("site") & "</td>"
			Response.Write "	  <td class=TableRow1 width='55%' vAlign=top rowspan=4><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;"& Rs("readme") &"</td>"
			Response.Write "	</tr>"
			Response.Write "	<tr>"
			Response.Write "	  <td class=TableRow2 align=right><b>银行帐号：</b></td>"
			Response.Write "	  <td class=TableRow2>"& Rs("code") &"</td>"
			Response.Write "	</tr>"
			Response.Write "	<tr>"
			Response.Write "	  <td class=TableRow1 align=right><b>收 款 人：</b></td>"
			Response.Write "	  <td class=TableRow1>"& Rs("payee") &"</td>"
			Response.Write "	</tr>"
			Response.Write "	<tr>"
			Response.Write "	  <td class=TableRow2 align=right><b>管理选项：</b></td>"
			Response.Write "	  <td class=TableRow2 align=center><a href='?action=edit&modeid="& Rs("modeid") &"'>修 改</a> | <a href='?action=del&modeid="& Rs("modeid") &"' onclick=""return confirm('您确定要删除此付款方式吗?')"">删 除</a></td>"
			Response.Write "	</tr>"
			Rs.movenext
		Loop
	End If
	Rs.Close:Set Rs = Nothing
	Response.Write "</table>"
End Sub

Sub EditPaymode()
	If Not IsNumeric(Request("modeid")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的系统参数，ID错误！</li>"
		Exit Sub
	End If
	Set Rs = enchiasp.Execute("SELECT modeid,modename,site,code,payee,url,readme FROM ECCMS_Paymode WHERE modeid="& CLng(Request("modeid")))
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=TableBorder>"
	Response.Write "	<tr>"
	Response.Write "	  <th colspan=4 noWrap>修改付款方式</th>"
	Response.Write "	</tr>"
	Response.Write "	<form name=myform method=post action=?action=modify>"
	Response.Write "	<input type=hidden name=modeid value='"& Rs("modeid") &"'>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow1 align=right noWrap><b>付款标题：</b></td>"
	Response.Write "	  <td class=TableRow1><input type=text name=modename size=35 value='"& Rs("modename") &"'></td>"
	Response.Write "	  <td class=TableRow2 rowspan=6 align=right noWrap><b>其它说明：</b><br>支持HTML</td>"
	Response.Write "	  <td class=TableRow1 rowspan=6><textarea name=readme rows=10 cols=45>"& Server.HTMLEncode(Rs("readme")) &"</textarea></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow2 align=right><b>开户银行：</b></td>"
	Response.Write "	  <td class=TableRow2><input type=text name=site size=35 value='"& Rs("site") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow1 align=right><b>银行帐号：</b></td>"
	Response.Write "	  <td class=TableRow1><input type=text name=code size=35 value='"& Rs("code") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow2 align=right><b>收 款 人：</b></td>"
	Response.Write "	  <td class=TableRow2><input type=text name=payee size=20 value='"& Rs("payee") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow1 align=right noWrap><b>银行 URL：</b></td>"
	Response.Write "	  <td class=TableRow1><input type=text name=url size=35 value='"& Rs("url") &"'></td>"
	Response.Write "	</tr>"
	Response.Write "	<tr>"
	Response.Write "	  <td class=TableRow2 align=center colspan=4><input type=submit value="" 确认提交 "" class=Button></td>"
	Response.Write "	</tr><form>"
	Response.Write "</table>"
	Response.Write "<br>"
	Rs.Close:Set Rs = Nothing
End Sub

Sub SavePaymode()
	SQL = "INSERT INTO ECCMS_Paymode (modename,site,code,payee,url,readme,modetype) VALUES ('"& enchiasp.CheckStr(Request("modename")) &"','"& enchiasp.CheckStr(Request("site")) &"','"& enchiasp.CheckStr(Request("code")) &"','"& enchiasp.CheckStr(Request("payee")) &"','"& enchiasp.CheckStr(Request("url")) &"','"& enchiasp.CheckStr(Request("readme")) &"',0)"
	enchiasp.Execute(SQL)
	Succeed("<li>恭喜您！添加付款方式成功。</li>")
End Sub

Sub ModifyPaymode()
	If Not IsNumeric(Request("modeid")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的系统参数，ID错误！</li>"
		Exit Sub
	End If
	SQL = "UPDATE [ECCMS_Paymode] SET modename='"& Request("modename") &"',site='"& Request("site") &"',code='"& Request("code") &"',payee='"& Request("payee") &"',url='"& Request("url") &"',readme='"& Request("readme") &"',modetype=0 WHERE modeid="& CLng(Request("modeid"))
	enchiasp.Execute(SQL)
	Succeed("<li>恭喜您！修改付款方式成功。</li>")
End Sub

Sub DelPaymode()
	If IsNumeric(Request("modeid")) Then
		enchiasp.Execute("DELETE FROM [ECCMS_Paymode] WHERE modeid="& CLng(Request("modeid")))
		Response.redirect (Request.ServerVariables("HTTP_REFERER"))
	Else
		FoundErr = True
		ErrMsg = ErrMsg + "<li>错误的系统参数，ID错误！</li>"
		Exit Sub
	End If
End Sub
%>
