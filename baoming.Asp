<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/cls_public.asp"-->
<!--#include file="inc/classmenu.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<%
'=====================================================================
' 软件名称：恩池网站管理系统--报名
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：94022511
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Dim HtmlContent,tempstr
enchiasp.LoadTemplates 9999, 9, 0

HtmlContent = enchiasp.HtmlContent


HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
HtmlContent = Replace(HtmlContent, "{$PageTitle}", "活动报名")



select case request.querystring("id")
case "1"
HtmlContent = Replace(HtmlContent,"{$TempContent}", enchiasp.HtmlSetting(1))
tempstr="自驾游1"
case "2"
HtmlContent = Replace(HtmlContent,"{$TempContent}", enchiasp.HtmlSetting(2))
tempstr="自驾游2"
case "3"
HtmlContent = Replace(HtmlContent,"{$TempContent}", enchiasp.HtmlSetting(3))
tempstr="自驾游3"
case "4"
HtmlContent = Replace(HtmlContent,"{$TempContent}", enchiasp.HtmlSetting(4))
tempstr="自驾游4"
case "5"
HtmlContent = Replace(HtmlContent,"{$TempContent}", enchiasp.HtmlSetting(5))
tempstr="平面模特"
case else
response.write "请勿非法访问"
response.end
end select
HtmlContent = Replace(HtmlContent,"{$id}", request.querystring("id"))



HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
HtmlContent = HTML.ReadAnnounceList(HtmlContent)
HtmlContent = HTML.ReadFriendLink(HtmlContent)

If enchiasp.CheckStr(LCase(Request.querystring("action"))) = "save" Then
	Call Savesq
End If

Response.Write HtmlContent

Sub Savesq()
	On Error Resume Next
	'不允许匿名发表

	
	If enchiasp.CheckPost = False Then
		ErrMsg = ErrMsg + "<li>您提交的数据不合法，请不要从外部提交。</li>"
		FoundErr = True
	End If
	If Trim(Request.Form("user_name")) = "" Then
		ErrMsg = ErrMsg + "姓名不能为空\n"
		Founderr = True
	End If
	If enchiasp.IsValidStr(Request.Form("user_name")) = False Then
		ErrMsg = ErrMsg + "姓名中含有非法字符\n"
		Founderr = True
	End If

	
	If Trim(Request.Form("shouji")) = "" Then
		ErrMsg = ErrMsg + "手机不能为空\n"
		Founderr = True
	End If
	If Trim(Request.Form("address")) = "" Then
		ErrMsg = ErrMsg + "联系地址不能为空\n"
		Founderr = True
	End If
	
	If Founderr = True Then
		Call OutAlertScript(ErrMsg)
		Exit Sub
	End If
	Call PreventRefresh  '防刷新

	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_baoming WHERE (id is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("baomingtypeid") = request.querystring("id")
		Rs("baomingtype") = tempstr
		Rs("xingming") =  enchiasp.ChkFormStr(Request.Form("user_name"))
		Rs("nianling") =  enchiasp.ChkFormStr(Request.Form("nianling"))
		Rs("xingbie") =  enchiasp.ChkFormStr(Request.Form("sex"))
		Rs("shouji") =  enchiasp.ChkNumeric(enchiasp.ChkFormStr(Request.Form("shouji")))
		Rs("dianhua") =  enchiasp.ChkFormStr(Request.Form("tel"))
		Rs("email") = enchiasp.ChkFormStr(Request.Form("email"))
		Rs("dizhi") =  enchiasp.ChkFormStr(Request.Form("address"))
		Rs("youbian") =  enchiasp.ChkFormStr(Request.Form("code"))
		Rs("jianli") =  enchiasp.ChkFormStr(Request.Form("jianli"))
	Rs.update
	Rs.Close:Set Rs = Nothing
	Response.Redirect("index.asp")
	
End Sub
Set HTML = Nothing
CloseConn

%>