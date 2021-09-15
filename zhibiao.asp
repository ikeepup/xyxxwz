<!--#include file="about/config.asp"-->
<!--#include file="inc/chkinput.asp"-->
<%
	dim rs,sql
	If enchiasp.CheckPost = False Then
		ErrMsg = ErrMsg + "<li>您提交的数据不合法，请不要从外部提交。</li>"
		FoundErr = True
	End If
	'If enchiasp.IsValidStr(Request.Form("xm")) or enchiasp.IsValidStr(Request.Form("lxdh"))  Then
		'ErrMsg = ErrMsg + "您提交的数据中含有非法字符\n"
		'Founderr = True
	'End If
	
	If trim(request.form("bh"))="" Then
		ErrMsg = ErrMsg + "请输入质保卡编号\n"
		Founderr = True
	end if
	
	If trim(request.form("cph"))="" Then
		ErrMsg = ErrMsg + "请输入车牌号\n"
		Founderr = True
	end if
		If Founderr = True Then
		Call OutAlertScript(ErrMsg)
		response.end
	End If
	Call PreventRefresh  '防刷新
	'数据库重新关闭，打开
	CloseConn
	ConnectionDatabase
	Set Rs = server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_zb where bh='"& request.form("bh") &"' and cph='"& request.form("cph") &"'   "
	Rs.Open SQL,Conn,1,1
	if rs.eof then
		Call OutputScript("未能查到质保卡号！","index.asp")

	else
		Call OutputScript("卡号与车牌号相符","index.asp")
	
	end if
	Rs.Close:Set Rs = Nothing
	CloseConn
%>