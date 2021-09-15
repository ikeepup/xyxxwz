<!--#include file="config.asp"-->
<!--#include file="../inc/chkinput.asp"-->
<%
	dim rs,sql
	If enchiasp.CheckPost = False Then
		ErrMsg = ErrMsg + "<li>您提交的数据不合法，请不要从外部提交。</li>"
		FoundErr = True
	End If
	'If not(enchiasp.IsValidStr(Request.Form("name")) and enchiasp.IsValidStr(Request.Form("birthday")) and enchiasp.IsValidStr(Request.Form("school")) and enchiasp.IsValidStr(Request.Form("studydegree")) and enchiasp.IsValidStr(Request.Form("specialty")) and enchiasp.IsValidStr(Request.Form("gradyear")) and enchiasp.IsValidStr(Request.Form("telephone")) and enchiasp.IsValidStr(Request.Form("email")) and enchiasp.IsValidStr(Request.Form("address")) and enchiasp.IsValidStr(Request.Form("ability")) and enchiasp.IsValidStr(Request.Form("resumes")) ) Then
		'ErrMsg = ErrMsg + "您提交的数据中含有非法字符\n"
		'Founderr = True
	'End If
	
	If trim(request.form("name"))="" Then
		ErrMsg = ErrMsg + "您没输入姓名\n"
		Founderr = True
	end if
	If trim(request.form("sex"))="" Then
		ErrMsg = ErrMsg + "您没输入性别\n"
		Founderr = True
	end if
	If trim(request.form("birthday"))="" Then
		ErrMsg = ErrMsg + "您没输入出生日期\n"
		Founderr = True
	else
		if not isdate(trim(request.form("birthday"))) then
			ErrMsg = ErrMsg + "出生日期数据格式有误，请您输入正确的日期数据\n"
			Founderr = True
		end if
	end if
	If trim(request.form("marry"))="" Then
		ErrMsg = ErrMsg + "您没输入婚姻状况\n"
		Founderr = True
	end if
	If trim(request.form("school"))="" Then
		ErrMsg = ErrMsg + "您没输入毕业院校\n"
		Founderr = True
	end if
	If trim(request.form("studydegree"))="" Then
		ErrMsg = ErrMsg + "您没输入学 历\n"
		Founderr = True
	end if
	If trim(request.form("specialty"))="" Then
		ErrMsg = ErrMsg + "您没输入专 业\n"
		Founderr = True
	end if
	If trim(request.form("gradyear"))="" Then
		ErrMsg = ErrMsg + "您没输入毕业时间\n"
		Founderr = True
	else
		if not isdate(trim(request.form("gradyear"))) then
			ErrMsg = ErrMsg + "毕业时间数据格式有误，请您输入正确的日期数据\n"
			Founderr = True
		end if

	end if
	If trim(request.form("telephone"))="" Then
		ErrMsg = ErrMsg + "您没输入电 话\n"
		Founderr = True
	else
		if not IsValidTel(trim(request.form("telephone"))) then
			ErrMsg = ErrMsg + "请输入正确的电话\n"
			Founderr = True
		end if
	end if
	If trim(request.form("email"))="" Then
		ErrMsg = ErrMsg + "您没输入E-mail\n"
		Founderr = True
	else
		If Not IsValidEmail(Request.Form("email")) Then
			ErrMsg = ErrMsg + "请正确填写您的邮箱\n"
			Founderr = True
		End If	
	end if
	If trim(request.form("address"))="" Then
		ErrMsg = ErrMsg + "您没输入联系地址\n"
		Founderr = True
	end if
	If trim(request.form("ability"))="" Then
		ErrMsg = ErrMsg + "您没输入水平与能力\n"
		Founderr = True
	end if
	If trim(request.form("resumes"))="" Then
		ErrMsg = ErrMsg + "您没输入个人简历\n"
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
	SQL = "SELECT * FROM ECCMS_Jobbook  "
	Rs.Open SQL,Conn,1,3
	
	Rs.Addnew
	
	rs("jobid")=Request.Form("jobid")
	rs("jobname")=request.form("jobname")
	rs("name")=request.form("name")
	rs("sex")=request.form("sex")
	rs("birthday")=request.form("birthday")
	rs("marry")=request.form("marry")
	rs("school")=request.form("school")
	rs("studydegree")=request.form("studydegree")
	rs("specialty")=request.form("specialty")
	rs("gradyear")=request.form("gradyear")
	rs("telephone")=request.form("telephone")
	rs("email")=request.form("email")
	rs("address")=request.form("address")
	rs("ability")=Html2Ubb(request.form("ability"))
	rs("resumes")=Html2Ubb(request.form("resumes"))
	rs("riqi")=now()
	rs.update
	Rs.Close:Set Rs = Nothing
	CloseConn
	Call OutputScript("您的简历已经提交成功，稍后我们会和您取得联系，谢谢您的参与！","index.asp")
%>