<!--#include file="config.asp"-->
<!--#include file="../inc/chkinput.asp"-->
<%
	dim rs,sql
	If enchiasp.CheckPost = False Then
		ErrMsg = ErrMsg + "<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ��</li>"
		FoundErr = True
	End If
	'If not(enchiasp.IsValidStr(Request.Form("name")) and enchiasp.IsValidStr(Request.Form("birthday")) and enchiasp.IsValidStr(Request.Form("school")) and enchiasp.IsValidStr(Request.Form("studydegree")) and enchiasp.IsValidStr(Request.Form("specialty")) and enchiasp.IsValidStr(Request.Form("gradyear")) and enchiasp.IsValidStr(Request.Form("telephone")) and enchiasp.IsValidStr(Request.Form("email")) and enchiasp.IsValidStr(Request.Form("address")) and enchiasp.IsValidStr(Request.Form("ability")) and enchiasp.IsValidStr(Request.Form("resumes")) ) Then
		'ErrMsg = ErrMsg + "���ύ�������к��зǷ��ַ�\n"
		'Founderr = True
	'End If
	
	If trim(request.form("name"))="" Then
		ErrMsg = ErrMsg + "��û��������\n"
		Founderr = True
	end if
	If trim(request.form("sex"))="" Then
		ErrMsg = ErrMsg + "��û�����Ա�\n"
		Founderr = True
	end if
	If trim(request.form("birthday"))="" Then
		ErrMsg = ErrMsg + "��û�����������\n"
		Founderr = True
	else
		if not isdate(trim(request.form("birthday"))) then
			ErrMsg = ErrMsg + "�����������ݸ�ʽ��������������ȷ����������\n"
			Founderr = True
		end if
	end if
	If trim(request.form("marry"))="" Then
		ErrMsg = ErrMsg + "��û�������״��\n"
		Founderr = True
	end if
	If trim(request.form("school"))="" Then
		ErrMsg = ErrMsg + "��û�����ҵԺУ\n"
		Founderr = True
	end if
	If trim(request.form("studydegree"))="" Then
		ErrMsg = ErrMsg + "��û����ѧ ��\n"
		Founderr = True
	end if
	If trim(request.form("specialty"))="" Then
		ErrMsg = ErrMsg + "��û����ר ҵ\n"
		Founderr = True
	end if
	If trim(request.form("gradyear"))="" Then
		ErrMsg = ErrMsg + "��û�����ҵʱ��\n"
		Founderr = True
	else
		if not isdate(trim(request.form("gradyear"))) then
			ErrMsg = ErrMsg + "��ҵʱ�����ݸ�ʽ��������������ȷ����������\n"
			Founderr = True
		end if

	end if
	If trim(request.form("telephone"))="" Then
		ErrMsg = ErrMsg + "��û����� ��\n"
		Founderr = True
	else
		if not IsValidTel(trim(request.form("telephone"))) then
			ErrMsg = ErrMsg + "��������ȷ�ĵ绰\n"
			Founderr = True
		end if
	end if
	If trim(request.form("email"))="" Then
		ErrMsg = ErrMsg + "��û����E-mail\n"
		Founderr = True
	else
		If Not IsValidEmail(Request.Form("email")) Then
			ErrMsg = ErrMsg + "����ȷ��д��������\n"
			Founderr = True
		End If	
	end if
	If trim(request.form("address"))="" Then
		ErrMsg = ErrMsg + "��û������ϵ��ַ\n"
		Founderr = True
	end if
	If trim(request.form("ability"))="" Then
		ErrMsg = ErrMsg + "��û����ˮƽ������\n"
		Founderr = True
	end if
	If trim(request.form("resumes"))="" Then
		ErrMsg = ErrMsg + "��û������˼���\n"
		Founderr = True
	end if
	If Founderr = True Then
		Call OutAlertScript(ErrMsg)
		response.end
	End If
	Call PreventRefresh  '��ˢ��
	'���ݿ����¹رգ���
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
	Call OutputScript("���ļ����Ѿ��ύ�ɹ����Ժ����ǻ����ȡ����ϵ��лл���Ĳ��룡","index.asp")
%>