<!--#include file="about/config.asp"-->
<!--#include file="inc/chkinput.asp"-->
<%
	dim rs,sql
	If enchiasp.CheckPost = False Then
		ErrMsg = ErrMsg + "<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ��</li>"
		FoundErr = True
	End If
	'If enchiasp.IsValidStr(Request.Form("xm")) or enchiasp.IsValidStr(Request.Form("lxdh"))  Then
		'ErrMsg = ErrMsg + "���ύ�������к��зǷ��ַ�\n"
		'Founderr = True
	'End If
	
	If trim(request.form("bh"))="" Then
		ErrMsg = ErrMsg + "�������ʱ������\n"
		Founderr = True
	end if
	
	If trim(request.form("cph"))="" Then
		ErrMsg = ErrMsg + "�����복�ƺ�\n"
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
	SQL = "SELECT * FROM ECCMS_zb where bh='"& request.form("bh") &"' and cph='"& request.form("cph") &"'   "
	Rs.Open SQL,Conn,1,1
	if rs.eof then
		Call OutputScript("δ�ܲ鵽�ʱ����ţ�","index.asp")

	else
		Call OutputScript("�����복�ƺ����","index.asp")
	
	end if
	Rs.Close:Set Rs = Nothing
	CloseConn
%>