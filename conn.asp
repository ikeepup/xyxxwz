
<%
'
'=====================================================================
' ������ƣ�������վ����ϵͳ
' ��ǰ�汾��enchicms Version 3.0.0
' �������ڣ�2005-03-25
' �ٷ���վ���˳��ж�������Ƽ��������޹�˾(www.enchi.com.cn) 
' ����֧�֣����Ʒ�
' ���䣺liuyunfan@163.com
' QQ��21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
'Option Explicit
Dim startime,Conn,db,comDomain 
Response.Buffer = True
startime = Timer()
Dim NowString
Dim ConnStr
'������ȷ����

'comDomain = "www.carnurse.com" '����COM���� 
'If Request.ServerVariables("SERVER_NAME") <> comDomain Then '����������������com������
            'Response.Status = 301 '��ʾ״̬�л���301
            'Response.AddHeader "Location","http://" & comDomain & "/"&Request.ServerVariables("HTTP_X_REWRITE_URL") '��ͷ�����Location�ֶΣ�ʹ�õ�ǰ�������·������COM��������ϳ��µĵ�ַ��
'End If


'�������ݿ����1ΪSQL���ݿ⣬0ΪAccess���ݿ�
const isSqlDataBase = 0

If IsSqlDataBase = 1 Then
	'-----------------------SQL���ݿ����Ӳ���---------------------------------------
	Dim SqlDatabaseName,SqlPassword,SqlUsername,SqlLocalName
	NowString = "GetDate()"
	SqlDatabaseName = "Ec_cms_tongda"     '���ݿ���
	SqlUsername = "sa"          '�û���
	SqlPassword = "avdswx"          '�û�����
	SqlLocalName = "."        '��������������local�������IP��
	ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " &  SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	'-------------------------------------------------------------------------------
Else
	'-----------------------ACCESS���ݿ�����----------------------------------------
	NowString = "Now()"
	'ACCESS���ݿ�����,��ʹ�ø�·�����߾���·��
	db = "/database/#ECCMS.mdb"
	Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ChkMapPath(db)
	'-------------------------------------------------------------------------------
End If

Sub ConnectionDatabase()
	On Error Resume Next
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Connstr
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "���ݿ����ӳ���"
		Response.End
	End If
	
	
	
End Sub

Dim DBPath
'-- �ɼ����ݿ�����·��
DBPath = "/database/#EC_Collection.mdb"		'-- ���ø����·��

Sub CloseConn()
	On Error Resume Next
	If IsObject(Conn) Then
		Conn.Close
		Set Conn = Nothing
	End If
	Set enchiasp = Nothing
End Sub
%>