<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="head.inc"-->
<%
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
Call InnerLocation("�û�����")

ErrMsg = ErrMsg + "<li>�Բ�����û�в鿴��ҳ��Ȩ�ޣ�����ʲô��������ϵ����Ա��</li>"
Founderr = True
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If
%>
<!--#include file="foot.inc"-->