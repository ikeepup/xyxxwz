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
Dim Rs,SQL
SQL = "select * from ECCMS_User where username='" & enchiasp.membername & "'"
Set Rs = enchiasp.Execute(SQL)
If Rs("usermsg") > 0 Then
	Response.Write "<bgsound src=""images/mail.wav"" border=0>"
End If
Call InnerLocation("��Ա������ҳ - ��ӭ <font color=red>" & enchiasp.membername & "</font> ��¼�������")
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=Usertableborder>
<tr>
	<th colspan=2>�û�������� -- ��ҳ</th>            
</tr>
<tr>
	<td width="50%" class=Usertablerow1><b class=userfont2>�û����ƣ�</b><font color=red><%=enchiasp.membername%></font></td>
	<td width="50%" class=Usertablerow1><b class=userfont2>�û���ݣ�</b><font color=red><%=enchiasp.membergroup%></font></td>
</tr>
<tr>
	<td class=Usertablerow2><b class=userfont2>�û��ǳƣ�</b><%=enchiasp.menbernickname%></td>
	<td class=Usertablerow2><b class=userfont2>��ʵ������</b><%=Rs("TrueName")%></td>
</tr>
<tr>
	<td class=Usertablerow1><b class=userfont2>�˻���</b>��<%=FormatNumber(Rs("usermoney"),,-1)%> Ԫ</td>
	<td class=Usertablerow1><b class=userfont2>�Ѿ����ѣ�</b>��<%=FormatNumber(Rs("prepaid"),,-1)%> Ԫ</td>
</tr>
<tr>
	<td class=Usertablerow2><b class=userfont2>���õ�����</b><%=Rs("userpoint")%> ��</td>
	<td class=Usertablerow2><b class=userfont2>�û����飺</b><%=Rs("experience")%></td>
</tr>
<tr>
	<td class=Usertablerow1><b class=userfont2>�û�������</b><%=Rs("charm")%></td>
	<td class=Usertablerow1><b class=userfont2>ע�����ڣ�</b><%=Rs("JoinTime")%></td>
</tr>
<tr>
	<td class=Usertablerow2><b class=userfont2>��Ա���ͣ�</b><%If Rs("UserGrade") = 999 Then
				Response.Write "����Ա"
			Else
				If Rs("UserClass") = 0 Then
					Response.Write "�Ƶ��Ա"
				ElseIf Rs("UserClass") = 1 Then
					Response.Write "��ʱ��Ա"
				Else
					Response.Write "���ڻ�Ա"
				End If
			End If%></td>
	<td class=Usertablerow2><b class=userfont2>����ʱ�䣺</b><%=Rs("ExpireTime")%></td>
</tr>
<tr>
	<td class=Usertablerow1><b class=userfont2>�ϴε�¼ʱ�䣺</b><%=Request.Cookies("enchiasp_net")("LastTimeDate")%></td>
	<td class=Usertablerow1><b class=userfont2>�ϴε�¼IP��</b><%=Request.Cookies("enchiasp_net")("LastTimeIP")%></td>
</tr>
<tr>
	<td class=Usertablerow2><b class=userfont2>��¼������</b><%=Rs("userlogin")%> ��</td>
	<td class=Usertablerow2>
<%
If Rs("UserClass") > 0 Then
	Response.Write "<b class=userfont2>������ʾ��</b>"
	If DateDiff("D", CDate(Rs("ExpireTime")), Now()) < 0 Then
		Response.Write "�����˺�ʹ��ʱ�޻��� <font color=red><b>"
		Response.Write DateDiff("D", Now(), CDate(Rs("ExpireTime")))
		Response.Write "</b></font> ��"
	Else
		Response.Write "<font color=red>�����˺��ѹ���,����ϵ����Ա��</font>"
	End If
End If
%>
	
	</td>
</tr>
<tr>
	<td class=Usertablerow1><a href="usersms.asp">�ҵ��ռ��� 
<%
If Rs("usermsg") > 0 Then
	Response.Write "(<b class=userfont1>" & Rs("usermsg") & "</b>)"
Else
	Response.Write "(<font color=""#808080"">" & Rs("usermsg") & "</font>)"
End If
%>
	</a></td>
	<td class=Usertablerow1><a href="changepsw.asp">�޸�����</a> | <a href="changeinfo.asp">�޸�����</a> | <a href="addmoney.asp">��Ա��ֵ</a></td>            
</tr>

</table>
<%
Rs.Close:Set Rs = Nothing
%>
<!--#include file="foot.inc"-->