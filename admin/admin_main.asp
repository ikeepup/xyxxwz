<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
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
Dim theInstalledObjects(4)
theInstalledObjects(0) = "Persits.Jpeg"
theInstalledObjects(1) = "Scripting.FileSystemObject"
theInstalledObjects(2) = "adodb.connection"

theInstalledObjects(3) = "JMail.SMTPMail"
theInstalledObjects(4) = "CDONTS.NewMail"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<title>����ҳ��</title>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<link rel="stylesheet" href="style.css" type="text/css">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="5" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<table cellpadding="2" cellspacing="1" border="0" width="96%" class="tableBorder" align=center>
<tr>
<th colspan=2 height=25>ϵͳ��Ϣͳ��</th>
</tr>
<tr>
<td class=BodyTitle colspan=2 height=25>ϵͳ��Ϣͳ�ơ���
<%
On Error Resume Next
Response.Write "�������أ�<font color=red><b>"
SQL = "SELECT SUM(DayHits) FROM ECCMS_SoftList WHERE isAccept>0 And Datediff('d',HitsTime,Now())=0"
Set Rs = enchiasp.Execute(SQL)
If Rs.BOF And Rs.EOF Then
	Response.Write 0
Else
	Response.Write enchiasp.CheckNumeric(Rs(0))
End If
Set Rs = Nothing
Response.Write "</b></font> �Ρ�"
Response.Write "���������<font color=red><b>"
SQL = "SELECT SUM(DayHits) FROM ECCMS_Article WHERE isAccept>0 And Datediff('d',HitsTime,Now())=0"
Set Rs = enchiasp.Execute(SQL)
If Rs.BOF And Rs.EOF Then
	Response.Write 0
Else
	Response.Write enchiasp.CheckNumeric(Rs(0))
End If
Set Rs = Nothing
Response.Write "</b></font> �Ρ�"
Response.Write "���չۿ���<font color=red><b>"
SQL = "SELECT SUM(DayHits) FROM ECCMS_FlashList WHERE isAccept>0 And Datediff('d',HitsTime,Now())=0"
Set Rs = enchiasp.Execute(SQL)
If Rs.BOF And Rs.EOF Then
	Response.Write 0
Else
	Response.Write enchiasp.CheckNumeric(Rs(0))
End If
Set Rs = Nothing
Response.Write "</b></font> �Ρ�"
Response.Write "��ǰ���ߣ�<font color=red><b>"
SQL = "SELECT COUNT(id) FROM ECCMS_Online WHERE DateDIff('s',lastTime,Now()) < 20*60"
Set Rs = enchiasp.Execute(SQL)
If Rs.BOF And Rs.EOF Then
	Response.Write 0
Else
	Response.Write enchiasp.CheckNumeric(Rs(0))
End If
Set Rs = Nothing
Response.Write "</b></font> �ˡ�"
%>
</td>
</tr>
<tr>
<td width="50%"  class="TableRow2" height=23>���������ͣ�<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
<td width="50%" class="TableRow1">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr>
<td width="50%" class="TableRow2" height=23>վ������·����<%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
<td width="50%" class="TableRow1">AspJpeg�����
<%If Not IsObjInstalled(theInstalledObjects(0)) Then%><font color="#FF0066"><b>��</b></font><%else%><b>��</b><%end if%>
</td>

</td>

</tr>
<tr>
<td width="50%" class="TableRow2" height=23>FSO�ı���д��<%If Not IsObjInstalled(theInstalledObjects(1)) Then%><font color="#FF0066"><b>��</b></font><%else%><b>��</b><%end if%></td>
<td width="50%" class="TableRow1">���ݿ�ʹ�ã�<%If Not IsObjInstalled(theInstalledObjects(2)) Then%><font color="#FF0066"><b>��</b></font><%else%><b style="color:blue"><%If IsSqlDataBase = 1 Then%>MS SQL<%else%>ACCESS<%end if%></b><%end if%></td>
</tr>
<tr>
<td width="50%" class="TableRow2" height=23>Jmail���֧�֣�<%If Not IsObjInstalled(theInstalledObjects(3)) Then%><font color="#FF0066"><b>��</b></font><%else%><b>��</b><%end if%></td>
<td width="50%" class="TableRow1">CDONTS���֧�֣�<%If Not IsObjInstalled(theInstalledObjects(4)) Then%><font color="#FF0066"><b>��</b></font><%else%><b>��</b><%end if%></td>
</tr>
</tr>
<tr><td colspan=2 class="TableRow1" height=25><B>��ݲ���ѡ�</B> <a href=admin_config.asp>��վ��������</a>&nbsp;
<a href=CleanCache.asp>�ؽ�ϵͳ����</a>&nbsp;
<a href=admin_user.asp>��Ա����</a>&nbsp;
<a href=admin_online.asp>��������ͳ��</a>&nbsp;
<a href=admin_template.asp>ģ����ʽ�ܹ���</a></td>
</tr>
</table>
<BR>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
      <tr>
	<th colspan=2 height=25>��վ����ϵͳ˵��</th>
	</tr>
	<tr>
	<td width="60" class="TableRow2" height=23>�߼�����Ա</td>
	<td class="TableRow1" style="LINE-HEIGHT: 150%">ӵ������Ȩ�ޡ�<BR>��һ��ʹ��ʱ�뵽<font color=Red>�û�����</font>-<a href=admin_password.asp target=main><font color=Red>����Ա�����޸�</font></A>�������ù�������</td>
      </tr>
      <tr>
	<td class="TableRow2" height=23>��ͨ����Ա</td>
	<td class="TableRow1" style="LINE-HEIGHT: 150%">��Ҫ�߼�����Ա����Ȩ�ޡ�<BR>ע:<a href=admin_master.asp><font color=Red>�ڹ���Ա��������Ȩ��!</font></A><br>
	   </td>
      </tr>
      <tr>
	<td class="TableRow2" height=23>ʹ������</td>
	<td class="TableRow1" style="LINE-HEIGHT: 150%">��һ��ʹ��<font color=Red>�������</font>��վ����ϵͳ<BR>
	 �����߹������˵��еġ�<a href=admin_config.asp><font color=Red>��������-��������</A></font>��<BR>�������վ��Ϣ��һЩ��վ���ò����������á�
	   </td>
      </tr>
    </table>
<BR>
    <BR>

    <table cellpadding="3" cellspacing="1" border="0" width="96%" class="tableBorder" align=center>
      <tr>
	<th colspan=2 height=25>���������վ����ϵͳ����</th>

      <tr>
	<td width="60" class="TableRow2" height=23>��������</td>
	<td class="TableRow1">�˳��ж�������Ƽ��������޹�˾</td>
      </tr>
      <tr>
	<td class="TableRow2" height=23>��ϵ��ʽ</td>
	<td class="TableRow1">E_mail��liuyunfan@163.com<br>QQ��21556923<br>�绰��0359-8698845<br>
	</td>
      </tr>
      <tr>
	<td class="TableRow2" height=23>������ҳ</td>
	<td class="TableRow1"><a href="http://www.enchi.com.cn/" target=_blank>www.enchi.com.cn</a>
	</td>
      </tr>
    </table>
<%
If CInt(enchiasp.VersionID) <> 0 Then
	enchiasp.Execute("UPDATE ECCMS_Config SET VersionID=0")
End If
Admin_footer
CloseConn
%>






















