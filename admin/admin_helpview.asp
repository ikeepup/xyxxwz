<!--#include file="setup.asp"-->
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
Admin_header
%>
<br>
<table border="0" cellspacing="1" cellpadding="5" class=tableBorder align=center style="width:90%">
<tr><th height="24"><%=enchiasp.SiteName%> --- �������</th></tr>
<tr><td width="100%" class=class=Forumrow >
<script>
document.write(opener.txtRun.value)
</script>
</td></tr>
<tr>
<td height="22" align=center class=FORUMROWHIGHLIGHT>
<input type="button" name="close" value="[��  ��]" onclick="window.close()" class=Button>
</td>
</tr>
</table>
<%
Admin_footer
CloseConn
%>