<!--#include file="setup.asp"-->
<%
'=====================================================================
' 软件名称：恩池网站管理系统
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Admin_header
%>
<br>
<table border="0" cellspacing="1" cellpadding="5" class=tableBorder align=center style="width:90%">
<tr><th height="24"><%=enchiasp.SiteName%> --- 管理帮助</th></tr>
<tr><td width="100%" class=class=Forumrow >
<script>
document.write(opener.txtRun.value)
</script>
</td></tr>
<tr>
<td height="22" align=center class=FORUMROWHIGHLIGHT>
<input type="button" name="close" value="[关  闭]" onclick="window.close()" class=Button>
</td>
</tr>
</table>
<%
Admin_footer
CloseConn
%>