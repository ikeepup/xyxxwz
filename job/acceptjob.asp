<!--#include file="config.asp"-->
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
Dim Rs, SQL
enchiasp.LoadTemplates ChannelID, 3, 0
HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
'--Ƶ��Ŀ¼
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","�ύ����")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
if Request("job")="" then
	Call OutputScript("����Ĳ������벻Ҫ��������һЩ������","index.asp")
end if
Set Rs = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM ECCMS_job where id="& Request("job") &" "
Set Rs = enchiasp.Execute(SQL)

If Rs.BOF And Rs.EOF Then
	response.write "<p align=center>�����˴���,</p>"
Else

	HtmlContent = Replace(HtmlContent,"{$jobid}", rs("id"))
	HtmlContent = Replace(HtmlContent,"{$jobname}", rs("duix"))
	HtmlContent = HTML.ReadAnnounceContent(HtmlContent, ChannelID)
	HtmlContent = HTML.ReadAnnounceList(HtmlContent)
	Response.Write HtmlContent
End If
Rs.Close:Set Rs = Nothing
Set HTML = Nothing
CloseConn
%>