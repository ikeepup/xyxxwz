<!--#include file="config.asp"-->
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
Dim Rs, SQL
enchiasp.LoadTemplates ChannelID, 3, 0
HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", ChannelID)
'--频道目录
HtmlContent = Replace(HtmlContent,"{$ChannelRootDir}", ChannelRootDir, 1, -1, 1)
HtmlContent = Replace(HtmlContent,"{$CurrentStation}",enchiasp.ChannelName)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","提交简历")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)
if Request("job")="" then
	Call OutputScript("错误的参数，请不要随意输入一些参数！","index.asp")
end if
Set Rs = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM ECCMS_job where id="& Request("job") &" "
Set Rs = enchiasp.Execute(SQL)

If Rs.BOF And Rs.EOF Then
	response.write "<p align=center>发生了错误,</p>"
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