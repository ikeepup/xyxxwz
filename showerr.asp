<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/classmenu.asp"-->
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
Dim action,Message
Dim HtmlContent,strHtml

enchiasp.LoadTemplates 9999, 7, 2
HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","出错啦~!")
HtmlContent = ReadClassMenu(HtmlContent)
HtmlContent = ReadClassMenubar(HtmlContent)


action = Trim(Replace(Request.Querystring("action"),"'", "''"))
'action = enchiasp.CheckStr("action")

Message = Trim(Replace(Request.Querystring("Message"),"'", "''"))
Select Case action
        Case "error"
                Call ErrorMsg()
        Case "stop"
            
                Call StopRemind()
              
        Case "estop"
        
                Call EstopRemind()
            
        Case "chanstop"
                Call ChannelStop()
        Case Else
                Call ErrorMsg()
           
End Select

HtmlContent = Replace(HtmlContent,"{$TempContent}", strHtml)

Response.Write HtmlContent
CloseConn
Sub ErrorMsg()
        strHtml = "<BR><table cellpadding=5 cellspacing=1 border=0 width=65% align=center class=tableBorder1>" & vbCrLf
        strHtml = strHtml & "  <tr><th>错误信息</th></tr>" & vbCrLf
        if message="" then
        	strHtml = strHtml & "  <tr><td width=100% class=tablerow>发生未知错误，请与系统管理员联系！</td></tr>" & vbCrLf
        else
       	 	strHtml = strHtml & "  <tr><td width=100% class=tablerow>" & Message & "</td></tr>" & vbCrLf
        end if
        strHtml = strHtml & "  <tr><td width=100% class=tablerow align=""right"" style='COLOR: Red;'>时间：" & Now() & "</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td align=center class=tablerow>[<a href=""javascript:history.go(-1)"">返回上一页</a>]　　[<a href=""javascript:window.close()"">关闭本窗口</a>]</td></tr>" & vbCrLf
        strHtml = strHtml & "</table><BR>" & vbCrLf
End Sub

Sub ChannelStop()
        strHtml = "<BR><table cellpadding=5 cellspacing=1 border=0 width=65% align=center class=tableBorder1>" & vbCrLf
        strHtml = strHtml & "  <tr><th>频道已经关闭!</th></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow>处理 URL 时服务器出错，请与系统管理员联系！</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow align=""right"" style='COLOR: Red;'>时间：" & Now() & "</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td align=center class=tablerow>[<a href=""javascript:window.close()"">关闭本窗口</a>]</td></tr>" & vbCrLf
        strHtml = strHtml & "</table><BR>" & vbCrLf
End Sub
Sub StopRemind()

        strHtml = "<BR><table cellpadding=5 cellspacing=1 border=0 width=65% align=center class=tableBorder1>" & vbCrLf
        strHtml = strHtml & "  <tr><th>系统维护中!</th></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow>" & enchiasp.StopReadme &"</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow align=""right"" style='COLOR: Red;'>时间：" & Now() & "</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td align=center class=tablerow>[<a href=""javascript:window.close()"">关闭本窗口</a>]</td></tr>" & vbCrLf
        strHtml = strHtml & "</table><BR>" & vbCrLf
End Sub
Sub EstopRemind()
        strHtml = "<BR><table cellpadding=5 cellspacing=1 border=0 width=65% align=center class=tableBorder1>" & vbCrLf
        strHtml = strHtml & "  <tr><th>非法操作!</th></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow>处理 URL 时服务器出错，请与系统管理员联系！</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow align=""right"" style='COLOR: Red;'>时间：" & Now() & "</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td align=center class=tablerow>[<a href=""javascript:window.close()"">关闭本窗口</a>]</td></tr>" & vbCrLf
        strHtml = strHtml & "</table><BR>" & vbCrLf
End Sub
%>