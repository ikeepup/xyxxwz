<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/classmenu.asp"-->
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
Dim action,Message
Dim HtmlContent,strHtml

enchiasp.LoadTemplates 9999, 7, 2
HtmlContent = enchiasp.HtmlContent
HtmlContent = Replace(HtmlContent,"{$InstallDir}", enchiasp.InstallDir)
HtmlContent = Replace(HtmlContent, "{$ChannelID}", 0)
HtmlContent = Replace(HtmlContent,"{$PageTitle}","������~!")
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
        strHtml = strHtml & "  <tr><th>������Ϣ</th></tr>" & vbCrLf
        if message="" then
        	strHtml = strHtml & "  <tr><td width=100% class=tablerow>����δ֪��������ϵͳ����Ա��ϵ��</td></tr>" & vbCrLf
        else
       	 	strHtml = strHtml & "  <tr><td width=100% class=tablerow>" & Message & "</td></tr>" & vbCrLf
        end if
        strHtml = strHtml & "  <tr><td width=100% class=tablerow align=""right"" style='COLOR: Red;'>ʱ�䣺" & Now() & "</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td align=center class=tablerow>[<a href=""javascript:history.go(-1)"">������һҳ</a>]����[<a href=""javascript:window.close()"">�رձ�����</a>]</td></tr>" & vbCrLf
        strHtml = strHtml & "</table><BR>" & vbCrLf
End Sub

Sub ChannelStop()
        strHtml = "<BR><table cellpadding=5 cellspacing=1 border=0 width=65% align=center class=tableBorder1>" & vbCrLf
        strHtml = strHtml & "  <tr><th>Ƶ���Ѿ��ر�!</th></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow>���� URL ʱ��������������ϵͳ����Ա��ϵ��</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow align=""right"" style='COLOR: Red;'>ʱ�䣺" & Now() & "</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td align=center class=tablerow>[<a href=""javascript:window.close()"">�رձ�����</a>]</td></tr>" & vbCrLf
        strHtml = strHtml & "</table><BR>" & vbCrLf
End Sub
Sub StopRemind()

        strHtml = "<BR><table cellpadding=5 cellspacing=1 border=0 width=65% align=center class=tableBorder1>" & vbCrLf
        strHtml = strHtml & "  <tr><th>ϵͳά����!</th></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow>" & enchiasp.StopReadme &"</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow align=""right"" style='COLOR: Red;'>ʱ�䣺" & Now() & "</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td align=center class=tablerow>[<a href=""javascript:window.close()"">�رձ�����</a>]</td></tr>" & vbCrLf
        strHtml = strHtml & "</table><BR>" & vbCrLf
End Sub
Sub EstopRemind()
        strHtml = "<BR><table cellpadding=5 cellspacing=1 border=0 width=65% align=center class=tableBorder1>" & vbCrLf
        strHtml = strHtml & "  <tr><th>�Ƿ�����!</th></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow>���� URL ʱ��������������ϵͳ����Ա��ϵ��</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td width=100% class=tablerow align=""right"" style='COLOR: Red;'>ʱ�䣺" & Now() & "</td></tr>" & vbCrLf
        strHtml = strHtml & "  <tr><td align=center class=tablerow>[<a href=""javascript:window.close()"">�رձ�����</a>]</td></tr>" & vbCrLf
        strHtml = strHtml & "</table><BR>" & vbCrLf
End Sub
%>