<!--#include file="setup.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>��̨���������ʾ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Style.css" type="text/css">
</head>
<body leftmargin=0 bottommargin=0 rightmargin=0 topmargin=0>
<p>&nbsp;</p>
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
action = Trim(Replace(Request("action"),"'", "''"))
If Request.Querystring("message") <> "" Then
        Message = Trim(Replace(Request.Querystring("message"),"'", ""))
End If
Select Case action
        Case "error"
                Call Error_Msg()
        Case "err"
                Call AdminError()
        Case "succeed"
                Call Succeed_Msg()
        Case "remind"
                Call Remind_Msg()
        Case "keyerr"
                Call KeyError()
        Case "genup"
                Call GenupMsg()
        Case Else
                Call AdminError()
End Select
Admin_Footer
CloseConn
Sub Error_Msg()
        response.write "<br><br><table width=""523""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"& vbCrLf
        response.write "  <tr>"& vbCrLf
        response.write "    <td><img src=""images/img_r2_c1.gif"" width=""523"" height=""55""></td>"& vbCrLf
        response.write "  </tr>"& vbCrLf
        response.write "  <tr>"& vbCrLf
        response.write "    <td height=""100"" background=""images/img_r2_c2.gif""><table width=""92%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"& vbCrLf
        response.write "      <tr>"& vbCrLf
        response.write "        <td width=""22%"" align=""center""><img src=""images/err.gif"" width=""95"" height=""97""></td>"& vbCrLf
        response.write "        <td width=""78%""><b>������������Ŀ���ԭ��</b><br>" & Message &"</td>"& vbCrLf
        response.write "      </tr>"& vbCrLf
        response.write "    </table></td>"& vbCrLf
        response.write "  </tr>"& vbCrLf
        response.write "  <tr>"& vbCrLf
        response.write "    <td align=""right"" background=""images/img_r2_c3.gif""><a href=""" & Request.ServerVariables("HTTP_REFERER") & """><img src=""images/confirm_r2.gif"" alt=""ȷ������"" width=""123"" height=""42"" border=""0""></a></td>"& vbCrLf
        response.write "  </tr>"& vbCrLf
        response.write "</table><p>&nbsp;</p>"& vbCrLf
End Sub
Sub AdminError()
        response.write "<br><br><table width=""523""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"& vbCrLf
        response.write "  <tr>"& vbCrLf
        response.write "    <td><img src=""images/img_r2_c1.gif"" width=""523"" height=""55""></td>"& vbCrLf
        response.write "  </tr>"& vbCrLf
        response.write "  <tr>"& vbCrLf
        response.write "    <td height=""100"" background=""images/img_r2_c2.gif""><table width=""92%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"& vbCrLf
        response.write "      <tr>"& vbCrLf
        response.write "        <td width=""22%"" align=""center""><img src=""images/err.gif"" width=""95"" height=""97""></td>"& vbCrLf
        response.write "        <td width=""78%""><b>��������Ŀ���ԭ��</b><br><li>ȷ�����ʧ�ܣ���û��ʹ�õ�ǰ���ܵ�Ȩ�ޡ�</li><li>��ǰ�����Ѽ�¼�������ʲô���⣬����ϵ����Ա��</li></td>"& vbCrLf
        response.write "      </tr>"& vbCrLf
        response.write "    </table></td>"& vbCrLf
        response.write "  </tr>"& vbCrLf
        response.write "  <tr>"& vbCrLf
        response.write "    <td align=""right"" background=""images/img_r2_c3.gif""><a href=""" & Request.ServerVariables("HTTP_REFERER") & """><img src=""images/confirm_r2.gif"" alt=""ȷ������"" width=""123"" height=""42"" border=""0""></a></td>"& vbCrLf
        response.write "  </tr>"& vbCrLf
        response.write "</table><p>&nbsp;</p>"& vbCrLf
End Sub
'********�ɹ���ʾ��Ϣ****************
Sub Succeed_Msg()
        Response.Write "<BR><BR><table align=""center"" border=""0"" cellpadding=""5"" cellspacing=""1"" class=""tableBorder1"">"& vbCrLf
        Response.Write "    <tr> "& vbCrLf
        Response.Write "      <th>�ɹ���ʾ��Ϣ!</th>"& vbCrLf
        Response.Write "    </tr>"& vbCrLf
        Response.Write "  <tr><td class=TableRow2 style=""padding-right: 8px; padding-left: 8px; padding-bottom: 5px; padding-top: 5px"">" & Message &"</td></tr>" & vbCrLf
        Response.Write "  <tr><td class=TableRow2 align=""right"" style='COLOR: Red;'>ʱ�䣺" & Now() & "</td></tr>" & vbCrLf
        Response.Write "  <tr><td align=center class=TableRow1><a href='" & Request.ServerVariables("HTTP_REFERER") & "'>������һҳ...</a></td></tr>" & vbCrLf
        response.Write " </table><p>&nbsp;</p>"& vbCrLf
End Sub
'********��ʾ��Ϣ****************
Sub Remind_Msg()
        Response.Write "<BR><BR><table cellpadding=5 cellspacing=1 border=0 align=center class=tableBorder1>" & vbCrLf
        Response.Write "  <tr><th>��ʾ!</th></tr>" & vbCrLf
        Response.Write "  <tr><td class=TableRow2 style=""padding-right: 8px; padding-left: 8px; padding-bottom: 5px; padding-top: 5px"">" & Message &"</td></tr>" & vbCrLf
        Response.Write "  <tr><td align=""right"" class=TableRow2 style='COLOR: Red;'>ʱ�䣺" & enchiasp.NowTime & "</td></tr>" & vbCrLf
        Response.Write "  <tr><td align=center class=TableRow1><a href='"&Request.ServerVariables("HTTP_REFERER")&"'>������һҳ...</a></td></tr>" & vbCrLf
        Response.Write "</table><p>&nbsp;</p>" & vbCrLf
End Sub
%>