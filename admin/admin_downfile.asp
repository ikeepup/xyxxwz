<!--#include file="setup.asp" -->
<!--#include file="check.asp"-->
<%
Server.ScriptTimeOut = 18000
Admin_header
Response.Write "<base target=""_self"">" & vbNewLine
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
Dim Action
If CInt(ChannelID) = 0 Then ChannelID = 2
Action = LCase(Request("action"))
If Action = "down" Then
	Call BeginDown
Else
	Call showmain
End If

If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
Public Sub showmain()
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>" & vbNewLine
	Response.Write " <tr>"
	Response.Write "   <th colspan=2>Զ���������</th>"
	Response.Write " </tr>"
	Response.Write "<form name=myform method=post action=admin_downfile.asp?action=down>" & vbNewLine
	Response.Write "<input type=hidden name=ChannelID value='" & ChannelID & "'>"
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<td class=tablerow1>Զ��URL��</td>" & vbNewLine
	Response.Write "	<td class=tablerow1><input type=text name=FileUrl size=60></td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<td class=tablerow2>����·����</td>" & vbNewLine
	Response.Write "	<td class=tablerow2><input type=text name=FilePath size=60 value='" & enchiasp.InstallDir & enchiasp.ChannelDir & "UploadFile/'><br><b>ע�⣺</b>������������·�����ļ���</td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<td class=tablerow1>�Ƿ���ʾ���ؽ��ȣ�</td>" & vbNewLine
	Response.Write "	<td class=tablerow1><input type=radio name=IsPross value='0'> ����ʾ&nbsp;&nbsp;<input type=radio name=IsPross value='1'> �ٶ�&nbsp;&nbsp;"
	Response.Write "	<input type=radio name=IsPross value='2' checked> ����&nbsp;&nbsp;<input type=radio name=IsPross value='3'> �ٶ�+����&nbsp;&nbsp; </td>" & vbNewLine
	Response.Write "</tr>" & vbNewLine
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<td class=tablerow1></td>" & vbNewLine
	Response.Write "	<td class=tablerow1><input type=button value=' �رմ��� ' onClick=""window.close();"" class=Button>&nbsp;&nbsp;<input type=submit value=' ��ʼ���� ' class=Button></td>" & vbNewLine
	Response.Write "</tr></form>" & vbNewLine
	Response.Write "</table>" & vbNewLine
End Sub

Public Sub BeginDown()
	Dim strFilePath,IsPross
	Dim FileUrl,FilePath,strFileName
	If Trim(Request.Form("FileUrl")) = "" Then
		Call AlertInform("Զ��URL����Ϊ�գ�","admin_downfile.asp")
		Exit Sub
	Else
		FileUrl = Request.Form("FileUrl")
	End If
	If Trim(Request.Form("FilePath")) = "" Then
		Call AlertInform("����·������Ϊ�գ�","admin_downfile.asp")
		Exit Sub
	Else
		FilePath = Trim(Request.Form("FilePath"))
	End If
	If Right(FilePath,1) = "/" Or Right(FilePath,1) = "\" Then
		Call AlertInform("������������·�����ļ����ƣ�","admin_downfile.asp")
		Exit Sub
	End If
	
	IsPross = enchiasp.ChkNumeric(Request.Form("IsPross"))
	Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>" & vbNewLine
	Response.Write " <tr>"
	Response.Write "   <th><span id=txt1>���ڲɼ������Ժ�....</span></th>"
	Response.Write " </tr>"
	Response.Write "<tr>" & vbNewLine
	Response.Write "	<td class=tablerow1 style=""line-height: 20px;"">" & vbNewLine
	
	If IsPross <> 0 Then
		If IsPross = 2 Or IsPross = 3 Then
			RevealProgress
		End If
		
		If IsPross = 1 Or IsPross = 3 Then
			Response.Write "<div><span id=Proess1 style=""color: #008800;""></span></div>"
		End If
	End If

	Response.Write "<br><strong><font color=blue>�����С��</font></strong><span id=txt2 style=""font-size:9pt;color:red"">0</span>"
	Response.Write "</td></tr>" & vbNewLine
	Response.Write "<tr><td class=tablerow1><div align=center><a href='admin_downfile.asp' class=showmenu>ֹͣ��������</a>"
	'Response.Write "<a href='#'  onClick=""window.close();"" class=showmenu>ֹͣ��������</a>"
	Response.Write "</div></td></tr></table>" & vbNewLine
	Response.Flush
	Dim Myenchicms

	On Error Resume Next
	Set Myenchicms = CreateObject("Gatherer.SoftCollection")
	If Err Then
		Response.Write "<br /><br /><br /><div align=""center"" style=""font-size:18px;color:red"">�ǳ��ź������ķ�������֧�ֲɼ������</div>"
		Err.Clear
		Set Myenchiasp = Nothing
		Response.End
	End If
	Myenchicms.Connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ChkMapPath(DBPath)
	Myenchicms.ChannelPath = enchiasp.InstallDir & enchiasp.ChannelDir
	Myenchicms.DBConnstr = Connstr
	Myenchicms.SqlDataBase = IsSqlDataBase
	strFilePath = Myenchicms.FileDownload(FileUrl,FilePath,IsPross,200,vbNullString)
	If len(strFilePath) > 0 Then
		strFileName = Mid(strFilePath, InStrRev(strFilePath, "\") + 1)
		Response.Write "<br><br><div align=center><a href='#' onClick=""window.returnValue='" & strFileName & "|" & Myenchicms.SoftSize & "';window.close();"" class=showmenu>�ļ�������ɣ�����رմ��ڽ��Զ������ļ���</a></div>"
		Response.Write "<div align=center>�����ļ�·����<input type=text name=sPath size=75 value='" & strFilePath & "'></div>"
		Response.Write "<div align=center><input type=button value=' �رմ��� ' onClick=""window.returnValue='" & strFileName & "|" & Myenchicms.SoftSize & "';window.close();"" class=Button></div><br><br>"
	Else
		Response.Write "<br><br><div align=center>�ļ�����ʧ��!</div>"
		Response.Write "<div align=center><input type=button value=' �رմ��� ' onClick=""window.close();"" class=Button></div><br><br>"
	End If
	
	
	Set Myenchicms = Nothing
End Sub

Sub RevealProgress()
	Response.Write "<div><table width='400' align=left border=0 cellspacing=1 cellpadding=1>" & vbCrLf
	Response.Write "<tr> " & vbCrLf
	Response.Write "<td style=""border: 1px #384780 solid ;background-color: #FFFFFF;"">" & vbCrLf
	Response.Write "<table width=0 id=tablePros name=tablePros border=0 cellspacing=0 cellpadding=0>" & vbCrLf
	Response.Write "<tr height=12><td bgcolor='#0650D2'>" & vbCrLf
	Response.Write "</td></tr></table></td></tr>" & vbCrLf
	Response.Write "</table><br></div>" & vbCrLf
	Response.Flush
End Sub

%>
