<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../api/cls_api.asp"-->
<%
Admin_header
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
If Not ChkAdmin("9999") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveConformify
	Case Else
		Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Function GetFormID()
	Dim i,sessionid
	sessionid = Session.SessionID
	For i=1 to Len(sessionid)
		GetFormID=GetFormID&Chr(Mid(sessionid,i,1)+97)
	Next
End Function
Sub showmain()
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder">
<form name="myform" method="post" action="?action=save">
<tr>
	<th colspan="2">��ϵͳ���Ͻӿ�����</th>
</tr>
<tr>
	<td class="TableRow1" width="20%" align="right"><u>�Ƿ�����ϵͳ���ϳ���</u>��</td>
	<td class="TableRow1" width="80%">
	<input type="radio" name="API_Enable" value="false"<%
	If Not API_Enable Then Response.Write " checked"
	%>> �ر�&nbsp;&nbsp;        
	<input type="radio" name="API_Enable" value="true"<%
	If API_Enable Then Response.Write " checked"
	%>> ����        
	</td>
</tr>
<tr>
	<td class="TableRow2" align="right"><u>����ϵͳ��Կ</u>��</td>
	<td class="TableRow2"><input type="text" name="API_ConformKey" size="35" value="<%=API_ConformKey%>"> 
		<font color="red">ϵͳ���ϣ����뱣֤������ϵͳ���õ���Կһ�¡�</font>
	</td>
</tr>
<tr>
	<td class="TableRow1" align="right"><u>�Ƿ����</u>��</td>
	<td class="TableRow1">
	<input type="radio" name="API_Debug" value="false"<%
	If Not API_Debug Then Response.Write " checked"
	%>> ��&nbsp;&nbsp;        
	<input type="radio" name="API_Debug" value="true"<%
	If API_Debug Then Response.Write " checked"
	%>> ��&nbsp;&nbsp;<font color="red">������ϵ���̳�����ENCHICMS���û����ݲ�ͬ���������ѡ���ǡ�</font>        
	</td>
</tr>
<tr>
	<td class="TableRow2" align="right"><u>���ϳ���Ľӿ��ļ�·��</u>��</td>
	<td class="TableRow2"><textarea name="API_Urls" rows="6" cols="70"><%=API_Urls%></textarea></td>
</tr>
<tr>
	<td class="TableRow1" align="right"><u>�����û���¼��ת��URL</u>��</td>
	<td class="TableRow1"><input type="text" name="API_LoginUrl" size="45" value="<%=API_LoginUrl%>"> 
		<font color="red">�����������롰0����</font>
	</td>
</tr>
<tr>
	<td class="TableRow2" align="right"><u>�����û�ע���ת��URL</u>��</td>
	<td class="TableRow2"><input type="text" name="API_ReguserUrl" size="45" value="<%=API_ReguserUrl%>"> 
		<font color="red">�����������롰0����</font>
	</td>
</tr>
<tr>
	<td class="TableRow1" align="right"><u>�����û�ע����ת��URL</u>��</td>
	<td class="TableRow1"><input type="text" name="API_LogoutUrl" size="45" value="<%=API_LogoutUrl%>"> 
		<font color="red">�����������롰0����</font>
	</td>
</tr>
<tr>
	<td class="TableRow2" align="right"></td>
	<td class="TableRow2"><input type="submit" value="��������" name="B1" class="Button"></td>
</tr>
</form>
<tr>
	<td class="TableRow1" colspan="2"><b>˵����</b><br /><font color="blue">����ж���������ϣ��ӿ�֮���ð��"|"�ָ�<br />���磺http://�����̳��ַ/dv_dpo.asp|http://�����վ��ַ/���Ͱ�װĿ¼/oblogresponse.asp;<br />
	��ϵͳ�Ľӿ�·����<font color="red"><%=enchiasp.SiteUrl%><%=enchiasp.InstallDir%>api/api_reponse.asp</font><br /></font></td>
</tr>
</table>

<%
End Sub

Sub SaveConformify()
	Dim XslDoc,XslNode,Xsl_Files
	Xsl_Files = API_Path & "api.config"
	Xsl_Files = Server.MapPath(Xsl_Files)
	Set XslDoc = Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	If Not XslDoc.Load(Xsl_Files) Then
		Response.Write "��ʼ���ݲ����ڣ�"
		Response.End
	Else
		Set XslNode = XslDoc.documentElement.selectSingleNode("rs:data/z:row")
		XslNode.attributes.getNamedItem("api_enable").text = Trim(Request.Form("API_Enable"))
		XslNode.attributes.getNamedItem("api_conformkey").text = ChkRequestForm("API_ConformKey")
		XslNode.attributes.getNamedItem("api_urls").text = ChkRequestForm("API_Urls")
		XslNode.attributes.getNamedItem("api_debug").text = ChkRequestForm("API_Debug")
		XslNode.attributes.getNamedItem("api_loginurl").text = ChkRequestForm("API_LoginUrl")
		XslNode.attributes.getNamedItem("api_reguserurl").text = ChkRequestForm("API_ReguserUrl")
		XslNode.attributes.getNamedItem("api_logouturl").text = ChkRequestForm("API_LogoutUrl")
		'XslNode.attributes.setNamedItem(XslDoc.createNode(2,"date","")).text = Now()
		'XslNode.appendChild(XslDoc.createNode(1,"pubDate","")).text = Now()
		XslDoc.save Xsl_Files
		Set XslNode = Nothing
	End If
	Set XslDoc = Nothing
	Succeed("<li>��ϲ�����������óɹ���</li>")
End Sub
Function ChkRequestForm(reform)
	Dim strForm
	strForm = Trim(Request.Form(reform))
	If IsNull(strForm) Then
		strForm = "0"
	Else
		strForm = Replace(strForm, Chr(0), vbNullString)
		strForm = Replace(strForm, Chr(34), vbNullString)
		strForm = Replace(strForm, "'", vbNullString)
		strForm = Replace(strForm, """", vbNullString)
	End If
	If strForm = "" Then strForm = "0"
	ChkRequestForm = strForm
End Function

%>