<!--#include file="setup.asp" -->
<!--#include file="check.asp"-->

<%
'=====================================================================
' ������ƣ�������վ����ϵͳ--ע�����
' ��ǰ�汾��enchicms Version 3.0.0
' �������ڣ�2005-03-25
' �ٷ���վ���˳��ж�������Ƽ��������޹�˾(www.enchi.com.cn) 
' ����֧�֣����Ʒ�
' ���䣺liuyunfan@163.com
' QQ��21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
Dim selAdminID
Dim i,Action,strClass
Admin_header
If Not ChkAdmin("999") Then
	Server.Transfer("showerr.asp")
	Response.End
End If
Response.Write "<table cellpadding=2 cellspacing=1 border=0 class=tableBorder align=center>" & vbCrLf
Response.Write " <tr>" & vbCrLf
Response.Write " <th height=22 colspan=6>���ע��</th>" & vbCrLf
Response.Write " </tr>" & vbCrLf
Response.Write " <tr>" & vbCrLf
Response.Write " <td class=TableRow1> <b>��ע��</b> Ϊ����֪ʶ��Ȩ��Ϊ���û��ϵͳ���񣬽�������������Ĳ�Ʒע�ᣬ�����޷���ȡ��ص��������񣬶��ɴ˶��������κι��Ͻ��ò��������ϵ�֧�֡����¸�������Ϊ�������������д�����в�����ط�����ѯ��˾������Ա��лл������"
Response.Write " </td>" & vbCrLf
Response.Write " </tr>" & vbCrLf
Response.Write " </table><br>" & vbCrLf


Action = LCase(Request("action"))
Select Case Trim(Action)
Case "reg"
	Call savereg
Case Else
	Call reginfo
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn

Private Sub reginfo()
	dim urlflag
	Response.Write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" class=""tableBorder"" align=center>" & vbCrLf
	Response.Write "<form action='?action=reg' method=post>" & vbCrLf

	Response.Write " <tr>" & vbCrLf
	Response.Write " <th height=22 colspan=2>ע��ѡ��</th>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
	Set Rs = enchiasp.Execute("select * from ECCMS_config")
	if rs.eof then
		response.write "���ݿ����ó������飡"
	else
		
		urlflag=rs("urlflag")
		Response.Write " <tr>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write "<B>ע����ַ</b>" & vbCrLf
		Response.Write "</td>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write " <input type=text name='url'"
		response.write "value='"
		response.write rs("url")
		response.write "'>" & vbCrLf	
        Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf
		
		
		Response.Write " <tr>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write "<B>ע������</b>" & vbCrLf
		Response.Write "</td>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write " <input type=text name='urldate'"
		response.write "value='"
		response.write rs("urldate")
		response.write "'>" & vbCrLf	
        Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf

		
		
		
		
		Response.Write " <tr>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write "<B>ע����</b>" & vbCrLf
		Response.Write "</td>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write " <input type=text name='urlman'"
		response.write "value='"
		response.write rs("urlman")
		response.write "'>" & vbCrLf	
        Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf
		
		
		
		Response.Write " <tr>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write "<B>ע��ģ��</b>" & vbCrLf
		Response.Write "</td>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		
		response.write "<input type='checkbox' name='urlflag' value='SiteConfig'"
		If InStr(urlflag, "SiteConfig") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">��������"
		
		

		response.write "<input type='checkbox' name='urlflag' value='yemian'"
		If InStr(urlflag, "yemian") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">��ҳ��ͼ��"

		response.write "<input type='checkbox' name='urlflag' value='Article'"
		If InStr(urlflag, "Article") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">����Ƶ��"
		
		
		response.write "<input type='checkbox' name='urlflag' value='soft'"
		If InStr(urlflag, "soft") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">����Ƶ��"
		
		response.write "<br>"

		
		response.write "<input type='checkbox' name='urlflag' value='flash'"
		If InStr(urlflag, "flash") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">����Ƶ��"
			
		response.write "<input type='checkbox' name='urlflag' value='shop'"
		If InStr(urlflag, "shop") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">��ƷƵ��"	
		
		response.write "<input type='checkbox' name='urlflag' value='order'"
		If InStr(urlflag, "order") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">����Ƶ��"	


		response.write "<input type='checkbox' name='urlflag' value='job'"
		If InStr(urlflag, "job") <> 0 Then 
			Response.Write " checked"
		end if
		response.write ">��ƸƵ��"

		
		
		
        Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf
		
		
		Response.Write " <tr>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write "<B>���к�</b>" & vbCrLf
		Response.Write "</td>" & vbCrLf
		Response.Write "<td Class=TableRow1>" & vbCrLf
		Response.Write " <input type=text name='urlreg' size='100'"
		response.write "value='"
		response.write rs("urlreg")
		response.write "'>" & vbCrLf	
        Response.Write "</td>" & vbCrLf
		Response.Write " </tr>" & vbCrLf


	end if
	
		
	
	Rs.Close
	Set Rs = Nothing
	
	Response.Write " <tr>" & vbCrLf
	Response.Write " <td colspan=""6"" align=center Class=TableRow1>" & vbCrLf
	Response.Write " <input type='submit' class=""button"" name=""Submit"" value=""  ע ��  "" >" & vbCrLf
	Response.Write " </td>" & vbCrLf
	Response.Write " </tr>" & vbCrLf
		Response.Write " </form>" & vbCrLf
	Response.Write " </table>" & vbCrLf
End Sub


Private Sub savereg()
	Dim adminuserid
	dim zcj
	If Not ChkAdmin("9999") Then
		ErrMsg = "<li>��û�д˲���Ȩ��!</li><li>����ʲô��������ϵվ����</li>"
		Founderr = True
		Exit Sub
	End If

	If Request.Form("url") = "" or Request.Form("urldate") = "" or Request.Form("urlman") = "" or Request.Form("urlreg") = ""Then
		ErrMsg = "��������ص������ټ�����"
		Founderr = True
		Exit Sub
	Else
		'
		if not isdate(Request.Form("urldate")) then
			ErrMsg = "������������ݣ���������ȷ���������ݸ�ʽ��"
			Founderr = True
			Exit Sub
		end if
		'
		if Request.Form("url") <>enchiasp.SiteUrl&enchiasp.InstallDir then
			ErrMsg = "�������վ��ַ����������ȷ����վ��ַ����HTTP����Ӧ��ע�����վΪ��<br>"&enchiasp.SiteUrl&enchiasp.InstallDir
			Founderr = True
			Exit Sub
		end if 
		
		'
		zcj=md5(request.form("url")&"liuyunfan")&"-" & md5("yunliufan")&md5(request.form("urldate"))&"-"&md5("liu")&md5(request.form("urlman")) & md5("fanyun")&"-"& md5( Replace(Replace(Request("urlflag"), "'", ""), " ", ""))
		if Request.Form("urlreg") =zcj then
			
			Set Rs = Server.CreateObject("adodb.recordset")
			SQL = "SELECT * FROM ECCMS_config"
			Rs.Open SQL, conn, 1, 3
			If Not (Rs.EOF And Rs.BOF) Then
				rs("url")=request.form("url")
				rs("urldate")=request.form("urldate")
				rs("urlreg")=request.form("urlreg")
				rs("urlman")=request.form("urlman")	
				Rs("urlflag") = Replace(Replace(Request("urlflag"), "'", ""), " ", "")
				Rs.update
			End If
			Rs.Close
			Set Rs = Nothing
			Succeed ("ע��ɹ���")
		else
			ErrMsg = "��������кţ���������ȷ�����кţ�"
			Founderr = True
			Exit Sub
		end if
	End If
	
	End Sub

%>