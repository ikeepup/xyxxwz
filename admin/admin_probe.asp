<%@ Language="VBScript" %>
<%' Option Explicit %>
<script language=javascript>
 <!--
 var startTime,endTime;
 var d=new Date();
 startTime=d.getTime();
 //-->
 </script>
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
if session("AdminName")="" then
	Response.Redirect "admin_klogin.asp"
	Response.End
end if
'��ʹ�������������ֱ�ӽ����н����ʾ�ڿͻ���
'Response.Buffer = False
Dim starttime
starttime=timer()*1000
'�������������
Dim ObjTotest(26,4)

ObjTotest(0,0) = "MSWC.AdRotator"
ObjTotest(1,0) = "MSWC.BrowserType"
ObjTotest(2,0) = "MSWC.NextLink"
ObjTotest(3,0) = "MSWC.Tools"
ObjTotest(4,0) = "MSWC.Status"
ObjTotest(5,0) = "MSWC.Counters"
ObjTotest(6,0) = "IISSample.ContentRotator"
ObjTotest(7,0) = "IISSample.PageCounter"
ObjTotest(8,0) = "MSWC.PermissionChecker"
ObjTotest(9,0) = "Scripting.FileSystemObject"
	ObjTotest(9,1) = "(FSO �ı��ļ���д)"
ObjTotest(10,0) = "adodb.connection"
	ObjTotest(10,1) = "(ADO ���ݶ���)"
'ObjTotest(11,0) = "enchicmsCMS.SiteMainObject"
	'ObjTotest(11,1) = "(������վ�������)"	
ObjTotest(12,0) = "SoftArtisans.FileUp"
	ObjTotest(12,1) = "(SA-FileUp �ļ��ϴ�)"
ObjTotest(13,0) = "SoftArtisans.FileManager"
	ObjTotest(13,1) = "(SoftArtisans �ļ�����)"
ObjTotest(14,0) = "LyfUpload.UploadFile"
	ObjTotest(14,1) = "(���Ʒ���ļ��ϴ����)"
ObjTotest(15,0) = "Persits.Upload.1"
	ObjTotest(15,1) = "(ASPUpload �ļ��ϴ�)"
ObjTotest(16,0) = "w3.upload"
	ObjTotest(16,1) = "(Dimac �ļ��ϴ�)"

ObjTotest(17,0) = "JMail.SmtpMail"
	ObjTotest(17,1) = "(Dimac JMail �ʼ��շ�) "
ObjTotest(18,0) = "CDONTS.NewMail"
	ObjTotest(18,1) = "(���� SMTP ����)"
ObjTotest(19,0) = "Persits.MailSender"
	ObjTotest(19,1) = "(ASPemail ����)"
ObjTotest(20,0) = "SMTPsvg.Mailer"
	ObjTotest(20,1) = "(ASPmail ����)"
ObjTotest(21,0) = "DkQmail.Qmail"
	ObjTotest(21,1) = "(dkQmail ����)"
ObjTotest(22,0) = "Geocel.Mailer"
	ObjTotest(22,1) = "(Geocel ����)"
ObjTotest(23,0) = "IISmail.Iismail.1"
	ObjTotest(23,1) = "(IISmail ����)"
ObjTotest(24,0) = "SmtpMail.SmtpMail.1"
	ObjTotest(24,1) = "(SmtpMail ����)"
	
ObjTotest(25,0) = "SoftArtisans.ImageGen"
	ObjTotest(25,1) = "(SA ��ͼ���д���)"
ObjTotest(26,0) = "W3Image.Image"
	ObjTotest(26,1) = "(Dimac ��ͼ���д���)"

public IsObj,VerObj,TestObj

'���Ԥ�����֧��������汾

dim i
for i=0 to 26
	on error resume next
	IsObj=false
	VerObj=""
	'dim TestObj
	TestObj=""
	set TestObj=server.CreateObject(ObjTotest(i,0))
	If -2147221005 <> Err then		
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if
	ObjTotest(i,2)=IsObj
	ObjTotest(i,3)=VerObj
next

'�������Ƿ�֧�ּ�����汾���ӳ���
sub ObjTest(strObj)
	on error resume next
	IsObj=false
	VerObj=""
	TestObj=""
	set TestObj=server.CreateObject (strObj)
	If -2147221005 <> Err then		'
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if	
End sub
%>
<html>
<head>
<title>�������������Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="90%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
    <tr> 
      <th width="27%" height=25  Class=TableTitle>�� �� �� �� Ϣ</th></tr>
    <tr>
      <td height=22 class="TableBody1">
<font class=fonts>��ķ������Ƿ�֧��ASP</font>
<br>���������������ʾ���Ŀռ䲻֧��ASP��
<br>1�����ʱ��ļ�ʱ��ʾ���ء�
<br>2�����ʱ��ļ�ʱ�������ơ�&lt;%@ Language="VBScript" %&gt;�������֡�
</td></tr></table><br>
	<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
          <tr><th colspan=2>�� ���������йز���</th></tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;��������</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;������IP</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;�������˿�</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;������ʱ��</td><td class="TableRow1">&nbsp;<%=now%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;IIS�汾</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;�ű���ʱʱ��</td><td class="TableRow1">&nbsp;<%=Server.ScriptTimeout%> ��</td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;���ļ�·��</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("PATH_TRANSLATED")%></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;������CPU����</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;��������������</td><td class="TableRow1">&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	  </tr>
	  <tr>
		<td height=20 class=TableRow2>&nbsp;����������ϵͳ</td><td class="TableRow1">&nbsp;<%=Request.ServerVariables("OS")%></td>
	  </tr>
	</table>
<br>
<%
Dim strClass
	strClass = Trim(Request.Form("classname"))
	If "" <> strClass then
	Response.Write "<br>��ָ��������ļ������"
	Dim Verobj1
	ObjTest(strClass)
	  If Not IsObj then 
		Response.Write "<br><font color=red>���ź����÷�������֧�� " & strclass & " �����</font>"
	  Else
		if VerObj="" or isnull(VerObj) then 
			Verobj1="�޷�ȡ�ø�����汾"
		Else
			Verobj1="������汾�ǣ�" & VerObj
		End If
		Response.Write "<br><font class=fonts>��ϲ���÷�����֧�� " & strclass & " �����" & verobj1 & "</font>"
	  End If
	  Response.Write "<br>"
	end if
	%>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=2>�� IIS�Դ���ASP���</th></tr>
	<tr height=23 align=center><td width=70% class=TableTitle>�� �� �� ��</td><td width=30% class=TableTitle>֧�ּ��汾</td></tr>
	<%For i=0 to 10%>
	<tr>
		<td height=20 class=TableRow2>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td class=TableRow1>&nbsp;<%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
	</tr>
	<%next%>
</table>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=2>�� �������ļ��ϴ��͹������</th></tr>
	<tr height=23 align=center><td width=70% class=TableTitle>�� �� �� ��</td><td width=30% class=TableTitle>֧�ּ��汾</td></tr>
	<%For i=12 to 16%>
	<tr height="18" class=backq>
		<td height=20 class=TableRow2>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td class=TableRow1>&nbsp;<%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
	</tr>
	<%next%>
</table>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=2>�� �������շ��ʼ����</th></tr>
	<tr height=23 align=center><td width=70% class=TableTitle>�� �� �� ��</td><td width=30% class=TableTitle>֧�ּ��汾</td></tr>
	<%For i=17 to 24%>
	<tr height="18" class=backq>
		<td height=20 class=TableRow2>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td class=TableRow1>&nbsp;<%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
	</tr>
	<%next%>
</table>
<BR>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=2>�� ͼ�������</th></tr>
	<tr height=23 align=center><td width=70% class=TableTitle>�� �� �� ��</td><td width=30% class=TableTitle>֧�ּ��汾</td></tr>
	<%For i=25 to 26%>
	<tr height="20">
		<td height=20 class=TableRow2>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td class=TableRow1>&nbsp;<%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
	</tr>
	<%next%>
</table>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th>�� �������֧��������</th></tr>
<tr height=23 align=center><td width=70% class=TableTitle>
��������������������Ҫ���������ProgId��ClassId��</td></tr>
<FORM action=<%=Request.ServerVariables("SCRIPT_NAME")%> method=post id=form1 name=form1>
	<tr height="23">
		<td align=center class=TableRow1 height=30><input class=kuang type=text value="" name="classname" size=40>
<INPUT type=submit value=" ȷ �� " class=kuang id=submit1 name=submit1>
<INPUT type=reset value=" �� �� " class=kuang id=reset1 name=reset1> 
</td>
	  </tr>
</FORM>
</table>

<%if ObjTest("Scripting.FileSystemObject") then

	set fsoobj=server.CreateObject("Scripting.FileSystemObject")

%>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=6>�� ������������Ϣ</th></tr>
  <tr height="20" align=center>
	<td width="100" class=TableTitle>�̷��ʹ�������</td>
	<td width="50" class=TableTitle>����</td>
	<td width="80" class=TableTitle>���</td>
	<td width="60" class=TableTitle>�ļ�ϵͳ</td>
	<td width="80" class=TableTitle>���ÿռ�</td>
	<td width="80" class=TableTitle>�ܿռ�</td>
  </tr>
<%

	' ���Դ�����Ϣ���뷨���ԡ�COCOON ASP ̽�롱
	
	set drvObj=fsoobj.Drives
	for each d in drvObj
%>
  <tr height="18" align=center>
	<td class=TableRow2 align="right"><%=cdrivetype(d.DriveType) & " " & d.DriveLetter%>:</td>
<%
	if d.DriveLetter = "A" then	'Ϊ��ֹӰ������������������
		Response.Write "<td class=TableRow1></td><td class=TableRow1></td><td class=TableRow1></td><td class=TableRow1></td><td class=TableRow1></td>"
	else
%>
	<td class=TableRow1><%=cIsReady(d.isReady)%></td>
	<td class=TableRow1><%=d.VolumeName%></td>
	<td class=TableRow1><%=d.FileSystem%></td>
	<td align="right" class=TableRow1><%=cSize(d.FreeSpace)%></td>
	<td align="right" class=TableRow1><%=cSize(d.TotalSize)%></td>
<%
	end if
%>
  </tr>
<%
	next
%>
</td></tr>
</table>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=5>�� ��ǰ�ļ�����Ϣ <%
	dPath = server.MapPath("./")
	set dDir = fsoObj.GetFolder(dPath)
	set dDrive = fsoObj.GetDrive(dDir.Drive)
%>
�ļ���: <%=dPath%></th></tr>
  <tr height="23" align="center">
	<td width="75" class=TableTitle>���ÿռ�</td>
	<td width="75" class=TableTitle>���ÿռ�</td>
	<td width="75" class=TableTitle>�ļ�����</td>
	<td width="75" class=TableTitle>�ļ���</td>
	<td width="150" class=TableTitle>����ʱ��</td>
  </tr>
  <tr height="20" align="center">
	<td class=TableRow1><%=cSize(dDir.Size)%></td>
	<td class=TableRow1><%=cSize(dDrive.AvailableSpace)%></td>
	<td class=TableRow1><%=dDir.SubFolders.Count%></td>
	<td class=TableRow1><%=dDir.Files.Count%></td>
	<td class=TableRow1><%=dDir.DateCreated%></td>
  </tr>
</td></tr>
</table>

<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=2>�� �����ļ������ٶȲ���</th></tr>
<tr height="20" align=center>
	<td colspan=2 class=TableRow1><%
	
	Response.Write "�����ظ�������д���ɾ���ı��ļ�50��..."

	dim thetime3,tempfile,iserr

iserr=false
	t1=timer
	tempfile=server.MapPath("./") & "\aspchecktest.txt"
	for i=1 to 50
		Err.Clear

		set tempfileOBJ = FsoObj.CreateTextFile(tempfile,true)
		if Err <> 0 then
			Response.Write "�����ļ�����"
			iserr=true
			Err.Clear
			exit for
		end if
		tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"
		if Err <> 0 then
			Response.Write "д���ļ�����"
			iserr=true
			Err.Clear
			exit for
		end if
		tempfileOBJ.close
		Set tempfileOBJ = FsoObj.GetFile(tempfile)
		tempfileOBJ.Delete 
		if Err <> 0 then
			Response.Write "ɾ���ļ�����"
			iserr=true
			Err.Clear
			exit for
		end if
		set tempfileOBJ=nothing
	next
	t2=timer
if iserr <> true then
	thetime3=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...����ɣ�<font color=red>" & thetime3 & "����</font>��"

%>
</td></tr>
  <tr height=18>
	<td height=20 class=TableRow2>&nbsp;<font color=red>���ķ�����: <%=Request.ServerVariables("SERVER_NAME")%></font>&nbsp;</td><td class=TableRow1>&nbsp;<font color=red><%=thetime3%></font></td>
  </tr>
</table>
<%
end if
	
	set fsoobj=nothing

end if%>
<br>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder">
<tr><th colspan=3>�� ASP�ű����ͺ������ٶȲ���</th></tr>
<tr height="20" align=center>
	<td colspan=3 class=TableRow1>
<%
	'��Ϊֻ����50��μ��㣬����ȥ�����Ƿ����ѡ���ֱ�Ӽ��
	
	Response.Write "����������ԣ����ڽ���50��μӷ�����..."
	dim t1,t2,lsabc,thetime,thetime2
	t1=timer
	for i=1 to 500000
		lsabc= 1 + 1
	next
	t2=timer
	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...����ɣ�<font color=red>" & thetime & "����</font>��<br>"


	Response.Write "����������ԣ����ڽ���20��ο�������..."
	t1=timer
	for i=1 to 200000
		lsabc= 2^0.5
	next
	t2=timer
	thetime2=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...����ɣ�<font color=red>" & thetime2 & "����</font>��<br>"
%>
</td></tr>
<tr height=18>
	<td height=20 class=TableRow2>&nbsp;<font color=red>���ķ�����: <%=Request.ServerVariables("SERVER_NAME")%></font>&nbsp;</td><td class=TableRow1 nowrap>&nbsp;<font color=red><%=thetime%></font></td><td  class=TableRow1 nowrap>&nbsp;<font color=red><%=thetime2%></font></td>
  </tr>
</table>
<BR>
<table border=0 width=90% align=center cellspacing=0 cellpadding=0>
<tr><td align=center style="LINE-HEIGHT: 150%">
ҳ��װ��ʱ�䣺<script language=javascript>d=new Date();endTime=d.getTime();document.write((endTime-startTime)/1000);</script> ��<BR> Copyright (c) 2002-2005 <a href="http://www.enchi.com.cn" target="_blank"><font color=#6C70AA><b>enchi<font color=#CC0000>.com</font></b></font></a>. All Rights Reserved .


</td></tr>
</table>
</BODY>
</HTML>

<%
function cdrivetype(tnum)
    Select Case tnum
        Case 0: cdrivetype = "δ֪"
        Case 1: cdrivetype = "���ƶ�����"
        Case 2: cdrivetype = "����Ӳ��"
        Case 3: cdrivetype = "�������"
        Case 4: cdrivetype = "CD-ROM"
        Case 5: cdrivetype = "RAM ����"
    End Select
end function

function cIsReady(trd)
    Select Case trd
		case true: cIsReady="<font class=fonts><b>��</b></font>"
		case false: cIsReady="<font color='red'><b>��</b></font>"
	End Select
end function

function cSize(tSize)
    if tSize>=1073741824 then
		cSize=int((tSize/1073741824)*1000)/1000 & " GB"
    elseif tSize>=1048576 then
    	cSize=int((tSize/1048576)*1000)/1000 & " MB"
    elseif tSize>=1024 then
		cSize=int((tSize/1024)*1000)/1000 & " KB"
	else
		cSize=tSize & "B"
	end if
end function
%>







