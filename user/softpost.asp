<!--#include file="config.asp"-->
<!--#include file="check.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
<!--#include file="head.inc"-->
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
Call InnerLocation("�������")
If CInt(GroupSetting(11)) = 0 Then
	ErrMsg = ErrMsg + "<li>�Բ������Ļ�ԱȨ�޲��㣬����Ҫ��Ȩ������ϵ����Ա��</li>"
	Founderr = True
End If


if CInt(enchiasp.membergrade)<CInt(enchiasp.postgrade) then
	ErrMsg = ErrMsg + "<li>�Բ�������Ȩ�޲���,��û�з��������Ȩ�ޣ�����Ҫ��Ȩ������ϵ����Ա��</li>"
	Founderr = True

end if


Dim Rs,SQL,i,SoftID
ChannelID = enchiasp.ChkNumeric(Request("ChannelID"))

if Request("ChannelID")="" then
	ErrMsg = ErrMsg + "<li>��������,�벻Ҫ�ֹ����ò�����</li>"
	Founderr = True
end if

If ChannelID < 2 Then ChannelID = 2
ChannelID = CLng(ChannelID)

Dim Action:Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveNewSoft
	Case "view"
		Call SoftView
	Case Else
		Call showmain
End Select
If Founderr = True Then
	Call Returnerr(ErrMsg)
End If

Sub showmain()
	If Founderr = True Then Exit Sub
	Dim Channel_Setting
	Channel_Setting = Split(enchiasp.Channel_Setting, "|||")
%>
<script language=JavaScript>
var _maxCount = '<%=CLng(GroupSetting(17))%>';
function ToRunsystem(addTitle) {
	var revisedTitle;
	var currentTitle;
	currentTitle = document.myform.RunSystem.value;
	revisedTitle = currentTitle+addTitle;
	document.myform.RunSystem.value=revisedTitle;
	document.myform.RunSystem.focus();
	return; 
}
function doSubmit(){
	if (document.myform.SoftName.value==""){
		alert("������Ʋ���Ϊ�գ�");
		return false;
	}
	if (document.myform.DownAddress.value==""){
		alert("������Ҫ��дһ�����ص�ַ�ɣ�");
		return false;
	}
	if (document.myform.SiteName.value==""){
		alert("�������Ʋ���Ϊ�գ�");
		return false;
	}
	if (document.myform.ClassID.value==""){
		alert("��һ�������Ѿ����������࣬��ѡ�����������࣡");
		return false;
	}
	if (document.myform.ClassID.value=="0"){
		alert("�÷������ⲿ���ӣ�����������ݣ�");
		return false;
	}
	if (document.myform.RunSystem.value==""){
		alert("������л�������Ϊ�գ�");
		return false;
	}
	if (document.myform.SoftType.value==""){
		alert("������Ͳ���Ϊ�գ�");
		return false;
	}
	if (document.myform.SoftSize.value==""){
		alert("�����С��û����д��");
		return false;
	}
	<%If CInt(GroupSetting(2)) = 1 Then%>
	if (document.myform.codestr.value==""){
		alert("����д��֤�룡");
		return false;
	}
	<%End If%>
	myform.content1.value = getHTML(); 
	MessageLength = Composition.document.body.innerHTML.length;
	if(MessageLength < 2){
		alert("�����鲻��С��2���ַ���");
		return false;
	}
	if(MessageLength > _maxCount){
		alert("�����鲻�ܳ���"+_maxCount+"���ַ���");
		return false;
	}
	document.myform.Submit1.disabled = true;
	document.myform.submit();
}
</script>
<div onkeydown=CtrlEnter()>
<table  border="0" align="center" cellpadding="3" cellspacing="1" class="UserTableBorder">
        <tr>
          <th colspan="4">&gt;&gt;�������&lt;&lt;</th>
        </tr>
	<form method=Post name="myform" action="softpost.asp">
	<input type="Hidden" name="action" value="save">
	<input type=hidden name=ChannelID value="<%=ChannelID%>">
        <tr>
          <td width="15%" align="right" nowrap class="UserTableRow2"><strong>��������</strong></td>
          <td width="85%" class="UserTableRow1">
<%
	Dim sClassSelect
	Response.Write "<select name=""ClassID"" id=""ClassID"">"
	sClassSelect = enchiasp.LoadSelectClass(ChannelID)
	Response.Write sClassSelect
	Response.Write "</select>"
%>
	  </td>
        </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>�������</strong></td>
          <td class="UserTableRow1"><input name="SoftName" type="text" id="SoftName" size="45" value=""> 
          <font color=red>*</font> <strong>����汾</strong><input name="SoftVer" type="text" id="SoftVer" size="20" value=""></td>
	  </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>������</strong></td>
          <td class="UserTableRow1"><input name="Related" type="text" id="Related" size="60" value=""> <font color=red>*</font></td>
        </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>���л���</strong></td>
          <td class="UserTableRow1"><input name="RunSystem" type="text" size="60" value="<%=Channel_Setting(1)%>"><br>
<%
	Dim RunSystem
	RunSystem = Split(Channel_Setting(0), "|")
	For i = 0 To UBound(RunSystem)
		Response.Write "<a href='javascript:ToRunsystem(""" & Trim(RunSystem(i)) & """)'><u>" & Trim(RunSystem(i)) & "</u></a>  "
		If i = 10 Then Response.Write "<br>"
	Next
%>
          </td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�������</strong></td>
          <td class="UserTableRow1">
<%
	Dim SoftType
	SoftType = Split(Channel_Setting(2), ",")
	For i = 0 To UBound(SoftType)
		Response.Write "<input type=""radio"" name=""SoftType"" value=""" & Trim(SoftType(i)) & """ "
		If i = 0 Then Response.Write " checked"
		Response.Write ">" & Trim(SoftType(i)) & " "
		If i = 6 Then Response.Write "<br>"
	Next
%>
	  </td>
        </tr>
        <tr>
          <td align="right" class="UserTableRow2"><strong>Ԥ��ͼƬ</strong></td>
          <td class="UserTableRow1"><input name="SoftImage" id="ImageUrl" type="text" size="60" value=""></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�ϴ�ͼƬ</strong></td>
          <td class="UserTableRow1"><iframe name="image" frameborder=0 width='100%' height=55 scrolling=no src=upload.asp?ChannelID=2></iframe></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�����С</strong></td>
          <td class="UserTableRow1">
	<input type="text" name="SoftSize" size="14" id="filesize" onkeyup="if(isNaN(this.value))this.value=''" value=''> <input name="SizeUnit" type="radio" value="KB" checked> KB <input type="radio" name="SizeUnit" value="MB"> MB <font color="#FF0000">��</font>
	<strong>��ѹ����</strong>
	<input type="text" name="Decode" size="15" maxlength="100" value=''> <font color="#808080">û��������</font>
          </td>
	</tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�������</strong></td>
          <td class="UserTableRow1">
<%
	Response.Write " <select name=""impower"">"
	Dim ImpowerStr
	ImpowerStr = Split(Channel_Setting(3), ",")
	For i = 0 To UBound(ImpowerStr)
		Response.Write " <option value=""" & ImpowerStr(i) & """>" & ImpowerStr(i) & "</option>"
	Next
	Response.Write " </select>&nbsp;&nbsp;"
	Response.Write " <select name=""Languages"">"
	Response.Write " "
	Dim LanguagesStr
	LanguagesStr = Split(Channel_Setting(4), ",")
	For i = 0 To UBound(LanguagesStr)
		Response.Write " <option value=""" & LanguagesStr(i) & """>" & LanguagesStr(i) & "</option>"
	Next
	Response.Write " </select>&nbsp;&nbsp;"
%>
		<select name="star">
		<option value=5>������</option>
          	<option value=4>�����</option>
          	<option value=3 selected>����</option>
		<option value=2>���</option>
		<option value=1>��</option>
          </select>&nbsp;&nbsp;
	  <strong><font color=blue>ע������ļ۸�</font></strong>
	  <input name="SoftPrice" type="text" size="10" onkeyup="if(isNaN(this.value))this.value=''" value="0"> Ԫ
	  </td>
	</tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>��ϵ��ʽ</strong></td>
          <td class="UserTableRow1">
		<input name="Contact" type="text" size="33"> 
		<strong>������ҳ</strong>
		<input name="Homepage" type="text" size="30">
	  </td>
	</tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�������</strong></td>
          <td class="UserTableRow1">
		<input name="Author" type="text" size="33">
		<strong>ע����ַ</strong>
		<input name="Regsite" type="text" size="30">
	  </td>
        <tr>
          <td align="right" class="UserTableRow2"><strong>������</strong></td>
          <td class="UserTableRow1"><textarea name='content1' id='content1' style='display:none'></textarea>
		<script Language=Javascript src="../editor/editor1.js"></script></td>
        </tr>
	        </tr>
<%
	If CInt(GroupSetting(2)) = 1 Then
%>
	<tr>
		<td class=UserTableRow2 align="right"><strong>��֤��</strong></td>
		<td class=UserTableRow1><input type="text" name="codestr" maxlength="4" size="4">&nbsp;<img src="../inc/getcode.asp"></td>
	</tr>
<%
	End If
%>
	<tr>
          <td align="right" class="UserTableRow2"><strong>��ֹ����</strong></td>
          <td class="UserTableRow1"><input name="ForbidEssay" type="checkbox" id="ForbidEssay" value="1">&nbsp;&nbsp;&nbsp;&nbsp;
	  <strong>��������</strong>
	  <input name="isAccept" type="checkbox" id="isAccept" value="1" disabled> �ǣ�<font color=blue>���ѡ�еĻ���ֱ�ӷ�����������˺���ܷ�����</font>��</td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>���ص�ַ</strong></td>
          <td class="UserTableRow1"><input name="SiteName" type="text" size="15" value="���ص�ַ1">
	  <input name="DownAddress" type="text" id="filePath" size="60" value=""> <font color=red>*</font></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>���ص�ַ</strong></td>
          <td class="UserTableRow1"><input name="SiteName" type="text" size="15" value="���ص�ַ2">
	  <input name="DownAddress" type="text" size="60" value=""></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>���ص�ַ</strong></td>
          <td class="UserTableRow1"><input name="SiteName" type="text" size="15" value="���ص�ַ3">
	  <input name="DownAddress" type="text" size="60" value=""></td>
        </tr>
	<tr>
          <td align="right" class="UserTableRow2"><strong>�ļ��ϴ�</strong></td>
          <td class="UserTableRow1"><iframe name="file1" frameborder=0 width='100%' height=60 scrolling=no src=upfile.asp?ChannelID=<%=ChannelID%>></iframe></td>
        </tr>
        <tr align="center">
          <td colspan="4" class="UserTableRow2">
	  <input type="button" name="Submit4" onclick="javascript:history.go(-1)" value="������һҳ" class=Button>
	  <input type="button" name="Submit1" value="���ڷ���" class=Button onclick="doSubmit();"></td>
        </tr></form>
      </table></div>
<%
End Sub

Function FormatDownAddress(ByVal str)
	If Trim(str) = ""  Or Trim(str) = "|||" Then
		FormatDownAddress = ""
		Exit Function
	End If
	Dim strDownAddress,sDownAddress,sDownSiteName
	Dim n,AddressNum,strAddress,strDownName,strTemp
	On Error Resume Next
	strDownAddress = Split(str, "|||")
	sDownAddress = Split(strDownAddress(1), "|")
	sDownSiteName = Split(strDownAddress(0), "|")
	If UBound(sDownAddress) < UBound(sDownSiteName) Then
		AddressNum = UBound(sDownAddress)
	Else
		AddressNum = UBound(sDownSiteName)
	End If
	strAddress = ""
	strDownName = ""
	For n = 0 To CInt(AddressNum)
		If Trim(sDownAddress(n)) <> "" And Trim(sDownSiteName(n)) <> "" Then
			strAddress = strAddress & Trim(sDownAddress(n)) & "|"
			strDownName = strDownName & Trim(sDownSiteName(n)) & "|"
		End If
	Next
	If Len(strDownName) > 0 Then strDownName = Left(strDownName, Len(strDownName) - 1)
	If Len(strAddress) > 0 Then strAddress = Left(strAddress, Len(strAddress) - 1)
	strTemp = strDownName & "|||" & strAddress
	FormatDownAddress = Trim(strTemp)
End Function
Sub SaveNewSoft()
	Dim TextContent,isAccept,ForbidEssay,DownAddress
	If CLng(UserToday(2)) => CLng(GroupSetting(14)) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ÿ�����ֻ�ܷ���<font color=red><b>" & GroupSetting(14) & "</b></font>������������Ҫ�������������������ɣ�</li>"
	End If
	'��ֹ�ⲿ�ύ
	If enchiasp.CheckPost=false Then
		ErrMsg = ErrMsg + Postmsg
		FoundErr = True
		Exit Sub
	End If
	If CInt(GroupSetting(2)) = 1 Then
		If Not enchiasp.CodeIsTrue() Then
			ErrMsg = ErrMsg + "<meta http-equiv=refresh content=""2;URL="&Request.ServerVariables("HTTP_REFERER")&"""><li>��֤��У��ʧ�ܣ��뷵��ˢ��ҳ�����ԡ�������Զ�����</li>"
			Founderr = True
			Exit Sub
		End If
		Session("GetCode") = ""
	End If

	If Trim(Request.Form("SoftName")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������Ʋ���Ϊ�գ�</li>"
	End If
	If Len(Request.Form("SoftName")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>������Ʋ��ܳ���200���ַ���</li>"
	End If
	If Len(Request.Form("Related")) => 200 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>���������ܳ���200���ַ���</li>"
	End If
	If Not IsNumeric(Request.Form("star")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>����Ǽ�����Ϊ�ա�</li>"
	End If
	If CLng(Request.Form("ClassID")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�÷������ⲿ���ӣ�������������</li>"
	End If
	If Not IsNumeric(Request.Form("ClassID")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��һ�������Ѿ����������࣬������������</li>"
	End If
	If Trim(Request.Form("SoftType")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ѡ��������ͣ�</li>"
	End If
	If Trim(Request.Form("impower")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ѡ�������Ȩ��ʽ��</li>"
	End If
	If Trim(Request.Form("Languages")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ѡ��������ԣ�</li>"
	End If
	If Trim(Request.Form("content1")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�����鲻��Ϊ�գ�</li>"
	End If
	TextContent = ""
	For i = 1 To Request.Form("content1").Count
		TextContent = TextContent & Request.Form("content1")(i)
	Next
	If Len(Request.Form("RunSystem")) = 0 Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>���л�������Ϊ�գ�</li>"
	End If
	If Not IsNumeric(Request.Form("SoftSize")) Then
		Founderr = True
		ErrMsg = ErrMsg + "<li>�����С������������</li>"
	End If
	If CInt(Request("isAccept")) = 1 Then
		isAccept = 1
	Else
		isAccept = 0
	End If
	If CInt(Request.Form("ForbidEssay")) = 1 Then
		ForbidEssay = 1
	Else
		ForbidEssay = 0
	End If
	'---- ��ȡ���ص�ַ���е�����
	Dim TempAddress,TempSiteName,TempDownAddress
	Dim strTempAddress,strTempSiteName
	If Trim(Request.Form("DownAddress")) <> "" And Trim(Request.Form("SiteName")) <> "" Then
		strTempAddress = ""
		For Each TempAddress In Request.Form("DownAddress")
			If LCase(TempAddress) <> "del" And Trim(TempAddress) <> "" Then
				strTempAddress = strTempAddress & Replace(TempAddress, "|", "") & "|"
			End If
		Next
		If Len(strTempAddress) > 0 Then strTempAddress = Left(strTempAddress, Len(strTempAddress) - 1)
		strTempSiteName = ""
		For Each TempSiteName In Request.Form("SiteName")
			If LCase(TempSiteName) <> "del" And Trim(TempSiteName) <> "" Then
				strTempSiteName = strTempSiteName & Replace(TempSiteName, "|", "") & "|"
			End If
		Next
		If Len(strTempSiteName) > 0 Then strTempSiteName = Left(strTempSiteName, Len(strTempSiteName) - 1)
		TempDownAddress = enchiasp.CheckStr(strTempSiteName &"|||"& strTempAddress)
	Else
		TempDownAddress = ""
	End If
	DownAddress = FormatDownAddress(TempDownAddress)
	If Founderr = True Then Exit Sub
	On Error Resume Next
	Call PreventRefresh  '��ˢ��
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM ECCMS_SoftList WHERE (SoftID is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("ChannelID") = ChannelID
		Rs("ClassID") = enchiasp.ChkNumeric(Request.Form("ClassID"))
		Rs("SpecialID") = 0
		Rs("SoftName") = enchiasp.ChkFormStr(Request.Form("SoftName"))
		Rs("SoftVer") = enchiasp.ChkFormStr(Request.Form("SoftVer"))
		Rs("ColorMode") = 0
		Rs("FontMode") = 0
		Rs("Related") = enchiasp.ChkFormStr(Request.Form("Related"))
		Rs("Content") = Html2Ubb(TextContent)
		Rs("Languages") = enchiasp.ChkFormStr(Request.Form("Languages"))
		Rs("SoftType") = enchiasp.ChkFormStr(Request.Form("SoftType"))
		Rs("RunSystem") = enchiasp.ChkFormStr(Request.Form("RunSystem"))
		Rs("impower") = enchiasp.ChkFormStr(Request.Form("impower"))
		If UCase(Request.Form("SizeUnit")) = "MB" Then
			Rs("SoftSize") = enchiasp.CheckNumeric(Request.Form("SoftSize") * 1024)
		Else
			Rs("SoftSize") = enchiasp.CheckNumeric(Request.Form("SoftSize"))
		End If
		Rs("star") = enchiasp.ChkNumeric(Request.Form("star"))
		Rs("Homepage") = enchiasp.ChkFormStr(Request.Form("Homepage"))
		Rs("Contact") = enchiasp.ChkFormStr(Request.Form("Contact"))
		Rs("Author") = enchiasp.ChkFormStr(Request.Form("Author"))
		Rs("Regsite") = enchiasp.ChkFormStr(Request.Form("Regsite"))
		Rs("showreg") = 0
		Rs("username") = Trim(enchiasp.MemberName)
		Rs("PointNum") = 0
		Rs("SoftPrice") = enchiasp.CheckNumeric(Request.Form("SoftPrice"))
		Rs("SoftTime") = Now()
		Rs("isTop") = 0
		Rs("AllHits") = 0
		Rs("DayHits") = 0
		Rs("WeekHits") = 0
		Rs("MonthHits") = 0
		Rs("HitsTime") = Now()
		Rs("HtmlFileDate") = Trim(enchiasp.HtmlRndFileName)
		Rs("SoftImage") = enchiasp.ChkFormStr(Request.Form("SoftImage"))
		Rs("Decode") = enchiasp.ChkFormStr(Request.Form("Decode"))
		Rs("isBest") = 0
		Rs("UserGroup") = 0
		Rs("isUpdate") = 1
		Rs("ErrCode") = 0
		Rs("downid") = 0
		Rs("DownFileName") = ""
		Rs("DownAddress") = enchiasp.ChkFormStr(DownAddress)
		Rs("isAccept") = isAccept
		Rs("ForbidEssay") = ForbidEssay
	Rs.update
	Rs.Close
	Rs.Open "select top 1 softid from ECCMS_SoftList where ChannelID=" & ChannelID & " order by softid desc", Conn, 1, 1
	SoftID = Rs("SoftID")
	Rs.Close:Set Rs = Nothing
	Dim strUserToday
	strUserToday = UserToday(0) &","& UserToday(1) &","& UserToday(2)+1 &","& UserToday(3) &","& UserToday(4) &","& UserToday(5)
	UpdateUserToday(strUserToday)
	Call Returnsuc("<li>��ϲ�����ύ�ɹ�����ȴ�����Ա��֤����ʽ������</li><li><a href=?action=view&ChannelID=" & ChannelID & "&SoftID=" & SoftID & ">����˴��鿴�����</a></li>")
End Sub
Private Sub SoftView()
	If Request("SoftID") = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry����ѡ���˴����ϵͳ������</li>"
		Exit Sub
	End If
	If ChannelID = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ָ��Ƶ����</li>"
		Exit Sub
	End If
	SQL = "SELECT * FROM ECCMS_SoftList WHERE ChannelID=" & ChannelID & " And username='" & enchiasp.MemberName & "' And SoftID=" & Request("SoftID")
	Set Rs = enchiasp.Execute(SQL)
	If Rs.bof And Rs.EOF Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>Sorry��û���ҵ��κ������������ѡ���˴����ϵͳ������</li>"
		Exit Sub
	Else
	Dim strDownAddress,sDownAddress
	strDownAddress = Split(Rs("DownAddress"), "|||")
	sDownAddress = Split(strDownAddress(1), "|")
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="UserTableBorder" style="table-layout:fixed;word-break:break-all">
	<tr>
	  <th colspan="2">&gt;&gt;�鿴�����Ϣ&lt;&lt;</th>
	</tr>
	<tr>
	  <td align="center" class="UserTableRow2" colspan="2"><font size=3 color=blue><%=enchiasp.ReadFontMode(Rs("SoftName"),Rs("ColorMode"),Rs("FontMode"))%>&nbsp;<%=Rs("SoftVer")%></font></td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>������л�����</strong> <%=Rs("RunSystem")%></td>
	  <td class="UserTableRow1"><strong>������ͣ�</strong> <%=Rs("SoftType")%></td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>�����С��</strong> <%=Rs("SoftSize")%></td>
	  <td class="UserTableRow1"><strong>����Ǽ���</strong> 
<%
Response.Write "<font color=red>"
For i = 1 to Rs("star")
	Response.Write "��"
Next
Response.Write "</font>"
%>
	  </td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>������ԣ�</strong> <%=Rs("Languages")%></td>
	  <td class="UserTableRow1"><strong>��Ȩ��ʽ��</strong> <%=Rs("impower")%></td>
	</tr>
	<tr>
	  <td class="UserTableRow1"><strong>����ʱ�䣺</strong> <%=Rs("SoftTime")%></td>
	  <td class="UserTableRow1"><strong>������ҳ��</strong> <%=Rs("Homepage")%></td>
	</tr>
	<tr>
	  <td colspan="2" class="UserTableRow1"><strong>�����飺</strong><br><%=UBBCode(Rs("content"))%></td>
	</tr>
	<tr>
	  <td colspan="2" class="UserTableRow1"><strong>���ص�ַ��</strong><br>
<%
	For i = 0 To UBound(sDownAddress)
		Response.Write "<li><a href=""" & sDownAddress(i) & """ target=_blank>" & sDownAddress(i) & "</a></li>" & vbNewLine
	Next

%>
	  </td>
	</tr>
	<tr>
	  <td align="center" colspan="2" class="UserTableRow2"><input type="button" onclick="javascript:window.close()" value="�رձ�����" name="B2" class=Button>&nbsp;&nbsp;
	  <input type="button" name="Submit1" onclick="javascript:location.href='#'" value="���ض���" class=button>&nbsp;&nbsp;
	  </td>
	</tr>
</table>
<%

	End If
	Rs.Close
	Set Rs = Nothing 
End Sub

%><!--#include file="foot.inc"-->