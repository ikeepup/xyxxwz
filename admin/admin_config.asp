<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
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
Response.Write "<table border=0 align=center cellspacing=1 class=TableBorder cellpadding=3>"
Response.Write "	<tr>"
Response.Write "		<th>��վ�������ù���</th>"
Response.Write "	</tr>"
Response.Write "	<tr>"
Response.Write "		<td class=TableRow><a href='admin_config.asp#setting1'>ϵͳ������Ϣ</a> |"
Response.Write "		<a href='admin_config.asp#setting2'>ϵͳ�ʼ�����</a> | "
Response.Write "		<a href='admin_config.asp#setting3'>ע���û�����</a> | "
Response.Write "		<a href='admin_config.asp#setting4'>ϵͳ��������</a> |"
Response.Write "		<a href='admin_config.asp#setting5'>�����ַ�����</a> |"
Response.Write "		<a href='admin_config.asp#setting6'>����Ա��ȫ����</a> |"
Response.Write "		<a href='?action=edit'><font color=blue>���ݹؼ�������</font></a> |"
Response.Write "		<a href='?action=reload'><font color=red>�ؽ�����</font></a></td>"
Response.Write "	</tr>"
Response.Write "</table>"
Response.Write "<br>"

Dim Action
Action = enchiasp.RemoveBadCharacters(Request("action"))
If Not ChkAdmin("SiteConfig") Then
	Server.Transfer("showerr.asp")
	Request.End
End If
Select Case LCase(Action)
	Case "save"
		Call SaveConfig
	Case "reload"
		Call ReloadCache
	Case "edit"
		Call EditContentKeyword
	Case "savedit"
		Call SaveContentKeyword
	Case Else
		Call MainPage
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn
Private Sub MainPage()
	Dim RootDirStr,SiteRootDir
	Dim ChinaeBankPay,PreviewSetting
	RootDirStr = Left(LCase(Request.ServerVariables("SCRIPT_NAME")), InStrRev(LCase(Request.ServerVariables("SCRIPT_NAME")), "/") - 1)
	SiteRootDir = Left(RootDirStr, InStrRev(RootDirStr, "/"))
	SiteRootDir = Trim(SiteRootDir)
	ChinaeBankPay = Split(enchiasp.ChinaeBank, "|||")
	PreviewSetting = Split(enchiasp.PreviewSetting, ",")
	If UBound(PreviewSetting) < 15 Then
		PreviewSetting = Split("999,1,2,110,90,www.enchiasp.cn,12,#FF0000,Arial,0,images/WaterMap.gif,0.8,#0066FF,100,35,0", ",")
	End If 
%>
<iframe width="260" height="165" id="colourPalette" src="include/selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
<div onkeydown=CtrlEnter()>
<table border="0" align="center" cellspacing="1" class="TableBorder" cellpadding="3">
	<tr>
		<th align="left"><img src="images/welcome.gif" width="16" height="17" align="absMiddle"> ��վ��������</th>
	</tr>
	<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  var checkOK = "0123456789-,";
  var checkStr = theForm.AddUserPoint.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("�� ���������� ���У�ֻ������ ���� �ַ���");
    theForm.AddUserPoint.focus();
    return (false);
  }

  var checkOK = "0123456789-,";
  var checkStr = theForm.ActionTime.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("�� ���������� ���У�ֻ������ ���� �ַ���");
    theForm.ActionTime.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form name=FrontPage_Form1 method="POST" action="?action=save" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
	<tr>
	<td class="TableRow1">
	<fieldset style="cursor: default"><legend>&nbsp;��վ������Ϣ<a name="setting2"></a>[<a href="#top">����</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
		
		<tr>
			<td class="TableRow1" width="35%"><div class="divbody">��վ���ƣ�</div></td>
			<td class="TableRow1" width="65%">
			<input type="text" name="SiteName" size="35" value="<%=enchiasp.SiteName%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">��վURL��</div></td>
			<td class="TableRow2"><input type="text" name="SiteUrl" size="35" value="<%=RootPath2DomainPath("")%>">
			<font color="#FF0000">ϵͳ�Զ���ȡ���벻Ҫ�޸�</font></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">����ԱEmail��</div></td>
			<td class="TableRow1">
			<input type="text" name="MasterMail" size="25" value="<%=enchiasp.MasterMail%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">��ҳ�ļ�����</div></td>
			<td class="TableRow2"><input type="text" name="IndexName" size="25" value="<%=enchiasp.IndexName%>"></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">ϵͳ���ڸ�Ŀ¼��</div></td>
			<td class="TableRow1">
			<input type="text" size="25" value="<%=enchiasp.InstallDir%>" disabled>
			<input type=hidden name="InstallDir" value="<%=SiteRootDir%>">&nbsp;* 
			<font color="#FF0000">ϵͳ�Զ���ȡ�������ֶ�����</font></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">վ��ؼ��֣�<br>
			������������������������վ�Ĺؼ�����</div></td>
			<td class="TableRow2"><textarea rows="3" name="keywords" cols="60"><%=enchiasp.keywords%></textarea></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">��վ��Ȩ��Ϣ��</div></td>
			<td class="TableRow1" width="65%">
			<textarea rows="5" name="Copyright" cols="60"><%=enchiasp.Copyright%></textarea></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">�Ƿ�ر���վ��</div></td>
			<td class="TableRow2">
			<input type="radio" name="IstopSite" value="1" <%If CInt(enchiasp.IstopSite) = 1 Then Response.Write (" checked")%>> �ر� 
			<input type="radio" name="IstopSite" value="0" <%If CInt(enchiasp.IstopSite) = 0 Then Response.Write (" checked")%>> ��</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">��վά��˵����<br>֧��HTML����,���ܳ���250���ַ�</div></td>
			<td class="TableRow1">
			<textarea rows="5" name="StopReadme" cols="60"><%=enchiasp.StopReadme%></textarea></td>
		</tr>
	</table></fieldset>
	<br>
	<fieldset style="cursor: default"><legend>&nbsp;ϵͳ�ʼ�����<a name="setting2"></a>[<a href="#top">����</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
		<tr>
			<td class="TableRow2" width="35%"><div class="divbody">�Ƿ�ر��ʼ����ܣ�</div></td>
			<td class="TableRow2">
			<input type="radio" name="IsCloseMail" value="1" <%If CInt(enchiasp.IsCloseMail) = 1 Then Response.Write (" checked")%>> �ر� 
			<input type="radio" name="IsCloseMail" value="0" <%If CInt(enchiasp.IsCloseMail) = 0 Then Response.Write (" checked")%>> ��</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">�����ʼ��������</div></td>
			<td class="TableRow1"><select size="1" name="SendMailType"  onChange="chkselect(options[selectedIndex].value,'know1');">
			<option value="0" <%If CInt(enchiasp.SendMailType) = 0 Then Response.Write (" selected")%>>��֧��</option>
			<option value="1" <%If CInt(enchiasp.SendMailType) = 1 Then Response.Write (" selected")%>>JMAIL</option>
			<option value="2" <%If CInt(enchiasp.SendMailType) = 2 Then Response.Write (" selected")%>>CDONTS</option>
			<option value="3" <%If CInt(enchiasp.SendMailType) = 3 Then Response.Write (" selected")%>>ASPEMAIL</option>
			</select><div id=know1></div></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">ϵͳ����ԱEmail��<br>
			���û������ʼ�ʱ����ʾ����ԴEmail��Ϣ</div></td>
			<td class="TableRow2">
			<input type="text" name="MailFrom" size="25" value="<%=enchiasp.MailFrom%>"></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">SMTP Server��ַ��</div></td>
			<td class="TableRow1">
			<input type="text" name="MailServer" size="25" value="<%=enchiasp.MailServer%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">�ʼ���¼�û�����</div></td>
			<td class="TableRow2">
			<input type="text" name="MailUserName" size="25" value="<%=enchiasp.MailUserName%>"></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">�ʼ���¼���룺</div></td>
			<td class="TableRow1">
			<input type="password" name="MailPassword" size="25" value="<%=enchiasp.MailPassword%>"></td>
		</tr>
	</table></fieldset>
	<br>
	<fieldset style="cursor: default"><legend>&nbsp;ע���û�����<a name="setting3"></a>[<a href="#top">����</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
		<tr>
			<td class="TableRow2" width="35%"><div class="divbody">�Ƿ��������û�ע�᣺</div></td>
			<td class="TableRow2">
			<input type="radio" name="CheckUserReg" value="0" <%If CInt(enchiasp.CheckUserReg) = 0 Then Response.Write (" checked")%>> �� 
			<input type="radio" name="CheckUserReg" value="1" <%If CInt(enchiasp.CheckUserReg) = 1 Then Response.Write (" checked")%>> ��</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">ע���Ա�Ƿ�Ҫ����Ա��֤��</div></td>
			<td class="TableRow1">
			<input type="radio" name="AdminCheckReg" value="0" <%If CInt(enchiasp.AdminCheckReg) = 0 Then Response.Write (" checked")%>> �� 
			<input type="radio" name="AdminCheckReg" value="1" <%If CInt(enchiasp.AdminCheckReg) = 1 Then Response.Write (" checked")%>> ��</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">ע���Ա�Ƿ�Email֪ͨ���룺<br>ȷ������վ��֧�ַ���mail������������Ϊϵͳ�������</div></td>
			<td class="TableRow2">
			<input type="radio" name="MailInformPass" value="0" <%If CInt(enchiasp.MailInformPass) = 0 Then Response.Write (" checked")%>> �ر� 
			<input type="radio" name="MailInformPass" value="1" <%If CInt(enchiasp.MailInformPass) = 1 Then Response.Write (" checked")%>> ��</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">һ��Emailֻ��ע��һ����Ա��</div></td>
			<td class="TableRow1">
			<input type="radio" name="ChkSameMail" value="0" <%If CInt(enchiasp.ChkSameMail) = 0 Then Response.Write (" checked")%>> �ر� 
			<input type="radio" name="ChkSameMail" value="1" <%If CInt(enchiasp.ChkSameMail) = 1 Then Response.Write (" checked")%>> ��</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">����ע���ʼ���Ϣ��<br>��ȷ���������ʼ�����</div></td>
			<td class="TableRow2">
			<input type="radio" name="SendRegMessage" value="0" <%If CInt(enchiasp.SendRegMessage) = 0 Then Response.Write (" checked")%>> �ر� 
			<input type="radio" name="SendRegMessage" value="1" <%If CInt(enchiasp.SendRegMessage) = 1 Then Response.Write (" checked")%>> ��</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">ע���Ա���ӵĵ�����<br>
			����������</div></td>
			<td class="TableRow1">
			&nbsp;<!--webbot bot="Validation" s-display-name="����������" s-data-type="Integer" s-number-separators="," --><input type="text" name="AddUserPoint" size="15" value="<%=enchiasp.AddUserPoint%>"></td>
		</tr>
	</table></fieldset>
	<br>
	<fieldset style="cursor: default"><legend>&nbsp;ϵͳ��������<a name="setting4"></a>[<a href="#top">����</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
		<tr>
			<td class="TableRow2" width="35%"><div class="divbody">�Ƿ���ȫ��������<br>
			ȫ������ռ�÷�������Դ�����鿪��</div></td>
			<td class="TableRow2">
			<input type="radio" name="FullContQuery" value="0" <%If CInt(enchiasp.FullContQuery) = 0 Then Response.Write (" checked")%>> �� 
			<input type="radio" name="FullContQuery" value="1" <%If CInt(enchiasp.FullContQuery) = 1 Then Response.Write (" checked")%>> ��</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">ɾ������û�ʱ�䣺<br>
			��λ�����ӣ�����������</div></td>
			<td class="TableRow1">
			&nbsp;<!--webbot bot="Validation" s-display-name="����������" s-data-type="Integer" s-number-separators="," --><input type="text" name="ActionTime" size="15" value="<%=enchiasp.ActionTime%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">�Ƿ���ʾҳ��ִ��ʱ�䣺</div></td>
			<td class="TableRow2">
			<input type="radio" name="IsRunTime" value="0" <%If CInt(enchiasp.IsRunTime) = 0 Then Response.Write (" checked")%>> �� 
			<input type="radio" name="IsRunTime" value="1" <%If CInt(enchiasp.IsRunTime) = 1 Then Response.Write (" checked")%>> ��</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">ѡȡ�ϴ������</div></td>
			<td class="TableRow1"><select size="1" name="UploadClass" onChange="chkselect(options[selectedIndex].value,'know2');">
			<option value="999" <%If CInt(enchiasp.UploadClass) = 999 Then Response.Write (" selected")%>>�ر��ϴ�����</option>
			<option value="0" <%If CInt(enchiasp.UploadClass) = 0 Then Response.Write (" selected")%>>������ϴ���</option>
			<option value="1" <%If CInt(enchiasp.UploadClass) = 1 Then Response.Write (" selected")%>>Aspupload3.0���</option>
			<option value="2" <%If CInt(enchiasp.UploadClass) = 2 Then Response.Write (" selected")%>>SA-FileUp 4.0���</option>
			</select><div id="know2"></div></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">�ϴ��ļ���С�� ��λ(KB)</div></td>
			<td class="TableRow2">
			<input type="text" name="UploadFileSize" size="15" value="<%=enchiasp.UploadFileSize%>"> <b>KB</></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">�ϴ��ļ����ͣ� ���á�|���ֿ�</div></td>
			<td class="TableRow1">
			<input type="text" name="UploadFileType" size="60" value="<%=enchiasp.UploadFileType%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">�Ƿ�ر������������룺</div></td>
			<td class="TableRow2">
			<input type="radio" name="StopApplyLink" value="1" <%If CInt(enchiasp.StopApplyLink) = 1 Then Response.Write (" checked")%>> �ر� 
			<input type="radio" name="StopApplyLink" value="0" <%If CInt(enchiasp.StopApplyLink) = 0 Then Response.Write (" checked")%>> ��</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">FSO��������ƣ�<br>
			FSOĬ�����ƣ�Scripting.FileSystemObject</div></td>
			<td class="TableRow1">
			<input type="text" name="FSO_ScriptName" size="35" value="<%=enchiasp.FSO_ScriptName%>"><br>
			ĳЩ��վΪ�˰�ȫ����FSO��������ƽ��и����Դﵽ����FSO��Ŀ�ġ���������վ���������ģ����ڴ�������Ĺ������ơ�</div></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">��ʼ������ɫ��</div><br>
			ÿ����ɫ��ֵ��ġ�,���ֿ�</td>
			<td class="TableRow2"><textarea rows="3" name="InitTitleColor" cols="60"><%=enchiasp.InitTitleColor%></textarea></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">��ѡ������֧�����ܣ�</div></td>
			<td class="TableRow1">
			<input type="radio" name="StopBankPay" value="0" <%If CInt(enchiasp.StopBankPay) = 0 Then Response.Write (" checked")%>> �ر�����֧��&nbsp;&nbsp;
			<input type="radio" name="StopBankPay" value="1" <%If CInt(enchiasp.StopBankPay) = 1 Then Response.Write (" checked")%>> NPS ����֧��&nbsp;&nbsp;
			<input type="radio" name="StopBankPay" value="2" <%If CInt(enchiasp.StopBankPay) = 2 Then Response.Write (" checked")%>> ��������֧��&nbsp;&nbsp;
			<br><font color=blue>* ����㻹û��ͨNPS����֧������<a href="http://www.nps.cn/merchant/join_agreement.jsp" target="_blank"><font color=RED><strong>��˿�ͨNPS����֧��</strong></font></a>��</font>
			<br><font color=blue>* ����㻹û��ͨ��������֧������<a href="http://www.chinaebank.cn/" target="_blank"><font color=RED><strong>��˿�ͨ�й���������֧��</strong></font></a>��</font>
			</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">����֧�� ID����<a href="http://www.nps.cn/" target="_blank">NPS</a>����<a href="http://www.chinaebank.cn/" target="_blank">��������</a>ID��</div></td>
			<td class="TableRow2">
			<input type="text" name="ChinaeBank1" size="30" value="<%=ChinaeBankPay(0)%>"></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">����֧��KEY����<a href="http://www.nps.cn/" target="_blank">NPS</a>����<a href="http://www.chinaebank.cn/" target="_blank">��������</a>��Կ��</div></td>
			<td class="TableRow1">
			<input type="password" name="ChinaeBank2" size="30" value="<%=ChinaeBankPay(1)%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">����֧�������ѣ� ��λ���ٷֱȣ�</div></td>
			<td class="TableRow2">
			<input type="text" name="ChinaeBank3" size="5" value="<%=ChinaeBankPay(2)%>"> ��&nbsp;&nbsp;
			<font color=blue>* ��������û�������Ӧ��������������0</font></td>
		</tr>
	</table></fieldset>
	<br>
	<fieldset style="cursor: default"><legend>&nbsp;ͼƬˮӡ����<a name="setting6"></a>[<a href="#top">����</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
	<tr>
			<td class="TableRow1"><u>ѡȡԤ��ͼƬ���</u></td>
			<td class="TableRow1"><select size="1" name="PreviewSetting" onChange="chkselect(options[selectedIndex].value,'know3');">
			<option value="999" <%If CInt(PreviewSetting(0)) = 999 Then Response.Write (" selected")%>>�ر�Ԥ��ͼƬ����</option>
			<option value="0" <%If CInt(PreviewSetting(0)) = 0 Then Response.Write (" selected")%>>CreatePreviewImage���</option>
			<option value="1" <%If CInt(PreviewSetting(0)) = 1 Then Response.Write (" selected")%>>AspJpeg���</option>
			<option value="2" <%If CInt(PreviewSetting(0)) = 2 Then Response.Write (" selected")%>>SA-ImgWriter���</option>
			</select><div id="know3"></div></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>����Ԥ��ͼƬ��С����ѡ��</u></td>
			<td class="TableRow2"><select size="1" name="PreviewSetting">
			<option value="0" <%If CInt(PreviewSetting(1)) = 0 Then Response.Write (" selected")%>>�̶���С</option>
			<option value="1" <%If CInt(PreviewSetting(1)) = 1 Then Response.Write (" selected")%>>�ȱ�����С</option>
			</select></td>
		</tr>
		<tr>
			<td class="TableRow1"><u>ͼƬˮӡ���ÿ���</u></td>
			<td class="TableRow1"><select size="1" name="PreviewSetting">
			<option value="0" <%If CInt(PreviewSetting(2)) = 0 Then Response.Write (" selected")%>>�ر�ˮӡЧ��</option>
			<option value="1" <%If CInt(PreviewSetting(2)) = 1 Then Response.Write (" selected")%>>ˮӡ����Ч��</option>
			<option value="2" <%If CInt(PreviewSetting(2)) = 2 Then Response.Write (" selected")%>>ˮӡͼƬЧ��</option>
			</select></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>����Ԥ��ͼƬ��С����(���|�߶�)</u><br></td>
			<td class="TableRow2">
			��ȣ�<input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(3))%>"> ���� |
			�߶ȣ�<input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(4))%>"> ����
			</td>
		</tr>
		<tr>
			<td class="TableRow1"><u>�ϴ�ͼƬ���ˮӡ������Ϣ����Ϊ�ջ�0��</u><br></td>
			<td class="TableRow1"><input type="text" name="PreviewSetting" size="30" value="<%=Trim(PreviewSetting(5))%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>�ϴ����ˮӡ�����С</u><br></td>
			<td class="TableRow2"><input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(6))%>"> <b>px<b></td>
		</tr>
		<tr>
			<td class="TableRow1"><u>�ϴ����ˮӡ������ɫ</u><br></td>
			<td class="TableRow1"><input type="text" name="PreviewSetting" id="PreviewSetting7" size="10" value="<%=Trim(PreviewSetting(7))%>">
			<img border=0 src="images/rect.gif" align="absMiddle" style="cursor:pointer;background-Color:<%=Trim(PreviewSetting(7))%>;" onclick="Getcolor(this,'PreviewSetting7');" title="ѡȡ��ɫ!"></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>�ϴ����ˮӡ��������</u></td>
			<td class="TableRow2"><select size="1" name="PreviewSetting">
			<option value="����" <%If Trim(PreviewSetting(8)) = "����" Then Response.Write (" selected")%>>����</option>
			<option value="����_GB2312" <%If Trim(PreviewSetting(8)) = "����_GB2312" Then Response.Write (" selected")%>>����</option>
			<option value="������" <%If Trim(PreviewSetting(8)) = "������" Then Response.Write (" selected")%>>������</option>
			<option value="����" <%If Trim(PreviewSetting(8)) = "����" Then Response.Write (" selected")%>>����</option>
			<option value="����" <%If Trim(PreviewSetting(8)) = "����" Then Response.Write (" selected")%>>����</option>
			<option value="Arial" <%If Trim(PreviewSetting(8)) = "Arial" Then Response.Write (" selected")%>>Arial</option>
			<option value="Georgia" <%If Trim(PreviewSetting(8)) = "Georgia" Then Response.Write (" selected")%>>Georgia</option>
			<option value="Impact" <%If Trim(PreviewSetting(8)) = "Impact" Then Response.Write (" selected")%>>Impact</option>
			<option value="Tahoma" <%If Trim(PreviewSetting(8)) = "Tahoma" Then Response.Write (" selected")%>>Tahoma</option>
			<option value="Stencil" <%If Trim(PreviewSetting(8)) = "Stencil" Then Response.Write (" selected")%>>Stencil</option>
			<option value="Verdana" <%If Trim(PreviewSetting(8)) = "Verdana" Then Response.Write (" selected")%>>Verdana</option>
			</select></td>
		</tr>
		<tr>
			<td class="TableRow1"><u>�ϴ�ˮӡ�����Ƿ����</u></td>
			<td class="TableRow1"><select size="1" name="PreviewSetting">
			<option value="0" <%If CInt(PreviewSetting(9)) = 0 Then Response.Write (" selected")%>>�Է��</option>
			<option value="1" <%If CInt(PreviewSetting(9)) = 1 Then Response.Write (" selected")%>>���ǡ�</option>
			</select></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>�ϴ�ͼƬ���ˮӡLOGOͼƬ��Ϣ</u><br>��дLOGO��ͼƬ���·��</td>
			<td class="TableRow2"><input type="text" name="PreviewSetting" size="30" value="<%=Trim(PreviewSetting(10))%>"></td>
		</tr>
		<tr>
			<td class="TableRow1"><u>�ϴ�ͼƬ���ˮӡ͸����</u><br></td>
			<td class="TableRow1"><input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(11))%>"> <font color=blue>��50%����д0.5</font></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>ˮӡͼƬȥ����ɫ</U><br>����Ϊ����ˮӡͼƬ��ȥ����ɫ</td>
			<td class="TableRow2"><input type="text" name="PreviewSetting" id="PreviewSetting12" size="10" value="<%=Trim(PreviewSetting(12))%>">
			<img border=0 src="images/rect.gif" align="absMiddle" style="cursor:pointer;background-Color:<%=Trim(PreviewSetting(12))%>;" onclick="Getcolor(this,'PreviewSetting12');" title="ѡȡ��ɫ!"></td>
		</tr>
		<tr>
			<td class="TableRow1"><u>ˮӡ���ֻ�ͼƬ�ĳ���������</u><br>��ˮӡͼƬ�Ŀ�Ⱥ͸߶�</td>
			<td class="TableRow1">
			��ȣ�<input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(13))%>"> ���� |
			�߶ȣ�<input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(14))%>"> ����
			</td>
		</tr>
		<tr>
			<td class="TableRow2"><u>�ϴ�ͼƬ���ˮӡLOGOλ������</u></td>
			<td class="TableRow2"><select size="1" name="PreviewSetting">
			<option value="0" <%If CInt(PreviewSetting(15)) = 0 Then Response.Write (" selected")%>>�����ϡ�</option>
			<option value="1" <%If CInt(PreviewSetting(15)) = 1 Then Response.Write (" selected")%>>�����¡�</option>
			<option value="2" <%If CInt(PreviewSetting(15)) = 2 Then Response.Write (" selected")%>>�Ծ��С�</option>
			<option value="3" <%If CInt(PreviewSetting(15)) = 3 Then Response.Write (" selected")%>>�����ϡ�</option>
			<option value="4" <%If CInt(PreviewSetting(15)) = 4 Then Response.Write (" selected")%>>�����¡�</option>
			</select></td>
		</tr>
	</table></fieldset>
	<br>
	<fieldset style="cursor: default"><legend>&nbsp;�����ַ�����<a name="setting5"></a>[<a href="#top">����</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
		<tr>
			<td class="TableRow1" width="35%"><div class="divbody">�����ַ���<br>
			���á�|���ֿ�</div></td>
			<td class="TableRow1"><textarea rows="5" name="Badwords" cols="60"><%=enchiasp.Badwords%></textarea></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">���˺���ַ���<br>
			���á�|���ֿ�</div></td>
			<td class="TableRow2"><textarea rows="5" name="Badwordr" cols="60"><%=enchiasp.Badwordr%></textarea></td>
		</tr>
	</table></fieldset>
	
	<fieldset style="cursor: default"><legend>&nbsp;����Ա��ȫ����<a name="setting6"></a>[<a href="#top">����</a>]</legend>
	<script language="JavaScript"> 
function pressKey(){
if(navigator.appName=='Netscape')
alert('�㰴���ı���ֵΪ��'+ event.which +
'\n��ASCII���ַ�Ϊ:'+ String.fromCharCode(event.which))
else
alert('�㰴���ı���ֵΪ��'+ event.keyCode +
'\n��ASCII���ַ�Ϊ:'+ String.fromCharCode(event.keyCode));}
</script>
	<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
<tr><td width=100 class="TableRow1">&nbsp;�������뿪�أ�</td><td class="TableRow1">
<input type="radio" name="ercilogin" value="0" <%if enchiasp.ercilogin =0 then %> checked <%end if%> >�ر�
<input type="radio" name="ercilogin" value="1" <%if enchiasp.ercilogin =1 then %> checked <%end if%>> ��<br><br>
</td></tr>
<tr><td width=100 class="TableRow1">&nbsp;����������룺</td><td class="TableRow1"><input type=text size="20" value='<%=enchiasp.mypass%>' name="mypass" class=ycenchi>&nbsp;����������������������<br>
</td></tr>
<tr><td width=100 class="TableRow1">&nbsp;�������뿪ֵ��</td><td class="TableRow1"><input type=text size="20" value='<%=enchiasp.mypasskey%>' name="mypasskey" class=ycenchi>&nbsp;ACCII����ֵ��<br>
</td></tr>
<tr><td width=100 class="TableRow1">&nbsp;��˵����</td><td style="line-height:150%" class="TableRow1">
������������˿��أ�������̨ʱ��Ҫ��֤�������룬�˶��������ɶ��λ������һ��������������ɣ���Ӧ������������û�����˵��й���Ա�������ã���ͬ����Ա�����趨��ͬ���������<br>
�ڶ��������ֹ�ڿͱ����ƽ�����Ĳ����� <br>
�۶������뿪ֵָACCII����ֵ,���ֵ����<a href="#" onKeyPress="pressKey()"><font color=red><b>����</b></font></a>�˴�,���н����ʱ�����������Ҫ����ӦACCII����ֵ,������������������������Ҫ�������ʱ���������õ�ACCII��Ӧ�ļ��̼��������Ż���ʾ���������
</td></tr>
</table>

		
		</fieldset>

	
	</td>
	</tr>
	<tr>
		<td class="TableRow1" align="center">
		<input type="submit" value="��������" name="B1" class=Button> <font color=red>��Ctrl+Enterֱ���ύ</font></td>
	</tr></form>
</table>
</div>
<%
	Dim InstalledObjects(10)
	Dim i
	Response.Write "<div id=""Issubport0"" style=""display:none"">��ѡ��EMAIL�����</div>" & vbCrLf
	Response.Write "<div id=""Issubport999"" style=""display:none"">��ѡ���ϴ������</div>" & vbCrLf
	
	InstalledObjects(1) = "JMail.Message"				'JMail 4.3
	InstalledObjects(2) = "CDONTS.NewMail"				'CDONTS
	InstalledObjects(3) = "Persits.MailSender"			'ASPEMAIL
	'-----------------------
	InstalledObjects(4) = "Adodb.Stream"				'Adodb.Stream
	InstalledObjects(5) = "Persits.Upload"				'Aspupload3.0
	InstalledObjects(6) = "SoftArtisans.FileUp"			'SA-FileUp 4.0
	InstalledObjects(7) = "DvFile.Upload"				'DvFile-Up V1.0
	'-----------------------
	InstalledObjects(8) = "CreatePreviewImage.cGvbox"		'CreatePreviewImage
	InstalledObjects(9) = "Persits.Jpeg"				'AspJpeg
	InstalledObjects(10) = "SoftArtisans.ImageGen"			'SoftArtisans ImgWriter V1.21
	For i = 1 To 10
		Response.Write "<div id=""Issubport" & i & """ style=""display:none"">"
		If Not IsObjInstalled(InstalledObjects(i)) Then
			Response.Write "<b>��</b>��������֧��!"
		Else 
			Response.Write "<font color=red><b>��</b>������֧��!</font>"
		End If
		Response.Write "</div>"
	Next
	Response.Write "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf
	Response.Write "<!--" & vbCrLf
	Response.Write "function chkselect(s,divid)" & vbCrLf
	Response.Write "{" & vbCrLf
	Response.Write "var divname='Issubport';" & vbCrLf
	Response.Write "var chkreport;" & vbCrLf
	Response.Write "s=Number(s)" & vbCrLf
	Response.Write "if (divid==""know1"")" & vbCrLf
	Response.Write "{" & vbCrLf
	Response.Write "divname=divname+s;" & vbCrLf
	Response.Write "}" & vbCrLf
	Response.Write "if (divid==""know2"")" & vbCrLf
	Response.Write "{" & vbCrLf
	Response.Write "s+=4;" & vbCrLf
	Response.Write "if (s==1003){s=999;}" & vbCrLf
	Response.Write "divname=divname+s;" & vbCrLf
	Response.Write "}" & vbCrLf
	Response.Write "if (divid==""know3"")" & vbCrLf
	Response.Write "{" & vbCrLf
	Response.Write "s+=8;" & vbCrLf
	Response.Write "if (s==1007){s=999;}" & vbCrLf
	Response.Write "divname=divname+s;" & vbCrLf
	Response.Write "}" & vbCrLf
	Response.Write "document.getElementById(divid).innerHTML=divname;" & vbCrLf
	Response.Write "chkreport=document.getElementById(divname).innerHTML;" & vbCrLf
	Response.Write "document.getElementById(divid).innerHTML=chkreport;" & vbCrLf
	Response.Write "}" & vbCrLf
	Response.Write "//-->"
	Response.Write "</SCRIPT>" & vbCrLf
End Sub

Sub SaveConfig()
	Dim strChinaeBank
	If Len(Request.Form("SiteName")) = 0 Or Len(Request.Form("SiteName")) => 50 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��վ���Ʋ���Ϊ�ջ��߳���50���ַ���</li>"
	End If
	If Len(Request.Form("keywords")) = 0 Or Len(Request.Form("keywords")) => 250 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>վ��ؼ��ֲ���Ϊ�ջ��߳���250���ַ���</li>"
	End If
	If Len(Request.Form("StopReadme")) = 0 Or Len(Request.Form("StopReadme")) => 250 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ά��˵������Ϊ�ջ��߳���250���ַ���</li>"
	End If
	If Len(Request.Form("Copyright")) = 0 Or Len(Request.Form("Copyright")) => 250 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��վ��Ȩ��Ϣ����Ϊ�ջ��߳���250���ַ���</li>"
	End If
	If Len(Request.Form("IndexName")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ҳ�ļ�������Ϊ�գ�</li>"
	End If
	If Not enchiasp.IsValidChar(Request.Form("IndexName")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ҳ�ļ����к��зǷ��ַ����������ַ���</li>"
	End If
	If Trim(Request.Form("StopApplyLink")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>�Ƿ�ر��������ӹ��ܣ�</li>"
	End If
	If Len(Request.Form("InitTitleColor")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��ʼ��������ɫ����Ϊ�գ�</li>"
	End If
	If Len(Request.Form("InstallDir")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>ϵͳ���ڸ�Ŀ¼����Ϊ�գ�</li>"
	End If
	If Len(Request.Form("ChinaeBank4")) > 20 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>����֧���������벻�ܴ���20���ַ���</li>"
	End If
	if Trim(Request.Form("ercilogin")) = "1"  then
	If Trim(Request.Form("mypass")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������ֵ����Ϊ�գ�</li>"
	End If

	If Trim(Request.Form("mypasskey")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>��������kaiֵ����Ϊ�գ�</li>"
	End If
	end if

	strChinaeBank = Trim(Request.Form("ChinaeBank1")) & "|||" & Trim(Request.Form("ChinaeBank2")) & "|||" & Trim(Request.Form("ChinaeBank3")) & "|||" & Trim(Request.Form("ChinaeBank4")) & "|||"
	
	If FoundErr = True Then Exit Sub
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from ECCMS_Config where id = 1"
	Rs.Open SQL,Conn,1,3
		Rs("SiteName") = Trim(Request.Form("SiteName"))
		Rs("SiteUrl") = Trim(Request.Form("SiteUrl"))
		Rs("MasterMail") = Trim(Request.Form("MasterMail"))
		Rs("keywords") = Trim(Request.Form("keywords"))
		Rs("Copyright") = Trim(Request.Form("Copyright"))
		Rs("InstallDir") = Trim(Request.Form("InstallDir"))
		Rs("IndexName") = Trim(Request.Form("IndexName"))
		Rs("IstopSite") = Trim(Request.Form("IstopSite"))
		Rs("StopReadme") = Trim(Request.Form("StopReadme"))
		Rs("IsCloseMail") = Trim(Request.Form("IsCloseMail"))
		Rs("SendMailType") = Trim(Request.Form("SendMailType"))
		Rs("MailFrom") = Trim(Request.Form("MailFrom"))
		Rs("MailServer") = Trim(Request.Form("MailServer"))
		Rs("MailUserName") = Trim(Request.Form("MailUserName"))
		Rs("MailPassword") = Trim(Request.Form("MailPassword"))
		Rs("CheckUserReg") = Trim(Request.Form("CheckUserReg"))
		Rs("AdminCheckReg") = Trim(Request.Form("AdminCheckReg"))
		Rs("MailInformPass") = Trim(Request.Form("MailInformPass"))
		Rs("ChkSameMail") = Trim(Request.Form("ChkSameMail"))
		Rs("AddUserPoint") = Trim(Request.Form("AddUserPoint"))
		Rs("SendRegMessage") = Trim(Request.Form("SendRegMessage"))
		Rs("FullContQuery") = Trim(Request.Form("FullContQuery"))
		Rs("ActionTime") = Trim(Request.Form("ActionTime"))
		Rs("IsRunTime") = Trim(Request.Form("IsRunTime"))
		Rs("UploadClass") = Trim(Request.Form("UploadClass"))
		Rs("UploadFileSize") = Trim(Request.Form("UploadFileSize"))
		Rs("UploadFileType") = Trim(Request.Form("UploadFileType"))
		Rs("StopApplyLink") = Trim(Request.Form("StopApplyLink"))
		Rs("FSO_ScriptName") = Trim(Request.Form("FSO_ScriptName"))
		Rs("InitTitleColor") = Trim(Request.Form("InitTitleColor"))
		Rs("StopBankPay") = Trim(Request.Form("StopBankPay"))
		Rs("ChinaeBank") = strChinaeBank
		Rs("Badwords") = Trim(Request.Form("Badwords"))
		Rs("Badwordr") = Trim(Request.Form("Badwordr"))
		Rs("PreviewSetting") = Replace(Replace(Request.Form("PreviewSetting"), " ", ""), "'", "")
		

		Rs("ercilogin") = Trim(Request.Form("ercilogin"))
		if Trim(Request.Form("ercilogin")) = "1"  then
			Rs("mypass") = Trim(Request.Form("mypass"))
			Rs("mypasskey") = Trim(Request.Form("mypasskey"))
		end if

	Rs.update
	Rs.close:set Rs = Nothing
	enchiasp.DelCahe("Config")
	enchiasp.DelCache("MyConfig")
	Succeed("<li>��ϲ�����������óɹ���</li>")
End Sub

Sub ReloadCache()
	Application.Contents.RemoveAll
	Response.Write "<script>alert('�ؽ�����ɹ���');javascript:history.back(1)</script>"
End Sub

Sub EditContentKeyword()
	Dim i,ContentKeywordStr,KeywordStr
	Set Rs = enchiasp.Execute("Select id,ContentKeyword From ECCMS_Config where id = 1")
	ContentKeywordStr = Split(Rs("ContentKeyword"), "@@@")
	
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan=2 align="left"><img src="images/welcome.gif" width="16" height="17" align="absMiddle"> �������ݹؼ�������</th>
</tr>
<form name=myform method=post action=?action=savedit>
<input type=hidden name=id value=''>
<tr>
	<td class=TableTitle align=center>���ݹؼ���</td>
	<td class=TableTitle align=center>�ؼ�������URL</td>
</tr>
<%
If Trim(Rs("ContentKeyword")) <> "" Then
	
	For i = 0 To UBound(ContentKeywordStr)
		
		If ContentKeywordStr(i) <> "" Then
			KeywordStr = Split(ContentKeywordStr(i), "$$$")
		Else
			KeywordStr(0) = "del"
			KeywordStr(1) = "del"
		End If
		
		
%>
<tr>
	<td class=tablerow1><input type=text name=KeywordStr size=45 value='<%=KeywordStr(0)%>'></td>
	<td class=tablerow2><input type=text name=KeywordUrl size=45 value='<%=KeywordStr(1)%>'></td>
</tr>
<%
	Next
Else
%>
<tr>
	<td class=tablerow1><input type=text name=KeywordStr size=45 value='del'></td>
	<td class=tablerow2><input type=text name=KeywordUrl size=45 value='del'></td>
</tr>
<%
End If
%>
<tr>
	<td class=tablerow2 colspan=2 align=center><input type=submit value='��������' class=Button></td>
</tr>
<tr>
	<td class=tablerow1 colspan=2><strong>˵��:</strong><br>&nbsp;&nbsp;�������������������ҵ��ʵ��Ĺؼ���, ��Ϊ�����ú��ʵ�������.
��: �����˹ؼ��� �������, ָ���ַ http://www.enchi.com.cn  ��ô�������г���"�������"��Щ��ʱ, ҳ�������Զ�Ϊ��Щ�ּ���ָ��http://www.enchi.com.cn�����ַ������.   ע��: ���������ظ��Ĺؼ���, Ҳ�������ð�����ͬ���ֵĹؼ���,���������ùؼ��֣����ڶ�Ӧ���ı���������"del"��
���Ҫ���ö���ؼ�����һ���������á�|���ֿ���</td>
</tr>
</form>
</table>

<%
	Rs.Close:Set Rs = Nothing
End Sub

Sub SaveContentKeyword()
	Dim TempKeywordStr,TempKeywordUrl
	Dim i
	Dim ContentKeyword,ContentKeywordStr
	TempKeywordStr = Request.Form("KeywordStr")
	TempKeywordUrl = Request.Form("KeywordUrl")
	TempKeywordStr = Split(TempKeywordStr, ",")
	TempKeywordUrl = Split(TempKeywordUrl, ",")
	For i = 0 To UBound(TempKeywordStr)
		If Trim(TempKeywordStr(i)) = "" Or Trim(TempKeywordUrl(i)) = "" Then
			FoundErr = True
			ErrMsg = ErrMsg + "<li>���ݹؼ��ֺ͹ؼ�������URL����Ϊ�գ�</li>"
			Exit Sub
		End If
		If LCase(Trim(TempKeywordStr(i))) <> "del" And LCase(Trim(TempKeywordUrl(i))) <> "del" Then
			ContentKeyword = ContentKeyword & Trim(TempKeywordStr(i)) & "$$$" & Trim(TempKeywordUrl(i)) & "@@@"
		End If
	Next
	ContentKeywordStr = enchiasp.CheckStr(ContentKeyword)
	enchiasp.Execute ("update [ECCMS_Config] set ContentKeyword ='" & ContentKeywordStr & "' where id = 1")
	enchiasp.DelCahe("Config")
	enchiasp.DelCache("MyConfig")
	OutHintScript("��ϲ�����������ݹؼ��ֳɹ���")
End Sub
%>