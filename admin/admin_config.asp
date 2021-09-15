<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header 
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
Response.Write "<table border=0 align=center cellspacing=1 class=TableBorder cellpadding=3>"
Response.Write "	<tr>"
Response.Write "		<th>网站基本配置管理</th>"
Response.Write "	</tr>"
Response.Write "	<tr>"
Response.Write "		<td class=TableRow><a href='admin_config.asp#setting1'>系统基本信息</a> |"
Response.Write "		<a href='admin_config.asp#setting2'>系统邮件设置</a> | "
Response.Write "		<a href='admin_config.asp#setting3'>注册用户设置</a> | "
Response.Write "		<a href='admin_config.asp#setting4'>系统基本设置</a> |"
Response.Write "		<a href='admin_config.asp#setting5'>过滤字符设置</a> |"
Response.Write "		<a href='admin_config.asp#setting6'>管理员安全设置</a> |"
Response.Write "		<a href='?action=edit'><font color=blue>内容关键字设置</font></a> |"
Response.Write "		<a href='?action=reload'><font color=red>重建缓存</font></a></td>"
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
		<th align="left"><img src="images/welcome.gif" width="16" height="17" align="absMiddle"> 网站基本设置</th>
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
    alert("在 请输入整数 域中，只能输入 数字 字符。");
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
    alert("在 请输入整数 域中，只能输入 数字 字符。");
    theForm.ActionTime.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form name=FrontPage_Form1 method="POST" action="?action=save" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
	<tr>
	<td class="TableRow1">
	<fieldset style="cursor: default"><legend>&nbsp;网站基本信息<a name="setting2"></a>[<a href="#top">顶部</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
		
		<tr>
			<td class="TableRow1" width="35%"><div class="divbody">网站名称：</div></td>
			<td class="TableRow1" width="65%">
			<input type="text" name="SiteName" size="35" value="<%=enchiasp.SiteName%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">网站URL：</div></td>
			<td class="TableRow2"><input type="text" name="SiteUrl" size="35" value="<%=RootPath2DomainPath("")%>">
			<font color="#FF0000">系统自动获取，请不要修改</font></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">管理员Email：</div></td>
			<td class="TableRow1">
			<input type="text" name="MasterMail" size="25" value="<%=enchiasp.MasterMail%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">首页文件名：</div></td>
			<td class="TableRow2"><input type="text" name="IndexName" size="25" value="<%=enchiasp.IndexName%>"></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">系统所在根目录：</div></td>
			<td class="TableRow1">
			<input type="text" size="25" value="<%=enchiasp.InstallDir%>" disabled>
			<input type=hidden name="InstallDir" value="<%=SiteRootDir%>">&nbsp;* 
			<font color="#FF0000">系统自动获取，无需手动输入</font></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">站点关键字：<br>
			将被搜索引擎用来搜索您网站的关键内容</div></td>
			<td class="TableRow2"><textarea rows="3" name="keywords" cols="60"><%=enchiasp.keywords%></textarea></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">网站版权信息：</div></td>
			<td class="TableRow1" width="65%">
			<textarea rows="5" name="Copyright" cols="60"><%=enchiasp.Copyright%></textarea></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">是否关闭网站：</div></td>
			<td class="TableRow2">
			<input type="radio" name="IstopSite" value="1" <%If CInt(enchiasp.IstopSite) = 1 Then Response.Write (" checked")%>> 关闭 
			<input type="radio" name="IstopSite" value="0" <%If CInt(enchiasp.IstopSite) = 0 Then Response.Write (" checked")%>> 打开</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">网站维护说明：<br>支持HTML方法,不能超过250个字符</div></td>
			<td class="TableRow1">
			<textarea rows="5" name="StopReadme" cols="60"><%=enchiasp.StopReadme%></textarea></td>
		</tr>
	</table></fieldset>
	<br>
	<fieldset style="cursor: default"><legend>&nbsp;系统邮件设置<a name="setting2"></a>[<a href="#top">顶部</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
		<tr>
			<td class="TableRow2" width="35%"><div class="divbody">是否关闭邮件功能：</div></td>
			<td class="TableRow2">
			<input type="radio" name="IsCloseMail" value="1" <%If CInt(enchiasp.IsCloseMail) = 1 Then Response.Write (" checked")%>> 关闭 
			<input type="radio" name="IsCloseMail" value="0" <%If CInt(enchiasp.IsCloseMail) = 0 Then Response.Write (" checked")%>> 打开</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">发送邮件的组件：</div></td>
			<td class="TableRow1"><select size="1" name="SendMailType"  onChange="chkselect(options[selectedIndex].value,'know1');">
			<option value="0" <%If CInt(enchiasp.SendMailType) = 0 Then Response.Write (" selected")%>>不支持</option>
			<option value="1" <%If CInt(enchiasp.SendMailType) = 1 Then Response.Write (" selected")%>>JMAIL</option>
			<option value="2" <%If CInt(enchiasp.SendMailType) = 2 Then Response.Write (" selected")%>>CDONTS</option>
			<option value="3" <%If CInt(enchiasp.SendMailType) = 3 Then Response.Write (" selected")%>>ASPEMAIL</option>
			</select><div id=know1></div></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">系统管理员Email：<br>
			给用户发送邮件时，显示的来源Email信息</div></td>
			<td class="TableRow2">
			<input type="text" name="MailFrom" size="25" value="<%=enchiasp.MailFrom%>"></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">SMTP Server地址：</div></td>
			<td class="TableRow1">
			<input type="text" name="MailServer" size="25" value="<%=enchiasp.MailServer%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">邮件登录用户名：</div></td>
			<td class="TableRow2">
			<input type="text" name="MailUserName" size="25" value="<%=enchiasp.MailUserName%>"></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">邮件登录密码：</div></td>
			<td class="TableRow1">
			<input type="password" name="MailPassword" size="25" value="<%=enchiasp.MailPassword%>"></td>
		</tr>
	</table></fieldset>
	<br>
	<fieldset style="cursor: default"><legend>&nbsp;注册用户设置<a name="setting3"></a>[<a href="#top">顶部</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
		<tr>
			<td class="TableRow2" width="35%"><div class="divbody">是否允许新用户注册：</div></td>
			<td class="TableRow2">
			<input type="radio" name="CheckUserReg" value="0" <%If CInt(enchiasp.CheckUserReg) = 0 Then Response.Write (" checked")%>> 否 
			<input type="radio" name="CheckUserReg" value="1" <%If CInt(enchiasp.CheckUserReg) = 1 Then Response.Write (" checked")%>> 是</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">注册会员是否要管理员认证：</div></td>
			<td class="TableRow1">
			<input type="radio" name="AdminCheckReg" value="0" <%If CInt(enchiasp.AdminCheckReg) = 0 Then Response.Write (" checked")%>> 否 
			<input type="radio" name="AdminCheckReg" value="1" <%If CInt(enchiasp.AdminCheckReg) = 1 Then Response.Write (" checked")%>> 是</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">注册会员是否Email通知密码：<br>确认您的站点支持发送mail，所包含密码为系统随机生成</div></td>
			<td class="TableRow2">
			<input type="radio" name="MailInformPass" value="0" <%If CInt(enchiasp.MailInformPass) = 0 Then Response.Write (" checked")%>> 关闭 
			<input type="radio" name="MailInformPass" value="1" <%If CInt(enchiasp.MailInformPass) = 1 Then Response.Write (" checked")%>> 打开</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">一个Email只能注册一个会员：</div></td>
			<td class="TableRow1">
			<input type="radio" name="ChkSameMail" value="0" <%If CInt(enchiasp.ChkSameMail) = 0 Then Response.Write (" checked")%>> 关闭 
			<input type="radio" name="ChkSameMail" value="1" <%If CInt(enchiasp.ChkSameMail) = 1 Then Response.Write (" checked")%>> 打开</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">发送注册邮件信息：<br>请确认您打开了邮件功能</div></td>
			<td class="TableRow2">
			<input type="radio" name="SendRegMessage" value="0" <%If CInt(enchiasp.SendRegMessage) = 0 Then Response.Write (" checked")%>> 关闭 
			<input type="radio" name="SendRegMessage" value="1" <%If CInt(enchiasp.SendRegMessage) = 1 Then Response.Write (" checked")%>> 打开</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">注册会员增加的点数：<br>
			请输入整数</div></td>
			<td class="TableRow1">
			&nbsp;<!--webbot bot="Validation" s-display-name="请输入整数" s-data-type="Integer" s-number-separators="," --><input type="text" name="AddUserPoint" size="15" value="<%=enchiasp.AddUserPoint%>"></td>
		</tr>
	</table></fieldset>
	<br>
	<fieldset style="cursor: default"><legend>&nbsp;系统基本设置<a name="setting4"></a>[<a href="#top">顶部</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
		<tr>
			<td class="TableRow2" width="35%"><div class="divbody">是否开启全文搜索：<br>
			全文搜索占用服务器资源不建议开启</div></td>
			<td class="TableRow2">
			<input type="radio" name="FullContQuery" value="0" <%If CInt(enchiasp.FullContQuery) = 0 Then Response.Write (" checked")%>> 否 
			<input type="radio" name="FullContQuery" value="1" <%If CInt(enchiasp.FullContQuery) = 1 Then Response.Write (" checked")%>> 是</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">删除不活动用户时间：<br>
			单位：分钟，请输入数字</div></td>
			<td class="TableRow1">
			&nbsp;<!--webbot bot="Validation" s-display-name="请输入整数" s-data-type="Integer" s-number-separators="," --><input type="text" name="ActionTime" size="15" value="<%=enchiasp.ActionTime%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">是否显示页面执行时间：</div></td>
			<td class="TableRow2">
			<input type="radio" name="IsRunTime" value="0" <%If CInt(enchiasp.IsRunTime) = 0 Then Response.Write (" checked")%>> 否 
			<input type="radio" name="IsRunTime" value="1" <%If CInt(enchiasp.IsRunTime) = 1 Then Response.Write (" checked")%>> 是</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">选取上传组件：</div></td>
			<td class="TableRow1"><select size="1" name="UploadClass" onChange="chkselect(options[selectedIndex].value,'know2');">
			<option value="999" <%If CInt(enchiasp.UploadClass) = 999 Then Response.Write (" selected")%>>关闭上传功能</option>
			<option value="0" <%If CInt(enchiasp.UploadClass) = 0 Then Response.Write (" selected")%>>无组件上传类</option>
			<option value="1" <%If CInt(enchiasp.UploadClass) = 1 Then Response.Write (" selected")%>>Aspupload3.0组件</option>
			<option value="2" <%If CInt(enchiasp.UploadClass) = 2 Then Response.Write (" selected")%>>SA-FileUp 4.0组件</option>
			</select><div id="know2"></div></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">上传文件大小： 单位(KB)</div></td>
			<td class="TableRow2">
			<input type="text" name="UploadFileSize" size="15" value="<%=enchiasp.UploadFileSize%>"> <b>KB</></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">上传文件类型： 请用“|”分开</div></td>
			<td class="TableRow1">
			<input type="text" name="UploadFileType" size="60" value="<%=enchiasp.UploadFileType%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">是否关闭友情连接申请：</div></td>
			<td class="TableRow2">
			<input type="radio" name="StopApplyLink" value="1" <%If CInt(enchiasp.StopApplyLink) = 1 Then Response.Write (" checked")%>> 关闭 
			<input type="radio" name="StopApplyLink" value="0" <%If CInt(enchiasp.StopApplyLink) = 0 Then Response.Write (" checked")%>> 打开</td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">FSO组件的名称：<br>
			FSO默认名称：Scripting.FileSystemObject</div></td>
			<td class="TableRow1">
			<input type="text" name="FSO_ScriptName" size="35" value="<%=enchiasp.FSO_ScriptName%>"><br>
			某些网站为了安全，将FSO组件的名称进行更改以达到禁用FSO的目的。如果你的网站是这样做的，请在此输入更改过的名称。</div></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">初始标题颜色：</div><br>
			每个颜色的值请的“,”分开</td>
			<td class="TableRow2"><textarea rows="3" name="InitTitleColor" cols="60"><%=enchiasp.InitTitleColor%></textarea></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">请选择在线支付功能：</div></td>
			<td class="TableRow1">
			<input type="radio" name="StopBankPay" value="0" <%If CInt(enchiasp.StopBankPay) = 0 Then Response.Write (" checked")%>> 关闭在线支付&nbsp;&nbsp;
			<input type="radio" name="StopBankPay" value="1" <%If CInt(enchiasp.StopBankPay) = 1 Then Response.Write (" checked")%>> NPS 在线支付&nbsp;&nbsp;
			<input type="radio" name="StopBankPay" value="2" <%If CInt(enchiasp.StopBankPay) = 2 Then Response.Write (" checked")%>> 网银在线支付&nbsp;&nbsp;
			<br><font color=blue>* 如果你还没开通NPS在线支付，请<a href="http://www.nps.cn/merchant/join_agreement.jsp" target="_blank"><font color=RED><strong>点此开通NPS在线支付</strong></font></a>吧</font>
			<br><font color=blue>* 如果你还没开通网银在线支付，请<a href="http://www.chinaebank.cn/" target="_blank"><font color=RED><strong>点此开通中国网银在线支付</strong></font></a>吧</font>
			</td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">在线支付 ID：（<a href="http://www.nps.cn/" target="_blank">NPS</a>或者<a href="http://www.chinaebank.cn/" target="_blank">网银在线</a>ID）</div></td>
			<td class="TableRow2">
			<input type="text" name="ChinaeBank1" size="30" value="<%=ChinaeBankPay(0)%>"></td>
		</tr>
		<tr>
			<td class="TableRow1"><div class="divbody">在线支付KEY：（<a href="http://www.nps.cn/" target="_blank">NPS</a>或者<a href="http://www.chinaebank.cn/" target="_blank">网银在线</a>密钥）</div></td>
			<td class="TableRow1">
			<input type="password" name="ChinaeBank2" size="30" value="<%=ChinaeBankPay(1)%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">在线支付手续费： 单位（百分比）</div></td>
			<td class="TableRow2">
			<input type="text" name="ChinaeBank3" size="5" value="<%=ChinaeBankPay(2)%>"> ％&nbsp;&nbsp;
			<font color=blue>* 如果不对用户增加相应的手续费请输入0</font></td>
		</tr>
	</table></fieldset>
	<br>
	<fieldset style="cursor: default"><legend>&nbsp;图片水印设置<a name="setting6"></a>[<a href="#top">顶部</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
	<tr>
			<td class="TableRow1"><u>选取预览图片组件</u></td>
			<td class="TableRow1"><select size="1" name="PreviewSetting" onChange="chkselect(options[selectedIndex].value,'know3');">
			<option value="999" <%If CInt(PreviewSetting(0)) = 999 Then Response.Write (" selected")%>>关闭预览图片功能</option>
			<option value="0" <%If CInt(PreviewSetting(0)) = 0 Then Response.Write (" selected")%>>CreatePreviewImage组件</option>
			<option value="1" <%If CInt(PreviewSetting(0)) = 1 Then Response.Write (" selected")%>>AspJpeg组件</option>
			<option value="2" <%If CInt(PreviewSetting(0)) = 2 Then Response.Write (" selected")%>>SA-ImgWriter组件</option>
			</select><div id="know3"></div></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>生成预览图片大小规则选项</u></td>
			<td class="TableRow2"><select size="1" name="PreviewSetting">
			<option value="0" <%If CInt(PreviewSetting(1)) = 0 Then Response.Write (" selected")%>>固定大小</option>
			<option value="1" <%If CInt(PreviewSetting(1)) = 1 Then Response.Write (" selected")%>>等比例缩小</option>
			</select></td>
		</tr>
		<tr>
			<td class="TableRow1"><u>图片水印设置开关</u></td>
			<td class="TableRow1"><select size="1" name="PreviewSetting">
			<option value="0" <%If CInt(PreviewSetting(2)) = 0 Then Response.Write (" selected")%>>关闭水印效果</option>
			<option value="1" <%If CInt(PreviewSetting(2)) = 1 Then Response.Write (" selected")%>>水印文字效果</option>
			<option value="2" <%If CInt(PreviewSetting(2)) = 2 Then Response.Write (" selected")%>>水印图片效果</option>
			</select></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>生成预览图片大小设置(宽度|高度)</u><br></td>
			<td class="TableRow2">
			宽度：<input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(3))%>"> 象素 |
			高度：<input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(4))%>"> 象素
			</td>
		</tr>
		<tr>
			<td class="TableRow1"><u>上传图片添加水印文字信息（可为空或0）</u><br></td>
			<td class="TableRow1"><input type="text" name="PreviewSetting" size="30" value="<%=Trim(PreviewSetting(5))%>"></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>上传添加水印字体大小</u><br></td>
			<td class="TableRow2"><input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(6))%>"> <b>px<b></td>
		</tr>
		<tr>
			<td class="TableRow1"><u>上传添加水印字体颜色</u><br></td>
			<td class="TableRow1"><input type="text" name="PreviewSetting" id="PreviewSetting7" size="10" value="<%=Trim(PreviewSetting(7))%>">
			<img border=0 src="images/rect.gif" align="absMiddle" style="cursor:pointer;background-Color:<%=Trim(PreviewSetting(7))%>;" onclick="Getcolor(this,'PreviewSetting7');" title="选取颜色!"></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>上传添加水印所用字体</u></td>
			<td class="TableRow2"><select size="1" name="PreviewSetting">
			<option value="宋体" <%If Trim(PreviewSetting(8)) = "宋体" Then Response.Write (" selected")%>>宋体</option>
			<option value="楷体_GB2312" <%If Trim(PreviewSetting(8)) = "楷体_GB2312" Then Response.Write (" selected")%>>楷体</option>
			<option value="新宋体" <%If Trim(PreviewSetting(8)) = "新宋体" Then Response.Write (" selected")%>>新宋体</option>
			<option value="黑体" <%If Trim(PreviewSetting(8)) = "黑体" Then Response.Write (" selected")%>>黑体</option>
			<option value="隶书" <%If Trim(PreviewSetting(8)) = "隶书" Then Response.Write (" selected")%>>隶书</option>
			<option value="Arial" <%If Trim(PreviewSetting(8)) = "Arial" Then Response.Write (" selected")%>>Arial</option>
			<option value="Georgia" <%If Trim(PreviewSetting(8)) = "Georgia" Then Response.Write (" selected")%>>Georgia</option>
			<option value="Impact" <%If Trim(PreviewSetting(8)) = "Impact" Then Response.Write (" selected")%>>Impact</option>
			<option value="Tahoma" <%If Trim(PreviewSetting(8)) = "Tahoma" Then Response.Write (" selected")%>>Tahoma</option>
			<option value="Stencil" <%If Trim(PreviewSetting(8)) = "Stencil" Then Response.Write (" selected")%>>Stencil</option>
			<option value="Verdana" <%If Trim(PreviewSetting(8)) = "Verdana" Then Response.Write (" selected")%>>Verdana</option>
			</select></td>
		</tr>
		<tr>
			<td class="TableRow1"><u>上传水印字体是否粗体</u></td>
			<td class="TableRow1"><select size="1" name="PreviewSetting">
			<option value="0" <%If CInt(PreviewSetting(9)) = 0 Then Response.Write (" selected")%>>≡否≡</option>
			<option value="1" <%If CInt(PreviewSetting(9)) = 1 Then Response.Write (" selected")%>>≡是≡</option>
			</select></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>上传图片添加水印LOGO图片信息</u><br>填写LOGO的图片相对路径</td>
			<td class="TableRow2"><input type="text" name="PreviewSetting" size="30" value="<%=Trim(PreviewSetting(10))%>"></td>
		</tr>
		<tr>
			<td class="TableRow1"><u>上传图片添加水印透明度</u><br></td>
			<td class="TableRow1"><input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(11))%>"> <font color=blue>如50%请填写0.5</font></td>
		</tr>
		<tr>
			<td class="TableRow2"><u>水印图片去除底色</U><br>保留为空则水印图片不去除底色</td>
			<td class="TableRow2"><input type="text" name="PreviewSetting" id="PreviewSetting12" size="10" value="<%=Trim(PreviewSetting(12))%>">
			<img border=0 src="images/rect.gif" align="absMiddle" style="cursor:pointer;background-Color:<%=Trim(PreviewSetting(12))%>;" onclick="Getcolor(this,'PreviewSetting12');" title="选取颜色!"></td>
		</tr>
		<tr>
			<td class="TableRow1"><u>水印文字或图片的长宽区域定义</u><br>如水印图片的宽度和高度</td>
			<td class="TableRow1">
			宽度：<input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(13))%>"> 象素 |
			高度：<input type="text" name="PreviewSetting" size="10" value="<%=Trim(PreviewSetting(14))%>"> 象素
			</td>
		</tr>
		<tr>
			<td class="TableRow2"><u>上传图片添加水印LOGO位置坐标</u></td>
			<td class="TableRow2"><select size="1" name="PreviewSetting">
			<option value="0" <%If CInt(PreviewSetting(15)) = 0 Then Response.Write (" selected")%>>≡左上≡</option>
			<option value="1" <%If CInt(PreviewSetting(15)) = 1 Then Response.Write (" selected")%>>≡左下≡</option>
			<option value="2" <%If CInt(PreviewSetting(15)) = 2 Then Response.Write (" selected")%>>≡居中≡</option>
			<option value="3" <%If CInt(PreviewSetting(15)) = 3 Then Response.Write (" selected")%>>≡右上≡</option>
			<option value="4" <%If CInt(PreviewSetting(15)) = 4 Then Response.Write (" selected")%>>≡右下≡</option>
			</select></td>
		</tr>
	</table></fieldset>
	<br>
	<fieldset style="cursor: default"><legend>&nbsp;过滤字符设置<a name="setting5"></a>[<a href="#top">顶部</a>]</legend><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
		<tr>
			<td class="TableRow1" width="35%"><div class="divbody">过滤字符：<br>
			请用“|”分开</div></td>
			<td class="TableRow1"><textarea rows="5" name="Badwords" cols="60"><%=enchiasp.Badwords%></textarea></td>
		</tr>
		<tr>
			<td class="TableRow2"><div class="divbody">过滤后的字符：<br>
			请用“|”分开</div></td>
			<td class="TableRow2"><textarea rows="5" name="Badwordr" cols="60"><%=enchiasp.Badwordr%></textarea></td>
		</tr>
	</table></fieldset>
	
	<fieldset style="cursor: default"><legend>&nbsp;管理员安全设置<a name="setting6"></a>[<a href="#top">顶部</a>]</legend>
	<script language="JavaScript"> 
function pressKey(){
if(navigator.appName=='Netscape')
alert('你按键的编码值为：'+ event.which +
'\n其ASCII码字符为:'+ String.fromCharCode(event.which))
else
alert('你按键的编码值为：'+ event.keyCode +
'\n其ASCII码字符为:'+ String.fromCharCode(event.keyCode));}
</script>
	<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="TableBorder2">
<tr><td width=100 class="TableRow1">&nbsp;二次密码开关：</td><td class="TableRow1">
<input type="radio" name="ercilogin" value="0" <%if enchiasp.ercilogin =0 then %> checked <%end if%> >关闭
<input type="radio" name="ercilogin" value="1" <%if enchiasp.ercilogin =1 then %> checked <%end if%>> 打开<br><br>
</td></tr>
<tr><td width=100 class="TableRow1">&nbsp;二次密码基码：</td><td class="TableRow1"><input type=text size="20" value='<%=enchiasp.mypass%>' name="mypass" class=ycenchi>&nbsp;您可以随意输入您的密码<br>
</td></tr>
<tr><td width=100 class="TableRow1">&nbsp;二次密码开值：</td><td class="TableRow1"><input type=text size="20" value='<%=enchiasp.mypasskey%>' name="mypasskey" class=ycenchi>&nbsp;ACCII编码值。<br>
</td></tr>
<tr><td width=100 class="TableRow1">&nbsp;简单说明：</td><td style="line-height:150%" class="TableRow1">
①如果您开启了开关，则进入后台时需要验证二次密码，此二次密码由二次基码根据一定的密码规则生成，相应的密码规则在用户管理菜单中管理员进行设置，不同管理员可以设定不同的密码规则<br>
②二次密码防止黑客暴力破解密码的补助器 <br>
③二次密码开值指ACCII编码值,相关值，请<a href="#" onKeyPress="pressKey()"><font color=red><b>单击</b></font></a>此处,在有焦点的时候请键入你想要的相应ACCII键盘值,如果您开启了这个功能您必须要进入管理时按下您设置的ACCII相应的键盘键。这样才会显示出来输入框。
</td></tr>
</table>

		
		</fieldset>

	
	</td>
	</tr>
	<tr>
		<td class="TableRow1" align="center">
		<input type="submit" value="保存设置" name="B1" class=Button> <font color=red>按Ctrl+Enter直接提交</font></td>
	</tr></form>
</table>
</div>
<%
	Dim InstalledObjects(10)
	Dim i
	Response.Write "<div id=""Issubport0"" style=""display:none"">请选择EMAIL组件！</div>" & vbCrLf
	Response.Write "<div id=""Issubport999"" style=""display:none"">请选择上传组件！</div>" & vbCrLf
	
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
			Response.Write "<b>×</b>服务器不支持!"
		Else 
			Response.Write "<font color=red><b>√</b>服务器支持!</font>"
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
		ErrMsg = ErrMsg + "<li>网站名称不能为空或者超过50个字符！</li>"
	End If
	If Len(Request.Form("keywords")) = 0 Or Len(Request.Form("keywords")) => 250 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>站点关键字不能为空或者超过250个字符！</li>"
	End If
	If Len(Request.Form("StopReadme")) = 0 Or Len(Request.Form("StopReadme")) => 250 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>维护说明不能为空或者超过250个字符！</li>"
	End If
	If Len(Request.Form("Copyright")) = 0 Or Len(Request.Form("Copyright")) => 250 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>网站版权信息不能为空或者超过250个字符！</li>"
	End If
	If Len(Request.Form("IndexName")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>首页文件名不能为空！</li>"
	End If
	If Not enchiasp.IsValidChar(Request.Form("IndexName")) Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>首页文件名中含有非法字符或者中文字符！</li>"
	End If
	If Trim(Request.Form("StopApplyLink")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>是否关闭友情连接功能！</li>"
	End If
	If Len(Request.Form("InitTitleColor")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>初始化标题颜色不能为空！</li>"
	End If
	If Len(Request.Form("InstallDir")) = 0 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>系统所在根目录不能为空！</li>"
	End If
	If Len(Request.Form("ChinaeBank4")) > 20 Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>在线支付加密密码不能大于20个字符！</li>"
	End If
	if Trim(Request.Form("ercilogin")) = "1"  then
	If Trim(Request.Form("mypass")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>二次密码值不能为空！</li>"
	End If

	If Trim(Request.Form("mypasskey")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg + "<li>二次密码kai值不能为空！</li>"
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
	Succeed("<li>恭喜您！保存设置成功。</li>")
End Sub

Sub ReloadCache()
	Application.Contents.RemoveAll
	Response.Write "<script>alert('重建缓存成功！');javascript:history.back(1)</script>"
End Sub

Sub EditContentKeyword()
	Dim i,ContentKeywordStr,KeywordStr
	Set Rs = enchiasp.Execute("Select id,ContentKeyword From ECCMS_Config where id = 1")
	ContentKeywordStr = Split(Rs("ContentKeyword"), "@@@")
	
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan=2 align="left"><img src="images/welcome.gif" width="16" height="17" align="absMiddle"> 文章内容关键字设置</th>
</tr>
<form name=myform method=post action=?action=savedit>
<input type=hidden name=id value=''>
<tr>
	<td class=TableTitle align=center>内容关键字</td>
	<td class=TableTitle align=center>关键字连接URL</td>
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
	<td class=tablerow2 colspan=2 align=center><input type=submit value='保存设置' class=Button></td>
</tr>
<tr>
	<td class=tablerow1 colspan=2><strong>说明:</strong><br>&nbsp;&nbsp;本功能用于在文章中找到适当的关键字, 并为其设置合适当的链接.
如: 设置了关键字 恩池软件, 指向地址 http://www.enchi.com.cn  那么当文章中出现"恩池软件"这些字时, 页面上能自动为这些字加上指向http://www.enchi.com.cn这个地址的链接.   注意: 不能设置重复的关键字, 也不能设置包含相同文字的关键字,如果想清除该关键字，请在对应的文本框中输入"del"，
如果要设置多个关键字用一个连接请用“|”分开。</td>
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
			ErrMsg = ErrMsg + "<li>内容关键字和关键字连接URL不能为空！</li>"
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
	OutHintScript("恭喜您！设置内容关键字成功。")
End Sub
%>